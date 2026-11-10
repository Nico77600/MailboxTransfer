// =============================================================================
//  Mailbox Transfer - engine, part 6: delegated sign-in (MSAL.NET)
// -----------------------------------------------------------------------------
//  Why MSAL.NET and not Connect-MgGraph: the engine calls Microsoft Graph with its own HttpClient (the answer of
//  exportItems must be streamed to disk, which Invoke-MgGraphRequest cannot do) and needs a bearer token for each
//  request, for hours. MSAL gives exactly that: AcquireTokenSilent returns the cached token or refreshes it with the
//  refresh token, and the cache survives between runs (DPAPI, current user). The library is the one shipped with the
//  Microsoft.Graph.Authentication module; it is compiled separately because the rest of the engine does not need it.
//  Interactive sign-in uses the system browser and the redirect URI http://localhost (public client, no secret), or, when
//  the browser cannot be used (no window handle, window not visible), the device code flow: the code and the URL are
//  written at once to the standard output ("DEVICE CODE: ...") and to a file, then the call waits until the code really
//  expires (about 15 minutes). The callback is C# (a PowerShell script block would run outside its runspace). The token
//  cache is the same for both modes: one sign-in serves every later run.
//  GraphAppAuth (end of the file) is the non-interactive sign-in of the application itself (client credentials with a certificate,
//  "app-only"): no user, no refresh token, no file cache; the token (claim roles, not scp) is kept in memory until shortly before it expires.
// =============================================================================
using System;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

namespace MailboxTransfer
{
    /// <summary>
    /// The token cache file shared by every process of the user: DPAPI-protected, written to a temporary file and swapped in with File.Replace (a reader
    /// never sees a half-written file), and a named semaphore held from the read of one MSAL access to its write, so that two processes do not overwrite each other.
    /// A read that fails is tried again; a file that stays unreadable is put aside (never deleted) and reported in <see cref="Warning"/>.
    /// </summary>
    public sealed class TokenCacheFile
    {
        static readonly byte[] CacheEntropy = Encoding.UTF8.GetBytes("MailboxTransfer token cache");
        readonly string _path;
        readonly object _lock = new object();
        readonly Semaphore _gate;
        bool _gateHeld;
        /// <summary>How often a read that fails is tried again (100 ms, 200 ms... between the tries).</summary>
        public int ReadAttempts = 5;
        /// <summary>Set when the cache file could not be read and was put aside (the user signs in again).</summary>
        public string Warning { get; private set; }

        public TokenCacheFile(string path)
        {
            _path = Path.GetFullPath(path);
            string name;
            using (var sha = SHA256.Create())
                name = "Local\\MailboxTransfer-token-" + BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(_path.ToLowerInvariant()))).Replace("-", "").Substring(0, 32);
            _gate = new Semaphore(1, 1, name);
        }

        /// <summary>Takes the cross-process gate (waits up to 30 s: a process that died holding it gives it back at once, the others are not stalled for ever).</summary>
        public void Enter()
        {
            bool got = false;
            try { got = _gate.WaitOne(TimeSpan.FromSeconds(30)); } catch (AbandonedMutexException) { got = true; }
            lock (_lock) _gateHeld = got;
        }

        public void Leave()
        {
            bool held;
            lock (_lock) { held = _gateHeld; _gateHeld = false; }
            if (held) { try { _gate.Release(); } catch (SemaphoreFullException) { } }
        }

        /// <summary>The decrypted content, or null when there is no cache (or it was unreadable and put aside).</summary>
        public byte[] Read()
        {
            lock (_lock)
            {
                Exception last = null;
                for (int i = 1; i <= Math.Max(1, ReadAttempts); i++)
                {
                    try
                    {
                        if (!File.Exists(_path)) return null;
                        return ProtectedData.Unprotect(File.ReadAllBytes(_path), CacheEntropy, DataProtectionScope.CurrentUser);
                    }
                    catch (Exception ex) when (ex is IOException || ex is CryptographicException || ex is UnauthorizedAccessException)
                    {
                        last = ex;
                        if (i < ReadAttempts) Task.Delay(100 * i).GetAwaiter().GetResult();
                    }
                }
                PutAside(last);
                return null;
            }
        }

        /// <summary>The file cannot be used (unreadable, or its content is refused): kept under another name, never deleted, and reported.</summary>
        public void PutAside(Exception why)
        {
            lock (_lock)
            {
                string aside = _path + ".unreadable";
                try { File.Move(_path, aside, true); } catch (Exception) { aside = null; }
                Warning = "The token cache " + _path + " could not be read (" + why.Message + ")" + (aside == null ? "" : "; it was kept as " + aside) + ": sign in again.";
            }
        }

        /// <summary>Replaces the cache with this content: temporary file in the same folder, then File.Replace.</summary>
        public void Write(byte[] plain)
        {
            lock (_lock)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_path));
                string tmp = _path + "." + Guid.NewGuid().ToString("N") + ".tmp";
                try
                {
                    File.WriteAllBytes(tmp, ProtectedData.Protect(plain, CacheEntropy, DataProtectionScope.CurrentUser));
                    for (int i = 1; ; i++)
                    {
                        try
                        {
                            if (File.Exists(_path)) File.Replace(tmp, _path, null); else File.Move(tmp, _path);
                            break;
                        }
                        catch (IOException) when (i < 5) { Task.Delay(50 * i).GetAwaiter().GetResult(); }   // a reader that does not know the gate has it open
                    }
                }
                finally { try { if (File.Exists(tmp)) File.Delete(tmp); } catch (Exception) { } }
            }
        }
    }

    public sealed class GraphAuth
    {
        readonly IPublicClientApplication _app;
        readonly string[] _scopes;
        readonly string _cachePath;
        readonly TokenCacheFile _cacheFile;
        readonly object _tokenLock = new object();
        /// <summary>Set when the token cache file was unreadable and put aside (see <see cref="TokenCacheFile.Warning"/>).</summary>
        public string CacheWarning { get { return _cacheFile == null ? null : _cacheFile.Warning; } }

        /// <summary>Account signed in, null until the first token.</summary>
        public string Account { get; private set; }
        public DateTimeOffset? ExpiresOn { get; private set; }
        /// <summary>Sign in with a device code instead of the browser (no time limit other than the life of the code).</summary>
        public bool DeviceCode { get; set; }
        /// <summary>File that receives the "DEVICE CODE: ..." line (overwritten each time); null: none.</summary>
        public string DeviceCodeFile { get; set; }

        public GraphAuth(string clientId, string tenantId, string[] scopes, string cachePath, string redirectUri = "http://localhost",
            bool deviceCode = false, string deviceCodeFile = null)
        {
            _scopes = scopes;
            _cachePath = cachePath;
            DeviceCode = deviceCode;
            DeviceCodeFile = deviceCodeFile;
            _app = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority(AzureCloudInstance.AzurePublic, tenantId)
                .WithRedirectUri(redirectUri)
                .Build();
            if (!string.IsNullOrEmpty(cachePath))
            {
                _cacheFile = new TokenCacheFile(cachePath);
                _app.UserTokenCache.SetBeforeAccess(args =>
                {
                    _cacheFile.Enter();
                    byte[] data = _cacheFile.Read();
                    if (data != null)
                    {
                        try { args.TokenCache.DeserializeMsalV3(data); }
                        catch (Exception ex) { _cacheFile.PutAside(ex); }
                    }
                });
                _app.UserTokenCache.SetAfterAccess(args =>
                {
                    try { if (args.HasStateChanged) _cacheFile.Write(args.TokenCache.SerializeMsalV3()); }
                    finally { _cacheFile.Leave(); }
                });
            }
        }

        /// <summary>
        /// A valid access token: from the cache, else renewed with the refresh token, else (first run, or the refresh token
        /// expired) after an interactive sign-in (browser, or device code). Safe to call from several threads.
        /// The account is the one of loginHint: another account of the cache is never used in its place.
        /// forceRefresh: the cached access token is not used, a new one is asked with the refresh token (the permissions granted since are then in it).
        /// </summary>
        public string GetToken(string loginHint = null, bool interactive = true, CancellationToken ct = default(CancellationToken), bool forceRefresh = false)
        {
            lock (_tokenLock)
            {
                AuthenticationResult result = null;
                var accounts = _app.GetAccountsAsync().GetAwaiter().GetResult().ToList();
                IAccount account = loginHint == null
                    ? accounts.FirstOrDefault()
                    : accounts.FirstOrDefault(a => string.Equals(a.Username, loginHint, StringComparison.OrdinalIgnoreCase));
                if (account != null)
                {
                    try { result = _app.AcquireTokenSilent(_scopes, account).WithForceRefresh(forceRefresh).ExecuteAsync(ct).GetAwaiter().GetResult(); }
                    catch (MsalUiRequiredException) { result = null; }
                }
                if (result == null)
                {
                    if (!interactive) throw new InvalidOperationException("Not signed in: run the script interactively once.");
                    if (DeviceCode)
                    {
                        try { result = _app.AcquireTokenWithDeviceCode(_scopes, ShowDeviceCode).ExecuteAsync(ct).GetAwaiter().GetResult(); }
                        catch (MsalServiceException ex) when (ex.Message.Contains("AADSTS7000218"))
                        {
                            throw new InvalidOperationException("The device code flow needs \"Allow public client flows\" on the application (Authentication blade, or run tools\\Grant-LabConsent.ps1 which turns it on). " + ex.Message, ex);
                        }
                        Account = result.Account == null ? null : result.Account.Username;
                        ExpiresOn = result.ExpiresOn;
                        return result.AccessToken;
                    }
                    var builder = _app.AcquireTokenInteractive(_scopes)
                        .WithUseEmbeddedWebView(false)
                        .WithPrompt(Prompt.SelectAccount);
                    if (loginHint != null) builder = builder.WithLoginHint(loginHint);
                    using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct))
                    {
                        timeout.CancelAfter(TimeSpan.FromMinutes(5));
                        result = builder.ExecuteAsync(timeout.Token).GetAwaiter().GetResult();
                    }
                }
                Account = result.Account == null ? null : result.Account.Username;
                ExpiresOn = result.ExpiresOn;
                return result.AccessToken;
            }
        }

        /// <summary>The line shown to the user: starts with "DEVICE CODE:".</summary>
        public static string DeviceCodeLine(string userCode, string verificationUrl, DateTimeOffset expiresOn)
        {
            return "DEVICE CODE: " + userCode + "  URL: " + verificationUrl + "  (valid until " + expiresOn.ToLocalTime().ToString("HH:mm:ss") + ")";
        }

        /// <summary>Writes the line to the standard output at once (flushed) and, if a file is given, to that file (overwritten).</summary>
        public static void PublishDeviceCode(string line, string file)
        {
            Console.Out.WriteLine(line);
            Console.Out.Flush();
            if (!string.IsNullOrEmpty(file))
            {
                string dir = Path.GetDirectoryName(Path.GetFullPath(file));
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                File.WriteAllText(file, line + Environment.NewLine);
            }
        }

        Task ShowDeviceCode(DeviceCodeResult d)
        {
            PublishDeviceCode(DeviceCodeLine(d.UserCode, d.VerificationUrl, d.ExpiresOn), DeviceCodeFile);
            return Task.CompletedTask;
        }

        /// <summary>The token source expected by GraphClient (a C# delegate: it also works from threads that have no PowerShell runspace).</summary>
        public Func<string> Provider(string loginHint = null)
        {
            return () => GetToken(loginHint, false);
        }

        /// <summary>Forgets the accounts of the cache (and the file).</summary>
        public void SignOut()
        {
            foreach (IAccount a in _app.GetAccountsAsync().GetAwaiter().GetResult().ToList())
                _app.RemoveAsync(a).GetAwaiter().GetResult();
            if (!string.IsNullOrEmpty(_cachePath) && File.Exists(_cachePath)) File.Delete(_cachePath);
        }
    }

    /// <summary>An application access token (client credentials): the token and when it expires.</summary>
    public sealed class AppToken
    {
        public string AccessToken { get; set; }
        public DateTimeOffset ExpiresOn { get; set; }
    }

    /// <summary>
    /// App-only sign-in: the application authenticates with its certificate (client credentials, scope https://graph.microsoft.com/.default) and
    /// acts with the APPLICATION permissions consented to it (MailboxItem.ImportExport.All ...), on any mailbox its Exchange RBAC scope allows.
    /// Nothing is written to disk: MSAL keeps the app token in memory, and this class reuses the last token until 5 minutes before it expires.
    /// </summary>
    public sealed class GraphAppAuth
    {
        static readonly TimeSpan Margin = TimeSpan.FromMinutes(5);

        readonly Func<string[], CancellationToken, AppToken> _acquire;
        readonly string[] _scopes;
        readonly object _tokenLock = new object();
        AppToken _current;

        public string ClientId { get; private set; }
        public string TenantId { get; private set; }
        public string CertificateThumbprint { get; private set; }
        /// <summary>No user is signed in: always null (the delegated GraphAuth has the account).</summary>
        public string Account { get { return null; } }
        public DateTimeOffset? ExpiresOn { get { return _current == null ? (DateTimeOffset?)null : _current.ExpiresOn; } }
        /// <summary>How many tokens were asked of the service (not served from the memory of this object).</summary>
        public int Acquisitions { get; private set; }

        public GraphAppAuth(string clientId, string tenantId, X509Certificate2 certificate, string[] scopes)
        {
            if (certificate == null) throw new ArgumentNullException("certificate");
            if (!certificate.HasPrivateKey) throw new ArgumentException("The certificate has no private key: the application cannot sign its sign-in request.", "certificate");
            ClientId = clientId; TenantId = tenantId; _scopes = scopes; CertificateThumbprint = certificate.Thumbprint;
            IConfidentialClientApplication app = ConfidentialClientApplicationBuilder.Create(clientId)
                .WithAuthority(AzureCloudInstance.AzurePublic, tenantId)
                .WithCertificate(certificate)
                .Build();
            _acquire = (s, ct) =>
            {
                AuthenticationResult r = app.AcquireTokenForClient(s).ExecuteAsync(ct).GetAwaiter().GetResult();
                return new AppToken { AccessToken = r.AccessToken, ExpiresOn = r.ExpiresOn };
            };
        }

        /// <summary>The token source is given (tests: no network).</summary>
        public GraphAppAuth(string clientId, string tenantId, string[] scopes, Func<string[], CancellationToken, AppToken> acquire)
        {
            ClientId = clientId; TenantId = tenantId; _scopes = scopes; _acquire = acquire;
        }

        /// <summary>A valid access token; safe to call from several threads. forceRefresh: ask a new one even if the last is still valid.</summary>
        public string GetToken(CancellationToken ct = default(CancellationToken), bool forceRefresh = false)
        {
            lock (_tokenLock)
            {
                if (forceRefresh || _current == null || _current.ExpiresOn - DateTimeOffset.UtcNow < Margin)
                {
                    try { _current = _acquire(_scopes, ct); Acquisitions++; }
                    catch (MsalServiceException ex) when (ex.Message.Contains("AADSTS700027"))
                    {
                        throw new InvalidOperationException("The certificate was refused: its public key is not uploaded to the application " + ClientId + " (Certificates & secrets), or it is another certificate. " + ex.Message, ex);
                    }
                    catch (MsalServiceException ex) when (ex.Message.Contains("AADSTS700016") || ex.Message.Contains("AADSTS90002"))
                    {
                        throw new InvalidOperationException("The application " + ClientId + " is not found in the tenant " + TenantId + " (check -ClientId and -TenantId). " + ex.Message, ex);
                    }
                }
                return _current.AccessToken;
            }
        }

        /// <summary>The token source expected by GraphClient.</summary>
        public Func<string> Provider() { return () => GetToken(); }

        /// <summary>Forgets the token held in memory.</summary>
        public void SignOut() { lock (_tokenLock) { _current = null; } }
    }
}
