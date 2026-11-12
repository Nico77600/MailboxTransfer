// =============================================================================
//  Mailbox Transfer - engine, part 6: delegated sign-in (MSAL.NET)
// -----------------------------------------------------------------------------
//  Why MSAL.NET and not Connect-MgGraph: the engine calls Microsoft Graph with its own HttpClient (the answer of
//  exportItems must be streamed to disk, which Invoke-MgGraphRequest cannot do) and needs a bearer token for each
//  request, for hours. MSAL gives exactly that: AcquireTokenSilent returns the cached token or refreshes it with the
//  refresh token, and the cache survives between runs (DPAPI, current user). The library is the one shipped with the
//  Microsoft.Graph.Authentication module; it is compiled separately because the rest of the engine does not need it.
//  Interactive sign-in is a Windows sign-in window (the Windows account broker, WAM) opened over the console: the operator types the
//  account, nothing is loaded into a browser profile. The broker comes from the same Microsoft.Graph.Authentication module as MSAL
//  (WindowsBroker). When it cannot be loaded, or when the operator asks for it, the device code flow is used: the code and the URL
//  are written at once to the standard output ("DEVICE CODE: ...") and to a file, then the call waits until the code really
//  expires (about 15 minutes). There is no system-browser sign-in. The callback is C# (a PowerShell script block would run outside
//  its runspace). A token obtained with a code is kept in the DPAPI cache file; the broker keeps its own.
//  GraphAppAuth (end of the file) is the non-interactive sign-in of the application itself (client credentials with a certificate,
//  "app-only"): no user, no refresh token, no file cache; the token (claim roles, not scp) is kept in memory until shortly before it expires.
// =============================================================================
using System;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
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

    /// <summary>
    /// The Windows account broker (WAM) of MSAL.NET. The assembly Microsoft.Identity.Client.Broker is not shipped here: it is loaded from the folder of the
    /// Microsoft.Graph.Authentication module that holds the MSAL in use (same version), with its native runtime (msalruntime). It is reached by reflection so that
    /// this file compiles, and the engine works with a device code, when the broker is not there.
    /// </summary>
    public static class WindowsBroker
    {
        public const string BrokerAssemblyName = "Microsoft.Identity.Client.Broker";
        static readonly object Gate = new object();
        static MethodInfo _withBroker;

        /// <summary>The redirect URI the Windows sign-in window needs on the app registration (public client).</summary>
        public static string RedirectUri(string clientId) { return "ms-appx-web://microsoft.aad.brokerplugin/" + clientId; }

        /// <summary>Loads the broker from the folder (an assembly already loaded in the process is used as it is). Returns null when it is ready, else the reason in plain words.</summary>
        public static string Prepare(string directory)
        {
            lock (Gate)
            {
                if (_withBroker != null) return null;
                try
                {
                    Assembly broker = AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(a => a.GetName().Name == BrokerAssemblyName);
                    if (broker == null)
                    {
                        if (string.IsNullOrEmpty(directory)) return "the Windows sign-in component was not found next to MSAL";
                        string file = Path.Combine(directory, BrokerAssemblyName + ".dll");
                        if (!File.Exists(file)) return "the installed Microsoft.Graph.Authentication module has no Windows sign-in component (" + BrokerAssemblyName + ".dll)";
                        string native = Path.Combine(directory, NativeRuntimeFile());
                        if (File.Exists(native)) { try { NativeLibrary.Load(native); } catch (Exception) { } }
                        string interop = Path.Combine(directory, "Microsoft.Identity.Client.NativeInterop.dll");
                        if (File.Exists(interop) && !AppDomain.CurrentDomain.GetAssemblies().Any(a => a.GetName().Name == "Microsoft.Identity.Client.NativeInterop")) Assembly.LoadFrom(interop);
                        broker = Assembly.LoadFrom(file);
                    }
                    Type ext = broker.GetType("Microsoft.Identity.Client.Broker.BrokerExtension", false);
                    MethodInfo m = ext == null ? null : ext.GetMethod("WithBroker", new[] { typeof(PublicClientApplicationBuilder), typeof(BrokerOptions) });
                    if (m == null) return "the Windows sign-in component of this Microsoft.Graph.Authentication module is not the expected one";
                    _withBroker = m;
                    return null;
                }
                catch (Exception ex) { return "the Windows sign-in component could not be loaded (" + ex.GetType().Name + ": " + ex.Message + ")"; }
            }
        }

        /// <summary>The native runtime file of this process architecture, next to the managed broker.</summary>
        public static string NativeRuntimeFile()
        {
            switch (RuntimeInformation.ProcessArchitecture)
            {
                case Architecture.X86: return "msalruntime_x86.dll";
                case Architecture.Arm64: return "msalruntime_arm64.dll";
                default: return "msalruntime.dll";
            }
        }

        /// <summary>The builder with the Windows broker switched on (Prepare must have returned null).</summary>
        public static PublicClientApplicationBuilder Enable(PublicClientApplicationBuilder builder, string title)
        {
            if (_withBroker == null) throw new InvalidOperationException("The Windows broker is not loaded.");
            var options = new BrokerOptions(BrokerOptions.OperatingSystems.Windows) { Title = title };
            try { return (PublicClientApplicationBuilder)_withBroker.Invoke(null, new object[] { builder, options }); }
            catch (TargetInvocationException ex) { throw ex.InnerException ?? ex; }
        }

        [DllImport("kernel32.dll")] static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] static extern IntPtr GetDesktopWindow();
        [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hWnd);

        /// <summary>The window the sign-in opens over, from a command line: the console window when it is visible (a terminal that hosts the console has a hidden one), else the foreground window.</summary>
        public static IntPtr ConsoleOrForegroundWindow()
        {
            try
            {
                IntPtr w = GetConsoleWindow();
                if (w != IntPtr.Zero && IsWindowVisible(w)) return w;
                w = GetForegroundWindow();
                if (w != IntPtr.Zero) return w;
                return GetDesktopWindow();
            }
            catch (Exception) { return IntPtr.Zero; }
        }

        /// <summary>The actionable message when Entra says the redirect URI is not registered on the application; null for any other error.</summary>
        public static string MissingRedirectUriMessage(string errorText, string clientId)
        {
            if (string.IsNullOrEmpty(errorText)) return null;
            bool missing = errorText.Contains("AADSTS50011") || errorText.Contains("AADSTS500113")
                || (errorText.IndexOf("redirect uri", StringComparison.OrdinalIgnoreCase) >= 0 && (errorText.IndexOf("does not match", StringComparison.OrdinalIgnoreCase) >= 0 || errorText.IndexOf("not registered", StringComparison.OrdinalIgnoreCase) >= 0));
            if (!missing) return null;
            return "The application " + clientId + " does not accept the Windows sign-in window yet. Ask the administrator to add this redirect URI to its registration "
                + "(Authentication, platform \"Mobile and desktop applications\"): " + RedirectUri(clientId) + ". Until then, sign in with a code (-DeviceCode, or \"Use a code instead\" in the console).";
        }
    }

    public sealed class GraphAuth
    {
        readonly IPublicClientApplication _app;
        readonly IPublicClientApplication _brokerApp;
        readonly string[] _scopes;
        readonly string _cachePath;
        readonly string _clientId;
        readonly TokenCacheFile _cacheFile;
        readonly object _tokenLock = new object();
        /// <summary>Set when the token cache file was unreadable and put aside (see <see cref="TokenCacheFile.Warning"/>).</summary>
        public string CacheWarning { get { return _cacheFile == null ? null : _cacheFile.Warning; } }

        /// <summary>Account signed in, null until the first token.</summary>
        public string Account { get; private set; }
        public DateTimeOffset? ExpiresOn { get; private set; }
        /// <summary>Sign in with a device code instead of the Windows sign-in window (no time limit other than the life of the code).</summary>
        public bool DeviceCode { get; set; }
        /// <summary>File that receives the "DEVICE CODE: ..." line (overwritten each time); null: none.</summary>
        public string DeviceCodeFile { get; set; }
        /// <summary>The window the Windows sign-in opens over (the console of the tool, set by the caller); zero: the console window, else the foreground window.</summary>
        public IntPtr ParentWindow { get; set; }
        /// <summary>True when the Windows sign-in window can be used.</summary>
        public bool BrokerEnabled { get { return _brokerApp != null; } }
        /// <summary>Why the Windows sign-in window cannot be used (a code is used instead); null when it can, or when no window was asked for.</summary>
        public string BrokerProblem { get; private set; }
        /// <summary>The application of the Windows sign-in window (null when there is none).</summary>
        public IPublicClientApplication WindowApp { get { return _brokerApp; } }
        /// <summary>The application that holds the DPAPI cache and runs the device code flow.</summary>
        public IPublicClientApplication CacheApp { get { return _app; } }
        /// <summary>How an interactive sign-in is done now: "window" (Windows sign-in window) or "code" (device code).</summary>
        public string SignInMethod { get { return DeviceCode || _brokerApp == null ? "code" : "window"; } }

        /// <summary>
        /// brokerDirectory: the folder that holds the Windows broker next to MSAL (see WindowsBroker); null: no window is asked for, the device code is used.
        /// brokerProblem: why the caller already knows the window cannot be used (kept in BrokerProblem).
        /// </summary>
        public GraphAuth(string clientId, string tenantId, string[] scopes, string cachePath, string redirectUri = "http://localhost",
            bool deviceCode = false, string deviceCodeFile = null, string brokerDirectory = null, string brokerProblem = null)
        {
            _scopes = scopes;
            _cachePath = cachePath;
            _clientId = clientId;
            DeviceCode = deviceCode;
            DeviceCodeFile = deviceCodeFile;
            BrokerProblem = brokerProblem;
            _app = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority(AzureCloudInstance.AzurePublic, tenantId)
                .WithRedirectUri(redirectUri)
                .Build();
            if (!string.IsNullOrEmpty(brokerDirectory))
            {
                string problem = WindowsBroker.Prepare(brokerDirectory);
                if (problem == null)
                {
                    try
                    {
                        _brokerApp = WindowsBroker.Enable(PublicClientApplicationBuilder.Create(clientId)
                            .WithAuthority(AzureCloudInstance.AzurePublic, tenantId)
                            .WithRedirectUri(WindowsBroker.RedirectUri(clientId)), "Mailbox Transfer").Build();
                        var concrete = _brokerApp as PublicClientApplication;
                        if (concrete != null && !concrete.IsBrokerAvailable()) { _brokerApp = null; problem = "Windows does not offer its sign-in broker to this session"; }
                    }
                    catch (Exception ex) { _brokerApp = null; problem = "the Windows sign-in component could not be started (" + ex.GetType().Name + ": " + ex.Message + ")"; }
                }
                if (_brokerApp == null) BrokerProblem = problem;
            }
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

        static IAccount FindAccount(IPublicClientApplication app, string loginHint)
        {
            var accounts = app.GetAccountsAsync().GetAwaiter().GetResult().ToList();
            return loginHint == null
                ? accounts.FirstOrDefault()
                : accounts.FirstOrDefault(a => string.Equals(a.Username, loginHint, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// A valid access token: from the DPAPI cache, else from the Windows broker, renewed with the refresh token, else (first run, or the refresh token
        /// expired) after an interactive sign-in: a Windows sign-in window over the console, or a device code. Safe to call from several threads.
        /// The account is the one of loginHint: another account of the cache is never used in its place.
        /// forceRefresh: the cached access token is not used, a new one is asked with the refresh token (the permissions granted since are then in it).
        /// </summary>
        public string GetToken(string loginHint = null, bool interactive = true, CancellationToken ct = default(CancellationToken), bool forceRefresh = false)
        {
            lock (_tokenLock)
            {
                AuthenticationResult result = null;
                IAccount account = FindAccount(_app, loginHint);
                if (account != null)
                {
                    try { result = _app.AcquireTokenSilent(_scopes, account).WithForceRefresh(forceRefresh).ExecuteAsync(ct).GetAwaiter().GetResult(); }
                    catch (MsalUiRequiredException) { result = null; }
                }
                if (result == null && _brokerApp != null)
                {
                    try
                    {
                        IAccount known = FindAccount(_brokerApp, loginHint);
                        if (known != null) result = _brokerApp.AcquireTokenSilent(_scopes, known).WithForceRefresh(forceRefresh).ExecuteAsync(ct).GetAwaiter().GetResult();
                        else if (loginHint != null) result = _brokerApp.AcquireTokenSilent(_scopes, loginHint).WithForceRefresh(forceRefresh).ExecuteAsync(ct).GetAwaiter().GetResult();
                    }
                    catch (MsalException) { result = null; }
                }
                if (result == null)
                {
                    if (!interactive) throw new InvalidOperationException("Not signed in: run the script interactively once.");
                    if (DeviceCode || _brokerApp == null)
                    {
                        result = AcquireWithDeviceCode(ct);
                    }
                    else
                    {
                        try { result = AcquireInWindow(loginHint, ct); }
                        catch (Exception ex) when (ex is DllNotFoundException || ex is BadImageFormatException || ex is PlatformNotSupportedException || ex is TypeInitializationException || ex is EntryPointNotFoundException)
                        {
                            BrokerProblem = "the Windows sign-in window could not be started (" + ex.GetType().Name + ": " + ex.Message + ")";
                            result = AcquireWithDeviceCode(ct);
                        }
                    }
                }
                Account = result.Account == null ? null : result.Account.Username;
                ExpiresOn = result.ExpiresOn;
                return result.AccessToken;
            }
        }

        AuthenticationResult AcquireWithDeviceCode(CancellationToken ct)
        {
            try { return _app.AcquireTokenWithDeviceCode(_scopes, ShowDeviceCode).ExecuteAsync(ct).GetAwaiter().GetResult(); }
            catch (MsalServiceException ex) when (ex.Message.Contains("AADSTS7000218"))
            {
                throw new InvalidOperationException("The device code flow needs \"Allow public client flows\" on the application (Authentication blade of the app registration). " + ex.Message, ex);
            }
        }

        AuthenticationResult AcquireInWindow(string loginHint, CancellationToken ct)
        {
            IntPtr parent = ParentWindow != IntPtr.Zero ? ParentWindow : WindowsBroker.ConsoleOrForegroundWindow();
            var builder = _brokerApp.AcquireTokenInteractive(_scopes)
                .WithParentActivityOrWindow(parent)
                .WithPrompt(Prompt.SelectAccount);
            if (loginHint != null) builder = builder.WithLoginHint(loginHint);
            using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct))
            {
                timeout.CancelAfter(TimeSpan.FromMinutes(5));
                try { return builder.ExecuteAsync(timeout.Token).GetAwaiter().GetResult(); }
                catch (MsalException ex)
                {
                    string message = WindowsBroker.MissingRedirectUriMessage(ex.Message, _clientId);
                    if (message == null) throw;
                    throw new InvalidOperationException(message, ex);
                }
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
            if (!DeviceCode && !string.IsNullOrEmpty(BrokerProblem))
            {
                Console.Out.WriteLine("The Windows sign-in window cannot be used (" + BrokerProblem + "). A code is used instead.");
                Console.Out.Flush();
            }
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
            if (_brokerApp != null)
            {
                try
                {
                    foreach (IAccount a in _brokerApp.GetAccountsAsync().GetAwaiter().GetResult().ToList())
                        _brokerApp.RemoveAsync(a).GetAwaiter().GetResult();
                }
                catch (MsalException) { }
            }
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
