// LabFxAuth.cs - MSAL.NET delegated sign-in for the lab tool New-LabFxCorpus.ps1 (adapted from the lab fill tool, FillAuth.cs).
// Compiled in memory by New-LabFxCorpus.ps1 with Microsoft.Identity.Client.dll of the module Microsoft.Graph.Authentication.
//
//  - Device code: the line "DEVICE CODE: ..." is written at once to the standard output (flushed) and to a text file.
//    The wait lasts until the real expiry of the code (about 15 minutes).
//  - Token cache encrypted with DPAPI (CurrentUser): a later run is silent.
//  - GetToken() always returns a token valid for at least a few minutes (AcquireTokenSilent renews it with the refresh token).
//  - Only the expected account is ever used.
using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

namespace LabFx
{
    public sealed class Auth
    {
        static readonly byte[] Entropy = Encoding.UTF8.GetBytes("New-LabFxCorpus token cache");

        readonly IPublicClientApplication _app;
        readonly string[] _scopes;
        readonly string _cachePath;
        readonly string _upn;
        readonly object _cacheLock = new object();
        readonly object _tokenLock = new object();
        string _token;
        DateTimeOffset _expires = DateTimeOffset.MinValue;

        /// <summary>Signed-in account (UPN), null before SignIn().</summary>
        public string Account { get; private set; }
        public DateTimeOffset ExpiresOn { get { return _expires; } }
        /// <summary>Sign in with a device code instead of the system browser.</summary>
        public bool DeviceCode { get; set; }
        /// <summary>File that receives the "DEVICE CODE: ..." line (overwritten); null: none.</summary>
        public string DeviceCodeFile { get; set; }

        public Auth(string clientId, string tenantId, string[] scopes, string cachePath, string userPrincipalName)
        {
            _scopes = scopes; _cachePath = cachePath; _upn = userPrincipalName;
            _app = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority(AzureCloudInstance.AzurePublic, tenantId)
                .WithRedirectUri("http://localhost")
                .Build();
            if (string.IsNullOrEmpty(cachePath)) return;
            _app.UserTokenCache.SetBeforeAccess(args =>
            {
                lock (_cacheLock)
                {
                    try
                    {
                        if (File.Exists(_cachePath))
                            args.TokenCache.DeserializeMsalV3(ProtectedData.Unprotect(File.ReadAllBytes(_cachePath), Entropy, DataProtectionScope.CurrentUser));
                    }
                    catch (Exception) { /* unreadable cache: start empty, sign in again */ }
                }
            });
            _app.UserTokenCache.SetAfterAccess(args =>
            {
                if (!args.HasStateChanged) return;
                lock (_cacheLock)
                {
                    string dir = Path.GetDirectoryName(Path.GetFullPath(_cachePath));
                    if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                    string tmp = _cachePath + ".tmp";
                    File.WriteAllBytes(tmp, ProtectedData.Protect(args.TokenCache.SerializeMsalV3(), Entropy, DataProtectionScope.CurrentUser));
                    if (File.Exists(_cachePath)) File.Delete(_cachePath);
                    File.Move(tmp, _cachePath);
                }
            });
        }

        /// <summary>Silent when the cache holds the expected account, else device code / browser. Checks the account.</summary>
        public string SignIn()
        {
            lock (_tokenLock)
            {
                AuthenticationResult r = Silent(false);
                if (r == null)
                {
                    try
                    {
                        if (DeviceCode) r = _app.AcquireTokenWithDeviceCode(_scopes, ShowDeviceCode).ExecuteAsync().GetAwaiter().GetResult();
                        else
                        {
                            var b = _app.AcquireTokenInteractive(_scopes).WithUseEmbeddedWebView(false).WithPrompt(Prompt.SelectAccount).WithLoginHint(_upn);
                            using (var cts = new CancellationTokenSource(TimeSpan.FromMinutes(10)))
                                r = b.ExecuteAsync(cts.Token).GetAwaiter().GetResult();
                        }
                    }
                    catch (MsalServiceException ex) when (ex.Message.Contains("AADSTS7000218"))
                    {
                        throw new InvalidOperationException("The device code flow needs \"Allow public client flows\" on the application. " + ex.Message, ex);
                    }
                }
                string who = r.Account == null ? null : r.Account.Username;
                if (!string.Equals(who, _upn, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Signed in as '" + who + "', expected '" + _upn + "'. Run again and sign in with the right account.");
                Keep(r);
                return r.AccessToken;
            }
        }

        /// <summary>A token valid for at least 5 minutes, renewed with the refresh token. Never prompts. Thread-safe.</summary>
        public string GetToken()
        {
            lock (_tokenLock)
            {
                if (_token != null && _expires - DateTimeOffset.UtcNow > TimeSpan.FromMinutes(6)) return _token;
                Exception last = null;
                for (int attempt = 1; attempt <= 3; attempt++)
                {
                    try
                    {
                        AuthenticationResult r = Silent(false);
                        if (r != null && r.ExpiresOn - DateTimeOffset.UtcNow < TimeSpan.FromMinutes(8))
                        {
                            AuthenticationResult f = Silent(true);
                            if (f != null) r = f;
                        }
                        if (r == null) throw new InvalidOperationException("Session expired or account missing from the cache: run the script again to sign in.");
                        Keep(r);
                        return r.AccessToken;
                    }
                    catch (InvalidOperationException) { throw; }
                    catch (Exception ex)
                    {
                        last = ex;
                        if (_token != null && _expires - DateTimeOffset.UtcNow > TimeSpan.FromMinutes(1)) return _token;
                        Thread.Sleep(TimeSpan.FromSeconds(5 * attempt));
                    }
                }
                throw new InvalidOperationException("Cannot renew the token: " + last.Message, last);
            }
        }

        void Keep(AuthenticationResult r) { Account = r.Account.Username; _token = r.AccessToken; _expires = r.ExpiresOn; }

        AuthenticationResult Silent(bool force)
        {
            IAccount acc = _app.GetAccountsAsync().GetAwaiter().GetResult()
                .FirstOrDefault(a => string.Equals(a.Username, _upn, StringComparison.OrdinalIgnoreCase));
            if (acc == null) return null;
            try { return _app.AcquireTokenSilent(_scopes, acc).WithForceRefresh(force).ExecuteAsync().GetAwaiter().GetResult(); }
            catch (MsalUiRequiredException) { return null; }
        }

        Task ShowDeviceCode(DeviceCodeResult d)
        {
            string line = "DEVICE CODE: " + d.UserCode + "  URL: " + d.VerificationUrl +
                          "  (valid until " + d.ExpiresOn.ToLocalTime().ToString("HH:mm:ss") + ")";
            Console.Out.WriteLine(line);
            Console.Out.Flush();
            if (!string.IsNullOrEmpty(DeviceCodeFile))
            {
                string dir = Path.GetDirectoryName(Path.GetFullPath(DeviceCodeFile));
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                File.WriteAllText(DeviceCodeFile, line + Environment.NewLine);
            }
            return Task.CompletedTask;
        }
    }
}
