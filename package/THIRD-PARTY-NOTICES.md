# Third-party components

The folder `lib\sqlite` contains the SQLite libraries used by the local store (`src\Engine.Store.cs`). They are loaded by
`Initialize-MbtEngine`; the versions below are those of the files.

| Component | Version | Files | License |
|---|---|---|---|
| Microsoft.Data.Sqlite.Core (Microsoft) | 10.0.12 | `Microsoft.Data.Sqlite.dll` | MIT — https://github.com/dotnet/efcore/blob/main/LICENSE.txt |
| SQLitePCLRaw.core (Eric Sink) | 2.1.12 | `SQLitePCLRaw.core.dll` | Apache-2.0 — https://github.com/ericsink/SQLitePCL.raw/blob/master/LICENSE.TXT |
| SQLitePCLRaw.provider.e_sqlite3 | 2.1.12 | `SQLitePCLRaw.provider.e_sqlite3.dll` | Apache-2.0 |
| SQLitePCLRaw.bundle_e_sqlite3 | 2.1.12 | `SQLitePCLRaw.batteries_v2.dll` | Apache-2.0 |
| SQLitePCLRaw.lib.e_sqlite3 | 2.1.12 | `runtimes\win-x64\e_sqlite3.dll`, `runtimes\win-arm64\e_sqlite3.dll` | Apache-2.0 (packaging); SQLite itself is in the public domain — https://sqlite.org/copyright.html |

Not shipped, loaded at run time from the installed module Microsoft.Graph.Authentication (Microsoft, MIT): MSAL.NET (`Microsoft.Identity.Client.dll`) and `Microsoft.IdentityModel.Abstractions.dll`, to sign you in. They are the only assemblies that the tool loads besides the ones of this table, the .NET libraries of PowerShell 7 (WPF for the console) and the Windows DPAPI (`System.Security.Cryptography.ProtectedData`) for the token cache.

The tool contains no telemetry and contacts no service other than Microsoft Graph (and the sign-in page of Microsoft Entra ID).
