#
#  Mailbox Transfer - configuration
#  --------------------------------------------------------------------------
#  Every key is optional: the values below are the built-in defaults.
#  The Tenant section holds your tenant, your application and the mailbox to work on: the console shows them on its Setup page and the command line
#  uses them when the matching parameter is not given (-TenantId, -ClientId, -UserPrincipalName ...). A file that still calls this section Lab is read the same way.
#  The Auth section holds the application that signs in with a certificate (no user). The Lab section only holds what the developer scripts (tools\*.ps1) need.
#  Do not write your own tenant, application or mailbox ids in this file. Put them in
#  config\MailboxTransfer.config.local.psd1 (same layout, only the keys you change): git ignores it, and it is read after this file.
#  What you change in the console for one run is never written to either file.
#
@{
    Graph = @{
        BaseUrl               = 'https://graph.microsoft.com/v1.0'
        # Retry of 429 / 503 / 504 (Retry-After honoured, else 2 s, 4 s, 8 s ... up to MaxBackoffSeconds).
        MaxAttempts           = 6
        BaseBackoffSeconds    = 2
        MaxBackoffSeconds     = 60
        # Time allowed for the headers of one answer (exportItems answers after it has read the items).
        RequestTimeoutMinutes = 15
        # Items per exportItems request: 20 is the maximum of the API.
        ExportBatchSize       = 20
        # $top of the item listings. 0: the service chooses its page size.
        ItemPageSize          = 0
    }

    Auth = @{
        # .default: the delegated permissions granted to the application (admin consent).
        Scopes      = @('https://graph.microsoft.com/.default')
        RedirectUri = 'http://localhost'
        # Token cache (DPAPI, current user). Empty: %LOCALAPPDATA%\MailboxTransfer\token.cache
        CacheFile   = ''
        # App-only sign-in with a certificate (-CertificateThumbprint / -CertificatePath of the scripts; see docs\MailboxTransfer-Guide.md, chapter 4.6). Empty: the parameter is needed.
        # The console shows these three values on its Setup page ("The application (certificate)").
        AppTenantId           = ''
        AppClientId           = ''
        CertificateThumbprint = ''
        CertificatePath       = ''
    }

    Output = @{
        Root = '.\out'
    }

    Jobs = @{
        # Days a finished job keeps its folder in <work folder>\jobs (its report, its events, its status). A job that can still be continued (stopped, not complete,
        # waiting for a decision) is never removed. The stores, the selections and the PST or EML files a job wrote are never removed. 0 keeps everything.
        RetentionDays = 90
    }

    # Your tenant, your application (signing in as a user) and the mailbox to work on.
    # Placeholders only: keep the real values in config\MailboxTransfer.config.local.psd1.
    Tenant = @{
        TenantId          = ''     # for instance '00000000-0000-0000-0000-000000000000' or 'contoso.onmicrosoft.com'
        ClientId          = ''     # application (client) id of the app registration
        UserPrincipalName = ''     # for instance 'adele.vance@contoso.com'
        # Physical mailbox ids MBX:<mailbox guid>@<tenant id>, from Exchange Online PowerShell:
        # Get-MailboxLocation -User <upn> (Primary, MainArchive; an auto-expanding archive adds AuxArchive locations).
        # Empty: the primary one is read with GET /users/{upn}/settings/exchange (needs User.Read.All), the archive one
        # has no API and must then be given.
        PrimaryMailboxId  = ''
        ArchiveMailboxId  = ''
    }

    # Only for the developer scripts (tools\*.ps1). The keys of Tenant are accepted here too.
    Lab = @{
        UserObjectId      = ''     # only for a developer script of tools\ when the directory cannot be read
        # Folders of the developer scripts (tools\Invoke-Phase0.ps1).
        SourceFolder      = ''
        TargetFolder      = ''
    }
}
