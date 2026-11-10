#
#  Mailbox Transfer - configuration
#  --------------------------------------------------------------------------
#  Every key is optional: the values below are the built-in defaults.
#  The Lab section only holds the defaults of the lab scripts (tools\*.ps1) and of the command line: every value is also a parameter of them.
#  Do not write your own tenant, application or mailbox ids in this file. Put them in
#  config\MailboxTransfer.config.local.psd1 (same layout, only the keys you change): git ignores it, and it is read after this file.
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
        # .default: the delegated permissions granted to the application (see tools\Grant-LabConsent.ps1).
        Scopes      = @('https://graph.microsoft.com/.default')
        RedirectUri = 'http://localhost'
        # Token cache (DPAPI, current user). Empty: %LOCALAPPDATA%\MailboxTransfer\token.cache
        CacheFile   = ''
        # App-only sign-in with a certificate (-CertificateThumbprint / -CertificatePath of the scripts; see docs\MailboxTransfer-Guide.md, chapter 4.6). Empty: the parameter is needed.
        AppTenantId           = ''
        AppClientId           = ''
        CertificateThumbprint = ''
        CertificatePath       = ''
    }

    Output = @{
        Root = '.\out'
    }

    # Defaults of the command line and of the lab scripts: your tenant, your application and the mailbox to work on.
    # Placeholders only: keep the real values in config\MailboxTransfer.config.local.psd1.
    Lab = @{
        TenantId          = ''     # for instance '00000000-0000-0000-0000-000000000000' or 'contoso.onmicrosoft.com'
        ClientId          = ''     # application (client) id of the app registration
        UserPrincipalName = ''     # for instance 'adele.vance@contoso.com'
        UserObjectId      = ''     # only for tools\Grant-LabConsent.ps1 when the directory cannot be read
        # Physical mailbox ids MBX:<mailbox guid>@<tenant id>, from Exchange Online PowerShell:
        # Get-MailboxLocation -User <upn> (Primary, MainArchive; an auto-expanding archive adds AuxArchive locations).
        # Empty: the primary one is read with GET /users/{upn}/settings/exchange (needs User.Read.All), the archive one
        # has no API and must then be given.
        PrimaryMailboxId  = ''
        ArchiveMailboxId  = ''
        # Folders of the lab scripts (tools\Invoke-Phase0.ps1).
        SourceFolder      = ''
        TargetFolder      = ''
    }
}
