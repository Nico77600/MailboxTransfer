#Requires -Version 7.6
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of -Mode Copy (a mailbox copied into another mailbox: the export into a store, then the import of that store, in one job): what is refused before anything is
    read, the plan (-WhatIf), the chained run against the fake service (FakeMailboxService holds both mailboxes), the resume after a stop in each half, the guarantee that a second run creates
    nothing twice, and the names the copy derives (store, import job). Nothing here reaches the network.

    Run:  Invoke-Pester -Path .\tests\Copy.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    $script:Entry = Join-Path $script:Root 'Invoke-MailboxTransfer.ps1'
    $script:SavedLocal = $env:MAILBOXTRANSFER_LOCAL_CONFIG
    $env:MAILBOXTRANSFER_LOCAL_CONFIG = Join-Path $TestDrive 'no-local-config.psd1'
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport
    . (Join-Path $PSScriptRoot 'support\Import-TestFixtures.ps1')
    Import-TestFixtures -TestsRoot $PSScriptRoot
    $script:Tmp = Join-Path ([IO.Path]::GetTempPath()) ('mbt-copy-tests-' + [guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $script:Tmp
    $script:Tenant = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
    $script:Src = "MBX:11111111-2222-3333-4444-555555555555@$($script:Tenant)"
    $script:Dst = "MBX:99999999-8888-7777-6666-555555555555@$($script:Tenant)"
    $script:DstArchive = "MBX:99999999-8888-7777-6666-5555555555aa@$($script:Tenant)"
    $script:Me = 'adele.vance@contoso.com'

    function New-Message([string]$Subject) {
        $b = [MailboxTransfer.Tests.FxBuilder]::new()
        $null = $b.Str(0x001A, 'IPM.Note').Str(0x0037, $Subject)
        return $b.ToArray()
    }
    function New-FakeClient {
        param([Parameter(Mandatory)]$Service)
        $o = [MailboxTransfer.GraphOptions]::new()
        $o.DelayScale = 0
        $o.UseJitter = $false
        return [MailboxTransfer.GraphClient]::new([MailboxTransfer.Tests.FixedToken]::Provider, $o, $Service)
    }
    # The source mailbox (Inbox 6 items, Inbox/Projects 4 items, Empty) and the two locations of the target (an Inbox each) in one service.
    function New-CopyService {
        param([int]$InboxItems = 6, [int]$ProjectItems = 4)
        $svc = [MailboxTransfer.Tests.FakeMailboxService]::new()
        $null = $svc.AddMailbox($script:Src)
        $null = $svc.AddFolder($script:Src, 'F-INBOX', 'Inbox')
        $null = $svc.AddFolder($script:Src, 'F-PROJ', 'Projects', 'F-INBOX')
        $null = $svc.AddFolder($script:Src, 'F-EMPTY', 'Empty')
        foreach ($i in 1..$InboxItems | Where-Object { $InboxItems -gt 0 }) { $null = $svc.AddItem($script:Src, 'F-INBOX', "I$i", (New-Message "Inbox $i"), "ck-I$i") }
        foreach ($i in 1..$ProjectItems | Where-Object { $ProjectItems -gt 0 }) { $null = $svc.AddItem($script:Src, 'F-PROJ', "P$i", (New-Message "Project $i"), "ck-P$i") }
        foreach ($m in $script:Dst, $script:DstArchive) { $null = $svc.AddMailbox($m); $null = $svc.AddFolder($m, 'D-INBOX', 'Inbox') }
        return $svc
    }
    function Get-FolderPaths($svc, $mailbox) {
        $folders = $svc.FoldersOf($mailbox)
        $byId = @{}; foreach ($f in $folders) { $byId[$f.Id] = $f }
        @($folders | ForEach-Object {
                $p = $_.Name; $c = $_
                while ($byId.ContainsKey($c.ParentId)) { $c = $byId[$c.ParentId]; $p = "$($c.Name)/$p" }
                $p
            }) | Sort-Object
    }
    # The entry script in this process, the sign-in replaced by the fake service (see the Mock in the Describe blocks).
    function Invoke-Entry {
        param([Parameter(Mandatory)]$Service, [hashtable]$Parameters = @{})
        # a new client at each sign-in, as the real Connect-MbtGraph makes one (the entry script may disconnect and connect again)
        $handler = $Service
        $global:MbtEntryFactory = {
            $o = [MailboxTransfer.GraphOptions]::new(); $o.DelayScale = 0; $o.UseJitter = $false
            [MailboxTransfer.GraphClient]::new([MailboxTransfer.Tests.FixedToken]::Provider, $o, $handler)
        }.GetNewClosure()
        $global:LASTEXITCODE = $null
        $text = & $script:Entry @Parameters *>&1 | Out-String
        $global:MbtEntryExit = $LASTEXITCODE
        return $text
    }
    # A delegated copy of another mailbox (the source, by its id) into the signed-in user's own Primary mailbox (the target).
    function Get-CopyParameters {
        param([string]$Name, [hashtable]$More = @{})
        $p = @{
            Mode = 'Copy'; Mailbox = 'Primary'; MailboxId = $script:Src; PrimaryMailboxId = $script:Dst; ArchiveMailboxId = $script:DstArchive; UserPrincipalName = $script:Me
            StorePath = (Join-Path $script:Tmp "store-$Name"); TempPath = (Join-Path $script:Tmp "tmp-$Name"); TargetRoot = 'Copied'; ReconcileDelaySeconds = 0; IntervalSeconds = 1
        }
        foreach ($k in $More.Keys) { $p[$k] = $More[$k] }
        return $p
    }
    function New-EventsFile([string]$Name) {
        $dir = Join-Path $script:Tmp "job-$Name"
        $null = New-Item -ItemType Directory -Force -Path $dir
        return (Join-Path $dir 'events.jsonl')
    }
}
AfterAll {
    Remove-Variable -Name MbtEntryFactory, MbtEntryExit -Scope Global -ErrorAction SilentlyContinue
    $env:MAILBOXTRANSFER_LOCAL_CONFIG = $script:SavedLocal
    if ($script:Tmp -and (Test-Path -LiteralPath $script:Tmp)) { Remove-Item -LiteralPath $script:Tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

Describe 'Copy: what is refused before anything is read' {
    BeforeEach {
        Mock Connect-MbtGraph {
            $session = Set-MbtSession -Graph (& $global:MbtEntryFactory)
            $session.Token = [pscustomobject]@{
                Account = 'adele.vance@contoso.com'; Name = 'Adele Vance'; TenantId = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'; AppId = '1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
                Audience = 'https://graph.microsoft.com'; Scopes = @('MailboxFolder.ReadWrite', 'MailboxItem.Read', 'MailboxItem.ImportExport'); ExpiresOn = [DateTimeOffset]::UtcNow.AddHours(1)
            }
            return $session
        }
    }
    It 'takes the Primary, the Archive or both (-Mailbox All) as the source: AuxArchive is refused, and so are the parameters that name a store, a PST or a folder of .eml files' {
        $svc = New-CopyService
        { Invoke-Entry $svc (Get-CopyParameters 'r1' @{ Mailbox = 'AuxArchive' }) } | Should -Throw '*is not for a copy*'
        { Invoke-Entry $svc (Get-CopyParameters 'r1' @{ Mailbox = 'All' }) } | Should -Throw '*-MailboxId is one location*'
        { Invoke-Entry $svc (Get-CopyParameters 'r1' @{ Mailbox = 'All'; MailboxId = ''; TargetMailbox = 'Archive' }) } | Should -Throw '*-TargetLayout Mirror, Primary or Archive*'
        { Invoke-Entry $svc (Get-CopyParameters 'r1' @{ TargetLayout = 'Mirror' }) } | Should -Throw '*-TargetLayout belongs to a copy of the Primary and the Archive*'
        { Invoke-Entry $svc (Get-CopyParameters 'r1' @{ TargetArchiveRoot = 'Old' }) } | Should -Throw '*-TargetArchiveRoot belongs to a copy of the Primary and the Archive*'
        { Invoke-Entry $svc (Get-CopyParameters 'r1' @{ TargetArchiveMailboxId = $script:DstArchive }) } | Should -Throw '*needs -TargetMailbox Archive*'
        { Invoke-Entry $svc (Get-CopyParameters 'r1' @{ SourceMailbox = 'Primary' }) } | Should -Throw '*it belongs to -Mode Import*'
        { Invoke-Entry $svc (Get-CopyParameters 'r1' @{ PstPath = 'x.pst' }) } | Should -Throw '*-PstPath is not for -Mode Copy*'
        { Invoke-Entry $svc (Get-CopyParameters 'r1' @{ IncludeFolder = 'Inbox' }) } | Should -Throw '*-IncludeFolder is not for -Mode Copy*'
        { Invoke-Entry $svc (Get-CopyParameters 'r1' @{ OutputPst = 'x.pst' }) } | Should -Throw '*-OutputPst belongs to*'
        { Invoke-Entry $svc (Get-CopyParameters 'r1' @{ Folder = 'Inbox'; SelectionFile = (Join-Path $script:Tmp 'x.json') }) } | Should -Throw '*exclusive*'
        $svc.Log.Count | Should -Be 0
    }
    It 'refuses the archive of another mailbox with a delegated sign-in, before any request: the address names the Primary, and the service lets an account read only its own archive' {
        $svc = New-CopyService
        { Invoke-Entry $svc (Get-CopyParameters 'oa' @{ Mailbox = 'Archive'; MailboxId = 'alex@contoso.com' }) } | Should -Throw '*archive of another mailbox*cannot be read with a delegated sign-in*'
        $svc.Log.Count | Should -Be 0
    }
    It 'says that the source and the target are the same location, by address and by id, before any request' {
        $svc = New-CopyService
        { Invoke-Entry $svc (Get-CopyParameters 'same' @{ MailboxId = $script:Me }) } | Should -Throw '*same mailbox location*'
        { Invoke-Entry $svc (Get-CopyParameters 'same' @{ MailboxId = $script:Dst; TargetMailboxId = $script:Dst }) } | Should -Throw '*same mailbox location*'
        $svc.Log.Count | Should -Be 0
        # the primary mailbox into the archive of the same user is another location: it passes this check (the address is not known to the fake service, so the run stops at the sign-in)
        $other = try { Invoke-Entry $svc (Get-CopyParameters 'same-ok' @{ MailboxId = $script:Me; TargetMailbox = 'Archive'; WhatIf = $true }); '' } catch { $_.Exception.Message }
        $other | Should -Not -Match 'same mailbox location'
    }
    It 'refuses after the sign-in, still before the first request of the export, a target that is the source under another name' {
        $svc = New-CopyService
        { Invoke-Entry $svc (Get-CopyParameters 'same-id' @{ MailboxId = $script:Dst }) } | Should -Throw '*same mailbox location*'
        $svc.RequestsOf('exportItems') | Should -Be 0
        $svc.ItemsOf($script:Dst).Count | Should -Be 0
    }
    It 'enforces the sign-in rule: a delegated sign-in writes into its own mailbox only, and the application has no mailbox of its own' {
        $svc = New-CopyService
        { Invoke-Entry $svc (Get-CopyParameters 'deleg' @{ TargetMailboxId = 'alex@contoso.com' }) } | Should -Throw '*writes into the mailbox of the signed-in account only*'
        { Invoke-Entry $svc (Get-CopyParameters 'deleg' @{ TargetMailboxId = "MBX:aaaaaaaa-0000-0000-0000-00000000000a@$($script:Tenant)" }) } | Should -Throw '*is not known as a mailbox of*'
        $app = @{ CertificateThumbprint = ('AB' * 20); ClientId = '1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d'; TenantId = $script:Tenant }
        { Invoke-Entry $svc (Get-CopyParameters 'app' $app) } | Should -Throw '*has no mailbox of its own*-TargetMailboxId*'
        $svc.Log.Count | Should -Be 0
    }
    It 'refuses a target that cannot be read, before the source is read' {
        $svc = New-CopyService
        $gone = "MBX:00000000-0000-0000-0000-000000000000@$($script:Tenant)"
        { Invoke-Entry $svc (Get-CopyParameters 'gone' @{ PrimaryMailboxId = $gone }) } | Should -Throw '*cannot be read, so nothing can be written into it*'
        $svc.RequestsOf('exportItems') | Should -Be 0
    }
    It 'takes the options of the target only in a copy, and the period in a copy' {
        $svc = New-CopyService
        { Invoke-Entry $svc @{ Mode = 'Import'; TargetMailbox = 'Archive'; StorePath = (Join-Path $script:Tmp 'x') } } | Should -Throw '*belong to -Mode Copy*'
        { Invoke-Entry $svc @{ Mode = 'Import'; ArrivedFrom = '2026-01-01'; StorePath = (Join-Path $script:Tmp 'x') } } | Should -Throw '*belong to the export modes*'
    }
    It 'takes the selection of one mailbox location only' {
        $svc = New-CopyService
        $primary = Join-Path $script:Tmp 'sel-primary.json'
        $null = Invoke-Entry $svc @{ Mode = 'Browse'; PrimaryMailboxId = $script:Src; Folder = 'Inbox'; IncludeSubfolders = $true; SaveSelection = $primary; StorePath = (Join-Path $script:Tmp 'none') }
        $other = Join-Path $script:Tmp 'sel-other.json'
        $null = Invoke-Entry $svc @{ Mode = 'Browse'; PrimaryMailboxId = $script:Dst; Folder = 'Inbox'; IncludeSubfolders = $true; SaveSelection = $other; StorePath = (Join-Path $script:Tmp 'none') }
        $two = Save-MbtBrowseSelectionSet -Selection @((Read-MbtBrowseSelection -Path $primary), (Read-MbtBrowseSelection -Path $other)) -Path (Join-Path $script:Tmp 'sel-two.json')
        { Invoke-Entry $svc (Get-CopyParameters 'sel' @{ MailboxId = ''; SelectionFile = $two; PrimaryMailboxId = $script:Dst }) } | Should -Throw '*names 2 mailbox locations*a copy takes the Primary mailbox, the Archive, or both*'
        { Invoke-Entry $svc (Get-CopyParameters 'sel' @{ SelectionFile = $primary; Mailbox = 'Archive' }) } | Should -Throw '*does not match the selection*'
    }
    It 'keeps the store of one copy for that copy: another choice, or an export store, is refused' {
        $svc = New-CopyService
        $p = Get-CopyParameters 'own' @{ Folder = 'Inbox/Projects' }
        $null = Invoke-Entry $svc $p
        $global:MbtEntryExit | Should -Be 0
        { Invoke-Entry $svc (Get-CopyParameters 'own' @{ Folder = 'Inbox' }) } | Should -Throw '*holds the copy of another choice*'
        $plain = Join-Path $script:Tmp 'store-plain'
        $null = Invoke-Entry $svc @{ Mode = 'Export'; StorePath = $plain; TempPath = (Join-Path $script:Tmp 'tmp-plain'); PrimaryMailboxId = $script:Src }
        { Invoke-Entry $svc (Get-CopyParameters 'plain' @{ StorePath = $plain }) } | Should -Throw '*was not made by a copy*'
    }
}

Describe 'Copy: the plan (-WhatIf)' {
    BeforeEach {
        Mock Connect-MbtGraph {
            $session = Set-MbtSession -Graph (& $global:MbtEntryFactory)
            $session.Token = [pscustomobject]@{
                Account = 'adele.vance@contoso.com'; Name = 'Adele Vance'; TenantId = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'; AppId = '1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
                Audience = 'https://graph.microsoft.com'; Scopes = @('MailboxFolder.ReadWrite', 'MailboxItem.Read', 'MailboxItem.ImportExport'); ExpiresOn = [DateTimeOffset]::UtcNow.AddHours(1)
            }
            return $session
        }
    }
    It 'is one plan for both halves: what is read, the store that stays, where it goes, the time, the sign-in; nothing is read, written or created' {
        $svc = New-CopyService
        $p = Get-CopyParameters 'plan' @{ WhatIf = $true }
        $text = Invoke-Entry $svc $p
        $global:MbtEntryExit | Should -Be 0
        $text | Should -Match 'The copy in one plan'
        $text | Should -Match 'read:  the Primary mailbox .*, the whole mailbox: 10 items'
        $text | Should -Match 'keep:  the export store .* stays on disk after the copy'
        $text | Should -Match "write: the Primary mailbox .*, below the folder 'Copied'"
        $text | Should -Match 'time:'
        $text | Should -Not -Match 'PST file'
        $text | Should -Not -Match 'if converted'
        $text | Should -Match 'sign-in: delegated as adele.vance@contoso.com: it reads what that account can read and writes into the mailbox of that account only'
        $text | Should -Match 'Copy planned \(WhatIf, nothing was written\)'
        $text | Should -Match 'Next steps\s+Run the same command without -WhatIf'
        Test-Path -LiteralPath (Join-Path $p.StorePath 'catalog.db') | Should -BeFalse
        $svc.RequestsOf('exportItems') | Should -Be 0
        $svc.RequestsOf('importItem') | Should -Be 0
        $svc.ItemsOf($script:Dst).Count | Should -Be 0
        (Get-FolderPaths $svc $script:Dst) | Should -Not -Contain 'Copied'
    }
    It 'names the default store after the source, its kind and the choice, in the work folder, and a period makes another store' {
        $svc = New-CopyService
        $p = Get-CopyParameters 'plan2' @{ WhatIf = $true; WorkFolder = (Join-Path $script:Tmp 'work-plan') }
        $p.Remove('StorePath')
        $name = 'copy-MBX_11111111-2222-3333-4444-555555555555@1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d-primary-[0-9a-f]{8}'
        $a = Invoke-Entry $svc $p
        $a | Should -Match ([regex]::Escape((Join-Path $script:Tmp 'work-plan\stores')) + '\\' + $name)
        $p.ArrivedFrom = '2026-01-01'
        $b = Invoke-Entry $svc $p
        $b | Should -Match 'arrived from 2026-01-01'
        ([regex]::Match($a, $name).Value) | Should -Not -Be ([regex]::Match($b, $name).Value)
        $b | Should -Match 'at most'
    }
    It 'estimates the time from the announced size at about 4 GiB per hour, and says so when the size is not announced' {
        (Get-MbtCopyEstimate -Bytes (8GB)) | Should -Be 7200
        (Get-MbtCopyEstimate -Bytes 0) | Should -Be 0
        (Format-MbtCopyDuration 7200) | Should -Be 'about 2 hours'
        (Format-MbtCopyDuration 600) | Should -Be 'about 10 minutes'
        (Format-MbtCopyDuration (3 * 86400)) | Should -Be 'about 3 days'
        $text = Show-MbtCopyPlan -Source 'alex@contoso.com' -SourceKind Primary -Target 'adele@contoso.com' -TargetKind Archive -StorePath 'D:\store' -Items 1000 -Bytes (8GB) 6>&1 | Out-String
        $text | Should -Match 'about 2 hours, dominated by the import at the observed rate of about 4 GiB per hour per target mailbox'
        $text | Should -Match 'with the same folder names as the source'
        $text | Should -Match 'the Archive mailbox adele@contoso.com'
        $none = Show-MbtCopyPlan -Source 'alex@contoso.com' -SourceKind Primary -Target 'adele@contoso.com' -TargetKind Primary -StorePath 'D:\store' -Items 10 -SizeUnknown 6>&1 | Out-String
        $none | Should -Match 'cannot be estimated'
    }
}

Describe 'Copy: the names it derives' {
    It 'gives the same key to the same choice whatever the case, and another to another choice' {
        $a = Get-MbtCopyChoiceKey -Kind Primary -Folder 'Inbox/Projects'
        (Get-MbtCopyChoiceKey -Kind primary -Folder 'inbox\PROJECTS') | Should -Be $a
        (Get-MbtCopyChoiceKey -Kind Primary -Folder 'Inbox') | Should -Not -Be $a
        (Get-MbtCopyChoiceKey -Kind Archive -Folder 'Inbox/Projects') | Should -Not -Be $a
        $period = ConvertTo-MbtArrivalPeriod -From '2026-01-01' -To $null
        (Get-MbtCopyChoiceKey -Kind Primary -Folder 'Inbox/Projects' -Period $period) | Should -Not -Be $a
        $a | Should -Match '^[0-9a-f]{8}$'
    }
    It 'names the default store after the work folder, the source and the key, and the import job after the target, the store and the folder, whatever the case' {
        (Get-MbtCopyStorePath -WorkFolder 'D:\work' -Source 'Alex@Contoso.com' -Kind Archive -Key 'ABCD1234') | Should -Be 'D:\work\stores\copy-Alex@Contoso.com-archive-abcd1234'
        $job = Get-MbtCopyJobName -Target 'adele@contoso.com' -TargetKind Primary -StorePath 'D:\work\stores\s1' -TargetRoot 'Copied' -SourceKind Primary
        $job | Should -Match '^copy-[0-9a-f]{12}$'
        (Get-MbtCopyJobName -Target 'ADELE@Contoso.com' -TargetKind primary -StorePath 'd:\WORK\stores\S1\' -TargetRoot 'copied/' -SourceKind PRIMARY) | Should -Be $job
        (Get-MbtCopyJobName -Target 'adele@contoso.com' -TargetKind Archive -StorePath 'D:\work\stores\s1' -TargetRoot 'Copied' -SourceKind Primary) | Should -Not -Be $job
        (Get-MbtCopyJobName -Target 'adele@contoso.com' -TargetKind Primary -StorePath 'D:\work\stores\s1' -TargetRoot 'Other' -SourceKind Primary) | Should -Not -Be $job
    }
}

Describe 'Copy: the chained run against the fake service' {
    BeforeEach {
        Mock Connect-MbtGraph {
            $session = Set-MbtSession -Graph (& $global:MbtEntryFactory)
            $session.Token = [pscustomobject]@{
                Account = 'adele.vance@contoso.com'; Name = 'Adele Vance'; TenantId = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'; AppId = '1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
                Audience = 'https://graph.microsoft.com'; Scopes = @('MailboxFolder.ReadWrite', 'MailboxItem.Read', 'MailboxItem.ImportExport'); ExpiresOn = [DateTimeOffset]::UtcNow.AddHours(1)
            }
            return $session
        }
    }
    It 'copies a mailbox into another in one job: the export into the store, then the import of the store, which stays; one plan of nine steps, one result, one report' {
        $svc = New-CopyService
        $events = New-EventsFile 'happy'
        $p = Get-CopyParameters 'happy' @{ EventsFile = $events }
        $text = Invoke-Entry $svc $p
        $global:MbtEntryExit | Should -Be 0
        $text | Should -Match 'export finished: 10 items exported'
        $text | Should -Match 'import finished: 10 items imported'
        $text | Should -Match 'Copy finished'
        $text | Should -Match 'Next steps\s+Open the Primary mailbox adele.vance@contoso.com in Outlook and look in the folder ''Copied'''
        $svc.ItemsOf($script:Dst).Count | Should -Be 10
        $svc.ItemsOf($script:Src).Count | Should -Be 10 -Because 'the source is only read'
        (Get-FolderPaths $svc $script:Dst) | Should -Contain 'Copied/Inbox/Projects'
        Test-Path -LiteralPath (Join-Path $p.StorePath 'catalog.db') | Should -BeTrue
        $lines = @(Get-Content -LiteralPath $events | ForEach-Object { $_ | ConvertFrom-Json -AsHashtable })
        $plan = @($lines | Where-Object { $_.e -eq 'plan' })
        $plan.Count | Should -Be 1
        $plan[0].task | Should -Be 'mailcopy'
        (@($plan[0].steps | ForEach-Object { $_.id }) -join ',') | Should -Be 'tree,list,download,verify,check,folders,send,reconcile,final'
        foreach ($id in 'tree', 'list', 'download', 'verify', 'check', 'folders', 'send', 'reconcile', 'final') {
            @($lines | Where-Object { $_.e -eq 'step' -and $_.id -eq $id -and $_.state -eq 'done' }).Count | Should -BeGreaterThan 0 -Because "step $id"
        }
        $result = @($lines | Where-Object { $_.e -eq 'result' })
        $result.Count | Should -Be 1
        $result[0].exit | Should -Be 0
        $result[0].counts.exported | Should -Be 10
        $result[0].counts.done | Should -Be 10
        $status = Get-Content -LiteralPath (Join-Path (Split-Path $events -Parent) 'status.json') -Raw | ConvertFrom-Json
        $status.task | Should -Be 'mailcopy'
        $status.mode | Should -Be 'Copy'
        Test-Path -LiteralPath (Join-Path (Split-Path $events -Parent) 'report.html') | Should -BeTrue
    }
    It 'continues where it stopped: the same command again does not read the source again and creates nothing twice' {
        $svc = New-CopyService
        $p = Get-CopyParameters 'again'
        $null = Invoke-Entry $svc $p
        $exports = $svc.RequestsOf('exportItems')
        $svc.ItemsOf($script:Dst).Count | Should -Be 10
        $text = Invoke-Entry $svc $p
        $global:MbtEntryExit | Should -Be 0
        $text | Should -Match 'the export half is already complete'
        $svc.RequestsOf('exportItems') | Should -Be $exports
        $svc.ItemsOf($script:Dst).Count | Should -Be 10
        $text | Should -Match 'job ''copy-[0-9a-f]{12}'' complete: 10 of 10 done'
    }
    It '-Force reads the source again: what arrived since is taken, and only that is imported' {
        $svc = New-CopyService
        $p = Get-CopyParameters 'force'
        $null = Invoke-Entry $svc $p
        $null = $svc.AddItem($script:Src, 'F-INBOX', 'I99', (New-Message 'Inbox 99'), 'ck-I99')
        $quiet = Invoke-Entry $svc $p
        $svc.ItemsOf($script:Dst).Count | Should -Be 10 -Because 'a finished export is not read again without -Force'
        $text = Invoke-Entry $svc ($p + @{ Force = $true })
        $global:MbtEntryExit | Should -Be 0
        $text | Should -Match 'export finished: 1 items exported'
        $svc.ItemsOf($script:Dst).Count | Should -Be 11
    }
    It 'stops in the export half and continues: nothing is written into the target, then the run ends complete' {
        $svc = New-CopyService
        $p = Get-CopyParameters 'stop1'
        $stop = Join-Path $script:Tmp 'stop1.request'
        Set-Content -LiteralPath $stop -Value 'stop'
        $text = Invoke-Entry $svc ($p + @{ StopFile = $stop; EventsFile = (New-EventsFile 'stop1') })
        $global:MbtEntryExit | Should -Be 2
        $text | Should -Match 'stop requested'
        $text | Should -Match 'Next steps\s+Run the same command again to continue'
        $svc.RequestsOf('importItem') | Should -Be 0
        $svc.ItemsOf($script:Dst).Count | Should -Be 0
        Remove-Item -LiteralPath $stop -Force
        $again = Invoke-Entry $svc ($p + @{ StopFile = $stop })
        $global:MbtEntryExit | Should -Be 0
        $svc.ItemsOf($script:Dst).Count | Should -Be 10
    }
    It 'stops in the import half and continues: the export is not read again and no item is created twice' {
        $svc = New-CopyService
        $p = Get-CopyParameters 'stop2' @{ Concurrency = 1 }
        $stop = Join-Path $script:Tmp 'stop2.request'
        # the import is slowed down once the export is over; the stop request comes after the third item (the stop file is looked at every second)
        $watcher = Start-ThreadJob -ArgumentList $svc, $stop -ScriptBlock {
            param($service, $file)
            $until = [DateTime]::UtcNow.AddSeconds(90)
            while ([DateTime]::UtcNow -lt $until) {
                if ($service.RequestsOf('createImportSession') -ge 1) { $service.LatencyMs = 400 }
                if ($service.RequestsOf('importItem') -ge 3) { [IO.File]::WriteAllText($file, 'stop'); return }
                Start-Sleep -Milliseconds 20
            }
        }
        try { $text = Invoke-Entry $svc ($p + @{ StopFile = $stop }) } finally { $null = Wait-Job $watcher -Timeout 5; Remove-Job $watcher -Force }
        # 2: not complete, run again; 3: a request that was in flight when the job stopped may have created its item (the next run looks for it first)
        $global:MbtEntryExit | Should -BeIn 2, 3
        $afterStop = $svc.ItemsOf($script:Dst).Count
        $afterStop | Should -BeGreaterThan 0
        $afterStop | Should -BeLessThan 10
        $exports = $svc.RequestsOf('exportItems')
        Remove-Item -LiteralPath $stop -Force
        $svc.LatencyMs = 0
        $again = Invoke-Entry $svc ($p + @{ StopFile = $stop })
        $global:MbtEntryExit | Should -Be 0
        $again | Should -Match 'the export half is already complete'
        $svc.RequestsOf('exportItems') | Should -Be $exports
        $svc.ItemsOf($script:Dst).Count | Should -Be 10
    }
    It 'imports what the export half could read, ends not complete, and the next run does only what is left' {
        $svc = New-CopyService
        $null = $svc.BadItemIds.Add('I3')
        $p = Get-CopyParameters 'bad'
        $text = Invoke-Entry $svc $p
        $global:MbtEntryExit | Should -Be 2
        $text | Should -Match 'exit code 2: the export is not complete'
        $svc.ItemsOf($script:Dst).Count | Should -Be 9
        $null = $svc.BadItemIds.Remove('I3')
        $again = Invoke-Entry $svc $p
        $global:MbtEntryExit | Should -Be 0
        $again | Should -Match 'export finished: 1 items exported'
        $svc.ItemsOf($script:Dst).Count | Should -Be 10
    }
    It 'copies one folder with -Folder, into the archive of the target with -TargetMailbox Archive, with the same folder names without -TargetRoot' {
        $svc = New-CopyService
        $p = Get-CopyParameters 'folder' @{ Folder = 'Inbox/Projects'; TargetMailbox = 'Archive' }
        $p.Remove('TargetRoot')
        $text = Invoke-Entry $svc $p
        $global:MbtEntryExit | Should -Be 0
        $text | Should -Match 'import finished: 4 items imported'
        $svc.ItemsOf($script:DstArchive).Count | Should -Be 4
        $svc.ItemsOf($script:Dst).Count | Should -Be 0
        (Get-FolderPaths $svc $script:DstArchive) | Should -Contain 'Inbox/Projects'
    }
    It 'follows a mapping file below the target folder' {
        $svc = New-CopyService
        $map = Join-Path $script:Tmp 'map.psd1'
        Set-Content -LiteralPath $map -Value "@{ 'Inbox/Projects' = 'Archive/Work' }"
        $null = Invoke-Entry $svc (Get-CopyParameters 'map' @{ MappingFile = $map })
        $global:MbtEntryExit | Should -Be 0
        (Get-FolderPaths $svc $script:Dst) | Should -Contain 'Copied/Archive/Work'
        { Invoke-Entry $svc (Get-CopyParameters 'map2' @{ MappingFile = (Join-Path $script:Tmp 'nothing.psd1') }) } | Should -Throw '*-MappingFile not found*'
    }
    It 'copies the same source into two targets from the one store: the second copy does not read the source again, its import job is its own' {
        $svc = New-CopyService
        $shared = Join-Path $script:Tmp 'store-two'
        $null = Invoke-Entry $svc (Get-CopyParameters 'two' @{ StorePath = $shared })
        $exports = $svc.RequestsOf('exportItems')
        $text = Invoke-Entry $svc (Get-CopyParameters 'two' @{ StorePath = $shared; TargetMailbox = 'Archive'; TargetRoot = 'FromAdele' })
        $global:MbtEntryExit | Should -Be 0
        $text | Should -Match 'the export half is already complete'
        $svc.RequestsOf('exportItems') | Should -Be $exports
        $svc.ItemsOf($script:Dst).Count | Should -Be 10
        $svc.ItemsOf($script:DstArchive).Count | Should -Be 10
    }
    It 'takes the items that arrived inside the period, leaves the others out and says so' {
        $svc = [MailboxTransfer.Tests.FakeMailboxService]::new()
        $null = $svc.AddMailbox($script:Src); $null = $svc.AddFolder($script:Src, 'F-INBOX', 'Inbox')
        foreach ($m in $script:Dst) { $null = $svc.AddMailbox($m); $null = $svc.AddFolder($m, 'D-INBOX', 'Inbox') }
        $null = $svc.AddItem($script:Src, 'F-INBOX', 'OLD1', (New-Message 'Old 1'), 'ck1', [datetime]::new(2025, 3, 1, 10, 0, 0, [DateTimeKind]::Utc))
        $null = $svc.AddItem($script:Src, 'F-INBOX', 'NEW1', (New-Message 'New 1'), 'ck2', [datetime]::new(2026, 3, 10, 10, 0, 0, [DateTimeKind]::Utc))
        $null = $svc.AddItem($script:Src, 'F-INBOX', 'NEW2', (New-Message 'New 2'), 'ck3', [datetime]::new(2026, 3, 20, 10, 0, 0, [DateTimeKind]::Utc))
        $text = Invoke-Entry $svc (Get-CopyParameters 'period' @{ ArrivedFrom = '2026-03-01'; ArrivedTo = '2026-03-31' })
        $global:MbtEntryExit | Should -Be 0
        $svc.ItemsOf($script:Dst).Count | Should -Be 2
        $text | Should -Match '1 outside the period'
    }
    It 'copies a selection of one mailbox location' {
        $svc = New-CopyService
        $sel = Join-Path $script:Tmp 'sel-copy.json'
        $null = Invoke-Entry $svc @{ Mode = 'Browse'; PrimaryMailboxId = $script:Src; Folder = 'Inbox/Projects'; IncludeSubfolders = $true; SaveSelection = $sel; StorePath = (Join-Path $script:Tmp 'none') }
        $p = Get-CopyParameters 'selection' @{ SelectionFile = $sel }
        $p.Remove('MailboxId'); $p.PrimaryMailboxId = $script:Dst
        $text = Invoke-Entry $svc $p
        $global:MbtEntryExit | Should -Be 0
        $svc.ItemsOf($script:Dst).Count | Should -Be 4
    }
}

Describe 'Copy: the Primary and the Archive of a mailbox together' {
    BeforeEach {
        Mock Connect-MbtGraph {
            $session = Set-MbtSession -Graph (& $global:MbtEntryFactory)
            $session.Token = [pscustomobject]@{
                Account = $null; Name = ''; TenantId = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'; AppId = '1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d'; Audience = 'https://graph.microsoft.com'
                Scopes = @(); Roles = @('MailboxFolder.ReadWrite.All', 'MailboxItem.Read.All', 'MailboxItem.ImportExport.All'); AppOnly = $true; ExpiresOn = [DateTimeOffset]::UtcNow.AddHours(1)
            }
            return $session
        }
    }
    BeforeAll {
        $script:SrcArchive = "MBX:11111111-2222-3333-4444-5555555555aa@$($script:Tenant)"
        # the source Primary (10 items) and Archive (3 items), the target Primary and Archive (an Inbox each)
        function New-BothService {
            $svc = New-CopyService
            $null = $svc.AddMailbox($script:SrcArchive)
            $null = $svc.AddFolder($script:SrcArchive, 'F-AINBOX', 'Inbox')
            $null = $svc.AddFolder($script:SrcArchive, 'F-OLD', 'Old', 'F-AINBOX')
            foreach ($i in 1..2) { $null = $svc.AddItem($script:SrcArchive, 'F-AINBOX', "A$i", (New-Message "Archive $i"), "ck-A$i") }
            $null = $svc.AddItem($script:SrcArchive, 'F-OLD', 'A3', (New-Message 'Archive 3'), 'ck-A3')
            return $svc
        }
        function Get-BothParameters {
            param([string]$Name, [hashtable]$More = @{})
            $p = @{
                Mode = 'Copy'; CertificateThumbprint = ('AB' * 20); ClientId = '1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d'; TenantId = $script:Tenant; UserPrincipalName = 'alex@contoso.com'
                Mailbox = 'All'; PrimaryMailboxId = $script:Src; ArchiveMailboxId = $script:SrcArchive; TargetMailboxId = $script:Dst; TargetArchiveMailboxId = $script:DstArchive
                StorePath = (Join-Path $script:Tmp "store-$Name"); TempPath = (Join-Path $script:Tmp "tmp-$Name"); TargetRoot = 'FromPrimary'; TargetArchiveRoot = 'FromArchive'; ReconcileDelaySeconds = 0; IntervalSeconds = 1
            }
            foreach ($k in $More.Keys) { $p[$k] = $More[$k] }
            return $p
        }
    }
    It 'reads both places into one store, then imports it once for each part: the Primary into the target Primary and the Archive into the target Archive (Mirror), with the steps of both parts in the events' {
        $svc = New-BothService
        $events = New-EventsFile 'both'
        $p = Get-BothParameters 'both' @{ EventsFile = $events }
        $text = Invoke-Entry $svc $p
        $global:MbtEntryExit | Should -Be 0
        $svc.ItemsOf($script:Dst).Count | Should -Be 10
        $svc.ItemsOf($script:DstArchive).Count | Should -Be 3
        $svc.ItemsOf($script:Src).Count | Should -Be 10 -Because 'the source is only read'
        (Get-FolderPaths $svc $script:Dst) | Should -Contain 'FromPrimary/Inbox/Projects'
        (Get-FolderPaths $svc $script:DstArchive) | Should -Contain 'FromArchive/Inbox/Old'
        (Get-FolderPaths $svc $script:Dst) | Should -Not -Contain 'FromArchive'
        $text | Should -Match 'Copy finished'
        $lines = @(Get-Content -LiteralPath $events | ForEach-Object { $_ | ConvertFrom-Json -AsHashtable })
        $plan = @($lines | Where-Object { $_.e -eq 'plan' })
        $plan.Count | Should -Be 1
        (@($plan[0].steps | ForEach-Object { $_.id }) -join ',') | Should -Be 'tree,list,download,verify,check,folders,send,reconcile,final,check-2,folders-2,send-2,reconcile-2,final-2'
        @($plan[0].steps | Where-Object { $_.id -eq 'send-2' })[0].title | Should -Be 'Send the items (Archive part)'
        foreach ($id in 'tree', 'list', 'download', 'verify', 'check', 'folders', 'send', 'reconcile', 'final', 'check-2', 'folders-2', 'send-2', 'reconcile-2', 'final-2') {
            @($lines | Where-Object { $_.e -eq 'step' -and $_.id -eq $id -and $_.state -eq 'done' }).Count | Should -BeGreaterThan 0 -Because "step $id"
        }
        $result = @($lines | Where-Object { $_.e -eq 'result' })
        $result.Count | Should -Be 1
        $result[0].exit | Should -Be 0
        $result[0].counts.exported | Should -Be 13
        $result[0].counts.done | Should -Be 13
        $result[0].counts.total | Should -Be 13
        $status = Get-Content -LiteralPath (Join-Path (Split-Path $events -Parent) 'status.json') -Raw | ConvertFrom-Json
        $status.task | Should -Be 'mailcopy'
        $status.destination | Should -Match 'Primary \+ Archive'
        $status.source | Should -Match 'Primary \+ Archive'
    }
    It 'reads the archive of another mailbox with the application: the address is the mailbox whose archive is looked up (here by -ArchiveMailboxId), not the Primary it names' {
        $svc = New-BothService
        $p = Get-BothParameters 'other-archive' @{ Mailbox = 'Archive'; MailboxId = 'bea@contoso.com'; UserPrincipalName = 'carl@contoso.com' }
        $p.Remove('TargetArchiveRoot'); $p.Remove('TargetArchiveMailboxId')
        $text = Invoke-Entry $svc $p
        $global:MbtEntryExit | Should -Be 0
        $svc.ItemsOf($script:Dst).Count | Should -Be 3 -Because 'the 3 items of the archive, not the 10 of the Primary'
        (Get-FolderPaths $svc $script:Dst) | Should -Contain 'FromPrimary/Inbox/Old'
    }
    It 'continues where it stopped: a part that is complete is not done again, an item that was read late is imported alone, and nothing is created twice' {
        $svc = New-BothService
        $null = $svc.BadItemIds.Add('A2')
        $p = Get-BothParameters 'both-resume'
        $text = Invoke-Entry $svc $p
        $global:MbtEntryExit | Should -Be 2
        $svc.ItemsOf($script:Dst).Count | Should -Be 10
        $svc.ItemsOf($script:DstArchive).Count | Should -Be 2
        $null = $svc.BadItemIds.Remove('A2')
        $again = Invoke-Entry $svc $p
        $global:MbtEntryExit | Should -Be 0
        $svc.ItemsOf($script:Dst).Count | Should -Be 10
        $svc.ItemsOf($script:DstArchive).Count | Should -Be 3
        $third = Invoke-Entry $svc $p
        $global:MbtEntryExit | Should -Be 0
        $third | Should -Match 'the export half is already complete'
        $svc.ItemsOf($script:Dst).Count | Should -Be 10
        $svc.ItemsOf($script:DstArchive).Count | Should -Be 3
    }
    It 'puts both parts into the target Primary (-TargetLayout Primary) or into the target Archive (-TargetLayout Archive), each below its own folder' {
        $svc = New-BothService
        $null = Invoke-Entry $svc (Get-BothParameters 'both-primary' @{ TargetLayout = 'Primary' })
        $global:MbtEntryExit | Should -Be 0
        $svc.ItemsOf($script:Dst).Count | Should -Be 13
        $svc.ItemsOf($script:DstArchive).Count | Should -Be 0
        (Get-FolderPaths $svc $script:Dst) | Should -Contain 'FromArchive/Inbox/Old'
        $svc2 = New-BothService
        $pa = Get-BothParameters 'both-archive' @{ TargetLayout = 'archive' }
        $pa.Remove('TargetArchiveRoot')
        $null = Invoke-Entry $svc2 $pa
        $global:MbtEntryExit | Should -Be 0
        $svc2.ItemsOf($script:Dst).Count | Should -Be 0
        $svc2.ItemsOf($script:DstArchive).Count | Should -Be 13
        (Get-FolderPaths $svc2 $script:DstArchive) | Should -Contain 'FromPrimary/Inbox/Old' -Because 'without -TargetArchiveRoot the Archive part goes below -TargetRoot too'
    }
    It 'refuses before any request a copy whose part would write into the place it reads, and a target without the id of its archive' {
        $svc = New-BothService
        { Invoke-Entry $svc (Get-BothParameters 'both-same' @{ TargetMailboxId = $script:Src }) } | Should -Throw '*same mailbox location*'
        { Invoke-Entry $svc (Get-BothParameters 'both-same2' @{ TargetArchiveMailboxId = $script:SrcArchive }) } | Should -Throw '*same mailbox location*'
        $svc.RequestsOf('exportItems') | Should -Be 0
        $svc.ItemsOf($script:Dst).Count | Should -Be 0
    }
    It 'plans both parts in one plan (-WhatIf) and creates nothing' {
        $svc = New-BothService
        $p = Get-BothParameters 'both-plan' @{ WhatIf = $true }
        $text = Invoke-Entry $svc $p
        $global:MbtEntryExit | Should -Be 0
        $text | Should -Match 'The copy in one plan'
        $text | Should -Match 'read:  the Primary \+ Archive mailbox'
        $text | Should -Match 'write: part 1, the Primary of the source into the Primary mailbox .*below the folder ''FromPrimary'''
        $text | Should -Match 'write: part 2, the Archive of the source into the Archive mailbox .*below the folder ''FromArchive'''
        $text | Should -Match 'Copy planned \(WhatIf, nothing was written\)'
        Test-Path -LiteralPath (Join-Path $p.StorePath 'catalog.db') | Should -BeFalse
        $svc.RequestsOf('importItem') | Should -Be 0
        $svc.ItemsOf($script:Dst).Count | Should -Be 0
    }
    It 'names the steps of the second part in the events of the job only while the second import runs' {
        $m = Get-Module MailboxTransfer
        & $m { $script:JobInfo = @{ Phase = 'import'; Part = 2 } }
        try {
            (& $m { Get-MbtJobStepKey 'send' }) | Should -Be 'send-2'
            (& $m { Get-MbtJobStepKey 'tree' }) | Should -Be 'tree'
            (& $m { Get-MbtJobStepKey 'send-2' }) | Should -Be 'send-2'
            (& $m { Get-MbtJobStepBase 'send-2' }) | Should -Be 'send'
            (& $m { Get-MbtJobStepBase 'final' }) | Should -Be 'final'
            & $m { $script:JobInfo.Part = 1 }
            (& $m { Get-MbtJobStepKey 'send' }) | Should -Be 'send'
        } finally { & $m { $script:JobInfo = $null } }
    }
}

Describe 'Copy: with the sign-in of the application' {
    BeforeEach {
        Mock Connect-MbtGraph {
            $session = Set-MbtSession -Graph (& $global:MbtEntryFactory)
            $session.Token = [pscustomobject]@{
                Account = $null; Name = ''; TenantId = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'; AppId = '1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d'; Audience = 'https://graph.microsoft.com'
                Scopes = @(); Roles = @('MailboxFolder.ReadWrite.All', 'MailboxItem.Read.All', 'MailboxItem.ImportExport.All'); AppOnly = $true; ExpiresOn = [DateTimeOffset]::UtcNow.AddHours(1)
            }
            return $session
        }
    }
    It 'copies the mailbox of one person into the mailbox of another when the application is in the scope of both' {
        $svc = New-CopyService
        $p = @{
            Mode = 'Copy'; CertificateThumbprint = ('AB' * 20); ClientId = '1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d'; TenantId = $script:Tenant; UserPrincipalName = 'alex@contoso.com'
            Mailbox = 'Primary'; MailboxId = $script:Src; TargetMailboxId = $script:Dst; StorePath = (Join-Path $script:Tmp 'store-app'); TempPath = (Join-Path $script:Tmp 'tmp-app')
            TargetRoot = 'FromAlex'; ReconcileDelaySeconds = 0; IntervalSeconds = 1
        }
        $text = Invoke-Entry $svc $p
        $global:MbtEntryExit | Should -Be 0
        $text | Should -Match 'signed in as the application'
        $svc.ItemsOf($script:Dst).Count | Should -Be 10
        $p.StorePath = Join-Path $script:Tmp 'store-app2'; $p.WhatIf = $true
        $plan = Invoke-Entry $svc $p
        $plan | Should -Match 'sign-in: as the application: it reads and writes the mailboxes of its scope'
    }
}
