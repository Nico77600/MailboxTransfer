#Requires -Version 7.6
<#
    The real adapters of the console (package\module\MailboxTransfer.Gui.Engine.ps1) on a worker that runs inline, over stubs of the engine calls that reach the cloud.
    The fake engine of the other console tests hands the pages ready objects, always several of them, and never calls these adapters; the worker of the console hands
    every result as an array. Here each adapter is called with 0, 1 and several results shaped like the real ones (the C# classes of the engine, the real functions of the
    module for everything that is offline), and what the page does with the answer is checked under StrictMode Latest: one object or an array, and the properties it reads.
    The functions are defined in the script scope of this script, where the console has them: the closures of the adapters find them the way they do in the console.
    Prints one JSON object: { Errors: [..], Checks: [ { Name; Ok; Detail } ] }. Nothing contacts the cloud. Run by tests\Gui.Engine.Tests.ps1 in its own pwsh.
#>
param([string]$Package, [string]$Work)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Import-StoreEngine.ps1')
Import-StoreEngine -PackageRoot $Package
foreach ($f in 'Core', 'Run', 'Jobs', 'Browse', 'State', 'Pages.Run', 'Engine') { . (Join-Path $Package "module\MailboxTransfer.Gui.$f.ps1") }
$script:GuiRoot = $Package
$script:Errors = [Collections.Generic.List[string]]::new()
$script:Checks = [Collections.Generic.List[object]]::new()

function Check([string]$CheckName, [scriptblock]$CheckBlock) {
    try { & $CheckBlock; $script:Checks.Add([pscustomobject]@{ Name = $CheckName; Ok = $true; Detail = '' }) }
    catch { $script:Checks.Add([pscustomobject]@{ Name = $CheckName; Ok = $false; Detail = ("{0} | {1}" -f ($_.Exception.Message -replace '\s+', ' '), ((@($_.ScriptStackTrace -split "`n") | Select-Object -First 2) -join ' < ')) }) }
}
function Assert-That([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function Assert-Equal($Actual, $Expected, [string]$What) {
    $a = ConvertTo-Json -InputObject @($Actual) -Compress -Depth 4; $e = ConvertTo-Json -InputObject @($Expected) -Compress -Depth 4
    if ($a -cne $e) { throw "$What : expected $e, got $a" }
}
function Assert-Array($Value, [string]$What) { if ($Value -isnot [array]) { throw "$What : an array was expected, got $(if ($null -eq $Value) { 'nothing' } else { $Value.GetType().Name })" } }
function Assert-One($Value, [string]$What) { if ($Value -is [array]) { throw "$What : one object was expected, got an array of $($Value.Count)" } }
function Assert-Reads($Object, [string[]]$Name, [string]$What) {
    # the properties a page reads, read as the console reads them (StrictMode Latest): a missing one is an error
    Set-StrictMode -Version Latest
    foreach ($n in $Name) { try { $null = $Object.$n } catch { throw "$What : the page reads .$n and it is not there ($($_.Exception.Message))" } }
}
function Invoke-Call {
    # Runs a call of the adapter. $Block gets the OnDone and the OnError the page would give. Returns what they received.
    param([Parameter(Mandatory)][scriptblock]$Block)
    $script:Calls = 0; $script:Got = $null; $script:Err = $null; $script:Failed = 0
    $ok = { param($x) $script:Calls++; $script:Got = $x }
    $bad = { param($m) $script:Failed++; $script:Err = $m }
    & $Block $ok $bad
    [pscustomobject]@{ Got = $script:Got; Calls = $script:Calls; Err = $script:Err; Failed = $script:Failed }
}

function New-TestLocation([string]$Kind) { New-MbtBrowseLocation -Kind $Kind -MailboxId "MBX:$($Kind.ToLowerInvariant())@contoso" -Source 'test' }
function New-TestFolder {
    param([string]$Id, [string]$Name, [int]$Count = 0, [int]$Children = 0, [string]$WellKnown, [long]$Size = -1, [switch]$NotAvailable)
    $f = [MailboxTransfer.BrowseFolder]::new()
    $f.Id = $Id; $f.DisplayName = $Name; $f.Path = $Name; $f.TotalItemCount = $Count; $f.ChildFolderCount = $Children; $f.WellKnownName = $WellKnown; $f.SizeBytes = $Size
    if ($NotAvailable) { $f.State = [MailboxTransfer.BrowseFolderState]::ContentNotAvailable }
    return $f
}
function New-TestItem {
    param([string]$Id, [string]$Subject = 'Quarterly report', [switch]$NoDetails)
    $i = [MailboxTransfer.BrowseItem]::new()
    $i.Id = $Id; $i.Size = 12345; $i.Source = $(if ($NoDetails) { [MailboxTransfer.MailboxBrowser]::ImportExportApi } else { [MailboxTransfer.MailboxBrowser]::MailApi })
    if (-not $NoDetails) { $i.Subject = $Subject; $i.FromName = 'Alex Wilber'; $i.FromAddress = 'alex.wilber@contoso.com'; $i.HasAttachments = $true }
    $i.Received = [datetimeoffset]::new(2026, 11, 3, 12, 0, 0, [TimeSpan]::Zero); $i.Created = $i.Received
    return $i
}
function New-TestPage([int]$Count, [string]$Next, [switch]$NoDetails, [switch]$NotAvailable) {
    $p = [MailboxTransfer.BrowsePage]::new()
    for ($n = 1; $n -le $Count; $n++) { $p.Items.Add((New-TestItem "i$n" "Subject $n" -NoDetails:$NoDetails)) }
    if ($Next) { $p.NextToken = $Next }
    if ($NotAvailable) { $p.State = [MailboxTransfer.BrowseFolderState]::ContentNotAvailable; $p.StateMessage = 'The service cannot read it yet.' }
    return $p
}
function New-TestPreview([int]$To, [int]$Attachments, [switch]$NoBody) {
    $p = [MailboxTransfer.BrowsePreview]::new()
    $p.Header = New-TestItem 'i1' 'Signed contract'
    for ($n = 1; $n -le $To; $n++) { $r = [MailboxTransfer.BrowseRecipient]::new(); $r.Name = "Person $n"; $r.Address = "p$n@contoso.com"; $p.To.Add($r) }
    for ($n = 1; $n -le $Attachments; $n++) { $a = [MailboxTransfer.BrowseAttachment]::new(); $a.Name = "file$n.pdf"; $a.Size = 4096; $p.Attachments.Add($a) }
    if (-not $NoBody) { $p.BodyText = 'Hello'; $p.BodyPreview = 'Hello' }
    return $p
}
function New-TestRoot {
    param([string]$Kind = 'primary', [switch]$HeadersOnly)
    $root = [MailboxTransferGui.FolderNode]::new("r-$Kind", 'Primary mailbox', 0, -1)
    $root.IsRoot = $true; $root.Kind = $Kind; $root.HeadersOnly = [bool]$HeadersOnly
    $root.Tag = @{ Location = (New-TestLocation 'Primary') }
    return $root
}
function New-TestFolderNode {
    # a root with one folder in it, as the browser builds them; returns the folder node
    param([int]$Count, [switch]$HeadersOnly)
    $root = New-TestRoot -HeadersOnly:$HeadersOnly
    $folder = [MailboxTransferGui.FolderNode]::new('f1', 'Inbox', $Count, 100)
    $folder.Tag = New-TestFolder 'f1' 'Inbox' -Count $Count
    $root.SetChildren([MailboxTransferGui.FolderNode[]]@($folder))
    return $folder
}
function New-TestSession([string[]]$Scopes = @(), [string[]]$Roles = @(), [switch]$AppOnly) {
    $session = [pscustomobject]@{ Token = [pscustomobject]@{ Account = 'adele.vance@contoso.com'; Name = 'Adele'; TenantId = $script:Guid; Scopes = $Scopes; Roles = $Roles; AppOnly = [bool]$AppOnly } }
    return { $session }.GetNewClosure()
}

$null = New-Item -ItemType Directory -Force -Path $Work
$script:Guid = '7c0d1e2f-4a5b-4c6d-8e9f-0a1b2c3d4e5f'
$entry = Join-Path $Package 'Invoke-MailboxTransfer.ps1'
$engine = New-MbtGuiRealEngine -PackageRoot $Package -EntryScript $entry -WorkFolder $Work
# the worker of this script runs the script at once, as the fake engine does; everything else is the adapter of the console
$engine.Worker.Inline = $true
$script:Gui = @{ Engine = $engine; S = (New-MbtGuiState -WorkFolder $Work); EntryScript = $entry; Pollers = [Collections.Generic.List[scriptblock]]::new(); Tick = 0; Browsers = @{}; Job = $null; Run = $null; CloseWhenStopped = $false; LastTickError = $null }
$script:Gui.S.Mailbox = 'adele.vance@contoso.com'
Initialize-MbtGuiTypes

# What the engine calls answer: the same parameters as the real functions (a parameter the real one has not is an error), a body that answers from $script:Stub.
# Set after the engine is made: Import-Module -Force (New-MbtGuiRealEngine) puts the real functions in this scope again.
$script:Stub = @{ Session = $null; Locations = @(); LocationsError = $null; ResolveFails = $false; Folders = @(); Page = $null; Preview = $null }
function Set-Stub([string]$Name, [scriptblock]$Body) {
    $real = Get-Command -Name $Name -CommandType Function -Module MailboxTransfer
    $params = foreach ($p in $real.Parameters.Values) {
        if ($p.Name -in [Management.Automation.PSCmdlet]::CommonParameters -or $p.Name -in [Management.Automation.PSCmdlet]::OptionalCommonParameters) { continue }
        if ($p.SwitchParameter) { "[switch]`$$($p.Name)" } else { "[$($p.ParameterType.FullName)]`$$($p.Name)" }
    }
    $bodyRef = $Body
    $stub = [scriptblock]::Create("[CmdletBinding()] param($($params -join ', ')) & `$bodyRef `$PSBoundParameters")
    Set-Item -Path "Function:script:$Name" -Value $stub.GetNewClosure()
}
Set-Stub Connect-MbtGraph { param($p) & $script:Stub.Session }
Set-Stub Get-MbtBrowseLocation { param($p) if ($script:Stub.LocationsError) { throw $script:Stub.LocationsError }; $script:Stub.Locations }
Set-Stub Resolve-MbtMailboxAddress { param($p) if ($script:Stub.ResolveFails) { throw 'The address is not known.' } else { 'MBX:resolved@contoso' } }
Set-Stub Get-MbtSession { param($p) [pscustomobject]@{ TenantId = 'contoso.onmicrosoft.com' } }
Set-Stub Get-MbtBrowseFolder { param($p) $script:Stub.Folders }
Set-Stub Get-MbtBrowseItem { param($p) $script:Stub.Page }
Set-Stub Get-MbtBrowsePreview { param($p) $script:Stub.Preview }
# the code of a device-code sign-in is not part of these checks
function Step-MbtGuiSignInCode { }
# no child process is started (the command is taken off the queue, as the console does when it starts it)
function Start-MbtGuiProcess { param($Job) if ($Job.Queue.Count -eq 0) { return $false }; $null = $Job.Queue.Dequeue(); return $true }


#region The worker of the console hands every result as an array ----------------------------------------------------------------------------------------------------
function Wait-TestWorker($Worker) {
    if ($Worker.Inline) { return }
    $until = [datetime]::UtcNow.AddSeconds(60)
    while (($Worker.Current -or $Worker.Queue.Count) -and [datetime]::UtcNow -lt $until) { Step-MbtGuiWork $Worker; Start-Sleep -Milliseconds 15 }
    Step-MbtGuiWork $Worker
}
$real = New-MbtGuiWorker
$inline = [pscustomobject]@{ Inline = $true }
foreach ($case in @(@{ N = 0; S = { } }, @{ N = 1; S = { 'one' } }, @{ N = 1; S = { , @('a', 'b', 'c') } }, @{ N = 3; S = { 'a'; 'b'; 'c' } })) {
    foreach ($w in @(@{ Name = 'in the runspace'; Worker = $real }, @{ Name = 'inline'; Worker = $inline })) {
        Check "worker: $($case.N) result(s) of a script arrive as an array of $($case.N) $($w.Name) ($($case.S.ToString().Trim()))" {
            $script:Seen = $null
            Start-MbtGuiWork -Worker $w.Worker -Script $case.S -OnDone { param($r) $script:Seen = $r }
            Wait-TestWorker $w.Worker
            Assert-Array $script:Seen 'the result'
            Assert-That ($script:Seen.Count -eq $case.N) "expected $($case.N) element(s), got $($script:Seen.Count)"
        }
    }
}
Check 'worker: a result that is one object arrives as an array of one, and the page takes the first' {
    $script:Seen = $null
    Start-MbtGuiWork -Worker $real -Script { [pscustomobject]@{ Missing = @() } } -OnDone { param($r) $script:Seen = $r }
    Wait-TestWorker $real
    Assert-Array $script:Seen 'the result'
    $threw = $false; try { Assert-Reads $script:Seen 'Missing' 'an array of one' } catch { $threw = $true }
    Assert-That $threw 'the property of the object was read through the array: member enumeration of an array of one whose property is empty is an error under StrictMode'
    Assert-Reads @($script:Seen)[0] 'Missing' 'the first element'
}
Check 'worker: the calls run one after the other, the answers come on the next ticks, in the order of the calls' {
    $script:Order = [Collections.Generic.List[string]]::new()
    Start-MbtGuiWork -Worker $real -Script { Start-Sleep -Milliseconds 300; 'first' } -OnDone { param($r) $script:Order.Add([string]$r[0]) }
    Start-MbtGuiWork -Worker $real -Script { 'second' } -OnDone { param($r) $script:Order.Add([string]$r[0]) }
    Start-MbtGuiWork -Worker $real -Script { throw 'third fails' } -OnDone { param($r) $script:Order.Add('done') } -OnError { param($m) $script:Order.Add('error: ' + $m) }
    Assert-That ($script:Order.Count -eq 0) 'an answer came at once'
    Wait-TestWorker $real
    Assert-Equal @($script:Order)[0..1] @('first', 'second') 'the order'
    Assert-That ($script:Order.Count -eq 3 -and $script:Order[2] -like 'error: *third fails*') "the third answer: $($script:Order -join ' / ')"
}
Check 'worker: a page step that raises in OnDone or in OnError is written down, and the next call still answers' {
    $script:Order = [Collections.Generic.List[string]]::new()
    Start-MbtGuiWork -Worker $real -Script { 1 } -OnDone { param($r) throw 'the page broke' }
    Start-MbtGuiWork -Worker $real -Script { throw 'x' } -OnError { param($m) throw 'the error page broke' }
    Start-MbtGuiWork -Worker $real -Script { 'after' } -OnDone { param($r) $script:Order.Add([string]$r[0]) }
    Wait-TestWorker $real
    Assert-Equal @($script:Order) @('after') 'the call after the two that raised'
}
Stop-MbtGuiWorker $real
#endregion

#region Sign-in ----------------------------------------------------------------------------------------------------------------------------------------------------
foreach ($case in @(
        @{ Name = 'every scope'; Scopes = @('MailboxFolder.Read', 'MailboxItem.Read', 'MailboxItem.ImportExport', 'Mail.Read'); Missing = 0; CanRead = $true }
        @{ Name = 'one scope missing'; Scopes = @('MailboxFolder.Read', 'MailboxItem.Read', 'Mail.Read'); Missing = 1; CanRead = $true }
        @{ Name = 'no scope at all'; Scopes = @(); Missing = 3; CanRead = $false }
        @{ Name = 'one scope'; Scopes = @('Mail.Read'); Missing = 3; CanRead = $true })) {
    Check "sign-in: ONE object (never the array of one) whose Missing the page reads, $($case.Name)" {
        $script:Stub.Session = New-TestSession -Scopes $case.Scopes
        $r = Invoke-Call { param($ok, $bad) & $engine.SignIn 'adele.vance@contoso.com' $false '' $ok $bad $true $script:Guid $script:Guid $null 0 '' $null }
        Assert-That (-not $r.Err) "OnError: $($r.Err)"
        Assert-That ($r.Calls -eq 1) "OnDone was called $($r.Calls) times"
        Assert-One $r.Got 'the sign-in'
        Assert-Reads $r.Got 'Account', 'TenantId', 'Scopes', 'Missing', 'CanReadMail', 'Access' 'the sign-in'
        Assert-That (@($r.Got.Missing).Count -eq $case.Missing) "Missing holds $(@($r.Got.Missing).Count), expected $($case.Missing)"
        Assert-That (@(Get-MbtGuiProp $r.Got 'Missing' @()).Count -eq $case.Missing) 'the page reads Missing through Get-MbtGuiProp'
        Assert-That ($r.Got.CanReadMail -eq $case.CanRead) 'CanReadMail'
    }
}
foreach ($case in @(
        @{ Name = 'no role (permissions come from Exchange RBAC)'; Roles = @(); Rbac = $true }
        @{ Name = 'one role'; Roles = @('Mail.Read'); Rbac = $false }
        @{ Name = 'several roles'; Roles = @('Mail.Read', 'MailboxItem.ImportExport.All', 'MailboxFolder.Read.All'); Rbac = $false })) {
    Check "sign-in as the application: $($case.Name)" {
        $script:Stub.Session = New-TestSession -Scopes @() -Roles $case.Roles -AppOnly
        $app = @{ Tenant = 'contoso.onmicrosoft.com'; ClientId = $script:Guid; Thumbprint = 'A1B2C3D4E5F60718293A4B5C6D7E8F9012345678' }
        $r = Invoke-Call { param($ok, $bad) & $engine.SignIn 'adele.vance@contoso.com' $false '' $ok $bad $true '' '' $app 0 '' $null }
        Assert-That (-not $r.Err) "OnError: $($r.Err)"
        Assert-One $r.Got 'the sign-in'
        Assert-Reads $r.Got 'Account', 'TenantId', 'Scopes', 'Missing', 'CanReadMail', 'Access' 'the sign-in'
        if ($case.Rbac) { Assert-That (@($r.Got.Missing).Count -eq 0 -and $r.Got.CanReadMail -and $r.Got.Access -like '*RBAC*') 'a token without a role is reported as granted by Exchange RBAC, nothing missing' }
    }
}
Check 'sign-in that fails answers OnError with the text and never OnDone' {
    $script:Stub.Session = { throw 'AADSTS50173: the grant has expired' }
    $r = Invoke-Call { param($ok, $bad) & $engine.SignIn 'adele.vance@contoso.com' $false '' $ok $bad $true $script:Guid $script:Guid $null 0 '' $null }
    Assert-That ($r.Calls -eq 0 -and $r.Failed -eq 1) "OnDone $($r.Calls), OnError $($r.Failed)"
    Assert-That ($r.Err -is [string] -and $r.Err -like '*AADSTS50173*') "the text: $($r.Err)"
}
#endregion

#region Locations, roots, folders ------------------------------------------------------------------------------------------------------------------------------------
foreach ($case in @(
        @{ Name = 'no location when the primary mailbox cannot be found either'; Found = @(); Resolve = $false; Expect = @() }
        @{ Name = 'the primary mailbox alone'; Found = @('Primary'); Resolve = $true; Expect = @('Primary mailbox') }
        @{ Name = 'the archive alone (the primary mailbox is found by its address)'; Found = @('Archive'); Resolve = $true; Expect = @('Primary mailbox', 'Archive') }
        @{ Name = 'three locations, the auxiliary one is being filled'; Found = @('Primary', 'Archive', 'AuxArchive'); Resolve = $true; Expect = @('Primary mailbox', 'Archive', 'Auxiliary archive 1') })) {
    Check "locations: an array, $($case.Name)" {
        $script:Stub.ResolveFails = -not $case.Resolve
        $script:Stub.Locations = @($case.Found | ForEach-Object { New-TestLocation $_ })
        $r = Invoke-Call { param($ok, $bad) & $engine.Locations 'adele.vance@contoso.com' '' $ok $bad '' $null }
        Assert-That (-not $r.Err) "OnError: $($r.Err)"
        Assert-That ($r.Calls -eq 1) "OnDone was called $($r.Calls) times"
        Assert-Array $r.Got 'the locations'
        Assert-Equal @($r.Got | ForEach-Object { $_.Label }) $case.Expect 'the labels'
        foreach ($l in $r.Got) {
            Assert-Reads $l 'Kind', 'Label', 'Hint', 'Root', 'MailboxId', 'Loc', 'Filling' 'a location'
            Assert-That ($l.Loc -is [MailboxTransfer.BrowseLocation]) 'Loc is the location of the engine'
        }
        Assert-That (@($r.Got | Where-Object { $_.Filling }).Count -eq $(if ('AuxArchive' -in $case.Found) { 1 } else { 0 })) 'only the auxiliary archive is being filled'
    }
}
Check 'locations: a shared mailbox is one location, found by its address' {
    $script:Stub.ResolveFails = $false
    $r = Invoke-Call { param($ok, $bad) & $engine.Locations 'adele.vance@contoso.com' '' $ok $bad 'projects@contoso.com' $null }
    Assert-Array $r.Got 'the locations'
    Assert-That ($r.Got.Count -eq 1 -and $r.Got[0].Label -eq 'Shared mailbox projects@contoso.com') "got $($r.Got.Count): $($r.Got[0].Label)"
}
Check 'locations: an engine that fails answers OnError' {
    $script:Stub.LocationsError = 'Graph said 403'
    try {
        $r = Invoke-Call { param($ok, $bad) & $engine.Locations 'adele.vance@contoso.com' '' $ok $bad '' $null }
        Assert-That ($r.Calls -eq 0 -and $r.Err -like '*403*') "OnDone $($r.Calls), OnError: $($r.Err)"
    } finally { $script:Stub.LocationsError = $null }
}
function New-TestLocationEntry([string]$Kind) { [pscustomobject]@{ Kind = $Kind; Label = $Kind; Hint = ''; Root = 'x'; MailboxId = "MBX:$Kind@contoso"; Loc = (New-TestLocation $Kind); Filling = ($Kind -eq 'AuxArchive') } }
foreach ($count in 0, 1, 3) {
    Check "roots: an array of $count root node(s)" {
        $kinds = @('Primary', 'Archive', 'AuxArchive')[0..([math]::Max(0, $count - 1))]
        $locs = @(if ($count) { $kinds | ForEach-Object { New-TestLocationEntry $_ } })
        $r = Invoke-Call { param($ok, $bad) & $engine.Roots $locs $ok $bad $false }
        Assert-That ($r.Calls -eq 1) "OnDone was called $($r.Calls) times"
        Assert-Array $r.Got 'the roots'
        Assert-That ($r.Got.Count -eq $count) "expected $count, got $($r.Got.Count)"
        foreach ($n in $r.Got) { Assert-That ($n -is [MailboxTransferGui.FolderNode] -and $n.IsRoot -and $n.Tag.Location) 'a root node with its location' }
    }
}
foreach ($count in 0, 1, 2) {
    Check "children: an array of $count folder node(s), a folder with subfolders is not read yet" {
        $script:Stub.Folders = @(for ($i = 1; $i -le $count; $i++) { New-TestFolder "f$i" "Folder $i" -Count 5 -Children 2 -Size 1000 -WellKnown $(if ($i -eq 1) { 'inbox' }) })
        $root = New-TestRoot
        $r = Invoke-Call { param($ok, $bad) & $engine.Children $root $ok $bad }
        Assert-That ($r.Calls -eq 1 -and -not $r.Err) "OnDone $($r.Calls), OnError $($r.Err)"
        Assert-Array $r.Got 'the folders (the page does SetChildren([FolderNode[]]@($kids)))'
        Assert-That ($r.Got.Count -eq $count) "expected $count, got $($r.Got.Count)"
        foreach ($n in $r.Got) { Assert-That ($n -is [MailboxTransferGui.FolderNode] -and $n.ChildrenNotRead) 'a folder node whose children are not read' }
        $root.SetChildren([MailboxTransferGui.FolderNode[]]@($r.Got))
    }
}
Check 'children: a folder that Microsoft is still moving is dimmed with its reason' {
    $script:Stub.Folders = @(New-TestFolder 'f1' 'Old Mail' -Count 9 -NotAvailable)
    $r = Invoke-Call { param($ok, $bad) & $engine.Children (New-TestRoot) $ok $bad }
    Assert-That ($r.Got.Count -eq 1 -and $r.Got[0].Dim -and $r.Got[0].Sub -like 'Not ready*') 'dim, with the reason'
}
#endregion

#region Messages and the reading pane -----------------------------------------------------------------------------------------------------------------------------------
foreach ($case in @(
        @{ Name = 'an empty folder'; Count = 0; Next = ''; Details = $true }
        @{ Name = 'a folder with one message'; Count = 1; Next = ''; Details = $true }
        @{ Name = 'a page of three and more to come'; Count = 3; Next = 'token-2'; Details = $true }
        @{ Name = 'one row of an auxiliary archive (headers only)'; Count = 1; Next = ''; Details = $false })) {
    Check "items: ONE object with the rows of $($case.Name)" {
        $script:Stub.Page = New-TestPage $case.Count $case.Next -NoDetails:(-not $case.Details)
        $folder = New-TestFolderNode $case.Count -HeadersOnly:(-not $case.Details)
        $r = Invoke-Call { param($ok, $bad) & $engine.Items $folder $null $ok $bad }
        Assert-That ($r.Calls -eq 1 -and -not $r.Err) "OnDone $($r.Calls), OnError $($r.Err)"
        Assert-One $r.Got 'the page of messages'
        Assert-Reads $r.Got 'Rows', 'Next', 'Total', 'Note' 'the page of messages'
        Assert-Array $r.Got.Rows 'Rows'
        Assert-That ($r.Got.Rows.Count -eq $case.Count -and $r.Got.Total -eq $case.Count) "rows $($r.Got.Rows.Count), total $($r.Got.Total)"
        Assert-That (($case.Next -and $r.Got.Next -eq $case.Next) -or (-not $case.Next -and -not $r.Got.Next)) 'the token of the next page'
        foreach ($row in $r.Got.Rows) { Assert-That ($row -is [MailboxTransferGui.MessageRow] -and $row.SubjectNa -eq (-not $case.Details) -and $row.Tag -is [MailboxTransfer.BrowseItem]) 'a row' }
    }
}
Check 'items: a folder the service cannot read yet says so in the note and has no row' {
    $script:Stub.Page = New-TestPage 0 '' -NotAvailable
    $folder = New-TestFolderNode 9
    $r = Invoke-Call { param($ok, $bad) & $engine.Items $folder $null $ok $bad }
    Assert-That ($r.Got.Rows.Count -eq 0 -and $r.Got.Note -like 'Not ready*') "note: $($r.Got.Note)"
}
Check 'items: a message without a date or a size is a row all the same' {
    $script:Stub.Page = New-TestPage 1 ''
    $script:Stub.Page.Items[0].Received = $null; $script:Stub.Page.Items[0].Created = $null
    $folder = New-TestFolderNode 1
    $r = Invoke-Call { param($ok, $bad) & $engine.Items $folder $null $ok $bad }
    Assert-That (-not $r.Err -and $r.Got.Rows[0].DateText -eq '') "error: $($r.Err)"
}
foreach ($case in @(@{ To = 0; Att = 0 }, @{ To = 1; Att = 1 }, @{ To = 2; Att = 3 })) {
    Check "preview: ONE object the reading pane reads, $($case.To) recipient(s) and $($case.Att) attachment(s)" {
        $script:Stub.Preview = New-TestPreview $case.To $case.Att
        $folder = New-TestFolderNode 1
        $row = [MailboxTransferGui.MessageRow]::new(); $row.Id = 'i1'; $row.Tag = New-TestItem 'i1'
        $r = Invoke-Call { param($ok, $bad) & $engine.Preview $row $folder $ok $bad }
        Assert-That ($r.Calls -eq 1 -and -not $r.Err) "OnDone $($r.Calls), OnError $($r.Err)"
        Assert-One $r.Got 'the preview'
        Assert-Reads $r.Got 'Subject', 'From', 'Address', 'To', 'Received', 'Location', 'Details', 'Attachment', 'Body', 'Available' 'the preview'
        Assert-That ($r.Got.Subject -eq 'Signed contract') 'the subject'
        Assert-That (@($r.Got.Attachment -split "`n" | Where-Object { $_ }).Count -eq $case.Att) 'the attachments'
        Assert-That (@($r.Got.To -split '; ' | Where-Object { $_ }).Count -eq $case.To) 'the recipients'
    }
}
Check 'preview: a message without a body, a date or recipients does not fail' {
    $script:Stub.Preview = New-TestPreview 0 0 -NoBody
    $script:Stub.Preview.Header.Received = $null
    $folder = New-TestFolderNode 1
    $row = [MailboxTransferGui.MessageRow]::new(); $row.Id = 'i1'; $row.Tag = New-TestItem 'i1'
    $r = Invoke-Call { param($ok, $bad) & $engine.Preview $row $folder $ok $bad }
    Assert-That (-not $r.Err -and $r.Got.Received -eq '' -and $r.Got.Body -eq '') "error: $($r.Err)"
}
#endregion

#region The selection file --------------------------------------------------------------------------------------------------------------------------------------------
function New-TestTick {
    # one ticked folder of a location, as Get-MbtGuiSelection gives it: an entry with its node
    param([string]$Kind, [string]$Id, [switch]$Picks)
    $root = New-TestRoot
    $folder = [MailboxTransferGui.FolderNode]::new($Id, "Folder $Id", 5, 1000); $folder.Tag = New-TestFolder $Id "Folder $Id" -Count 5 -Size 1000
    $root.SetChildren([MailboxTransferGui.FolderNode[]]@($folder))
    if ($Picks) { $folder.PickRows['i1'] = (New-TestItem 'i1') }
    [pscustomobject]@{ Root = $root; Entry = @{ Node = $folder; Kind = $Kind; Items = 1 } }
}
foreach ($case in @(@{ Name = 'one folder ticked alone'; Ticks = @('tree') }, @{ Name = 'one message ticked alone'; Ticks = @('picks') }, @{ Name = "a folder's own messages"; Ticks = @('own') }, @{ Name = 'a folder and a message picked'; Ticks = @('tree', 'picks') })) {
    Check "selection file: ONE file for $($case.Name), the path comes in an array of one" {
        $path = Join-Path $Work ('sel-' + [guid]::NewGuid().ToString('N') + '.json')
        $ticks = @($case.Ticks | ForEach-Object { New-TestTick $_ "f-$_" -Picks:($_ -eq 'picks') })
        $groups = @($ticks | ForEach-Object { @{ Root = $_.Root; Entries = @($_.Entry) } })
        $r = Invoke-Call { param($ok, $bad) & $engine.SaveSelections $groups $path $ok $bad }
        Assert-That ($r.Calls -eq 1 -and -not $r.Err) "OnDone $($r.Calls), OnError $($r.Err)"
        Assert-That ([string]$r.Got -eq $path) "the path: $($r.Got)"
        Assert-That (Test-Path -LiteralPath $path) 'the file was written'
    }
}
foreach ($count in 1, 2) {
    Check "selection file: the page gets the path in an array of one for $count entr$(if ($count -eq 1) { 'y' } else { 'ies' })" {
        $path = Join-Path $Work ('sel-' + [guid]::NewGuid().ToString('N') + '.json')
        $selection = [pscustomobject]@{ Entries = @(1..$count | ForEach-Object { (New-TestTick 'tree' "g$_").Entry }) }
        $r = Invoke-Call { param($ok, $bad) Save-MbtGuiSelectionFile -Selection $selection -Path $path -OnDone $ok -OnError $bad }
        Assert-That ($r.Calls -eq 1 -and -not $r.Err) "OnDone $($r.Calls), OnError $($r.Err)"
        Assert-Array $r.Got 'the files'
        Assert-That ($r.Got.Count -eq 1 -and $r.Got[0] -eq $path -and (Test-Path -LiteralPath $path)) 'one file, the one asked for'
    }
}
Check 'selection file: nothing ticked gives an empty array and writes no file' {
    $path = Join-Path $Work 'none.json'
    $r = Invoke-Call { param($ok, $bad) Save-MbtGuiSelectionFile -Selection ([pscustomobject]@{ Entries = @() }) -Path $path -OnDone $ok -OnError $bad }
    Assert-Array $r.Got 'the files'
    Assert-That ($r.Calls -eq 1 -and $r.Got.Count -eq 0 -and -not (Test-Path -LiteralPath $path)) 'nothing written'
}
#endregion

#region The list of jobs, the lights of this computer, the settings ---------------------------------------------------------------------------------------------------------
function New-TestJobFolder([string]$WorkFolder, [string]$Name, [string]$Json) {
    $d = Join-Path (Join-Path $WorkFolder 'jobs') $Name
    $null = New-Item -ItemType Directory -Force -Path $d
    Set-Content -LiteralPath (Join-Path $d 'status.json') -Value $Json -Encoding utf8
}
foreach ($case in @(
        @{ Count = 0; Json = @() }
        @{ Count = 1; Json = @('{"v":1,"jobId":"a1","task":"export","title":"Export x","state":"complete","updated":"2026-11-10T10:00:00.000Z"}') }
        @{ Count = 4; Json = @(
                '{"v":1,"jobId":"a1","task":"export","title":"Export x","state":"complete","updated":"2026-11-10T10:00:00.000Z"}'
                '{"v":1,"jobId":"a2","task":"import","state":"error","updated":"2026-11-10T11:00:00.000Z"}'
                '{"v":1,"jobId":"a3","task":"import","title":"Import y","state":"needs-decision","updated":"2026-11-10T12:00:00.000Z"}'
                '{"v":1,"jobId":"a4","task":"import","title":"Import z","state":"stopped","updated":"2026-11-10T13:00:00.000Z","counts":{"done":12}}') })) {
    Check "jobs: an array of $($case.Count) job(s), none of them empty, with what the page reads" {
        $w = Join-Path $Work ('jobs-' + [guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Force -Path $w
        $n = 0; foreach ($j in $case.Json) { $n++; New-TestJobFolder $w "20261110-100$n-job" $j }
        $jobs = @(& $engine.Jobs $w)
        Assert-That ($jobs.Count -eq $case.Count) "expected $($case.Count), got $($jobs.Count)"
        foreach ($j in $jobs) {
            Assert-That ($null -ne $j) 'a job that is nothing'
            Assert-Reads $j 'Task', 'Title', 'Sub', 'State', 'Button', 'Folder' 'a job'
            Assert-That ([bool]$j.Sub) 'a line under the title'
        }
    }
}
Check 'jobs: a batch is one line with its numbers and its report; the jobs of a batch that is complete are not listed, those of a batch that is not complete carry its name, three at most' {
    $w = Join-Path $Work ('jobs-' + [guid]::NewGuid().ToString('N'))
    foreach ($n in 1..2) { New-TestJobFolder $w "20261110-100$n-job" ('{"v":1,"jobId":"d' + $n + '","task":"export","title":"Export done' + $n + '","batch":"Done","state":"complete","updated":"2026-11-10T10:0' + $n + ':00.000Z"}') }
    foreach ($n in 3..7) { New-TestJobFolder $w "20261110-100$n-job" ('{"v":1,"jobId":"o' + $n + '","task":"export","title":"Export open' + $n + '","batch":"Open","state":"stopped","updated":"2026-11-10T10:0' + $n + ':00.000Z","counts":{"done":4}}') }
    New-TestJobFolder $w '20261110-1008-job' '{"v":1,"jobId":"plain","task":"import","title":"Import plain","state":"stopped","updated":"2026-11-10T10:08:00.000Z"}'
    $rowOf = { param($Id, $State, $Job) [ordered]@{ n = 1; upn = "$Id@contoso.com"; location = 'Primary'; selection = ''; output = ''; store = ''; state = $State; job = $Job; attempts = 1 } }
    $batchOf = { param($Id, $Rows) [ordered]@{ v = 1; id = $Id; name = $Id; created = '2026-11-10T09:00:00.000Z'; updated = '2026-11-10T09:00:00.000Z'; state = ''; mode = 'Export'; csv = ''; parallel = 1; options = @{}; counts = $null; rows = $Rows } }
    Save-MbtBatch -Batch (& $batchOf 'Done' @((& $rowOf 'a' 'complete' '20261110-1001-job'), (& $rowOf 'b' 'complete' '20261110-1002-job'))) -Folder (Join-Path $w 'batches\Done')
    Save-MbtBatch -Batch (& $batchOf 'Open' @((& $rowOf 'c' 'partial' '20261110-1003-job'), (& $rowOf 'd' 'failed' '20261110-1004-job'), (& $rowOf 'e' 'pending' ''))) -Folder (Join-Path $w 'batches\Open')
    $jobs = @(& $engine.Jobs $w)
    $lines = @($jobs | Where-Object Task -EQ 'batch')
    Assert-That ($lines.Count -eq 2) "two batch lines, got $($lines.Count)"
    foreach ($l in $lines) { Assert-Reads $l 'Task', 'Title', 'Sub', 'State', 'Button', 'Folder', 'Report', 'Batch' 'a batch line' }
    $open = $lines | Where-Object Batch -EQ 'Open'
    Assert-That ($open.Sub -like '0 of 3 mailboxes complete, 1 not complete, 1 failed, 1 not started*command line*') "the numbers: $($open.Sub)"
    Assert-That ($open.State -eq 'warn' -and $open.Button -eq 'Open the report') 'a batch that is not complete is a warning with a button that opens the report'
    $done = $lines | Where-Object Batch -EQ 'Done'
    Assert-That ($done.State -eq 'ok' -and $done.Sub -like '2 of 2 mailboxes complete*' -and $done.Sub -notlike '*command line*') "a complete batch: $($done.Sub)"
    $others = @($jobs | Where-Object Task -NE 'batch')
    Assert-That (-not @($others | Where-Object { $_.Title -like 'Export done*' }).Count) 'the jobs of a complete batch are left out'
    Assert-That (@($others | Where-Object { $_.Sub -like 'batch Open - *' }).Count -eq 3) 'the jobs of a batch that is not complete: three, with the name of the batch'
    Assert-That (@($others | Where-Object Title -EQ 'Import plain').Count -eq 1) 'the other jobs are listed'
    Assert-That ($jobs.Count -le 8) 'eight lines at most'
}
Check 'jobs: a job that stopped with an empty count says when it stopped' {
    $w = Join-Path $Work ('jobs-' + [guid]::NewGuid().ToString('N'))
    New-TestJobFolder $w '20261110-1000-job' '{"v":1,"jobId":"a1","task":"import","title":"t","state":"stopped","updated":"2026-11-10T10:00:00.000Z","counts":{}}'
    $jobs = @(& $engine.Jobs $w)
    Assert-That ($jobs.Count -eq 1 -and $jobs[0].Sub -like 'stopped *' -and $jobs[0].Button -eq 'Continue') "got: $($jobs[0].Sub)"
}
Check 'readiness: the lights the Setup page reads' {
    $r = & $engine.Readiness $Work
    Assert-That (@($r.Lights).Count -gt 3) 'the lights'
    foreach ($l in $r.Lights) { Assert-Reads ([pscustomobject]$l) 'Name', 'Detail', 'Ok' 'a light' }
}
Check 'settings: what the Setup page reads is there' {
    Assert-Reads (& $engine.SignInSettings) 'TenantId', 'ClientId', 'AppTenantId', 'AppClientId', 'CertificateThumbprint' 'the settings'
    Assert-That ([bool](& $engine.ConfigPath)) 'the configuration file'
    Assert-That ((& $engine.FreeSpace $Work) -gt 0) 'the free space'
    Assert-Reads (& $engine.WindowSignIn) 'Available', 'Reason' 'the sign-in window'
}
#endregion

#region A PST, an export store, a folder of .eml files, opened and read -----------------------------------------------------------------------------------------------------
$P = [MailboxTransfer.FxProperty]
function New-TestMessage([string]$Subject) {
    $m = [MailboxTransfer.FxMessage]::new()
    $m.Properties.Add($P::FromString(0x001A, 'IPM.Note'))
    $m.Properties.Add($P::FromString(0x0037, $Subject))
    $m.Properties.Add($P::FromString(0x1000, 'plain body of the message'))
    $m.Properties.Add($P::FromString(0x0C1A, 'Alice Martin'))
    $m.Properties.Add($P::FromString(0x5D01, 'alice@contoso.example'))
    $m.Properties.Add($P::FromTime(0x0039, [datetime]::new(2026, 10, 21, 9, 30, 0, [DateTimeKind]::Utc)))
    $m.Properties.Add($P::FromInt32(0x0E07, 1))
    return $m
}
function New-TestPst([string]$Name, [string[]]$Folders) {
    # Folders: "name:messages"
    $path = Join-Path $Work "$Name.pst"
    $w = [MailboxTransfer.PstWriter]::Create($path, 'Test archive')
    foreach ($spec in $Folders) {
        $n, $c = $spec -split ':'
        $f = $w.CreateFolder($null, $n, 'IPF.Note')
        for ($i = 1; $i -le [int]$c; $i++) { [void]$w.AddMessage($f, (New-TestMessage "$n message $i")) }
    }
    $w.Close()
    return $path
}
function Open-Test([string]$Kind, [string]$Path) { Invoke-Call { param($ok, $bad) & $engine.OpenSource $Kind $Path $ok $bad } }

foreach ($parts in 1, 2, 3) {
    Check "pst: the parts of a file are found whatever their number ($parts part(s))" {
        $dir = Join-Path $Work ('parts-' + [guid]::NewGuid().ToString('N')); $null = New-Item -ItemType Directory -Path $dir
        Set-Content -LiteralPath (Join-Path $dir 'archive.pst') -Value 'x'
        for ($i = 2; $i -le $parts; $i++) { Set-Content -LiteralPath (Join-Path $dir ('archive-{0:000}.pst' -f $i)) -Value 'x' }
        Set-Content -LiteralPath (Join-Path $dir 'other.pst') -Value 'x'
        $found = @(Find-MbtGuiPstParts (Join-Path $dir 'archive.pst'))
        Assert-That ($found.Count -eq $parts) "expected $parts, found $($found.Count)"
    }
}
foreach ($case in @(
        @{ Name = 'one folder with one message (one part)'; Folders = @('Inbox:1'); Messages = 1; Top = 2 }
        @{ Name = 'one empty folder'; Folders = @('Inbox:0'); Messages = 0; Top = 2 }
        @{ Name = 'three folders'; Folders = @('Inbox:2', 'Sent Items:1', 'Archive:3'); Messages = 6; Top = 4 })) {
    Check "pst: opens $($case.Name), reads its first folder and a message" {
        $path = New-TestPst ('open-' + [guid]::NewGuid().ToString('N').Substring(0, 8)) $case.Folders
        $r = Open-Test 'pst' $path
        Assert-That (-not $r.Err -and $r.Calls -eq 1) "OnDone $($r.Calls), OnError: $($r.Err)"
        Assert-One $r.Got 'the answer of the check'
        Assert-Reads $r.Got 'Ok', 'Level', 'Message', 'Items', 'Bytes', 'Folders', 'Parts', 'Roots' 'the answer of the check'
        Assert-That ($r.Got.Ok -and $r.Got.Parts -eq 1 -and $r.Got.Items -eq $case.Messages) "ok $($r.Got.Ok), parts $($r.Got.Parts), items $($r.Got.Items)"
        Assert-Array $r.Got.Roots 'Roots'
        Assert-That ($r.Got.Roots.Count -eq 1) 'one root'
        $kids = @($r.Got.Roots[0].Children | Where-Object { $_.Id -ne "`0placeholder" })
        Assert-That ($kids.Count -eq $case.Top) "folders: $($kids.Count) (the ones given and Deleted Items)"
        $node = @($kids | Where-Object { $_.Name -eq (($case.Folders[0] -split ':')[0]) })[0]
        $items = Invoke-Call { param($ok, $bad) & $engine.SourceItems $r.Got.Roots[0] $node $null $ok $bad }
        Assert-That (-not $items.Err) "items: $($items.Err)"
        Assert-One $items.Got 'the page of messages'
        Assert-Reads $items.Got 'Rows', 'Next', 'Total', 'Note' 'the page of messages'
        Assert-That ($items.Got.Rows.Count -eq $node.Count) "rows $($items.Got.Rows.Count), folder $($node.Count)"
        if ($items.Got.Rows.Count) {
            $prev = Invoke-Call { param($ok, $bad) & $engine.SourcePreview $r.Got.Roots[0] $node $items.Got.Rows[0] $ok $bad }
            Assert-That (-not $prev.Err) "preview: $($prev.Err)"
            Assert-One $prev.Got 'the preview'
            Assert-Reads $prev.Got 'Subject', 'From', 'Address', 'To', 'Received', 'Location', 'Details', 'Attachment', 'Body', 'Available' 'the preview'
        }
    }
}
Check 'pst: a PST whose file name is not *.pst opens too (no part is found by its name)' {
    $pst = New-TestPst ('odd-' + [guid]::NewGuid().ToString('N').Substring(0, 8)) @('Inbox:1')
    $odd = [IO.Path]::ChangeExtension($pst, '.bak'); Move-Item -LiteralPath $pst -Destination $odd
    $r = Open-Test 'pst' $odd
    Assert-That (-not $r.Err -and $r.Calls -eq 1 -and $r.Got.Ok) "OnDone $($r.Calls), OnError: $($r.Err)"
}
Check 'pst: a file that is not there is said not to be there' {
    $r = Open-Test 'pst' (Join-Path $Work 'missing.pst')
    Assert-That ($r.Calls -eq 0 -and $r.Err -like '*not found*') "OnDone $($r.Calls), OnError: $($r.Err)"
}
foreach ($case in @(
        @{ Name = 'one mailbox with one folder'; Mailboxes = @(@{ Id = 'mb1'; Kind = 'Primary'; Folders = @('inbox') }) }
        @{ Name = 'one mailbox without a folder'; Mailboxes = @(@{ Id = 'mb1'; Kind = 'Primary'; Folders = @() }) }
        @{ Name = 'two mailboxes and a subfolder'; Mailboxes = @(@{ Id = 'mb1'; Kind = 'Primary'; Folders = @('inbox', 'inbox/sub') }, @{ Id = 'mb2'; Kind = 'Archive'; Folders = @('old') }) })) {
    Check "store: opens an export store of $($case.Name)" {
        $dir = Join-Path $Work ('store-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        $store = [MailboxTransfer.ExportStore]::Create($dir, 'test', 'tests')
        try {
            foreach ($m in $case.Mailboxes) {
                $store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = $m.Id; Kind = $m.Kind; Upn = "$($m.Id)@contoso.test"; DisplayName = $m.Id }) | Out-Null
                foreach ($f in $m.Folders) {
                    $parent = if ($f.Contains('/')) { $f.Substring(0, $f.LastIndexOf('/')) } else { $null }
                    $store.UpsertFolders(@([MailboxTransfer.StoreFolder]@{ MailboxId = $m.Id; FolderId = $f; ParentFolderId = $parent; DisplayName = ($f -split '/')[-1]; SourceItemCount = 3; SourceSize = 3000 })) | Out-Null
                }
            }
        } finally { $store.Dispose() }
        $r = Open-Test 'store' $dir
        Assert-That (-not $r.Err -and $r.Calls -eq 1) "OnDone $($r.Calls), OnError: $($r.Err)"
        Assert-One $r.Got 'the answer of the check'
        Assert-Reads $r.Got 'Ok', 'Level', 'Message', 'Items', 'Bytes', 'Folders', 'Parts', 'Roots' 'the answer of the check'
        Assert-Array $r.Got.Roots 'Roots'
        $folders = 0; foreach ($m in $case.Mailboxes) { $folders += $m.Folders.Count }
        Assert-That ($r.Got.Roots.Count -eq $case.Mailboxes.Count -and $r.Got.Folders -eq $folders -and $r.Got.Items -eq (3 * $folders)) "roots $($r.Got.Roots.Count), folders $($r.Got.Folders), items $($r.Got.Items)"
    }
}
Check 'store: a store of 5,000 folders, some under a parent that is not there, opens in seconds with the right tree' {
    $dir = Join-Path $Work ('store-big-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    $store = [MailboxTransfer.ExportStore]::Create($dir, 'test', 'tests')
    try {
        $store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = 'mb1'; Kind = 'Primary'; Upn = 'mb1@contoso.test'; DisplayName = 'mb1' }) | Out-Null
        # folder 0 is a top folder; folder i is below folder (i - 1) / 4 (six levels), except every 500th, whose parent is not in the store
        $rows = foreach ($i in 0..4999) {
            $parent = if ($i -eq 0) { $null } elseif ($i % 500 -eq 0) { "gone-$i" } else { 'f' + [int][math]::Floor(($i - 1) / 4) }
            [MailboxTransfer.StoreFolder]@{ MailboxId = 'mb1'; FolderId = "f$i"; ParentFolderId = $parent; DisplayName = "Folder $i"; SourceItemCount = 2; SourceSize = 2000 }
        }
        $store.UpsertFolders(@($rows)) | Out-Null
    } finally { $store.Dispose() }
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $r = Open-Test 'store' $dir
    $watch.Stop()
    Assert-That (-not $r.Err -and $r.Calls -eq 1) "OnDone $($r.Calls), OnError: $($r.Err)"
    Assert-That ($watch.Elapsed.TotalSeconds -lt 15) "the store took $([math]::Round($watch.Elapsed.TotalSeconds, 1)) s to open"
    $root = $r.Got.Roots[0]
    $nodes = @($root.Descendants() | Where-Object { $_ -ne $root })
    $byId = @{}; foreach ($n in $nodes) { $byId[$n.Id] = $n }
    Assert-That ($r.Got.Roots.Count -eq 1 -and $r.Got.Folders -eq 5000 -and $r.Got.Items -eq 10000) "roots $($r.Got.Roots.Count), folders $($r.Got.Folders), items $($r.Got.Items)"
    Assert-That ($root.Children.Count -eq 10) "top folders: $($root.Children.Count) (folder 0 and the 9 whose parent is not there)"
    Assert-That ($nodes.Count -eq 5000 -and $byId.Count -eq 5000) "nodes in the tree: $($nodes.Count)"
    Assert-That ($byId['f500'].Parent -eq $root -and $byId['f0'].Parent -eq $root) 'a folder whose parent is not there is a top folder'
    Assert-That ($byId['f4999'].Parent.Id -eq 'f1249' -and $byId['f1249'].Parent.Id -eq 'f312' -and $byId['f312'].Parent.Id -eq 'f77') "f4999 is below $($byId['f4999'].Parent.Id)"
}
foreach ($case in @(@{ Name = 'one file at the top'; Files = @('a.eml') }, @{ Name = 'one file in one subfolder'; Files = @('Inbox/a.eml') }, @{ Name = 'two files in two subfolders'; Files = @('Inbox/a.eml', 'Sent/b.eml') })) {
    Check "eml: opens a folder of .eml files, $($case.Name)" {
        $dir = Join-Path $Work ('eml-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        foreach ($f in $case.Files) { $p = Join-Path $dir $f; $null = New-Item -ItemType Directory -Force -Path (Split-Path $p -Parent); Set-Content -LiteralPath $p -Value 'Subject: x' }
        $r = Open-Test 'eml' $dir
        Assert-That (-not $r.Err -and $r.Calls -eq 1) "OnDone $($r.Calls), OnError: $($r.Err)"
        Assert-One $r.Got 'the answer of the check'
        Assert-Reads $r.Got 'Ok', 'Level', 'Message', 'Items', 'Bytes', 'Folders', 'Parts', 'Roots' 'the answer of the check'
        Assert-That ($r.Got.Items -eq $case.Files.Count -and $r.Got.Roots.Count -eq 1) "items $($r.Got.Items)"
    }
}
Check 'eml: a folder without any .eml file is refused in plain words' {
    $dir = Join-Path $Work 'eml-empty'; $null = New-Item -ItemType Directory -Force -Path $dir
    $r = Open-Test 'eml' $dir
    Assert-That ($r.Calls -eq 0 -and $r.Err -like '*no .eml file*') "OnDone $($r.Calls), OnError: $($r.Err)"
}
#endregion

foreach ($count in 0, 1, 2) {
    Check "store: the messages of a folder of $count message(s) are one page, and a message is read" {
        $dir = Join-Path $Work ('storeitems-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        $store = [MailboxTransfer.ExportStore]::Create($dir, 'test', 'tests')
        try {
            $store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = 'mb1'; Kind = 'Primary'; Upn = 'mb1@contoso.test'; DisplayName = 'mb1' }) | Out-Null
            $store.UpsertFolders(@([MailboxTransfer.StoreFolder]@{ MailboxId = 'mb1'; FolderId = 'inbox'; DisplayName = 'Inbox'; SourceItemCount = $count; SourceSize = 3000 })) | Out-Null
            $entries = [Collections.Generic.List[MailboxTransfer.StoreItemEntry]]::new()
            for ($i = 1; $i -le $count; $i++) { $entries.Add([MailboxTransfer.StoreItemEntry]@{ MailboxId = 'mb1'; ItemId = "item$i"; FolderId = 'inbox'; ChangeKey = "c$i"; ItemType = 'IPM.Note'; Size = 1000 * $i }) }
            if ($count) { $store.UpsertItems($entries) | Out-Null }
        } finally { $store.Dispose() }
        $opened = Open-Test 'store' $dir
        Assert-That (-not $opened.Err) "open: $($opened.Err)"
        $root = $opened.Got.Roots[0]
        $node = @($root.Children | Where-Object { $_.Id -eq 'inbox' })[0]
        $items = Invoke-Call { param($ok, $bad) & $engine.SourceItems $root $node $null $ok $bad }
        Assert-That (-not $items.Err) "items: $($items.Err)"
        Assert-One $items.Got 'the page of messages'
        Assert-Reads $items.Got 'Rows', 'Next', 'Total', 'Note' 'the page of messages'
        Assert-Array $items.Got.Rows 'Rows'
        Assert-That ($items.Got.Rows.Count -eq $count -and $items.Got.Total -eq $count) "rows $($items.Got.Rows.Count), total $($items.Got.Total)"
        if ($count) {
            $ro = [MailboxTransfer.ExportStore]::OpenReadOnly($dir)
            try { $ro.SetPreview([long]$items.Got.Rows[0].Tag.Key, [MailboxTransfer.StorePreview]@{ ItemKey = [long]$items.Got.Rows[0].Tag.Key; Subject = 'A subject'; From = 'Alice'; To = 'Bob'; MessageClass = 'IPM.Note'; AttachmentNames = ''; BodyPreview = 'Hello' }) } catch { } finally { $ro.Dispose() }
        }
    }
}
Check 'store: the preview of a message the store holds is one object the pane reads' {
    $dir = Join-Path $Work ('storeprev-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    $store = [MailboxTransfer.ExportStore]::Create($dir, 'test', 'tests')
    try {
        $store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = 'mb1'; Kind = 'Primary'; Upn = 'mb1@contoso.test'; DisplayName = 'mb1' }) | Out-Null
        $store.UpsertFolders(@([MailboxTransfer.StoreFolder]@{ MailboxId = 'mb1'; FolderId = 'inbox'; DisplayName = 'Inbox'; SourceItemCount = 1; SourceSize = 1000 })) | Out-Null
        $store.UpsertItems(@([MailboxTransfer.StoreItemEntry]@{ MailboxId = 'mb1'; ItemId = 'item1'; FolderId = 'inbox'; ChangeKey = 'c1'; ItemType = 'IPM.Note'; Size = 1000 })) | Out-Null
        $key = [long]($store.GetItems([long]($store.GetFolders('mb1')[0].Key), 0, 50, [MailboxTransfer.StoreItemSort]::ReceivedDate, $true).Items[0].Key)
        $store.SetPreview($key, [MailboxTransfer.StorePreview]@{ ItemKey = $key; Subject = 'A subject'; From = 'Alice'; To = 'Bob'; MessageClass = 'IPM.Note'; AttachmentNames = ''; BodyPreview = 'Hello' })
    } finally { $store.Dispose() }
    $opened = Open-Test 'store' $dir
    $root = $opened.Got.Roots[0]
    $node = @($root.Children | Where-Object { $_.Id -eq 'inbox' })[0]
    $items = Invoke-Call { param($ok, $bad) & $engine.SourceItems $root $node $null $ok $bad }
    Assert-That (-not $items.Err -and $items.Got.Rows.Count -eq 1) "items: $($items.Err)"
    $prev = Invoke-Call { param($ok, $bad) & $engine.SourcePreview $root $node $items.Got.Rows[0] $ok $bad }
    Assert-That (-not $prev.Err) "preview: $($prev.Err)"
    Assert-One $prev.Got 'the preview'
    Assert-Reads $prev.Got 'Subject', 'From', 'Address', 'To', 'Received', 'Location', 'Details', 'Attachment', 'Body', 'Available' 'the preview'
    Assert-That ($prev.Got.Subject -eq 'A subject' -and $prev.Got.Body -eq 'Hello') "subject: $($prev.Got.Subject)"
}
Check 'eml: a folder of .eml files has no message to list before the import, said in a note' {
    $dir = Join-Path $Work ('emlitems-' + [guid]::NewGuid().ToString('N').Substring(0, 8)); $null = New-Item -ItemType Directory -Force -Path (Join-Path $dir 'Inbox'); Set-Content -LiteralPath (Join-Path $dir 'Inbox\a.eml') -Value 'Subject: x'
    $opened = Open-Test 'eml' $dir
    $root = $opened.Got.Roots[0]
    $node = @($root.Children | Where-Object { $_.Name -eq 'Inbox' })[0]
    $items = Invoke-Call { param($ok, $bad) & $engine.SourceItems $root $node $null $ok $bad }
    Assert-One $items.Got 'the page of messages'
    Assert-Reads $items.Got 'Rows', 'Next', 'Total', 'Note' 'the page of messages'
    Assert-That ($items.Got.Rows.Count -eq 0 -and $items.Got.Note) 'a note and no row'
}
Check 'static members: every [Type]::Member of the engine and of the view models that the console writes exists (an enum member that is not there is an error under StrictMode)' {
    $missing = [Collections.Generic.List[string]]::new()
    $files = @(Get-ChildItem -LiteralPath (Join-Path $Package 'module') -Filter 'MailboxTransfer.Gui*.ps1') + (Get-Item -LiteralPath (Join-Path $Package 'Start-MailboxTransferConsole.ps1'))
    foreach ($file in $files) {
        $tokens = $null; $errors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
        foreach ($m in $ast.FindAll({ param($n) $n -is [Management.Automation.Language.MemberExpressionAst] -and $n.Static -and $n.Expression -is [Management.Automation.Language.TypeExpressionAst] -and $n.Member -is [Management.Automation.Language.StringConstantExpressionAst] }, $true)) {
            $typeName = $m.Expression.TypeName.FullName
            if ($typeName -notmatch '^MailboxTransfer(Gui)?\.' -or $m.Member.Value -eq 'new') { continue }
            $type = $typeName -as [type]
            if (-not $type) { $missing.Add("$($file.Name):$($m.Extent.StartLineNumber) the type $typeName"); continue }
            if (-not $type.GetMember($m.Member.Value, [Reflection.BindingFlags]'Public,Static,FlattenHierarchy,IgnoreCase')) { $missing.Add("$($file.Name):$($m.Extent.StartLineNumber) $typeName::$($m.Member.Value)") }
        }
    }
    Assert-That ($missing.Count -eq 0) ($missing -join '; ')
}

#region The check of a PST: the result the page reads, whatever the number of findings -----------------------------------------------------------------------------------
function Invoke-TestPstCheck {
    param([string[]]$Lines)
    $g = $script:Gui
    $jobsDir = Join-Path $Work 'jobs'
    $before = @(Get-ChildItem -LiteralPath $jobsDir -Directory -ErrorAction SilentlyContinue | ForEach-Object Name)
    $g.Pollers.Clear()
    $script:Calls = 0; $script:Got = $null; $script:Err = $null
    Start-MbtGuiPstCheckJob 'D:\PST\a.pst' { param($r) $script:Calls++; $script:Got = $r } { param($m) $script:Err = $m }
    $job = Get-ChildItem -LiteralPath $jobsDir -Directory | Where-Object { $_.Name -notin $before } | Select-Object -First 1
    Set-Content -LiteralPath (Join-Path $job.FullName 'events.jsonl') -Value $Lines -Encoding utf8
    for ($pass = 0; $pass -lt 5 -and $g.Pollers.Count; $pass++) { for ($i = $g.Pollers.Count - 1; $i -ge 0; $i--) { if (& $g.Pollers[$i]) { $g.Pollers.RemoveAt($i) } } }
    [pscustomobject]@{ Got = $script:Got; Calls = $script:Calls; Err = $script:Err; Left = $g.Pollers.Count }
}
$script:Gui.S.RpPath = 'D:\PST\a.pst'
function New-CheckLines([int]$Findings, [int]$Exit) {
    $l = @()
    for ($i = 1; $i -le $Findings; $i++) { $l += ('{{"e":"item","step":"check","key":"MSG-NOT-LISTED-{0}","state":"{1}","msg":"finding {0}","total":{0}}}' -f $i, $(if ($i % 2) { 'error' } else { 'warning' })) }
    $l += ('{{"e":"result","exit":{0},"counts":{{"fileBytes":4096,"errors":{1},"warnings":0}}}}' -f $Exit, $Findings)
    return $l
}
foreach ($case in @(@{ Count = 0; Exit = 0 }, @{ Count = 1; Exit = 1 }, @{ Count = 2; Exit = 1 }, @{ Count = 5; Exit = 1 })) {
    Check "pst check: ONE object with an array of $($case.Count) finding(s)" {
        $r = Invoke-TestPstCheck (New-CheckLines $case.Count $case.Exit)
        Assert-That (-not $r.Err -and $r.Calls -eq 1) "OnDone $($r.Calls), OnError: $($r.Err)"
        Assert-That ($r.Left -eq 0) 'the poller ends with the answer'
        Assert-One $r.Got 'the check'
        Assert-Reads $r.Got 'Path', 'SizeBytes', 'Elapsed', 'Errors', 'Warnings', 'ExitCode', 'Recommended', 'Findings' 'the check'
        Assert-Array $r.Got.Findings 'Findings'
        Assert-That ($r.Got.Findings.Count -eq $case.Count -and $r.Got.SizeBytes -eq 4096) "findings $($r.Got.Findings.Count)"
        foreach ($f in $r.Got.Findings) { Assert-Reads $f 'Severity', 'Text', 'Code', 'Count', 'Level' 'a finding' }
    }
}
Check 'pst check: a check that fails answers OnError, once, and the poller ends' {
    $r = Invoke-TestPstCheck @('{"e":"error","code":"x","msg":"The file cannot be opened.","tech":"IOException"}', '{"e":"result","exit":1}')
    Assert-That ($r.Calls -eq 0 -and $r.Err -like '*cannot be opened*' -and $r.Left -eq 0) "OnDone $($r.Calls), OnError: $($r.Err), pollers left $($r.Left)"
}
Check 'pst check: a page that raises when it shows the result is told the check failed, not left waiting' {
    $g = $script:Gui
    $g.Pollers.Clear()
    $script:Told = $null
    Start-MbtGuiPstCheckJob 'D:\PST\a.pst' { param($r) throw 'the page broke' } { param($m) $script:Told = $m }
    $job = Get-ChildItem -LiteralPath (Join-Path $Work 'jobs') -Directory | Sort-Object LastWriteTime | Select-Object -Last 1
    Set-Content -LiteralPath (Join-Path $job.FullName 'events.jsonl') -Value (New-CheckLines 1 1) -Encoding utf8
    for ($pass = 0; $pass -lt 5 -and $g.Pollers.Count; $pass++) { for ($i = $g.Pollers.Count - 1; $i -ge 0; $i--) { if (& $g.Pollers[$i]) { $g.Pollers.RemoveAt($i) } } }
    Assert-That ($script:Told -like '*the page broke*' -and $g.Pollers.Count -eq 0) "told: $($script:Told), pollers left $($g.Pollers.Count)"
}
Check 'tick: a poller that raises is written down and dropped, one that is not finished stays, one that is finished goes' {
    $g = $script:Gui
    $g.Pollers.Clear()
    $g.Pollers.Add({ throw 'a poller broke' })
    $g.Pollers.Add({ $false })
    $g.Pollers.Add({ $true })
    Step-MbtGuiTick
    Assert-That ($g.Pollers.Count -eq 1) "pollers left: $($g.Pollers.Count), the tick raised: $($g.LastTickError)"
    $log = Get-ChildItem -LiteralPath (Join-Path $Work 'logs') -Filter 'MailboxTransfer_*.log' | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }
    Assert-That ("$log" -like '*a poller broke*') 'the failure is in the log of the work folder'
    $g.Pollers.Clear()
}
#endregion

$engine.Worker.Runspace.Dispose()
[pscustomobject]@{ Errors = @($script:Errors); Checks = @($script:Checks) } | ConvertTo-Json -Depth 4 -Compress
