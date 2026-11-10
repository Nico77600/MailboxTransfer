#
#  Mailbox Transfer - module manifest
#  --------------------------------------------------------------------------
#  Author  : Nicolas Fabert
#  Version : see ModuleVersion
#
#  Loaded by Invoke-MailboxTransfer.ps1 and the lab scripts (Import-Module by path).
#
@{
    RootModule        = 'MailboxTransfer.psm1'
    ModuleVersion     = '1.1.0'
    GUID              = 'e3b6a1f4-5d0c-4a8e-9b2f-7c41d8a05e63'
    Author            = 'Nicolas Fabert'
    Description       = 'Mailbox Transfer: exports Exchange Online mailboxes (primary and archive) to local storage and imports them into another mailbox with the Microsoft Graph mailbox import and export API. Phase 0: feasibility (folder tree, exportItems, FastTransfer reader, importItem).'
    PowerShellVersion = '7.4'

    # Functions called by the entry script, the lab scripts and the tests.
    FunctionsToExport = @(
        'Write-MbtStep', 'Write-MbtItem', 'Write-MbtDetail', 'Write-MbtBanner', 'Write-MbtSummary', 'Write-MbtLog', 'Start-MbtLog', 'Stop-MbtLog', 'Get-MbtLogPath', 'Format-MbtBytes', 'Format-MbtNumber', 'Format-MbtDuration', 'Format-MbtRate'
        'Import-MbtConfiguration', 'Initialize-MbtEngine', 'Initialize-MbtAuthEngine', 'Find-MbtMsal'
        'Connect-MbtGraph', 'Get-MbtAccessToken', 'Disconnect-MbtGraph', 'Get-MbtSession', 'Set-MbtSession', 'Get-MbtTokenInfo', 'Get-MbtCertificate', 'Get-MbtAppPermission', 'Get-MbtMissingAppPermission', 'Test-MbtAppOnly', 'Get-MbtSessionUser', 'Get-MbtExchangeSettingsPath'
        'Test-MbtMailboxId', 'ConvertTo-MbtMailboxId', 'Get-MbtPrimaryMailboxId', 'Test-MbtMailboxAddress', 'Resolve-MbtMailboxAddress', 'Get-MbtAccessHint', 'Get-MbtMailboxLocation', 'Read-MbtMailboxLocationFile', 'Get-MbtGraphException', 'Get-MbtMailboxTree', 'Resolve-MbtFolder', 'Get-MbtFolderItems', 'New-MbtFolderPath'
        'Export-MbtItems', 'Open-MbtExportStore', 'Save-MbtFolderTree', 'Save-MbtItemList', 'Export-MbtItemsToStore', 'Get-MbtStoreItems', 'Import-MbtStoreItem'
        'Read-MbtFxStream', 'Show-MbtFxReport', 'Import-MbtItem'
        'Resolve-MbtMailboxId', 'Wait-MbtTransferJob', 'Export-MbtMailbox', 'Import-MbtMapping', 'Import-MbtMailbox', 'Import-MbtPst', 'Export-MbtPstImportReport', 'Show-MbtPstImportPlan', 'Get-MbtImportUnknown', 'Resolve-MbtImportUnknown', 'Get-MbtStoreStatus', 'Show-MbtStoreStatus'
        'New-MbtBrowseLocation', 'Get-MbtBrowseLocation', 'Get-MbtBrowser', 'Get-MbtBrowseFolder', 'Find-MbtBrowseFolder', 'Get-MbtBrowseItem', 'Get-MbtBrowsePreview'
        'New-MbtBrowseSelection', 'Add-MbtBrowseSelection', 'Save-MbtBrowseSelection', 'Read-MbtBrowseSelection', 'Resolve-MbtBrowsePlan', 'Add-MbtBrowseMoveInfo', 'Format-MbtBrowseTotals'
        'Show-MbtBrowseTree', 'Show-MbtBrowseItems', 'Show-MbtBrowsePreview', 'Show-MbtExportPlan'
        'Test-MbtPst', 'Repair-MbtPst', 'Export-MbtPst', 'Copy-MbtPst', 'Get-MbtPstPreview', 'Show-MbtPstExportPlan', 'Resolve-MbtPstDestination', 'Show-MbtLiveExportPlan',
        'Export-MbtEml', 'Import-MbtEml', 'Resolve-MbtEmlUnknown', 'Compare-MbtEml',
        'Open-MbtEventLog', 'Close-MbtEventLog', 'Write-MbtEvent', 'Write-MbtJobPlan', 'Write-MbtJobStep', 'Write-MbtJobObject', 'Write-MbtJobResult', 'Write-MbtJobError', 'ConvertTo-MbtPlainError', 'Write-MbtJobStatus', 'Get-MbtJob'
        'Start-MbtStopWatch', 'Stop-MbtStopWatch', 'Test-MbtStopRequested', 'Get-MbtStopToken', 'Start-MbtKeepAwake', 'Stop-MbtKeepAwake', 'Get-MbtCapability'
        'Get-MbtConfiguredMailboxId', 'Assert-MbtSignInSettings'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
