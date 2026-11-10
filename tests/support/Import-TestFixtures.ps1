# Compiles the C# fixtures of the tests (fake HTTP handlers, FastTransfer builder, fake mailbox service, fake Mail API) once per process.
# Every test file of a Pester run goes through here: an in-memory assembly cannot be compiled twice under the same names.
function Import-TestFixtures {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$TestsRoot)
    if ('MailboxTransfer.Tests.FakeHandler' -as [type]) { return }
    $refs = 'System.Net.Http', 'System.Collections', 'System.Linq', 'System.Runtime', 'System.Threading.Tasks', 'System.Net.Primitives',
        'System.Memory', 'netstandard', 'System.Text.Encoding.Extensions', 'System.Runtime.Extensions', 'System.IO.FileSystem', 'System.Private.Uri',
        'System.Threading', 'System.Text.Json', 'System.Collections.Concurrent', 'System.Text.Encoding.CodePages', 'System.Text.RegularExpressions'
    Add-Type -LiteralPath @((Join-Path $TestsRoot 'Fixtures.cs'), (Join-Path $TestsRoot 'FakeMailboxService.cs'), (Join-Path $TestsRoot 'FakeMailApi.cs'), (Join-Path $TestsRoot 'EmlFakes.cs')) -ReferencedAssemblies $refs -IgnoreWarnings -WarningAction SilentlyContinue
}
