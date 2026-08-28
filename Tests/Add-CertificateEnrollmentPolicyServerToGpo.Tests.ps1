<#
.SYNOPSIS
    Pester suite for Add-CertificateEnrollmentPolicyServerToGpo.ps1. Requires Pester 5+.

.DESCRIPTION
    Three tiers (no Lab tier - writing to a real GPO is out of scope for CI):

      -Tag Unit    Pure helpers extracted from the script by AST (so the REAL code runs, never a
                   copy) and exercised in-process: the registry.pol binary parser (against a
                   hand-built .pol byte stream) and the entry/value extractors (against synthetic
                   record arrays). No GroupPolicy module, no GPO, no AD.
      -Tag Static  The script parses and its comment-based help binds. No module, no GPO.
      -Tag Guard   Parameter-conflict validation. These invocations throw BEFORE the script
                   resolves the GPO (Get-GPO), so they contact no GPO/AD and change nothing - but
                   the script's own '#Requires -Modules GroupPolicy' means it only loads where that
                   module is present, so the tier is skipped when GroupPolicy is unavailable.

.EXAMPLE
    Invoke-Pester -Path .\Tests\Add-CertificateEnrollmentPolicyServerToGpo.Tests.ps1

.EXAMPLE
    # Parser/extractor unit tests + parse/help only - runs without the GroupPolicy module:
    Invoke-Pester -Path .\Tests\Add-CertificateEnrollmentPolicyServerToGpo.Tests.ps1 -Tag Unit,Static
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '',
    Justification = 'the container parameter is consumed inside Pester Describe/BeforeAll scriptblocks, which the analyzer cannot see through')]
param(
    [string] $ScriptPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'Add-CertificateEnrollmentPolicyServerToGpo.ps1')
)

BeforeDiscovery {
    # -Skip is evaluated during discovery, so the gate must be set here.
    $script:HasGP = [bool](Get-Module -ListAvailable GroupPolicy)
}

Describe 'Add-CertificateEnrollmentPolicyServerToGpo' {

    BeforeAll {
        $script:Gpo = $ScriptPath
        $script:Gpo | Should -Exist

        # --- AST-extract the pure helpers so the Unit tier exercises the REAL code ------------
        # (The script has mandatory params and '#Requires -Modules GroupPolicy', so it cannot be
        # dot-sourced wholesale; extracting the function bodies runs them with no side effects.)
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:Gpo, [ref]$null, [ref]$null)
        foreach ($name in 'Read-PolRecords', 'Get-PolEntries', 'Get-PolValue') {
            $def = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
            if ($def) { . ([scriptblock]::Create($def[0].Extent.Text)) }
        }
        # Get-PolEntries closes over $relBase - mirror the script's definition.
        $script:relBase = 'Software\Policies\Microsoft\Cryptography\PolicyServers'
        $script:leaf    = 'dc032f3a68521c2445e1e161da81503bddce17a7'
        $script:entryKey = "$script:relBase\$script:leaf"

        # --- byte builder for a synthetic Registry.pol ([MS-GPREG]) --------------------------
        function New-Uni([string]$s) { , [System.Text.Encoding]::Unicode.GetBytes($s) }   # unary comma: return the byte[] intact, not unrolled
        function New-PolRecord {
            param([string]$Key, [string]$Value, [uint32]$Type, [byte[]]$Data)
            $b   = [System.Collections.Generic.List[byte]]::new()
            $nul = [byte[]](0, 0)
            $b.AddRange((New-Uni '['))
            $b.AddRange((New-Uni $Key));   $b.AddRange($nul); $b.AddRange((New-Uni ';'))
            $b.AddRange((New-Uni $Value)); $b.AddRange($nul); $b.AddRange((New-Uni ';'))
            $b.AddRange([BitConverter]::GetBytes([uint32]$Type));        $b.AddRange((New-Uni ';'))
            $b.AddRange([BitConverter]::GetBytes([uint32]$Data.Length)); $b.AddRange((New-Uni ';'))
            $b.AddRange($Data)
            $b.AddRange((New-Uni ']'))
            , $b.ToArray()
        }
        function New-Sz([string]$s)   { (New-Uni $s) + [byte[]](0, 0) }        # REG_SZ: NUL-terminated
        function New-Dword([uint32]$v) { [BitConverter]::GetBytes([uint32]$v) }

        $script:PolPath = Join-Path $TestDrive 'registry.pol'
        $pol = [System.Collections.Generic.List[byte]]::new()
        $pol.AddRange([byte[]](0x50, 0x52, 0x65, 0x67, 1, 0, 0, 0))          # "PReg" + version 1
        $pol.AddRange((New-PolRecord -Key $script:entryKey -Value 'URL'      -Type 1 -Data (New-Sz 'https://pki.example.net/ejbca/msae/CEPService?alias')))
        $pol.AddRange((New-PolRecord -Key $script:entryKey -Value 'PolicyID' -Type 1 -Data (New-Sz '241064013')))
        $pol.AddRange((New-PolRecord -Key $script:entryKey -Value 'Flags'    -Type 4 -Data (New-Dword 0x14)))
        $pol.AddRange((New-PolRecord -Key $script:relBase  -Value ''         -Type 1 -Data (New-Sz '241064013')))   # (Default) marker
        $pol.AddRange((New-PolRecord -Key $script:relBase  -Value 'Flags'    -Type 4 -Data (New-Dword 0x4)))        # root DISABLE bit
        [System.IO.File]::WriteAllBytes($script:PolPath, $pol.ToArray())
    }

    Context 'Unit: registry.pol parser and extractors' -Tag 'Unit' {

        It 'Read-PolRecords parses every record from a Registry.pol byte stream' {
            $recs = @(Read-PolRecords -Path $script:PolPath)
            $recs.Count | Should -Be 5
        }

        It 'Read-PolRecords decodes REG_SZ (trimmed) and REG_DWORD values by type' {
            $recs = Read-PolRecords -Path $script:PolPath
            $url = $recs | Where-Object { $_.Key -eq $script:entryKey -and $_.ValueName -eq 'URL' }
            $url.Data | Should -BeExactly 'https://pki.example.net/ejbca/msae/CEPService?alias'
            $flags = $recs | Where-Object { $_.Key -eq $script:entryKey -and $_.ValueName -eq 'Flags' }
            [uint32]$flags.Data | Should -Be 0x14
        }

        It 'Read-PolRecords returns no records for a missing file' {
            @(Read-PolRecords -Path (Join-Path $TestDrive 'nope.pol')).Count | Should -Be 0
        }

        It 'Get-PolEntries groups values into one entry per 40-hex subkey' {
            $recs = Read-PolRecords -Path $script:PolPath
            $entries = @(Get-PolEntries $recs)
            $entries.Count | Should -Be 1
            $entries[0].Key      | Should -BeExactly $script:leaf
            $entries[0].URL      | Should -BeExactly 'https://pki.example.net/ejbca/msae/CEPService?alias'
            $entries[0].PolicyID | Should -BeExactly '241064013'
        }

        It 'Get-PolEntries ignores GP deletion records (**del*) so they raise no phantom entry' {
            $recs = @(Read-PolRecords -Path $script:PolPath) + [pscustomobject]@{
                Key = "$script:relBase\ffffffffffffffffffffffffffffffffffffffff"; ValueName = '**del.URL'; Type = 1; Data = ' '; Index = 99
            }
            @(Get-PolEntries $recs).Count | Should -Be 1
        }

        It 'Get-PolValue reads a root value and is case-sensitive on the value name' {
            $recs = Read-PolRecords -Path $script:PolPath
            (Get-PolValue $recs $script:relBase '')      | Should -BeExactly '241064013'   # (Default) marker
            [uint32](Get-PolValue $recs $script:relBase 'Flags') | Should -Be 0x4
            (Get-PolValue $recs $script:relBase 'flags') | Should -BeNullOrEmpty            # wrong case -> no match
        }
    }

    Context 'Static: parse and help' -Tag 'Static' {

        It 'parses without errors' {
            $errs = $null
            $null = [System.Management.Automation.Language.Parser]::ParseFile($script:Gpo, [ref]$null, [ref]$errs)
            $errs | Should -BeNullOrEmpty
        }

        It 'comment-based help binds (Synopsis is present)' {
            (Get-Help $script:Gpo).Synopsis.Trim() | Should -Not -BeNullOrEmpty
        }

        It 'documents every non-common parameter' {
            $cmd = Get-Command $script:Gpo
            $common = [System.Management.Automation.PSCmdlet]::CommonParameters + [System.Management.Automation.PSCmdlet]::OptionalCommonParameters
            $documented = @((Get-Help $script:Gpo).parameters.parameter.name)
            foreach ($p in $cmd.Parameters.Keys | Where-Object { $_ -notin $common }) {
                $documented | Should -Contain $p -Because "parameter -$p should have a .PARAMETER help entry"
            }
        }
    }

    Context 'Guard: parameter-conflict validation (throws before Get-GPO)' -Tag 'Guard' -Skip:(-not $script:HasGP) {

        It 'rejects -SetAsDefault with -ClearDefault' {
            { & $script:Gpo -GpoName 'x' -Url 'https://y/' -PolicyName 'z' -SetAsDefault -ClearDefault } |
                Should -Throw -ExpectedMessage '*mutually exclusive*'
        }

        It 'rejects -DisableUserConfigured with -EnableUserConfigured' {
            { & $script:Gpo -GpoName 'x' -Url 'https://y/' -PolicyName 'z' -DisableUserConfigured -EnableUserConfigured } |
                Should -Throw -ExpectedMessage '*mutually exclusive*'
        }

        It 'rejects an Auto-Enrollment tuning switch without -EnableAutoEnrollmentPolicy' {
            { & $script:Gpo -GpoName 'x' -Url 'https://y/' -PolicyName 'z' -AEPolicy 5 } |
                Should -Throw -ExpectedMessage '*only has effect together with -EnableAutoEnrollmentPolicy*'
        }
    }
}
