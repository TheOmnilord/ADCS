<#
.SYNOPSIS
    Pester suite for Add-CertificateEnrollmentPolicyServerToGpo.ps1. Requires Pester 5+.

.DESCRIPTION
    Three always-on tiers plus one opt-in tier:

      -Tag Unit    Pure helpers extracted from the script by AST (so the REAL code runs, never a
                   copy) and exercised in-process: the registry.pol binary parser (against a
                   hand-built .pol byte stream), the entry/value extractors (against synthetic
                   record arrays), the DWORD converter, and the write/verify helpers (Invoke-GPWrite,
                   Test-GpoEntry) against mocked cmdlets. No GroupPolicy module, no GPO, no AD.
      -Tag Static  The script parses and its comment-based help binds. No module, no GPO.
      -Tag Guard   Parameter-conflict validation (throws BEFORE the script resolves the GPO), plus
                   the GPO-resolution branches, the post-resolution input validation and the
                   pre-write root-Flags refusal against a MOCKED Get-GPO and Read-PolRecords (the
                   mock reaches the script because Pester defines it as a script-scope alias, and
                   an alias wins over the script's own function). No GPO/AD is contacted and
                   nothing is written - but the
                   script imports the GroupPolicy module up front, so the tier is skipped when the
                   module is unavailable (detection dual-probes ListAvailable AND the Windows
                   PowerShell module path, since PowerShell 7 loads it via the WinPSCompat shim).
      -Tag Lab     Live GPO round-trips. Skipped unless -RunLab is passed; needs the GroupPolicy
                   module, a domain-joined machine, and permission to create GPOs. Surgical by
                   construction: it creates ONE throwaway GPO named PESTER-<hex> and NEVER links
                   it (an unlinked GPO applies to zero clients), authors and removes CEP entries
                   only inside that GPO, and deletes it in teardown by its exact tracked GUID
                   (with a run-prefix-scoped Get-GPO sweep as backstop). Pre-existing GPOs are
                   never touched. As a bonus, the tier re-reads the GPO's REAL registry.pol with
                   the suite's extracted parser - validating the parser against a genuine
                   Set-GPRegistryValue-authored file, not only the hand-built one.

.EXAMPLE
    Invoke-Pester -Path .\Tests\Add-CertificateEnrollmentPolicyServerToGpo.Tests.ps1 -ExcludeTag Lab

.EXAMPLE
    # Parser/extractor unit tests + parse/help only - runs without the GroupPolicy module:
    Invoke-Pester -Path .\Tests\Add-CertificateEnrollmentPolicyServerToGpo.Tests.ps1 -Tag Unit,Static

.EXAMPLE
    # Full run including the live GPO round-trip (throwaway unlinked GPO, removed afterwards):
    $cfg = New-PesterContainer -Path .\Tests\Add-CertificateEnrollmentPolicyServerToGpo.Tests.ps1 -Data @{ RunLab = $true }
    Invoke-Pester -Container $cfg
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '',
    Justification = 'container parameters are consumed inside Pester Describe/BeforeAll scriptblocks, which the analyzer cannot see through')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingEmptyCatchBlock', '',
    Justification = 'best-effort teardown paths (AfterAll GPO removal and backstop sweep) deliberately swallow per-item errors')]
param(
    [bool]   $RunLab     = $false,
    [string] $ScriptPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'Add-CertificateEnrollmentPolicyServerToGpo.ps1')
)

BeforeDiscovery {
    # -Skip is evaluated during discovery, so the gates must be set here. The domain check runs
    # only when -RunLab is passed (short-circuit), keeping ordinary runs free of CIM calls.
    # GroupPolicy detection must dual-probe: on PowerShell 7, Get-Module -ListAvailable does NOT
    # see the module (it lives only in the Windows PowerShell module path), yet Import-Module
    # loads it there via the WinPSCompatSession shim - so also probe that path directly.
    $script:HasGP = [bool](Get-Module -ListAvailable GroupPolicy) -or
                    (Test-Path "$env:windir\System32\WindowsPowerShell\v1.0\Modules\GroupPolicy")
    $script:GpoLabReady = $RunLab -and $script:HasGP -and (Get-CimInstance Win32_ComputerSystem).PartOfDomain
}

Describe 'Add-CertificateEnrollmentPolicyServerToGpo' {

    BeforeAll {
        $script:Gpo = $ScriptPath
        $script:Gpo | Should -Exist

        # --- AST-extract the pure helpers so the Unit tier exercises the REAL code ------------
        # (The script has mandatory params and imports GroupPolicy on load, so it cannot be
        # dot-sourced wholesale; extracting the function bodies runs them with no side effects.)
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:Gpo, [ref]$null, [ref]$null)
        foreach ($name in 'Read-PolRecords', 'Get-PolEffectiveValues', 'Get-PolEntries', 'Get-PolValue', 'Test-EntryRecordsPresent', 'Test-PolEntryUsable',
                          'Test-PolDeletionOrder', 'Test-GpoEntry', 'Invoke-GPWrite', 'ConvertTo-PolDwordValue', 'Get-PolRawRecord',
                          'Get-PolDwordRecordReason', 'Get-RootFlagsDisplay') {
            $def = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
            if ($def) { . ([scriptblock]::Create($def[0].Extent.Text)) }
        }
        # The PolicyID derivation is an inline block in the script's main body (the EJBCA MSAE
        # Java String.hashCode). Extract that if-statement so the oracle test runs the real code.
        $script:PidBlock = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.IfStatementAst] -and $n.Extent.Text -match 'String\.hashCode' }, $true)
        $script:PidBlock.Count | Should -Be 1
        function script:Get-DerivedPolicyId([string]$PolicyName, [string]$PolicyId) {
            # The block reads $PolicyName and assigns $PolicyId when it is empty - the script's own code.
            . ([scriptblock]::Create($script:PidBlock[0].Extent.Text))
            $PolicyId
        }
        # Get-PolEntries closes over $relBase - mirror the script's definition.
        $script:relBase = 'Software\Policies\Microsoft\Cryptography\PolicyServers'
        $script:leaf    = 'dc032f3a68521c2445e1e161da81503bddce17a7'
        $script:entryKey = "$script:relBase\$script:leaf"

        # --- byte builder for a synthetic Registry.pol ([MS-GPREG]) --------------------------
        function New-Uni([string]$s) { , [System.Text.Encoding]::Unicode.GetBytes($s) }   # unary comma: return the byte[] intact, not unrolled
        function New-PolRecord {
            # -CorruptSeparator N (1..4) writes 'X' in place of the Nth ';' only (after the key
            # name, the value name, the type, the size); -Closer replaces the closing ']'. Both
            # build a damaged record on purpose, with every other framing character intact.
            param([string]$Key, [string]$Value, [uint32]$Type, [byte[]]$Data, [int]$CorruptSeparator = 0, [string]$Closer = ']')
            $b   = [System.Collections.Generic.List[byte]]::new()
            $nul = [byte[]](0, 0)
            $sep = { param([int]$n) if ($n -eq $CorruptSeparator) { New-Uni 'X' } else { New-Uni ';' } }
            $b.AddRange((New-Uni '['))
            $b.AddRange((New-Uni $Key));   $b.AddRange($nul); $b.AddRange((& $sep 1))
            $b.AddRange((New-Uni $Value)); $b.AddRange($nul); $b.AddRange((& $sep 2))
            $b.AddRange([BitConverter]::GetBytes([uint32]$Type));        $b.AddRange((& $sep 3))
            $b.AddRange([BitConverter]::GetBytes([uint32]$Data.Length)); $b.AddRange((& $sep 4))
            if ($Data.Length) { $b.AddRange($Data) }
            $b.AddRange((New-Uni $Closer))
            , $b.ToArray()
        }
        function New-Sz([string]$s)   { (New-Uni $s) + [byte[]](0, 0) }        # REG_SZ: NUL-terminated
        function New-Dword([uint32]$v) { [BitConverter]::GetBytes([uint32]$v) }
        $script:PolHeader = [byte[]](0x50, 0x52, 0x65, 0x67, 1, 0, 0, 0)     # "PReg" + version 1

        $script:PolPath = Join-Path $TestDrive 'registry.pol'
        $pol = [System.Collections.Generic.List[byte]]::new()
        $pol.AddRange($script:PolHeader)
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

        It 'Read-PolRecords decodes a **-instruction''s data as a string whatever its type field says (a REG_BINARY **DeleteKeys still deletes)' {
            $p = Join-Path $TestDrive 'deletekeys-binary.pol'
            $pol = [System.Collections.Generic.List[byte]]::new()
            $pol.AddRange([System.IO.File]::ReadAllBytes($script:PolPath))
            # the same UTF-16 key list, but typed 3 (REG_BINARY) instead of 1 (REG_SZ)
            $pol.AddRange((New-PolRecord -Key '' -Value '**DeleteKeys' -Type 3 -Data (New-Sz $script:entryKey)))
            [System.IO.File]::WriteAllBytes($p, $pol.ToArray())
            $recs = @(Read-PolRecords -Path $p)
            ($recs | Where-Object { $_.ValueName -eq '**DeleteKeys' }).Data | Should -Be $script:entryKey
            (Get-PolValue $recs $script:entryKey 'URL') | Should -BeNullOrEmpty -Because 'clients delete the key regardless of the type field; a $null-decoded instruction would leave the entry looking present'
            @(Get-PolEntries $recs).Count | Should -Be 0
        }

        It 'Read-PolRecords keeps the declared type of a **soft.<name> record (it writes a value; a DWORD Cost stays a number)' {
            $p = Join-Path $TestDrive 'soft-dword.pol'
            $pol = [System.Collections.Generic.List[byte]]::new()
            $pol.AddRange([System.IO.File]::ReadAllBytes($script:PolPath))
            $pol.AddRange((New-PolRecord -Key $script:entryKey -Value '**soft.Cost' -Type 4 -Data (New-Dword 5)))
            [System.IO.File]::WriteAllBytes($p, $pol.ToArray())
            $recs = @(Read-PolRecords -Path $p)
            [uint32](Get-PolValue $recs $script:entryKey 'Cost') | Should -Be 5 -Because 'string-decoding a DWORD 5 would yield the control character U+0005'
        }

        It 'Test-PolEntryUsable requires the operational values, not just URL + PolicyID (an interrupted write leaves only those two)' {
            $url = 'https://pki.example.net/ejbca/msae/CEPService?alias'
            $full = @{ URL = $url; PolicyID = '241064013'; FriendlyName = 'Example'; Flags = [uint32]0x14; AuthFlags = [uint32]2; Cost = [uint32]0x7FFFFFFD }
            Test-PolEntryUsable $full $url '241064013' | Should -BeTrue
            Test-PolEntryUsable @{ URL = $url; PolicyID = '241064013' } $url '241064013' | Should -BeFalse -Because 'URL and PolicyID are written first; a run interrupted after them leaves a row no client can use'
            # An EMPTY URL or PolicyID must be rejected even when the caller passes that same empty
            # value as the expected one (the root-Flags gate does exactly this for arbitrary siblings):
            # an otherwise-complete row with no PolicyID is not a usable server.
            $noPid = $full.Clone(); $noPid['PolicyID'] = ''
            Test-PolEntryUsable $noPid $url '' | Should -BeFalse -Because 'empty PolicyID must not pass the empty-equals-empty comparison'
            $noUrl = $full.Clone(); $noUrl['URL'] = ''
            Test-PolEntryUsable $noUrl '' '241064013' | Should -BeFalse -Because 'empty URL must not pass'
            # URL/PolicyID/FriendlyName must be REG_SZ: a PolicyID decoded as a DWORD (uint32) stringifies
            # to a matching value but is not the string a client reads, so it is not a usable server.
            $dwordPid = $full.Clone(); $dwordPid['PolicyID'] = [uint32]241064013
            Test-PolEntryUsable $dwordPid '241064013' '241064013' | Should -BeFalse -Because 'a REG_DWORD PolicyID is not the REG_SZ a client reads'
            $dwordFn = $full.Clone(); $dwordFn['FriendlyName'] = [uint32]1
            Test-PolEntryUsable $dwordFn $url '241064013' | Should -BeFalse
            $dwordUrl = $full.Clone(); $dwordUrl['URL'] = [uint32]1
            Test-PolEntryUsable $dwordUrl 1 '241064013' | Should -BeFalse
            foreach ($missing in 'FriendlyName', 'Flags', 'AuthFlags', 'Cost') {
                $h = $full.Clone(); $h.Remove($missing)
                Test-PolEntryUsable $h $url '241064013' | Should -BeFalse -Because "$missing is required"
            }
            $wrongType = $full.Clone(); $wrongType['AuthFlags'] = 'Kerberos'
            Test-PolEntryUsable $wrongType $url '241064013' | Should -BeFalse
            # a REG_SZ that merely spells a number is still the wrong registry type
            $numericString = $full.Clone(); $numericString['Flags'] = '20'; $numericString['Cost'] = '5'
            Test-PolEntryUsable $numericString $url '241064013' | Should -BeFalse -Because 'clients need a DWORD, not a string that looks like one'
            # a REG_QWORD (Int64 as GPMC returns it) is not a DWORD either
            $qword = $full.Clone(); $qword['AuthFlags'] = [long]2
            Test-PolEntryUsable $qword $url '241064013' | Should -BeFalse -Because 'a QWORD AuthFlags is the wrong registry type'
            # Get-GPRegistryValue returns a high-bit DWORD as a signed Int32: Cost 0xFFFFFFFF reads as -1
            $signed = $full.Clone(); $signed['Cost'] = [int]-1
            Test-PolEntryUsable $signed $url '241064013' | Should -BeTrue -Because 'a live read of Cost 0xFFFFFFFF is -1 and must still count as the DWORD it is'
            Test-PolEntryUsable $full $url '999' | Should -BeFalse
            Test-PolEntryUsable $null $url '241064013' | Should -BeFalse
        }

        It 'Read-PolRecords returns no records for a missing file' {
            @(Read-PolRecords -Path (Join-Path $TestDrive 'nope.pol')).Count | Should -Be 0
        }

        It 'Read-PolRecords refuses a file without the PReg/version-1 header, and trailing junk after the records' {
            $bad = Join-Path $TestDrive 'bad-header.pol'
            [System.IO.File]::WriteAllBytes($bad, [byte[]](0x4D, 0x5A, 0, 0, 1, 0, 0, 0) + [byte[]](1..16))
            { Read-PolRecords -Path $bad } | Should -Throw -ExpectedMessage '*PReg*'
            $junk = Join-Path $TestDrive 'trailing-junk.pol'
            [System.IO.File]::WriteAllBytes($junk, [System.IO.File]::ReadAllBytes($script:PolPath) + [byte[]](0x41, 0x00, 0x42, 0x00))
            { Read-PolRecords -Path $junk } | Should -Throw -ExpectedMessage '*truncated or corrupt*'
            # a header-only file is a legitimately EMPTY policy, not an error
            $empty = Join-Path $TestDrive 'empty.pol'
            [System.IO.File]::WriteAllBytes($empty, [byte[]](0x50, 0x52, 0x65, 0x67, 1, 0, 0, 0))
            @(Read-PolRecords -Path $empty).Count | Should -Be 0
            # a file shorter than the 8-byte header is TRUNCATED (one byte of the version field
            # here), not empty: it is refused, so a damaged policy never passes for an empty one
            $stub = Join-Path $TestDrive 'five-bytes.pol'
            [System.IO.File]::WriteAllBytes($stub, [byte[]](0x50, 0x52, 0x65, 0x67, 1))
            { Read-PolRecords -Path $stub } | Should -Throw -ExpectedMessage '*PReg*'
            # a zero-byte file is a policy that was never written: no records, no error
            $zero = Join-Path $TestDrive 'zero-bytes.pol'
            [System.IO.File]::WriteAllBytes($zero, [byte[]]@())
            @(Read-PolRecords -Path $zero).Count | Should -Be 0
        }

        It 'Read-PolRecords refuses damaged record framing: a missing or wrong closing bracket, and ONE leftover byte' {
            # The 1.0.1 "trailing junk refused" check only caught two or more bytes; a file cut
            # exactly at the end of a record's data (no ']') parsed as intact.
            $intact = [System.IO.File]::ReadAllBytes($script:PolPath)
            $noBracket = Join-Path $TestDrive 'no-close-bracket.pol'
            [System.IO.File]::WriteAllBytes($noBracket, $intact[0..($intact.Length - 3)])
            { Read-PolRecords -Path $noBracket } | Should -Throw -ExpectedMessage '*truncated or corrupt*' -Because 'the last record has no closing ]'
            $oneByte = Join-Path $TestDrive 'one-trailing-byte.pol'
            [System.IO.File]::WriteAllBytes($oneByte, $intact + [byte[]](0x41))
            { Read-PolRecords -Path $oneByte } | Should -Throw -ExpectedMessage '*truncated or corrupt*' -Because 'one odd byte after the last record is corruption, not padding'
            # A wrong closing character is a separate check from a MISSING one: the record is
            # complete in length, so only the ']' comparison itself can catch it.
            $wrongClose = Join-Path $TestDrive 'wrong-close-bracket.pol'
            [System.IO.File]::WriteAllBytes($wrongClose, $script:PolHeader + (New-PolRecord -Key $script:entryKey -Value 'URL' -Type 1 -Data (New-Sz 'https://x/') -Closer '}'))
            { Read-PolRecords -Path $wrongClose } | Should -Throw -ExpectedMessage '*truncated or corrupt*' -Because '} in place of the closing ] is not a record'
            # ...and the intact file still parses after the stricter checks
            @(Read-PolRecords -Path $script:PolPath).Count | Should -Be 5
        }

        It 'Read-PolRecords checks each of the four ; separators on its own (the <Name>)' -TestCases @(
            @{ N = 1; Name = 'one after the key name' }
            @{ N = 2; Name = 'one after the value name' }
            @{ N = 3; Name = 'one after the type' }
            @{ N = 4; Name = 'one after the size' }
        ) {
            # One separator corrupted at a time, every field before it valid: parsing stops at the
            # first bad character, so a record with all four replaced would prove only the first
            # check. Each case passes only when ITS check exists.
            $p = Join-Path $TestDrive "x-separator-$N.pol"
            [System.IO.File]::WriteAllBytes($p, $script:PolHeader + (New-PolRecord -Key $script:entryKey -Value 'URL' -Type 1 -Data (New-Sz 'https://x/') -CorruptSeparator $N))
            { Read-PolRecords -Path $p } | Should -Throw -ExpectedMessage '*truncated or corrupt*' -Because "X in place of separator $N ($Name) is not a record"
            # the same record with every separator intact is a record
            $ok = Join-Path $TestDrive "ok-separator-$N.pol"
            [System.IO.File]::WriteAllBytes($ok, $script:PolHeader + (New-PolRecord -Key $script:entryKey -Value 'URL' -Type 1 -Data (New-Sz 'https://x/')))
            @(Read-PolRecords -Path $ok).Count | Should -Be 1
        }

        It 'Read-PolRecords decodes the unusual shapes: size-0 data, a short DWORD, and a REG_BINARY ordinary value' {
            $p = Join-Path $TestDrive 'shapes.pol'
            $pol = [System.Collections.Generic.List[byte]]::new()
            $pol.AddRange($script:PolHeader)
            $pol.AddRange((New-PolRecord -Key $script:relBase -Value 'EmptySz'    -Type 1 -Data ([byte[]]@())))
            $pol.AddRange((New-PolRecord -Key $script:relBase -Value 'EmptyDword' -Type 4 -Data ([byte[]]@())))
            $pol.AddRange((New-PolRecord -Key $script:relBase -Value 'ShortDword' -Type 4 -Data ([byte[]](1, 0))))
            $pol.AddRange((New-PolRecord -Key $script:relBase -Value 'Binary'     -Type 3 -Data ([byte[]](1, 2, 3))))
            $pol.AddRange((New-PolRecord -Key $script:relBase -Value 'Flags'      -Type 4 -Data (New-Dword 6)))
            [System.IO.File]::WriteAllBytes($p, $pol.ToArray())
            $recs = @(Read-PolRecords -Path $p)
            $recs.Count | Should -Be 5
            $byName = @{}; foreach ($r in $recs) { $byName[$r.ValueName] = $r }
            $byName['EmptySz'].Data    | Should -BeExactly '' -Because 'a REG_SZ with no data is the empty string'
            $byName['EmptyDword'].Data | Should -BeNullOrEmpty -Because 'a DWORD needs 4 data bytes'
            $byName['ShortDword'].Data | Should -BeNullOrEmpty
            $byName['Binary'].Data     | Should -BeNullOrEmpty -Because 'REG_BINARY is not a type this script reads'
            [uint32]$byName['Flags'].Data | Should -Be 6 -Because 'the records after the odd shapes still parse'
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

        It 'Get-PolValue reads a root value; value names compare case-insensitively (registry semantics)' {
            $recs = Read-PolRecords -Path $script:PolPath
            (Get-PolValue $recs $script:relBase '')      | Should -BeExactly '241064013'   # (Default) marker
            [uint32](Get-PolValue $recs $script:relBase 'Flags') | Should -Be 0x4
            [uint32](Get-PolValue $recs $script:relBase 'flags') | Should -Be 0x4          # the registry is case-insensitive on value names
            (Get-PolValue $recs $script:relBase 'Nope')  | Should -BeNullOrEmpty
        }

        It 'Get-PolValue returns the EFFECTIVE (last) record when a value appears more than once' {
            # Registry.pol is applied in order: a later record for the same name replaces the
            # earlier one, so state decisions must not be based on the obsolete first record.
            $recs = @(Read-PolRecords -Path $script:PolPath) + [pscustomobject]@{
                Key = $script:relBase; ValueName = 'Flags'; Type = 4; Data = [uint32]0; Index = 99
            }
            [uint32](Get-PolValue $recs $script:relBase 'Flags') | Should -Be 0
        }

        It 'Get-PolValue honours a later **del.<name> / **delvals. record (value deleted -> $null)' {
            $base = @(Read-PolRecords -Path $script:PolPath)
            $del1 = $base + [pscustomobject]@{ Key = $script:relBase; ValueName = '**del.Flags'; Type = 1; Data = ' '; Index = 99 }
            (Get-PolValue $del1 $script:relBase 'Flags') | Should -BeNullOrEmpty
            (Get-PolValue $del1 $script:relBase '')      | Should -BeExactly '241064013'   # untouched sibling survives
            $del2 = $base + [pscustomobject]@{ Key = $script:entryKey; ValueName = '**delvals.'; Type = 1; Data = ' '; Index = 99 }
            (Get-PolValue $del2 $script:entryKey 'URL')  | Should -BeNullOrEmpty
            @(Get-PolEntries $del2).Count | Should -Be 0 -Because 'an entry whose values are all deleted later is not one a client ends up with'
        }

        It 'Get-PolEffectiveValues honours **DeleteValues: each name in its ;-separated data is removed, the others stay' {
            $recs = @(Read-PolRecords -Path $script:PolPath) + [pscustomobject]@{ Key = $script:entryKey; ValueName = '**DeleteValues'; Type = 1; Data = 'url; Nope'; Index = 99 }
            $eff = Get-PolEffectiveValues $recs $script:entryKey
            $eff.Values.ContainsKey('URL') | Should -BeFalse -Because 'value names compare case-insensitively'
            @($eff.Deleted.Keys) | Should -Contain 'URL'
            $eff.Values['PolicyID'] | Should -BeExactly '241064013' -Because 'a name the instruction does not list survives'
            $eff.Deleted.ContainsKey('Nope') | Should -BeFalse -Because 'a name that was never written is not lost'
        }

        It 'Get-PolEffectiveValues honours **DeleteKeys as the GP engine does: FULL key paths in the data, record key ignored' {
            # Semantics per the LGPO author's corrections to MS-GPREG: the data is a ';'-separated
            # list of full key paths (no hive); the record's own key field is ignored entirely.
            $base = @(Read-PolRecords -Path $script:PolPath)
            # full path of this entry, recorded under an unrelated key (even an empty one)
            $byPath = $base + [pscustomobject]@{ Key = ''; ValueName = '**DeleteKeys'; Type = 1; Data = "Software\Policies\Other;$script:entryKey"; Index = 99 }
            (Get-PolValue $byPath $script:entryKey 'URL') | Should -BeNullOrEmpty
            @(Get-PolEntries $byPath).Count | Should -Be 0 -Because 'a client applying this file ends up without the entry'
            (Get-PolValue $byPath $script:relBase 'Flags') | Should -Not -BeNullOrEmpty -Because 'the PolicyServers key itself is not named'
            # naming an ANCESTOR (the whole PolicyServers key) deletes the entry and the root values
            $byAncestor = $base + [pscustomobject]@{ Key = 'whatever'; ValueName = '**DeleteKeys'; Type = 1; Data = $script:relBase; Index = 99 }
            (Get-PolValue $byAncestor $script:entryKey 'URL') | Should -BeNullOrEmpty
            (Get-PolValue $byAncestor $script:relBase 'Flags') | Should -BeNullOrEmpty
            # the documented-but-wrong parent-relative reading must NOT delete anything: a bare leaf
            # name is the key "<leaf>" at the hive root, not our entry
            $bareLeaf = $base + [pscustomobject]@{ Key = $script:relBase; ValueName = '**DeleteKeys'; Type = 1; Data = $script:leaf; Index = 99 }
            (Get-PolValue $bareLeaf $script:entryKey 'URL') | Should -Not -BeNullOrEmpty
            # a path that merely PREFIXES ours is not an ancestor (PolicyServers2 vs PolicyServers)
            $prefix = $base + [pscustomobject]@{ Key = ''; ValueName = '**DeleteKeys'; Type = 1; Data = "${script:relBase}2"; Index = 99 }
            (Get-PolValue $prefix $script:entryKey 'URL') | Should -Not -BeNullOrEmpty
            # ...and a DeleteKeys BEFORE the entry's records does not delete what is written afterwards
            $before = @([pscustomobject]@{ Key = ''; ValueName = '**DeleteKeys'; Type = 1; Data = $script:entryKey; Index = -1 }) + $base
            (Get-PolValue $before $script:entryKey 'URL') | Should -Not -BeNullOrEmpty
        }

        It 'Get-PolEffectiveValues skips **Comment:/**SecureKey and applies **soft. only when the value is absent' {
            $base = @(Read-PolRecords -Path $script:PolPath)
            $recs = $base + @(
                [pscustomobject]@{ Key = $script:entryKey; ValueName = '**Comment: from GPO X'; Type = 1; Data = ' '; Index = 90 }
                [pscustomobject]@{ Key = $script:entryKey; ValueName = '**SecureKey';           Type = 4; Data = 1;   Index = 91 }
                [pscustomobject]@{ Key = $script:entryKey; ValueName = '**soft.URL';            Type = 1; Data = 'https://soft.example/'; Index = 92 }   # URL exists -> ignored
                [pscustomobject]@{ Key = $script:entryKey; ValueName = '**soft.Cost';           Type = 4; Data = [uint32]5; Index = 93 }                 # Cost absent -> written
            )
            $eff = (Get-PolEffectiveValues $recs $script:entryKey).Values
            $eff['URL']  | Should -BeExactly 'https://pki.example.net/ejbca/msae/CEPService?alias'
            [uint32]$eff['Cost'] | Should -Be 5
            @($eff.Keys | Where-Object { $_ -like '`*`**' }).Count | Should -Be 0 -Because 'instructions never surface as values'
        }

        It 'Test-EntryRecordsPresent sees a damaged entry (values wiped by a later **delvals.) that the effective view hides - so -Remove can clear it' {
            $damaged = @(Read-PolRecords -Path $script:PolPath) + [pscustomobject]@{ Key = $script:entryKey; ValueName = '**delvals.'; Type = 1; Data = ' '; Index = 99 }
            @(Get-PolEntries $damaged).Count | Should -Be 0 -Because 'clients never see the entry'
            Test-EntryRecordsPresent $damaged $script:entryKey | Should -BeTrue -Because 'its records still occupy the key and must be removable'
            Test-EntryRecordsPresent $damaged "$script:relBase\ffffffffffffffffffffffffffffffffffffffff" | Should -BeFalse
            Test-EntryRecordsPresent @() $script:entryKey | Should -BeFalse
        }

        It 'Get-PolEffectiveValues fails CLOSED on a deletion instruction it does not understand' {
            $recs = @(Read-PolRecords -Path $script:PolPath) + [pscustomobject]@{ Key = $script:entryKey; ValueName = '**delsomethingnew'; Type = 1; Data = ' '; Index = 99 }
            { Get-PolValue $recs $script:entryKey 'URL' } | Should -Throw -ExpectedMessage '*does not understand*'
        }

        It 'Get-PolEffectiveValues: a deletion BETWEEN an obsolete record and its replacement loses nothing' {
            # old value -> **del.Flags -> new value: the client ends with the new value. Only a
            # deletion AFTER a value's last record is damage (the Deleted set drives the warning).
            $recs = @(
                [pscustomobject]@{ Key = $script:relBase; ValueName = 'Flags';       Type = 4; Data = [uint32]4; Index = 0 }
                [pscustomobject]@{ Key = $script:relBase; ValueName = '**del.Flags'; Type = 1; Data = ' ';       Index = 1 }
                [pscustomobject]@{ Key = $script:relBase; ValueName = 'Flags';       Type = 4; Data = [uint32]0; Index = 2 }
            )
            $eff = Get-PolEffectiveValues $recs $script:relBase
            [uint32]$eff.Values['Flags'] | Should -Be 0
            $eff.Deleted.Count | Should -Be 0 -Because 'the replacement survives, so nothing is lost'

            # ...whereas a deletion after the LAST record really does lose the value.
            $eff2 = Get-PolEffectiveValues ($recs[2], $recs[1]) $script:relBase
            $eff2.Values.ContainsKey('Flags') | Should -BeFalse
            @($eff2.Deleted.Keys) | Should -Contain 'Flags'
        }

        It 'Test-PolDeletionOrder names the lost values, (Default) included, and stays quiet for a deletion between a record and its replacement' {
            # The helper closes over the script''s $notes list; dynamic scoping gives it this one.
            $notes = New-Object System.Collections.Generic.List[string]
            $damaged = @(
                [pscustomobject]@{ Key = $script:relBase; ValueName = '';           Type = 1; Data = '241064013'; Index = 0 }
                [pscustomobject]@{ Key = $script:relBase; ValueName = 'Flags';      Type = 4; Data = [uint32]4;  Index = 1 }
                [pscustomobject]@{ Key = $script:relBase; ValueName = '**delvals.'; Type = 1; Data = ' ';        Index = 2 }
            )
            Test-PolDeletionOrder $damaged $script:relBase 'the PolicyServers root key'
            $notes.Count | Should -Be 1
            $notes[0] | Should -BeLike 'DAMAGED registry.pol ordering for the PolicyServers root key:*(Default), Flags*'
            $notes.Clear()
            $healthy = @(
                [pscustomobject]@{ Key = $script:relBase; ValueName = 'Flags';       Type = 4; Data = [uint32]4; Index = 0 }
                [pscustomobject]@{ Key = $script:relBase; ValueName = '**del.Flags'; Type = 1; Data = ' ';       Index = 1 }
                [pscustomobject]@{ Key = $script:relBase; ValueName = 'Flags';       Type = 4; Data = [uint32]0; Index = 2 }
            )
            Test-PolDeletionOrder $healthy $script:relBase 'the PolicyServers root key'
            $notes.Count | Should -Be 0 -Because 'the replacement survives the deletion in between'
        }

        It 'ConvertTo-PolDwordValue accepts a DWORD or a numeric string and refuses text, an empty string and out-of-range numbers' {
            $c = ConvertTo-PolDwordValue ([uint32]4);  $c.Value | Should -Be 4; $c.Value | Should -BeOfType [long]; $c.Reason | Should -BeExactly ''
            $c = ConvertTo-PolDwordValue '4';          $c.Value | Should -Be 4; $c.Reason | Should -BeExactly ''
            $c = ConvertTo-PolDwordValue '0x6';        $c.Value | Should -Be 6 -Because 'a hex string converted before and still does'
            $c = ConvertTo-PolDwordValue ([int]-1);    $c.Value | Should -Be 4294967295 -Because 'Get-GPRegistryValue returns 0xFFFFFFFF as Int32 -1'
            $c = ConvertTo-PolDwordValue ([uint32]::MaxValue); $c.Value | Should -Be 4294967295
            $c = ConvertTo-PolDwordValue ([long]5);    $c.Value | Should -Be 5
            $c = ConvertTo-PolDwordValue 'abc';        $c.Value | Should -BeNullOrEmpty; $c.Reason | Should -BeLike '*text, not a number*'
            $c = ConvertTo-PolDwordValue '';           $c.Value | Should -BeNullOrEmpty; $c.Reason | Should -BeLike '*empty string*'
            $c = ConvertTo-PolDwordValue $null;        $c.Value | Should -BeNullOrEmpty; $c.Reason | Should -BeExactly '' -Because 'an absent value has no reason'
            $c = ConvertTo-PolDwordValue ([long]4294967296); $c.Value | Should -BeNullOrEmpty; $c.Reason | Should -BeLike '*outside the DWORD range*'
            $c = ConvertTo-PolDwordValue '-1';         $c.Value | Should -BeNullOrEmpty; $c.Reason | Should -BeLike '*outside the DWORD range*'
            $c = ConvertTo-PolDwordValue ([byte[]](1, 2)); $c.Value | Should -BeNullOrEmpty; $c.Reason | Should -BeLike '*of type Byte`[`]*'
        }

        It 'Get-RootFlagsDisplay shows the value in hex, (absent) when missing, and "unusable: reason" instead of crashing on text' {
            Get-RootFlagsDisplay (Read-PolRecords -Path $script:PolPath) | Should -BeExactly '0x4'
            Get-RootFlagsDisplay @() | Should -BeExactly '(absent)'
            $text = @([pscustomobject]@{ Key = $script:relBase; ValueName = 'Flags'; Type = 1; Data = 'abc'; Index = 0 })
            Get-RootFlagsDisplay $text | Should -BeLike 'unusable: *text, not a number*'
        }

        It 'Get-PolRawRecord returns the NEWEST raw record by name (case-insensitive), ignores deletion instructions, and $null when none' {
            $recs = @(Read-PolRecords -Path $script:PolPath)
            (Get-PolRawRecord $recs $script:relBase 'flags').Type | Should -Be 4
            [uint32](Get-PolRawRecord $recs $script:relBase 'Flags').Data | Should -Be 4
            $later = $recs + [pscustomobject]@{ Key = $script:relBase; ValueName = 'Flags'; Type = 3; Data = $null; Index = 99 }
            (Get-PolRawRecord $later $script:relBase 'Flags').Type | Should -Be 3 -Because 'the last record in file order is the one a client applies last'
            $deleted = $later + [pscustomobject]@{ Key = $script:relBase; ValueName = '**del.Flags'; Type = 1; Data = ' '; Index = 100 }
            (Get-PolRawRecord $deleted $script:relBase 'Flags') | Should -BeNullOrEmpty -Because 'a later deletion record removes the value on the client, so no record establishes it'
            (Get-PolRawRecord $recs $script:relBase 'Nope') | Should -BeNullOrEmpty
            (Get-PolRawRecord $recs "$script:relBase\other" 'Flags') | Should -BeNullOrEmpty -Because 'the key must match exactly'
            (Get-PolRawRecord @() $script:relBase 'Flags') | Should -BeNullOrEmpty
            # **soft.Flags writes Flags with its declared type ONLY when no Flags exists yet
            # ([MS-GPREG] soft write). With a literal Flags present the soft record is ignored;
            # without one, the soft record is the write path and the raw view returns it for 'Flags'.
            $soft = $recs + [pscustomobject]@{ Key = $script:relBase; ValueName = '**soft.Flags'; Type = 11; Data = $null; Index = 101 }
            (Get-PolRawRecord $soft $script:relBase 'Flags').Type | Should -Be 4 -Because 'a literal Flags exists, so the soft write does not apply'
            $onlySoft = @($recs | Where-Object { "$($_.ValueName)" -ne 'Flags' }) + [pscustomobject]@{ Key = $script:relBase; ValueName = '**soft.Flags'; Type = 11; Data = $null; Index = 101 }
            (Get-PolRawRecord $onlySoft $script:relBase 'Flags').Type | Should -Be 11 -Because 'with no literal Flags the soft record is the write path'
            (Get-PolRawRecord $onlySoft $script:relBase 'Flags').ValueName | Should -BeExactly '**soft.Flags'
        }

        It 'Get-RootFlagsDisplay shows "unusable: ..." for a root Flags record of another type (<Name>) that the parser reads as no value' -TestCases @(
            @{ Name = 'REG_BINARY, 3 bytes';   Type = 3;  Data = [byte[]](1, 2, 3);                 Expect = 'unusable: *type 3 (REG_BINARY)*not REG_DWORD*' }
            @{ Name = 'REG_QWORD, 8 bytes';    Type = 11; Data = [BitConverter]::GetBytes([uint64]4); Expect = 'unusable: *type 11 (REG_QWORD)*not REG_DWORD*' }
            @{ Name = 'REG_MULTI_SZ';          Type = 7;  Data = [byte[]](0x34, 0, 0, 0, 0, 0);     Expect = 'unusable: *type 7 (REG_MULTI_SZ)*not REG_DWORD*' }   # "4" NUL NUL in UTF-16LE
            @{ Name = 'REG_DWORD, 2 bytes';    Type = 4;  Data = [byte[]](1, 0);                    Expect = 'unusable: *REG_DWORD with 2 data byte*not exactly 4*' }
            @{ Name = 'REG_DWORD, 5 bytes';    Type = 4;  Data = [byte[]](4, 0, 0, 0, 0xAA);         Expect = 'unusable: *REG_DWORD with 5 data byte*not exactly 4*' }   # [MS-GPREG]: a DWORD is 32 bits; a longer payload is not decoded from its first 4 bytes
            @{ Name = 'unnamed type 99';       Type = 99; Data = [byte[]](1);                       Expect = 'unusable: *type 99, not REG_DWORD*' }
        ) {
            # Read-PolRecords stores Data = $null for these, so the EFFECTIVE value is $null - the
            # same as an absent value. The display (and the step-0a refusal) must judge the raw
            # record's type instead, or step 3 overwrites a value it never understood with DWORD 0.
            $p = Join-Path $TestDrive "root-flags-type-$Type.pol"
            $pol = [System.Collections.Generic.List[byte]]::new()
            $pol.AddRange($script:PolHeader)
            $pol.AddRange((New-PolRecord -Key $script:relBase -Value 'Flags' -Type $Type -Data $Data))
            [System.IO.File]::WriteAllBytes($p, $pol.ToArray())
            $recs = @(Read-PolRecords -Path $p)
            (Get-PolValue $recs $script:relBase 'Flags') | Should -BeNullOrEmpty -Because 'the effective view cannot tell this record from an absent value'
            Get-PolDwordRecordReason (Get-PolRawRecord $recs $script:relBase 'Flags') | Should -Not -BeNullOrEmpty
            Get-RootFlagsDisplay $recs | Should -BeLike $Expect
            # a later deletion record removes the value on the client: absent, not unusable
            $deleted = $recs + [pscustomobject]@{ Key = $script:relBase; ValueName = '**del.Flags'; Type = 1; Data = ' '; Index = 99 }
            Get-RootFlagsDisplay $deleted | Should -BeExactly '(absent)'
        }

        It 'Get-RootFlagsDisplay shows "unusable: ..." for a lone **soft.Flags record of another type (the same write path as Flags)' {
            # A **soft.<name> record WRITES <name> with its declared type when no value of that name
            # exists yet, so a lone REG_BINARY **soft.Flags lands a REG_BINARY Flags on every client.
            # The parser stores Data = $null for it and the effective view sees no Flags at all.
            $p = Join-Path $TestDrive 'root-flags-soft-binary.pol'
            $pol = [System.Collections.Generic.List[byte]]::new()
            $pol.AddRange($script:PolHeader)
            $pol.AddRange((New-PolRecord -Key $script:relBase -Value '**soft.Flags' -Type 3 -Data ([byte[]](1, 2, 3))))
            [System.IO.File]::WriteAllBytes($p, $pol.ToArray())
            $recs = @(Read-PolRecords -Path $p)
            (Get-PolValue $recs $script:relBase 'Flags') | Should -BeNullOrEmpty -Because 'the effective view cannot tell this record from an absent value'
            (Get-PolRawRecord $recs $script:relBase 'Flags').Type | Should -Be 3
            Get-PolDwordRecordReason (Get-PolRawRecord $recs $script:relBase 'Flags') | Should -BeLike '*type 3 (REG_BINARY)*'
            Get-RootFlagsDisplay $recs | Should -BeLike 'unusable: *type 3 (REG_BINARY)*not REG_DWORD*'
        }

        It 'Get-PolRawRecord applies the [MS-GPREG] soft-write rule: a literal Flags record wins over any **soft.Flags record, whatever the order' {
            # A **soft.<name> record writes only when no value of that name exists. So a literal
            # Flags decides the type even when a **soft.Flags follows it, and a **soft.Flags that
            # precedes a literal Flags does not.
            $dword  = [pscustomobject]@{ Key = $script:relBase; ValueName = 'Flags';        Type = 4; Data = [uint32]4; Index = 0 }
            $softB  = [pscustomobject]@{ Key = $script:relBase; ValueName = '**soft.Flags'; Type = 3; Data = $null;     Index = 1 }
            $binary = [pscustomobject]@{ Key = $script:relBase; ValueName = 'Flags';        Type = 3; Data = $null;     Index = 0 }
            $softD  = [pscustomobject]@{ Key = $script:relBase; ValueName = '**soft.Flags'; Type = 4; Data = [uint32]4; Index = 1 }
            # literal DWORD then soft BINARY: the soft write does not apply, the value stays usable
            (Get-PolRawRecord @($dword, $softB) $script:relBase 'Flags').ValueName | Should -BeExactly 'Flags'
            Get-PolDwordRecordReason (Get-PolRawRecord @($dword, $softB) $script:relBase 'Flags') | Should -BeExactly ''
            Get-RootFlagsDisplay @($dword, $softB) | Should -BeExactly '0x4'
            # literal BINARY then soft DWORD: the binary value exists, so the soft DWORD never applies
            (Get-PolRawRecord @($binary, $softD) $script:relBase 'Flags').ValueName | Should -BeExactly 'Flags'
            Get-PolDwordRecordReason (Get-PolRawRecord @($binary, $softD) $script:relBase 'Flags') | Should -BeLike '*type 3 (REG_BINARY)*'
            Get-RootFlagsDisplay @($binary, $softD) | Should -BeLike 'unusable: *type 3 (REG_BINARY)*'
            # soft BINARY then literal DWORD: the literal record decides
            (Get-PolRawRecord @($softB, $dword) $script:relBase 'Flags').ValueName | Should -BeExactly 'Flags'
            Get-RootFlagsDisplay @($softB, $dword) | Should -BeExactly '0x4'
            # two soft writes: the FIRST establishes the value, the second does nothing
            (Get-PolRawRecord @($softD, $softB) $script:relBase 'Flags').Type | Should -Be 4 -Because 'the first soft write established a DWORD'
            Get-RootFlagsDisplay @($softD, $softB) | Should -BeExactly '0x4'
            (Get-PolRawRecord @($softB, $softD) $script:relBase 'Flags').Type | Should -Be 3 -Because 'the first soft write established a REG_BINARY the second cannot replace'
            Get-RootFlagsDisplay @($softB, $softD) | Should -BeLike 'unusable: *type 3 (REG_BINARY)*'
            # a deletion between two writes: the value after the deletion is the one that counts
            $del = [pscustomobject]@{ Key = $script:relBase; ValueName = '**del.Flags'; Type = 1; Data = ' '; Index = 5 }
            (Get-PolRawRecord @($binary, $del, $softD) $script:relBase 'Flags').Type | Should -Be 4 -Because 'the deletion made the value absent, so the soft DWORD applied'
            Get-RootFlagsDisplay @($binary, $del, $softD) | Should -BeExactly '0x4'
            (Get-PolRawRecord @($dword, $del) $script:relBase 'Flags') | Should -BeNullOrEmpty
            Get-RootFlagsDisplay @($dword, $del) | Should -BeExactly '(absent)'
            # **DeleteKeys removed the KEY: a soft write does not recreate it, a literal write does
            $delKeys = [pscustomobject]@{ Key = ''; ValueName = '**DeleteKeys'; Type = 1; Data = $script:relBase; Index = 5 }
            (Get-PolRawRecord @($dword, $delKeys, $softB) $script:relBase 'Flags') | Should -BeNullOrEmpty -Because 'the key is gone and a soft write cannot recreate it'
            Get-RootFlagsDisplay @($dword, $delKeys, $softB) | Should -BeExactly '(absent)'
            (Get-PolRawRecord @($binary, $delKeys, $dword, $softB) $script:relBase 'Flags').Type | Should -Be 4 -Because 'the literal write recreated the key and the value'
            Get-RootFlagsDisplay @($binary, $delKeys, $dword, $softB) | Should -BeExactly '0x4'
            # a soft write that restores a deleted value clears its "deleted" mark for the ordering scan
            $eff = Get-PolEffectiveValues @($dword, $del, $softD) $script:relBase
            [uint32]$eff.Values['Flags'] | Should -Be 4
            $eff.Deleted.ContainsKey('Flags') | Should -BeFalse -Because 'the value survives, so the deletion-order scan must not report it as lost'
            $eff = Get-PolEffectiveValues @($dword, $del) $script:relBase
            $eff.Deleted.ContainsKey('Flags') | Should -BeTrue
            # a plain write under a DESCENDANT key recreates the deleted parent (missing ancestors
            # are created), so a soft write into the parent applies again; an instruction under
            # the descendant creates nothing
            $childWrite = [pscustomobject]@{ Key = "$script:relBase\$script:leaf"; ValueName = 'URL';        Type = 1; Data = 'https://x/'; Index = 6 }
            $childInstr = [pscustomobject]@{ Key = "$script:relBase\$script:leaf"; ValueName = '**del.URL';  Type = 1; Data = ' ';          Index = 6 }
            (Get-PolRawRecord @($dword, $delKeys, $childWrite, $softD) $script:relBase 'Flags').Type | Should -Be 4 -Because 'the child write recreated the parent key'
            Get-RootFlagsDisplay @($dword, $delKeys, $childWrite, $softD) | Should -BeExactly '0x4'
            (Get-PolRawRecord @($dword, $delKeys, $childWrite, $softB) $script:relBase 'Flags').Type | Should -Be 3 -Because 'the soft BINARY applied into the recreated key and must be refused'
            (Get-PolRawRecord @($dword, $delKeys, $childInstr, $softD) $script:relBase 'Flags') | Should -BeNullOrEmpty -Because 'an instruction under the child creates no key'
        }

        It 'Get-PolDwordRecordReason accepts $null, a REG_SZ and a complete REG_DWORD (ConvertTo-PolDwordValue judges those)' {
            Get-PolDwordRecordReason $null | Should -BeExactly ''
            Get-PolDwordRecordReason ([pscustomobject]@{ Key = $script:relBase; ValueName = 'Flags'; Type = 1; Data = 'abc'; Index = 0 }) | Should -BeExactly '' -Because 'text is refused by the value converter, not by the type check'
            Get-PolDwordRecordReason ([pscustomobject]@{ Key = $script:relBase; ValueName = 'Flags'; Type = 4; Data = [uint32]4; Index = 0 }) | Should -BeExactly ''
            Get-PolDwordRecordReason ([pscustomobject]@{ Key = $script:relBase; ValueName = 'Flags'; Type = 4; Data = $null; Index = 0 }) | Should -BeLike '*REG_DWORD with*data byte*not exactly 4*'
        }

        It 'PolicyID derivation matches the published Java String.hashCode values' {
            # 'hello' -> 99162322 is the value every Java reference gives; the EJBCA MSAE default
            # policy name wraps negative. Both run through the script's own extracted block.
            script:Get-DerivedPolicyId -PolicyName 'hello' -PolicyId ''             | Should -BeExactly '99162322'
            script:Get-DerivedPolicyId -PolicyName 'EJBCA MSAE Policy' -PolicyId '' | Should -BeExactly '-1941035357'
            script:Get-DerivedPolicyId -PolicyName 'hello' -PolicyId 'given'        | Should -BeExactly 'given' -Because 'an explicit -PolicyId is never overwritten'
        }
    }

    Context 'Unit: write and verification helpers (mocked cmdlets)' -Tag 'Unit' {

        BeforeAll {
            # Test-GpoEntry splats $wr and calls Get-GPRegistryValue. Pester needs a command of that
            # name to exist before it can be mocked; this stub lives only in this Context's scope,
            # so the Lab tier further down sees the real cmdlet again. (The mock itself is a
            # script-scope alias, which wins over any function, so the stub never shadows it.)
            function Get-GPRegistryValue { param($Guid, $Key, $Domain, $Server) throw 'stub: Get-GPRegistryValue must be mocked' }
            $script:wr = @{}
        }

        It 'Invoke-GPWrite retries a sharing violation twice with a back-off, then rethrows' {
            Mock Start-Sleep { }
            $calls = New-Object System.Collections.Generic.List[int]
            { Invoke-GPWrite { $calls.Add(1); throw 'The process cannot access the file (0x80070020)' } } | Should -Throw -ExpectedMessage '*0x80070020*'
            $calls.Count | Should -Be 3 -Because 'three attempts in total'
            Should -Invoke Start-Sleep -Times 2 -Exactly
            Should -Invoke Start-Sleep -Times 1 -Exactly -ParameterFilter { $Milliseconds -eq 400 }
            Should -Invoke Start-Sleep -Times 1 -Exactly -ParameterFilter { $Milliseconds -eq 800 }
        }

        It 'Invoke-GPWrite treats "was not found" as success only with -TolerateNotFound' {
            Mock Start-Sleep { }
            $calls = New-Object System.Collections.Generic.List[int]
            { Invoke-GPWrite -TolerateNotFound { $calls.Add(1); throw 'The value was not found' } } | Should -Not -Throw
            $calls.Count | Should -Be 1
            $calls.Clear()
            { Invoke-GPWrite { $calls.Add(1); throw 'The value was not found' } } | Should -Throw -ExpectedMessage '*was not found*'
            $calls.Count | Should -Be 1 -Because 'not a transient error: no retry'
            Should -Invoke Start-Sleep -Times 0 -Exactly
        }

        It 'Invoke-GPWrite rethrows any other error after ONE attempt, and returns after a first-try success' {
            Mock Start-Sleep { }
            $calls = New-Object System.Collections.Generic.List[int]
            { Invoke-GPWrite { $calls.Add(1); throw 'simulated: something else' } } | Should -Throw -ExpectedMessage '*something else*'
            $calls.Count | Should -Be 1
            $calls.Clear()
            { Invoke-GPWrite { $calls.Add(1) } } | Should -Not -Throw
            $calls.Count | Should -Be 1
            Should -Invoke Start-Sleep -Times 0 -Exactly
        }

        It 'Test-GpoEntry names every missing value in its error' {
            Mock Get-GPRegistryValue { [pscustomobject]@{ ValueName = 'URL'; Value = 'LDAP:' } }
            { Test-GpoEntry 'HKLM\X' @{ URL = 'LDAP:'; PolicyID = '{G}'; Cost = [uint32]1 } } | Should -Throw -ExpectedMessage '*Post-write verification failed for value(s)*'
            try { Test-GpoEntry 'HKLM\X' @{ URL = 'LDAP:'; PolicyID = '{G}'; Cost = [uint32]1 } } catch { $msg = $_.Exception.Message }
            $msg | Should -BeLike '*PolicyID*'
            $msg | Should -BeLike '*Cost*'
            $msg | Should -Not -BeLike '*URL*under*' -Because 'a value that matches is not listed'
        }

        It 'Test-GpoEntry compares DWORDs bit-exactly (expected 4294967295, live Int32 -1) and accepts a [long] on either side' {
            Mock Get-GPRegistryValue {
                [pscustomobject]@{ ValueName = 'Cost';  Value = [int]-1 }
                [pscustomobject]@{ ValueName = 'Flags'; Value = [long]0x14 }
                [pscustomobject]@{ ValueName = 'AuthFlags'; Value = [int]2 }
            }
            { Test-GpoEntry 'HKLM\X' @{ Cost = [uint32]4294967295; Flags = [int]0x14; AuthFlags = [long]2 } } | Should -Not -Throw
            { Test-GpoEntry 'HKLM\X' @{ Cost = [uint32]4294967294 } } | Should -Throw -ExpectedMessage '*Cost*'
            { Test-GpoEntry 'HKLM\X' @{ Flags = [int]0x15 } } | Should -Throw -ExpectedMessage '*Flags*'
        }

        It 'Test-GpoEntry refuses a live [long] a DWORD cannot hold, and a string that differs in case' {
            Mock Get-GPRegistryValue {
                [pscustomobject]@{ ValueName = 'Flags'; Value = [long]4294967296 }
                [pscustomobject]@{ ValueName = 'URL';   Value = 'ldap:' }
            }
            { Test-GpoEntry 'HKLM\X' @{ Flags = [int]0 } } | Should -Throw -ExpectedMessage '*Flags*' -Because 'the conversion to UInt32 fails and counts as a mismatch'
            { Test-GpoEntry 'HKLM\X' @{ URL = 'LDAP:' } } | Should -Throw -ExpectedMessage '*URL*' -Because 'strings compare case-sensitively'
            { Test-GpoEntry 'HKLM\X' @{ URL = 'ldap:' } } | Should -Not -Throw
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

        It 'carries a PSScriptInfo header (Test-ScriptFileInfo parses it; Version is semver)' {
            $info = Test-ScriptFileInfo -Path $script:Gpo -ErrorAction Stop
            $info.Version | Should -Match '^\d+\.\d+\.\d+$'
            $info.Guid    | Should -Not -BeNullOrEmpty
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

    # The two Guard contexts need the GroupPolicy module (the script imports it before the
    # guards run). They are discovered only when it is present: a discovery condition, not
    # -Skip, because CI fails on a skipped test.
    if ($script:HasGP) {

    Context 'Guard: parameter-conflict validation (throws before Get-GPO)' -Tag 'Guard' {

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

    # -------------------------------------------------------------------------------------------
    # Guard tier with a MOCKED Get-GPO: the mock is defined in this file's script scope (as a
    # Pester alias), and the script - invoked with & from a test - resolves commands through this
    # scope chain before it reaches the GroupPolicy module, on both engines. The fake GPO carries
    # literal values only: $script: inside a mock body resolves to the INVOKED script's scope.
    # The "proceeds" cases run with -WhatIf -SkipADPolicy -Server localhost, so no AD lookup, no
    # write, and a registry.pol path that fails Test-Path locally at once.
    # -------------------------------------------------------------------------------------------
    Context 'Guard: GPO resolution and input validation against a mocked Get-GPO' -Tag 'Guard' {

        BeforeAll {
            Import-Module GroupPolicy -ErrorAction Stop -WarningAction SilentlyContinue
            $script:FakeId = [guid]'11111111-2222-3333-4444-555555555555'
            $script:Proceed = @{ SkipADPolicy = $true; Server = 'localhost'; WhatIf = $true }
        }

        It 'refuses to reinterpret a GUID-shaped name as an ID when a GPO with that display name exists but the name lookup failed' {
            Mock Get-GPO {
                if ($Name) { throw 'simulated: RPC server unavailable' }
                if ($All)  { return [pscustomobject]@{ DisplayName = '11111111-2222-3333-4444-555555555555'; Id = [guid]'11111111-2222-3333-4444-555555555555'; DomainName = 'pester.invalid' } }
                throw 'unexpected Get-GPO call'
            }
            { & $script:Gpo -GpoName $script:FakeId.ToString() -Url 'https://y/' -PolicyName 'z' } | Should -Throw -ExpectedMessage '*refusing to reinterpret*'
            Should -Invoke Get-GPO -Times 0 -Exactly -ParameterFilter { $null -ne $Guid -and $Guid -ne [guid]::Empty } -Because 'the GUID must never be tried'
        }

        It 'reports that neither a GPO with that name nor one with that ID exists' {
            Mock Get-GPO {
                if ($Name) { throw 'simulated: not found by name' }
                if ($All)  { return @() }
                throw 'simulated: not found by ID'
            }
            { & $script:Gpo -GpoName $script:FakeId.ToString() -Url 'https://y/' -PolicyName 'z' } | Should -Throw -ExpectedMessage '*no GPO has that ID either*'
        }

        It 'rethrows the original name-lookup error when the name is not GUID-shaped' {
            Mock Get-GPO { throw 'simulated: access is denied to the GPO' }
            { & $script:Gpo -GpoName 'PKI - Not A Guid' -Url 'https://y/' -PolicyName 'z' } | Should -Throw -ExpectedMessage '*access is denied to the GPO*'
            Should -Invoke Get-GPO -Times 1 -Exactly
        }

        It 'falls back to the GPO ID when no GPO carries the display name, and proceeds' {
            Mock Get-GPO {
                if ($Name) { throw 'simulated: not found by name' }
                if ($All)  { return @() }
                if ($Guid) { return [pscustomobject]@{ DisplayName = 'PESTER fake'; Id = [guid]'11111111-2222-3333-4444-555555555555'; DomainName = 'pester.invalid' } }
                throw 'unexpected Get-GPO call'
            }
            $o = & $script:Gpo -GpoName $script:FakeId.ToString() -Url 'https://y/' -PolicyName 'z' @script:Proceed 3>$null
            $o.Mode  | Should -BeExactly 'Add'
            $o.GpoId | Should -Be $script:FakeId
            $o.Gpo   | Should -BeExactly 'PESTER fake'
            $o.EntryApplied | Should -BeFalse -Because '-WhatIf writes nothing'
            Should -Invoke Get-GPO -Times 1 -Exactly -ParameterFilter { $Guid -eq [guid]'11111111-2222-3333-4444-555555555555' }
        }

        It 'validates the inputs after the GPO is resolved: URL scheme, control characters, and the whitespace warning' {
            Mock Get-GPO { [pscustomobject]@{ DisplayName = 'PESTER fake'; Id = [guid]'11111111-2222-3333-4444-555555555555'; DomainName = 'pester.invalid' } }
            { & $script:Gpo -GpoName 'PESTER fake' -Url 'ftp://x/' -PolicyName 'z' } | Should -Throw -ExpectedMessage '*absolute http/https URI*'
            { & $script:Gpo -GpoName 'PESTER fake' -Url "https://x/a`tb" -PolicyName 'z' } | Should -Throw -ExpectedMessage 'Url contains control characters.'
            { & $script:Gpo -GpoName 'PESTER fake' -Url 'https://x/' -PolicyName "a`nb" } | Should -Throw -ExpectedMessage 'PolicyName contains control characters.'
            { & $script:Gpo -GpoName 'PESTER fake' -Url 'https://x/' -PolicyName 'z' -EnableAutoEnrollmentPolicy -AEStore "a`0b" } | Should -Throw -ExpectedMessage 'AEStore contains control characters.'
            # a control character in the URL is refused in Remove mode too (before the http check, which Remove skips)
            { & $script:Gpo -GpoName 'PESTER fake' -Url "https://x/a`tb" -Remove } | Should -Throw -ExpectedMessage 'Url contains control characters.'
            $w = $null
            $o = & $script:Gpo -GpoName 'PESTER fake' -Url 'https://x/' -PolicyName ' z ' @script:Proceed -WarningAction SilentlyContinue -WarningVariable w
            @($w | Where-Object { "$_" -like '*leading/trailing whitespace*' }).Count | Should -Be 1
            $o.FriendlyName | Should -BeExactly ' z ' -Because 'the name is used verbatim'
            $o.PolicyID | Should -BeExactly '34566' -Because 'the hash covers the spaces as well: 32*31^2 + 122*31 + 32'
        }

        # Step 0a against a MOCKED registry.pol reader: Read-PolRecords is the script's own
        # function, but the Pester alias wins over it (alias before function in command lookup),
        # so the script reads the records the mock returns - the same record shapes the real
        # parser produces for these values. The run is NOT -WhatIf: a refusal that came too late
        # would reach the mocked Set-GPRegistryValue, and that mock counts every call.
        It 'refuses an existing root Flags record of another type (<Name>) BEFORE any write' -TestCases @(
            @{ Name = 'REG_BINARY';        Type = 3;  ValueName = 'Flags';        Expect = '*root Flags*not a usable DWORD*type 3 (REG_BINARY)*Nothing was written*' }
            @{ Name = 'REG_QWORD';         Type = 11; ValueName = 'Flags';        Expect = '*root Flags*not a usable DWORD*type 11 (REG_QWORD)*Nothing was written*' }
            @{ Name = 'short REG_DWORD';   Type = 4;  ValueName = 'Flags';        Expect = '*root Flags*not a usable DWORD*REG_DWORD with*data byte*not exactly 4*Nothing was written*' }
            # **soft.Flags writes Flags with its declared type when no Flags exists yet: the same
            # write path, judged the same way.
            @{ Name = '**soft.Flags REG_BINARY'; Type = 3; ValueName = '**soft.Flags'; Expect = '*root Flags*not a usable DWORD*type 3 (REG_BINARY)*Nothing was written*' }
        ) {
            Mock Get-GPO { [pscustomobject]@{ DisplayName = 'PESTER fake'; Id = [guid]'11111111-2222-3333-4444-555555555555'; DomainName = 'pester.invalid' } }
            Mock Set-GPRegistryValue { throw 'simulated: Set-GPRegistryValue must not be reached' }
            # literal values only, plus ONE environment variable for the record type: the mock body
            # runs in the invoked script's scope, where the test case's $Type is not a variable to
            # rely on
            Mock Read-PolRecords {
                @(
                    [pscustomobject]@{ Key = 'Software\Policies\Microsoft\Cryptography\PolicyServers'; ValueName = "$env:PESTER_MOCK_ROOT_FLAGS_NAME"; Type = [uint32]$env:PESTER_MOCK_ROOT_FLAGS_TYPE; Data = $null; Index = 0 }
                    # an unrelated instruction after it: a **Comment: is not a value and changes nothing
                    [pscustomobject]@{ Key = 'Software\Policies\Microsoft\Cryptography\PolicyServers'; ValueName = '**Comment:pester'; Type = 1; Data = ' '; Index = 1 }
                )
            }
            $env:PESTER_MOCK_ROOT_FLAGS_TYPE = "$Type"
            $env:PESTER_MOCK_ROOT_FLAGS_NAME = $ValueName
            try {
                { & $script:Gpo -GpoName 'PESTER fake' -Url 'https://x/' -PolicyName 'z' -SkipADPolicy -Server localhost -Confirm:$false 3>$null } |
                    Should -Throw -ExpectedMessage $Expect
            } finally {
                Remove-Item -Path Env:\PESTER_MOCK_ROOT_FLAGS_TYPE -ErrorAction SilentlyContinue
                Remove-Item -Path Env:\PESTER_MOCK_ROOT_FLAGS_NAME -ErrorAction SilentlyContinue
            }
            Should -Invoke Read-PolRecords -Times 1 -Exactly -Because 'the pre-write read is the only read before the refusal'
            Should -Invoke Set-GPRegistryValue -Times 0 -Exactly -Because 'the refusal comes before the first write'
        }

        It 'a root Flags record the parser reads as no value is refused, but a genuinely absent value proceeds' {
            Mock Get-GPO { [pscustomobject]@{ DisplayName = 'PESTER fake'; Id = [guid]'11111111-2222-3333-4444-555555555555'; DomainName = 'pester.invalid' } }
            Mock Read-PolRecords { @() }
            $o = & $script:Gpo -GpoName 'PESTER fake' -Url 'https://x/' -PolicyName 'z' @script:Proceed 3>$null
            $o.Mode | Should -BeExactly 'Add'
            $o.RootFlags | Should -BeLike '(absent)*' -Because 'no record at all is an absent value, not an unusable one'
        }
    }

    }   # if ($script:HasGP): end of the two Guard contexts

    # -------------------------------------------------------------------------------------------
    # Lab tier: live round-trip inside ONE throwaway, never-linked GPO. Opt-in (-RunLab).
    # Tests are SEQUENTIAL: add -> verify -> root Flags -> -Server/-Domain -> default+AE ->
    # AE percent (rewritten, then restored to 10%) -> entry switches (-Authentication,
    # -NoAutoEnroll, -NoClientId, -AllowUntrustedIssuer: each on its OWN fresh entry that the case
    # removes again, so the main entry keeps its state) -> redundant endpoint -> replace-sibling
    # -> remove -> -ClearDefault -> double remove -> User scope, mirroring a real rollout and
    # teardown inside the same GPO.
    # -------------------------------------------------------------------------------------------
    Context 'Lab: live GPO round-trip (throwaway unlinked GPO)' -Tag 'Lab' -Skip:(-not $script:GpoLabReady) {

        BeforeAll {
            # Prefix FIRST - before anything that can throw. AfterAll runs even when BeforeAll
            # dies, and its backstop sweep must never see an unset (= match-everything) prefix.
            $script:Prefix = "PESTER-$([guid]::NewGuid().ToString('N').Substring(0,8))"
            Import-Module GroupPolicy -ErrorAction Stop

            # ONE throwaway GPO, created UNLINKED (New-GPO without New-GPLink): it applies to zero
            # computers/users, so nothing this tier authors can reach a client. Tracked by GUID.
            $script:LabGpo = New-GPO -Name "$script:Prefix CEP" -Comment 'Pester throwaway - safe to delete'
            $script:LabGpo | Should -Not -BeNullOrEmpty

            # .invalid is RFC 2606-reserved - the URL can never reach a real endpoint.
            $script:UrlBase = "https://$($script:Prefix.ToLower()).lab.invalid/ejbca/msae/CEPService"
            $script:LabUrl  = "$script:UrlBase`?alias"
            $script:LabName = "$script:Prefix Policy"

            # Independent oracles (reference implementations, not the script's code).
            $sha1 = [System.Security.Cryptography.SHA1]::Create()
            $script:ExpectedKey = -join ($sha1.ComputeHash([System.Text.Encoding]::Unicode.GetBytes($script:LabUrl.ToLowerInvariant())) |
                                         ForEach-Object { $_.ToString('x2') })
            $h = [int64]0
            foreach ($c in $script:LabName.ToCharArray()) { $h = ($h * 31 + [int64]$c) -band 4294967295 }
            if ($h -ge 2147483648) { $h -= 4294967296 }
            $script:ExpectedPid = "$h"

            # Expected AD-row PolicyID: the domain object's objectGUID, as the script formats it.
            $dn = ([ADSI]'LDAP://RootDSE').defaultNamingContext.Value
            $script:ExpectedAdPid = '{' + (New-Object Guid (, ([byte[]]([ADSI]"LDAP://$dn").Properties['objectGUID'][0]))).ToString().ToUpper() + '}'

            $script:AdKey     = '37c9dc30f207f27f61a2f7c3aed598a6e2920b54'
            $script:GpoParams = @{ GpoName = "$script:Prefix CEP" }
            $script:RootKey   = 'HKLM\SOFTWARE\Policies\Microsoft\Cryptography\PolicyServers'
            $sysvol = "\\$($script:LabGpo.DomainName)\SYSVOL\$($script:LabGpo.DomainName)\Policies\{$($script:LabGpo.Id)}"
            $script:PolMachine = "$sysvol\Machine\registry.pol"
            $script:PolUser    = "$sysvol\User\registry.pol"
            $script:AeKey      = 'HKLM\SOFTWARE\Policies\Microsoft\Cryptography\AutoEnrollment'
            $script:relAe      = 'Software\Policies\Microsoft\Cryptography\AutoEnrollment'

            # Read this GPO's REAL registry.pol with the extracted parser, with a short retry in
            # case SYSVOL is a beat behind the GroupPolicy cmdlets (same-box PDC: normally instant).
            function script:Read-LabPol {
                param([string]$Path, [int]$MinRecords = 1, [int]$TimeoutSec = 10)
                $deadline = (Get-Date).AddSeconds($TimeoutSec)
                do {
                    $recs = @(Read-PolRecords -Path $Path)
                    if ($recs.Count -ge $MinRecords) { return $recs }
                    Start-Sleep -Milliseconds 500
                } while ((Get-Date) -lt $deadline)
                return @(Read-PolRecords -Path $Path)
            }
            function script:Get-LabKey([string]$Url) {
                $sha1 = [System.Security.Cryptography.SHA1]::Create()
                -join ($sha1.ComputeHash([System.Text.Encoding]::Unicode.GetBytes($Url.ToLowerInvariant())) | ForEach-Object { $_.ToString('x2') })
            }
            # One entry-switch round-trip on its OWN fresh entry: a second URL under the run's
            # .invalid host (suffix = the case), so the main entry keeps the state the earlier cases
            # left. Adds the entry with the extra parameters, captures it through the GPMC view AND
            # the registry.pol replay, then removes it again BEFORE the caller asserts - so the entry
            # is gone whatever the assertions decide. It only ever lives inside the throwaway GPO.
            function script:Invoke-LabSwitchRoundTrip([string]$Suffix, [hashtable]$Extra) {
                $url  = "$script:UrlBase`?$Suffix"
                $name = "$script:Prefix $Suffix"
                $key  = script:Get-LabKey $url
                $summary = & $script:Gpo @script:GpoParams @Extra -Url $url -PolicyName $name -Confirm:$false 3>$null
                $gpmc = @{}
                foreach ($v in @(Get-GPRegistryValue -Guid $script:LabGpo.Id -Key "$script:RootKey\$key")) { $gpmc[$v.ValueName] = $v.Value }
                $recs = script:Read-LabPol -Path $script:PolMachine
                $pol  = (Get-PolEffectiveValues $recs "$script:relBase\$key").Values
                $keysWith = @((Get-PolEntries $recs).Key)
                $removed  = & $script:Gpo @script:GpoParams -Url $url -Remove -Confirm:$false 3>$null
                [pscustomobject]@{
                    Key = $key; Url = $url; Summary = $summary; Gpmc = $gpmc; Pol = $pol; KeysWith = $keysWith
                    Removed   = $removed
                    KeysAfter = @((Get-PolEntries (script:Read-LabPol -Path $script:PolMachine)).Key)
                }
            }
        }

        AfterAll {
            # Surgical: delete ONLY the throwaway GPO, by its exact tracked GUID.
            if ($script:LabGpo) {
                try { Remove-GPO -Guid $script:LabGpo.Id -Confirm:$false } catch { Write-Warning "Failed to remove lab GPO $($script:LabGpo.Id): $_" }
            }
            # Backstop, scoped to THIS run's fresh-GUID prefix: REPORT ONLY. A prefix is not proof
            # that this run created a GPO, so a leftover is named for manual review, never deleted
            # here; the only deletion above is by the GUID New-GPO returned to this run.
            # STRUCTURAL GUARD: the sweep runs only when the prefix has its full PESTER-<hex8>
            # shape - an unset/empty prefix would otherwise degenerate the filter to -like "*".
            # Never widen this.
            if ($script:Prefix -match '^PESTER-[0-9a-f]{8}$') {
                $ownId = if ($script:LabGpo) { "$($script:LabGpo.Id)" } else { '' }
                foreach ($g in @(Get-GPO -All | Where-Object { $_.DisplayName -like "$script:Prefix*" -and "$($_.Id)" -ne $ownId })) {
                    Write-Warning "AfterAll backstop found a GPO with this run's prefix that this run did not create: $($g.DisplayName) ($($g.Id)). Review and remove it manually."
                }
            } else {
                Write-Warning "Backstop sweep skipped: run prefix is unset or malformed ('$script:Prefix')."
            }
        }

        It 'Add authors the CEP entry and the AD policy row into the GPO' {
            $o = & $script:Gpo @script:GpoParams -Url $script:LabUrl -PolicyName $script:LabName -Confirm:$false 3>$null
            $o.EntryApplied | Should -BeTrue
            $o.ADPolicyRow  | Should -BeExactly 'applied'
            $o.GpoId        | Should -Be $script:LabGpo.Id
            $o.Key          | Should -BeExactly "$script:RootKey\$script:ExpectedKey"
            $o.RootFlags    | Should -BeExactly '0x0 (applied)' -Because 'a fresh GPO has no root Flags; the first run writes 0'
        }

        It 'the GPMC API (Get-GPRegistryValue) sees every authored value' {
            $vals = Get-GPRegistryValue -Guid $script:LabGpo.Id -Key "$script:RootKey\$script:ExpectedKey"
            $map = @{}; foreach ($v in $vals) { $map[$v.ValueName] = $v.Value }
            $map['URL']          | Should -BeExactly $script:LabUrl
            $map['PolicyID']     | Should -BeExactly $script:ExpectedPid
            $map['FriendlyName'] | Should -BeExactly $script:LabName
            [int]$map['Flags']     | Should -Be 0x14
            [int]$map['AuthFlags'] | Should -Be 0x2
        }

        It 'the REAL registry.pol parses with the extracted parser: entry, AD row, and Cost round-trip' {
            $recs = script:Read-LabPol -Path $script:PolMachine -MinRecords 12   # 2 entries x 6 values
            $entries = @(Get-PolEntries $recs)
            $entries.Count | Should -Be 2
            ($entries | Where-Object Key -eq $script:ExpectedKey).PolicyID | Should -BeExactly $script:ExpectedPid
            $ad = $entries | Where-Object Key -eq $script:AdKey
            $ad.URL      | Should -BeExactly 'LDAP:'
            $ad.PolicyID | Should -BeExactly $script:ExpectedAdPid
            # Cost survives as full-range unsigned DWORDs in the .pol data.
            [uint32](Get-PolValue $recs "$script:relBase\$script:ExpectedKey" 'Cost') | Should -Be ([uint32]0x7FFFFFFD)
            [uint32](Get-PolValue $recs "$script:relBase\$script:AdKey" 'Cost')       | Should -Be ([uint32]4294967295)
        }

        It '-DisableUserConfigured sets root Flags 0x4, a plain rerun preserves it, -EnableUserConfigured clears it' {
            $o = & $script:Gpo @script:GpoParams -Url $script:LabUrl -PolicyName $script:LabName -DisableUserConfigured -Confirm:$false 3>$null
            $o.RootFlags | Should -BeExactly '0x4 (applied)'
            [uint32](Get-PolValue (script:Read-LabPol -Path $script:PolMachine) $script:relBase 'Flags') | Should -Be 4
            $o = & $script:Gpo @script:GpoParams -Url $script:LabUrl -PolicyName $script:LabName -Confirm:$false 3>$null
            $o.RootFlags | Should -BeExactly '0x4 (unchanged)' -Because 'a rerun without the switch keeps the bit'
            $o = & $script:Gpo @script:GpoParams -Url $script:LabUrl -PolicyName $script:LabName -EnableUserConfigured -Confirm:$false 3>$null
            $o.RootFlags | Should -BeExactly '0x0 (applied)'
            [uint32](Get-PolValue (script:Read-LabPol -Path $script:PolMachine) $script:relBase 'Flags') | Should -Be 0
        }

        It 'a root Flags bit 0x2 authored by another tool is cleared with a warning; the other bits are preserved' {
            Set-GPRegistryValue -Guid $script:LabGpo.Id -Key $script:RootKey -ValueName Flags -Type DWord -Value 6 | Out-Null
            $w = $null
            $o = & $script:Gpo @script:GpoParams -Url $script:LabUrl -PolicyName $script:LabName -Confirm:$false -WarningAction SilentlyContinue -WarningVariable w
            $o.RootFlags | Should -BeExactly '0x4 (applied)'
            @($w | Where-Object { "$_" -like '*bit 0x2 set*' }).Count | Should -Be 1
            [uint32](Get-PolValue (script:Read-LabPol -Path $script:PolMachine) $script:relBase 'Flags') | Should -Be 4
        }

        It 'refuses to run when the existing root Flags is text, BEFORE any write; -Remove still reports it' {
            Set-GPRegistryValue -Guid $script:LabGpo.Id -Key $script:RootKey -ValueName Flags -Type String -Value 'abc' | Out-Null
            (Get-PolValue (script:Read-LabPol -Path $script:PolMachine) $script:relBase 'Flags') | Should -BeExactly 'abc'
            { & $script:Gpo @script:GpoParams -Url "$script:UrlBase`?never" -PolicyName "$script:LabName never" -Confirm:$false 3>$null } |
                Should -Throw -ExpectedMessage '*root Flags*not a usable DWORD*Nothing was written*'
            (Get-PolEntries (script:Read-LabPol -Path $script:PolMachine)).Key | Should -Not -Contain (script:Get-LabKey "$script:UrlBase`?never") -Because 'the refusal happens before the CEP entry write'
            $o = & $script:Gpo @script:GpoParams -Url "$script:UrlBase`?never" -Remove -Confirm:$false 3>$null
            $o.RemovedEntry | Should -BeFalse
            $o.RootFlags    | Should -BeLike 'unusable: *' -Because 'the summary shows the problem instead of crashing'
            @($o.Notes) -match 'nothing to remove' | Should -Not -BeNullOrEmpty
            # repair: back to a DWORD 4, as the previous test left it
            Set-GPRegistryValue -Guid $script:LabGpo.Id -Key $script:RootKey -ValueName Flags -Type DWord -Value 4 | Out-Null
            $o = & $script:Gpo @script:GpoParams -Url $script:LabUrl -PolicyName $script:LabName -Confirm:$false 3>$null
            $o.RootFlags | Should -BeExactly '0x4 (unchanged)'
        }

        It 'refuses to run when the existing root Flags is a REG_BINARY (a type the parser reads as no value), BEFORE any write' {
            Set-GPRegistryValue -Guid $script:LabGpo.Id -Key $script:RootKey -ValueName Flags -Type Binary -Value ([byte[]](1, 2, 3)) | Out-Null
            $recs = script:Read-LabPol -Path $script:PolMachine
            (Get-PolRawRecord $recs $script:relBase 'Flags').Type | Should -Be 3 -Because 'GPMC writes the value as a REG_BINARY record'
            (Get-PolValue $recs $script:relBase 'Flags') | Should -BeNullOrEmpty -Because 'the effective view cannot tell it from an absent value'
            $neverUrl = "$script:UrlBase`?never-binary"
            { & $script:Gpo @script:GpoParams -Url $neverUrl -PolicyName "$script:LabName never" -Confirm:$false 3>$null } |
                Should -Throw -ExpectedMessage '*root Flags*not a usable DWORD*type 3 (REG_BINARY)*Nothing was written*'
            $recs = script:Read-LabPol -Path $script:PolMachine
            (Get-PolEntries $recs).Key | Should -Not -Contain (script:Get-LabKey $neverUrl) -Because 'the refusal happens before the CEP entry write'
            Test-EntryRecordsPresent $recs "$script:relBase\$(script:Get-LabKey $neverUrl)" | Should -BeFalse -Because 'no record of the entry exists at all'
            (Get-PolRawRecord $recs $script:relBase 'Flags').Type | Should -Be 3 -Because 'step 3 did not replace the record with DWORD 0'
            $o = & $script:Gpo @script:GpoParams -Url $neverUrl -Remove -Confirm:$false 3>$null
            $o.RemovedEntry | Should -BeFalse
            $o.RootFlags    | Should -BeLike 'unusable: *type 3 (REG_BINARY)*' -Because 'the summary names the type instead of (absent)'
            # remove the value, then restore the DWORD 4 the next tests expect
            Remove-GPRegistryValue -Guid $script:LabGpo.Id -Key $script:RootKey -ValueName Flags | Out-Null
            (Get-PolRawRecord (script:Read-LabPol -Path $script:PolMachine) $script:relBase 'Flags') | Should -BeNullOrEmpty
            Set-GPRegistryValue -Guid $script:LabGpo.Id -Key $script:RootKey -ValueName Flags -Type DWord -Value 4 | Out-Null
            $o = & $script:Gpo @script:GpoParams -Url $script:LabUrl -PolicyName $script:LabName -Confirm:$false 3>$null
            $o.RootFlags | Should -BeExactly '0x4 (unchanged)'
        }

        It '-Server (the PDC emulator) and -Domain give the same result as the default replica' {
            $pdc = (Get-ADDomain).PDCEmulator
            $o = & $script:Gpo @script:GpoParams -Url $script:LabUrl -PolicyName $script:LabName -Server $pdc -Confirm:$false 3>$null
            $o.EntryApplied | Should -BeTrue
            $o.ADPolicyRow  | Should -BeExactly 'applied'
            $o.Key          | Should -BeExactly "$script:RootKey\$script:ExpectedKey"
            $o.RootFlags    | Should -BeExactly '0x4 (unchanged)'
            $o = & $script:Gpo @script:GpoParams -Url $script:LabUrl -PolicyName $script:LabName -Domain $env:USERDNSDOMAIN -Confirm:$false 3>$null
            $o.EntryApplied | Should -BeTrue
            $o.ADPolicyRow  | Should -BeExactly 'applied'
            $o.Key          | Should -BeExactly "$script:RootKey\$script:ExpectedKey"
            $o.RootFlags    | Should -BeExactly '0x4 (unchanged)'
            @(Get-PolEntries (script:Read-LabPol -Path $script:PolMachine)).Count | Should -Be 2 -Because 'reruns are idempotent'
        }

        It '-SetAsDefault and -EnableAutoEnrollmentPolicy author the marker and AE values' {
            $o = & $script:Gpo @script:GpoParams -Url $script:LabUrl -PolicyName $script:LabName `
                     -SetAsDefault -EnableAutoEnrollmentPolicy -Confirm:$false 3>$null
            $o.DefaultChanged | Should -BeTrue
            $recs = script:Read-LabPol -Path $script:PolMachine -MinRecords 16
            (Get-PolValue $recs $script:relBase '') | Should -BeExactly $script:ExpectedPid
            $ae = Get-GPRegistryValue -Guid $script:LabGpo.Id -Key 'HKLM\SOFTWARE\Policies\Microsoft\Cryptography\AutoEnrollment'
            $map = @{}; foreach ($v in $ae) { $map[$v.ValueName] = $v.Value }
            [int]$map['AEPolicy']                 | Should -Be 7
            [int]$map['OfflineExpirationPercent'] | Should -Be 10
            $map['OfflineExpirationStoreNames']   | Should -BeExactly 'MY'
        }

        It '-AEExpirationPercent 25 rewrites OfflineExpirationPercent (with a change warning); a rerun with the default restores 10' {
            $w = $null
            $o = & $script:Gpo @script:GpoParams -Url $script:LabUrl -PolicyName $script:LabName -EnableAutoEnrollmentPolicy -AEExpirationPercent 25 `
                     -Confirm:$false -WarningAction SilentlyContinue -WarningVariable w
            $o.EntryApplied   | Should -BeTrue
            $o.AutoEnrollment | Should -BeExactly "AEPolicy=7, 25%, 'MY' (applied=True)"
            @($w | Where-Object { "$_" -like "*already carries Auto-Enrollment settings (AEPolicy=7, 10%, 'MY')*(AEPolicy=7, 25%, 'MY')*" }).Count |
                Should -Be 1 -Because 'the previous case left 10%, and the script warns before changing an existing AE value'
            $ae = Get-GPRegistryValue -Guid $script:LabGpo.Id -Key $script:AeKey
            $map = @{}; foreach ($v in $ae) { $map[$v.ValueName] = $v.Value }
            [int]$map['AEPolicy']                 | Should -Be 7
            [int]$map['OfflineExpirationPercent'] | Should -Be 25
            $map['OfflineExpirationStoreNames']   | Should -BeExactly 'MY'
            $recs = script:Read-LabPol -Path $script:PolMachine
            [int](Get-PolValue $recs $script:relAe 'OfflineExpirationPercent') | Should -Be 25 -Because 'the registry.pol replay must agree with the GPMC view'
            [int](Get-PolValue $recs $script:relAe 'AEPolicy')                 | Should -Be 7
            (Get-PolValue $recs $script:relAe 'OfflineExpirationStoreNames')   | Should -BeExactly 'MY'
            (Get-PolValue $recs $script:relBase '') | Should -BeExactly $script:ExpectedPid -Because 'a rerun without -SetAsDefault leaves the marker alone'
            # Restore the 10% the previous case established, so the later cases continue from that state.
            $o = & $script:Gpo @script:GpoParams -Url $script:LabUrl -PolicyName $script:LabName -EnableAutoEnrollmentPolicy -Confirm:$false 3>$null
            $o.AutoEnrollment | Should -BeExactly "AEPolicy=7, 10%, 'MY' (applied=True)"
            [int](Get-PolValue (script:Read-LabPol -Path $script:PolMachine) $script:relAe 'OfflineExpirationPercent') | Should -Be 10
            [int]((Get-GPRegistryValue -Guid $script:LabGpo.Id -Key $script:AeKey -ValueName OfflineExpirationPercent).Value) | Should -Be 10
        }

        It '-Authentication Certificate writes AuthFlags 8 on a fresh entry (GPMC view and registry.pol replay agree); the entry is removed again' {
            $r = script:Invoke-LabSwitchRoundTrip -Suffix 'authcert' -Extra @{ Authentication = 'Certificate' }
            $r.Summary.EntryApplied   | Should -BeTrue
            $r.Summary.Authentication | Should -BeExactly 'Certificate (0x8)'
            $r.Summary.Flags          | Should -BeExactly '0x14' -Because 'the entry Flags keep their default'
            $r.Summary.RootFlags      | Should -BeExactly '0x4 (unchanged)'
            $r.Gpmc['URL']            | Should -BeExactly $r.Url
            [int]$r.Gpmc['AuthFlags'] | Should -Be 8
            [int]$r.Gpmc['Flags']     | Should -Be 0x14
            $r.KeysWith               | Should -Contain $r.Key
            [int]$r.Pol['AuthFlags']  | Should -Be 8
            [int]$r.Pol['Flags']      | Should -Be 0x14
            $r.Removed.RemovedEntry   | Should -BeTrue
            $r.Removed.DefaultCleared | Should -BeFalse -Because 'the marker points at the main entry, which stays'
            $r.KeysAfter              | Should -Not -Contain $r.Key
            $r.KeysAfter              | Should -Contain $script:ExpectedKey -Because 'the main entry is untouched'
        }

        It '-NoAutoEnroll clears bit 0x10 on a fresh entry (Flags 0x4 in both views); the entry is removed again' {
            $r = script:Invoke-LabSwitchRoundTrip -Suffix 'noautoenroll' -Extra @{ NoAutoEnroll = $true }
            $r.Summary.EntryApplied   | Should -BeTrue
            $r.Summary.Flags          | Should -BeExactly '0x4'
            $r.Summary.Authentication | Should -BeExactly 'Kerberos (0x2)' -Because 'the default authentication is untouched'
            [int]$r.Gpmc['Flags']     | Should -Be 0x4
            [int]$r.Gpmc['AuthFlags'] | Should -Be 2
            [int]$r.Pol['Flags']      | Should -Be 0x4
            [int]$r.Pol['AuthFlags']  | Should -Be 2
            $r.Removed.RemovedEntry   | Should -BeTrue
            $r.KeysAfter              | Should -Not -Contain $r.Key
            $r.KeysAfter              | Should -Contain $script:ExpectedKey
        }

        It '-NoClientId clears bit 0x4 on a fresh entry (Flags 0x10 in both views); the entry is removed again' {
            $r = script:Invoke-LabSwitchRoundTrip -Suffix 'noclientid' -Extra @{ NoClientId = $true }
            $r.Summary.EntryApplied | Should -BeTrue
            $r.Summary.Flags        | Should -BeExactly '0x10'
            [int]$r.Gpmc['Flags']   | Should -Be 0x10
            [int]$r.Pol['Flags']    | Should -Be 0x10
            $r.Removed.RemovedEntry | Should -BeTrue
            $r.KeysAfter            | Should -Not -Contain $r.Key
            $r.KeysAfter            | Should -Contain $script:ExpectedKey
        }

        It '-AllowUntrustedIssuer sets bit 0x20 on a fresh entry (Flags 0x34 in both views); the entry is removed again' {
            $r = script:Invoke-LabSwitchRoundTrip -Suffix 'untrusted' -Extra @{ AllowUntrustedIssuer = $true }
            $r.Summary.EntryApplied | Should -BeTrue
            $r.Summary.Flags        | Should -BeExactly '0x34'
            [int]$r.Gpmc['Flags']   | Should -Be 0x34
            [int]$r.Pol['Flags']    | Should -Be 0x34
            $r.Removed.RemovedEntry | Should -BeTrue
            $r.KeysAfter            | Should -Not -Contain $r.Key
            $r.KeysAfter            | Should -Contain $script:ExpectedKey
        }

        It 'removing one of two endpoints that share a PolicyID keeps the (Default) marker' {
            $redUrl = "$script:UrlBase`?redundant"
            $w = $null
            $o = & $script:Gpo @script:GpoParams -Url $redUrl -PolicyName $script:LabName -PolicyId $script:ExpectedPid -Confirm:$false -WarningAction SilentlyContinue -WarningVariable w
            $o.EntryApplied | Should -BeTrue
            @($w | Where-Object { "$_" -like '*shares PolicyID*' }).Count | Should -Be 1 -Because 'without -ReplaceExisting the script only warns'
            $o = & $script:Gpo @script:GpoParams -Url $redUrl -Remove -Confirm:$false 3>$null
            $o.RemovedEntry   | Should -BeTrue
            $o.DefaultCleared | Should -BeFalse
            @($o.Notes) -match 'marker kept' | Should -Not -BeNullOrEmpty
            $recs = script:Read-LabPol -Path $script:PolMachine
            (Get-PolValue $recs $script:relBase '') | Should -BeExactly $script:ExpectedPid -Because 'the first endpoint still serves the PolicyID'
            (Get-PolEntries $recs).Key | Should -Not -Contain (script:Get-LabKey $redUrl)
        }

        It '-ReplaceExisting removes a same-PolicyID sibling (verified from registry.pol) but never the AD row' {
            $sibUrl = "$script:UrlBase`?stale"
            $null = & $script:Gpo @script:GpoParams -Url $sibUrl -PolicyName $script:LabName -PolicyId $script:ExpectedPid -Confirm:$false 3>$null
            $sibKey = script:Get-LabKey $sibUrl
            (Get-PolEntries (script:Read-LabPol -Path $script:PolMachine)).Key | Should -Contain $sibKey

            $o = & $script:Gpo @script:GpoParams -Url $script:LabUrl -PolicyName $script:LabName -ReplaceExisting -Confirm:$false 3>$null
            @($o.DuplicatesRemoved) | Should -Contain $sibUrl -Because 'the removal is reported only after registry.pol confirms it'
            $recs = script:Read-LabPol -Path $script:PolMachine
            $keysNow = @((Get-PolEntries $recs).Key)
            $keysNow | Should -Not -Contain $sibKey
            Test-EntryRecordsPresent $recs "$script:relBase\$sibKey" | Should -BeFalse -Because 'no physical record of the sibling remains'
            $keysNow | Should -Contain $script:AdKey
            $keysNow | Should -Contain $script:ExpectedKey
        }

        It '-Remove deletes the entry, clears the orphaned marker, and reports what remains' {
            $o = & $script:Gpo @script:GpoParams -Url $script:LabUrl -Remove -Confirm:$false 3>$null
            $o.RemovedEntry   | Should -BeTrue
            $o.DefaultCleared | Should -BeTrue
            $recs = script:Read-LabPol -Path $script:PolMachine
            $keysNow = @((Get-PolEntries $recs).Key)
            $keysNow | Should -Not -Contain $script:ExpectedKey
            $keysNow | Should -Contain $script:AdKey                            # AD row is left alone
            (Get-PolValue $recs $script:relBase '') | Should -BeNullOrEmpty     # marker cleared
            @($o.Notes) -match 'Auto-Enrollment' | Should -Not -BeNullOrEmpty   # AE flagged as remaining
            @($o.Notes) -match 'Remaining entries' | Should -Not -BeNullOrEmpty
        }

        It '-Remove -ClearDefault clears a marker that points at ANOTHER entry' {
            $null = & $script:Gpo @script:GpoParams -Url $script:LabUrl -PolicyName $script:LabName -SetAsDefault -Confirm:$false 3>$null
            $otherUrl = "$script:UrlBase`?other"
            $null = & $script:Gpo @script:GpoParams -Url $otherUrl -PolicyName "$script:Prefix Other" -Confirm:$false 3>$null
            (Get-PolValue (script:Read-LabPol -Path $script:PolMachine) $script:relBase '') | Should -BeExactly $script:ExpectedPid
            $o = & $script:Gpo @script:GpoParams -Url $otherUrl -Remove -ClearDefault -Confirm:$false 3>$null
            $o.RemovedEntry   | Should -BeTrue
            $o.DefaultCleared | Should -BeTrue -Because 'the marker did not point at the removed entry, so only -ClearDefault clears it'
            $o.DefaultMarker  | Should -BeExactly ''
            $recs = script:Read-LabPol -Path $script:PolMachine
            (Get-PolValue $recs $script:relBase '') | Should -BeNullOrEmpty
            (Get-PolEntries $recs).Key | Should -Contain $script:ExpectedKey -Because 'the other entry is left alone'
        }

        It 'a second -Remove of the same URL reports nothing to remove' {
            $o = & $script:Gpo @script:GpoParams -Url $script:LabUrl -Remove -Confirm:$false 3>$null
            $o.RemovedEntry   | Should -BeTrue
            $o.DefaultCleared | Should -BeFalse -Because 'no marker was set'
            $o = & $script:Gpo @script:GpoParams -Url $script:LabUrl -Remove -Confirm:$false 3>$null
            $o.RemovedEntry   | Should -BeFalse
            $o.DefaultCleared | Should -BeFalse
            @($o.Notes) -match 'nothing to remove' | Should -Not -BeNullOrEmpty
        }

        It 'User scope with -SkipADPolicy: after -Remove no entry remains and the root-values note is raised' {
            $o = & $script:Gpo @script:GpoParams -Url $script:LabUrl -PolicyName $script:LabName -Scope User -SkipADPolicy -DisableUserConfigured -Confirm:$false 3>$null
            $o.EntryApplied | Should -BeTrue
            $o.ADPolicyRow  | Should -BeExactly 'skipped (-SkipADPolicy)'
            $o.RootFlags    | Should -BeExactly '0x4 (applied)'
            @(Get-PolEntries (script:Read-LabPol -Path $script:PolUser)).Count | Should -Be 1
            $o = & $script:Gpo @script:GpoParams -Url $script:LabUrl -Scope User -Remove -Confirm:$false 3>$null
            $o.RemovedEntry | Should -BeTrue
            $o.RootFlags    | Should -BeExactly '0x4'
            @($o.Notes) -match 'No entries remain' | Should -Not -BeNullOrEmpty -Because 'the root Flags value still announces GP CEP configuration'
            @(Get-PolEntries (Read-PolRecords -Path $script:PolUser)).Count | Should -Be 0
        }

        It 'User scope: add and remove round-trip in the User half of the GPO' {
            $o = & $script:Gpo @script:GpoParams -Url $script:LabUrl -PolicyName $script:LabName -Scope User -Confirm:$false 3>$null
            $o.EntryApplied | Should -BeTrue
            $o.Key | Should -BeExactly "HKCU\SOFTWARE\Policies\Microsoft\Cryptography\PolicyServers\$script:ExpectedKey"
            (Get-PolEntries (script:Read-LabPol -Path $script:PolUser)).Key | Should -Contain $script:ExpectedKey

            $o = & $script:Gpo @script:GpoParams -Url $script:LabUrl -Scope User -Remove -Confirm:$false 3>$null
            $o.RemovedEntry | Should -BeTrue
            @((Get-PolEntries (Read-PolRecords -Path $script:PolUser)).Key) | Should -Not -Contain $script:ExpectedKey
        }
    }
}
