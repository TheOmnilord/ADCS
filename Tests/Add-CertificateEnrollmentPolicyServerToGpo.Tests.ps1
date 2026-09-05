<#
.SYNOPSIS
    Pester suite for Add-CertificateEnrollmentPolicyServerToGpo.ps1. Requires Pester 5+.

.DESCRIPTION
    Three always-on tiers plus one opt-in tier:

      -Tag Unit    Pure helpers extracted from the script by AST (so the REAL code runs, never a
                   copy) and exercised in-process: the registry.pol binary parser (against a
                   hand-built .pol byte stream) and the entry/value extractors (against synthetic
                   record arrays). No GroupPolicy module, no GPO, no AD.
      -Tag Static  The script parses and its comment-based help binds. No module, no GPO.
      -Tag Guard   Parameter-conflict validation. These invocations throw BEFORE the script
                   resolves the GPO (Get-GPO), so they contact no GPO/AD and change nothing - but
                   the script imports the GroupPolicy module up front, so the tier is skipped when
                   the module is unavailable (detection dual-probes ListAvailable AND the Windows
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
        # (The script has mandatory params and '#Requires -Modules GroupPolicy', so it cannot be
        # dot-sourced wholesale; extracting the function bodies runs them with no side effects.)
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:Gpo, [ref]$null, [ref]$null)
        foreach ($name in 'Read-PolRecords', 'Get-PolEffectiveValues', 'Get-PolEntries', 'Get-PolValue', 'Test-EntryRecordsPresent', 'Test-PolEntryUsable') {
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

    # -------------------------------------------------------------------------------------------
    # Lab tier: live round-trip inside ONE throwaway, never-linked GPO. Opt-in (-RunLab).
    # Tests are SEQUENTIAL: add -> verify -> default+AE -> replace-sibling -> remove, mirroring
    # a real rollout and teardown inside the same GPO.
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
            $sysvol = "\\$($script:LabGpo.DomainName)\SYSVOL\$($script:LabGpo.DomainName)\Policies\{$($script:LabGpo.Id)}"
            $script:PolMachine = "$sysvol\Machine\registry.pol"
            $script:PolUser    = "$sysvol\User\registry.pol"

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
        }

        AfterAll {
            # Surgical: delete ONLY the throwaway GPO, by its exact tracked GUID.
            if ($script:LabGpo) {
                try { Remove-GPO -Guid $script:LabGpo.Id -Confirm:$false } catch { Write-Warning "Failed to remove lab GPO $($script:LabGpo.Id): $_" }
            }
            # Backstop, scoped to THIS run's fresh-GUID prefix (cannot match a pre-existing GPO).
            # STRUCTURAL GUARD: the sweep runs only when the prefix has its full PESTER-<hex8>
            # shape - an unset/empty prefix would otherwise degenerate the filter to -like "*"
            # and delete every GPO in the domain. Never widen this.
            if ($script:Prefix -match '^PESTER-[0-9a-f]{8}$') {
                foreach ($g in @(Get-GPO -All | Where-Object { $_.DisplayName -like "$script:Prefix*" })) {
                    try { Remove-GPO -Guid $g.Id -Confirm:$false } catch { }
                    Write-Warning "AfterAll backstop removed an untracked test GPO: $($g.DisplayName)"
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
            $o.Key          | Should -BeExactly "HKLM\SOFTWARE\Policies\Microsoft\Cryptography\PolicyServers\$script:ExpectedKey"
        }

        It 'the GPMC API (Get-GPRegistryValue) sees every authored value' {
            $vals = Get-GPRegistryValue -Guid $script:LabGpo.Id -Key "HKLM\SOFTWARE\Policies\Microsoft\Cryptography\PolicyServers\$script:ExpectedKey"
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

        It '-ReplaceExisting removes a same-PolicyID sibling but never the AD row' {
            $sibUrl = "$script:UrlBase`?stale"
            $null = & $script:Gpo @script:GpoParams -Url $sibUrl -PolicyName $script:LabName -PolicyId $script:ExpectedPid -Confirm:$false 3>$null
            $sha1 = [System.Security.Cryptography.SHA1]::Create()
            $sibKey = -join ($sha1.ComputeHash([System.Text.Encoding]::Unicode.GetBytes($sibUrl.ToLowerInvariant())) |
                             ForEach-Object { $_.ToString('x2') })
            (Get-PolEntries (script:Read-LabPol -Path $script:PolMachine)).Key | Should -Contain $sibKey

            $o = & $script:Gpo @script:GpoParams -Url $script:LabUrl -PolicyName $script:LabName -ReplaceExisting -Confirm:$false 3>$null
            @($o.DuplicatesRemoved) | Should -Contain $sibUrl
            $keysNow = @((Get-PolEntries (script:Read-LabPol -Path $script:PolMachine)).Key)
            $keysNow | Should -Not -Contain $sibKey
            $keysNow | Should -Contain $script:AdKey
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
