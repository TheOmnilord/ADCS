<#
.SYNOPSIS
    Pester suite for Add-CertificateEnrollmentPolicyServerOffline.ps1. Requires Pester 5+.

.DESCRIPTION
    Three always-on tiers plus one opt-in tier:

      -Tag Unit    Exercises the REAL derivations (SHA-1/UTF-16LE subkey, EJBCA String.hashCode()
                   PolicyID, Flags/AuthFlags math) by invoking the script under -WhatIf against
                   the per-user hive (-Location LocalUser, which needs no elevation) and reading
                   the emitted summary object. -WhatIf makes every registry write a no-op, so the
                   tier changes nothing. No AD, no modules, no elevation.
      -Tag Static  The script parses and its comment-based help binds. No AD, no modules.
      -Tag Guard   Parameter-conflict validation. These invocations throw BEFORE the elevation
                   check and before any registry write, so they need neither elevation nor a
                   reachable registry hive and make no changes.
      -Tag Lab     Live registry round-trips in the USER-CONFIGURED stores (LocalUser always;
                   LocalMachine when the session is elevated). Skipped unless -RunLab is passed.
                   Surgical by construction: entries use a per-run PESTER-<hex> policy name and a
                   URL under the RFC-reserved .invalid TLD (can never reach anything real), an
                   entry key is tracked by exact path (and removed in teardown) ONLY when the
                   script's own summary reports that it CREATED it (EntryAction = Created, which
                   the script derives from the REG_CREATED_NEW_KEY disposition of its own
                   RegCreateKeyExW call) - a key the script merely updated pre-existed and is
                   foreign, and a key that is merely present after a throw is never tracked; both
                   are only named for manual review - a later case that modifies or removes an
                   entry an earlier case created first proves that the entry is tracked
                   (Assert-OwnedEntry) and fails without touching it otherwise, and every script
                   call that can remove an entry CONSUMES the ownership of the paths it can delete
                   BEFORE it runs (Invoke-RemovalOwned): a path is tracked again only when the
                   script's own summary reports that it did NOT remove it and the key still exists,
                   so a key the script deleted stays untracked whatever exists at its path
                   afterwards (a registry key carries no identity: a key another writer creates at
                   that path between the script's delete and the reconciliation is foreign), also
                   when the script throws; every teardown
                   delete is LEAF-ONLY (RegDeleteKeyW, which fails on a key with subkeys; never
                   Remove-Item, whose registry provider deletes the whole subtree even without
                   -Recurse), so a child that appears after the check survives and the key is
                   named in a warning; a pre-existing (Default) marker is snapshotted and
                   restored only when the current value is the value this suite last asked the
                   script to write (the PolicyId it passed with -SetAsDefault, never a value read
                   back afterwards), or when the marker is absent because the suite's latest
                   transition cleared that value (any other value is left and named); and a
                   PolicyServers base key is removed only if the script's own summary reports
                   that this run created it (BaseKeyCreated, from the same disposition) and it
                   ends the run empty. The GP-hive locations (GPMachine/GPUser) are NEVER
                   written: on a domain member they tattoo pseudo-policy backed by no GPO (the
                   script itself warns) - the GPO suite covers Group Policy delivery against a
                   throwaway unlinked GPO. One Lab case runs -Location GPUser under -WhatIf only,
                   to exercise the AD objectGUID lookup, and proves the GP hive unchanged
                   afterwards. Extracted-helper cases (the path checker, Test-RegEntryUsable) use
                   throwaway keys directly under HKCU:\Software; each is created with the native
                   RegCreateKeyExW and claimed only on its REG_CREATED_NEW_KEY disposition, and
                   deleted (leaf-only) only by the test that made it.

    Oracle constants (independently derived; the ldap: subkey matches the script's own AD_KEY
    documentation, which cross-checks the SHA-1/UTF-16LE method):
      * PolicyID  Java String.hashCode('Example PKI Service')                     = 241064013
      * Subkey    SHA-1(UTF-16LE(lowercased URL)) of the sample EJBCA CEP URL     = dc032f3a...
    The Lab tier re-derives both per run with reference implementations, so live writes are
    checked against an independent oracle, not against the script's own output.

.EXAMPLE
    Invoke-Pester -Path .\Tests\Add-CertificateEnrollmentPolicyServerOffline.Tests.ps1 -ExcludeTag Lab

.EXAMPLE
    # Parse/help/derivation only - no registry access at all:
    Invoke-Pester -Path .\Tests\Add-CertificateEnrollmentPolicyServerOffline.Tests.ps1 -Tag Unit,Static

.EXAMPLE
    # Full run including live registry round-trips (LocalUser; LocalMachine too when elevated):
    $cfg = New-PesterContainer -Path .\Tests\Add-CertificateEnrollmentPolicyServerOffline.Tests.ps1 -Data @{ RunLab = $true }
    Invoke-Pester -Container $cfg
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '',
    Justification = 'container parameters are consumed inside Pester Describe/BeforeAll scriptblocks, which the analyzer cannot see through')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingEmptyCatchBlock', '',
    Justification = 'best-effort teardown paths (AfterAll marker restore and key removal) deliberately swallow per-item errors')]
param(
    [bool]   $RunLab     = $false,
    [string] $ScriptPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'Add-CertificateEnrollmentPolicyServerOffline.ps1')
)

BeforeDiscovery {
    # -Skip conditions are evaluated during discovery, so anything they reference must be set here.
    $script:LabReady    = $RunLab
    $script:LabElevated = $RunLab -and ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
                          ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

Describe 'Add-CertificateEnrollmentPolicyServerOffline' {

    BeforeAll {
        $script:Cep = $ScriptPath
        $script:Cep | Should -Exist

        $script:KnownUrl    = 'https://pki.example.net/ejbca/msae/CEPService?alias'
        $script:KnownName   = 'Example PKI Service'
        $script:KnownPid    = '241064013'
        $script:KnownSubkey = 'dc032f3a68521c2445e1e161da81503bddce17a7'

        # Invoke under -WhatIf against the per-user hive (no elevation, no writes) and return the
        # summary object. Streams 3-6 are silenced (warnings/verbose/debug/information); the
        # "What if:" lines themselves go straight to the host and are NOT redirectable - they are
        # accepted as benign console noise, and no assertion depends on them.
        function Invoke-Cep {
            param([hashtable]$Params)
            & $script:Cep @Params -Location LocalUser -WhatIf 3>$null 4>$null 5>$null 6>$null
        }

        # Encodes one datum of a registry-hive snapshot line so that the field delimiter (|), the
        # line delimiter (;) and the element delimiter (,) cannot collide with data, and the
        # elements of a REG_MULTI_SZ or a byte array keep their boundaries ('a,b'+'c' differs
        # from 'a'+'b,c'). A backslash escape has no length limit: [uri]::EscapeDataString throws
        # on Windows PowerShell 5.1 above about 65,000 characters, and a REG_SZ under the GP hive
        # can be that long. Script scope so the Unit tier can call it without a registry.
        function script:ConvertTo-ShapeToken($Value) {
            if ($Value -is [Array]) { return '[' + (@($Value | ForEach-Object { script:ConvertTo-ShapeToken "$_" }) -join ',') + ']' }
            $s = "$Value"
            # Backslash first, so the escapes added below are not escaped again.
            $s = $s -replace '\\', '\\'
            $s = $s -replace '\|', '\|'
            $s = $s -replace ';', '\;'
            $s = $s -replace ',', '\,'
            $s = $s -replace "`r", '\r'
            $s = $s -replace "`n", '\n'
            $s
        }

        # The oracle for the script's root-Flags contract, independent of the script's helper and
        # shared by the Unit GPUser -WhatIf case and the Lab GP-hive snapshot case so the two
        # cannot drift. Returns the RootFlags string the summary must show for a usable value,
        # '(absent)' for no value, and $null for an UNUSABLE value (the script's step 0a refuses
        # such a value BEFORE its -WhatIf gates, so that run throws instead of returning a
        # summary). Usable: a DWORD; a QWORD that fits an [int]; a non-empty REG_SZ or
        # REG_EXPAND_SZ that the PowerShell [int] cast accepts (decimal, 0x hex, a sign, leading
        # zeros, an exponent). GetValue returns a REG_EXPAND_SZ as a string, and the script
        # converts it exactly like a REG_SZ. Everything else (an empty or non-numeric string,
        # REG_BINARY, REG_MULTI_SZ) is unusable.
        function script:Get-ExpectedRootFlags($Raw, $Kind) {
            if ($null -eq $Raw) { return '(absent)' }
            $dword  = [Microsoft.Win32.RegistryValueKind]::DWord
            $qword  = [Microsoft.Win32.RegistryValueKind]::QWord
            $string = [Microsoft.Win32.RegistryValueKind]::String
            $expand = [Microsoft.Win32.RegistryValueKind]::ExpandString
            if ($Kind -eq $dword) { return '0x{0:X}' -f [int]$Raw }
            if ($Kind -eq $qword) { try { return '0x{0:X}' -f [int]$Raw } catch { return $null } }
            if ($Kind -eq $string -or $Kind -eq $expand) {
                $s = "$Raw".Trim()
                if ($s.Length -eq 0) { return $null }
                try { return '0x{0:X}' -f [int]$s } catch { return $null }
            }
            $null
        }

        # Reads the raw root Flags value and its kind from a PolicyServers key. READ-ONLY: the GP
        # hive can carry the user's real configuration and is never written by this suite.
        function script:Read-RootFlags([string]$Path) {
            if (-not (Test-Path -LiteralPath $Path)) { return [pscustomobject]@{ Raw = $null; Kind = $null } }
            $k = Get-Item -LiteralPath $Path
            $raw  = $k.GetValue('Flags')
            $kind = if ($null -ne $raw) { $k.GetValueKind('Flags') } else { $null }
            [pscustomobject]@{ Raw = $raw; Kind = $kind }
        }
    }

    Context 'Unit: real derivations via -WhatIf (no writes)' -Tag 'Unit' {

        It 'PolicyID = Java String.hashCode() of the policy name (EJBCA MSAE)' {
            $o = Invoke-Cep @{ Url = $script:KnownUrl; PolicyName = $script:KnownName }
            $o.PolicyID | Should -BeExactly $script:KnownPid
        }

        It 'subkey = SHA-1 over UTF-16LE of the invariant-lowercased URL' {
            $o = Invoke-Cep @{ Url = $script:KnownUrl; PolicyName = $script:KnownName }
            (Split-Path $o.Path -Leaf) | Should -BeExactly $script:KnownSubkey
        }

        It 'subkey derivation is case-insensitive on the URL (invariant-lowercased first)' {
            $lower = Invoke-Cep @{ Url = $script:KnownUrl;            PolicyName = $script:KnownName }
            $upper = Invoke-Cep @{ Url = $script:KnownUrl.ToUpper();  PolicyName = $script:KnownName }
            (Split-Path $upper.Path -Leaf) | Should -BeExactly (Split-Path $lower.Path -Leaf)
        }

        It 'an explicit -PolicyId overrides the computed hash' {
            $o = Invoke-Cep @{ Url = $script:KnownUrl; PolicyName = $script:KnownName; PolicyId = '{ABCD}' }
            $o.PolicyID | Should -BeExactly '{ABCD}'
        }

        It 'default Flags = 0x14 (0x10 autoenroll | 0x4 ClientId), matching the GPO editor' {
            $o = Invoke-Cep @{ Url = $script:KnownUrl; PolicyName = $script:KnownName }
            $o.Flags | Should -BeExactly '0x14'
        }

        It '-NoClientId clears bit 0x4 (Flags -> 0x10)' {
            $o = Invoke-Cep @{ Url = $script:KnownUrl; PolicyName = $script:KnownName; NoClientId = $true }
            $o.Flags | Should -BeExactly '0x10'
        }

        It '-NoAutoEnroll clears bit 0x10 (Flags -> 0x4)' {
            $o = Invoke-Cep @{ Url = $script:KnownUrl; PolicyName = $script:KnownName; NoAutoEnroll = $true }
            $o.Flags | Should -BeExactly '0x4'
        }

        It '-AllowUntrustedIssuer sets bit 0x20 (Flags -> 0x34)' {
            $o = Invoke-Cep @{ Url = $script:KnownUrl; PolicyName = $script:KnownName; AllowUntrustedIssuer = $true }
            $o.Flags | Should -BeExactly '0x34'
        }

        It 'authentication maps to the correct AuthFlags bit (Certificate -> 0x8)' {
            $o = Invoke-Cep @{ Url = $script:KnownUrl; PolicyName = $script:KnownName; Authentication = 'Certificate' }
            $o.Authentication | Should -BeExactly 'Certificate (0x8)'
        }

        It 'Cost defaults to the dialog default 0x7FFFFFFD and is echoed as hex' {
            $o = Invoke-Cep @{ Url = $script:KnownUrl; PolicyName = $script:KnownName }
            $o.Cost | Should -BeExactly '0x7FFFFFFD'
        }

        It 'PolicyID wraps to a negative Int32 like Java (EJBCA MSAE Policy -> -1941035357; hello -> 99162322)' {
            # Oracles: 'hello' is the published Java String.hashCode() value; the EJBCA name was
            # recomputed with two independent implementations (int64 mask and unchecked Int32 wrap).
            (Invoke-Cep @{ Url = $script:KnownUrl; PolicyName = 'EJBCA MSAE Policy' }).PolicyID | Should -BeExactly '-1941035357'
            (Invoke-Cep @{ Url = $script:KnownUrl; PolicyName = 'hello' }).PolicyID              | Should -BeExactly '99162322'
        }

        It 'Remove mode under -WhatIf for an unknown URL reports nothing to remove' {
            $o = & $script:Cep -Url 'https://nothing-here.invalid/ejbca/msae/CEPService?alias' -Location LocalUser -Remove -WhatIf 3>$null
            $o.Mode         | Should -BeExactly 'Remove'
            $o.RemovedEntry | Should -BeFalse
            (@($o.Notes) -join ' ') | Should -BeLike '*nothing to remove*'
            $o.RootFlags    | Should -BeExactly 'n/a (local location)'
        }

        It 'GPUser under -WhatIf with -SkipADPolicy -DisableUserConfigured selects the GP hive, previews the root Flags and writes nothing' {
            # Covers the GP hive selection, the elevation branch under -WhatIf, Get-RootFlagsDisplay and
            # the root-Flags block. -SkipADPolicy keeps the AD objectGUID lookup out, so no AD is needed.
            # Warnings are captured (3>&1) so the elevation branch can be asserted either way.
            # The REAL root Flags of this user's GP hive decide the expected outcome (read-only, never
            # written): a usable value gives a summary; an unusable one (REG_SZ 'abc', an empty
            # string, REG_BINARY) makes step 0a refuse even under -WhatIf, which is correct behaviour
            # and must not fail this case. The shared oracle keeps this in step with the Lab case.
            $gp = 'HKCU:\Software\Policies\Microsoft\Cryptography\PolicyServers'
            $rf = script:Read-RootFlags $gp
            $expectedFlags = script:Get-ExpectedRootFlags $rf.Raw $rf.Kind
            if ($null -eq $expectedFlags) {
                { & $script:Cep -Url $script:KnownUrl -PolicyName $script:KnownName -Location GPUser -SkipADPolicy -DisableUserConfigured -WhatIf 3>$null 4>$null 5>$null 6>$null } |
                    Should -Throw -ExpectedMessage '*root Flags*cannot be used*Nothing was written*'
                return
            }
            $out = & $script:Cep -Url $script:KnownUrl -PolicyName $script:KnownName -Location GPUser -SkipADPolicy -DisableUserConfigured -WhatIf 3>&1 4>$null 5>$null 6>$null
            $warnings = @($out | Where-Object { $_ -is [System.Management.Automation.WarningRecord] })
            $o = @($out | Where-Object { $_ -isnot [System.Management.Automation.WarningRecord] })[-1]
            $o.Location     | Should -BeExactly 'GPUser'
            $o.Path         | Should -BeLike "$gp\*"
            $o.EntryApplied | Should -BeFalse
            $o.EntryAction  | Should -BeExactly 'Declined' -Because '-WhatIf declines the entry write at the confirmation gate'
            $o.ADPolicyRow  | Should -BeExactly 'skipped (-SkipADPolicy)'
            $o.RootFlags    | Should -Match '^(\(absent\)|0x[0-9A-F]+)$'
            $o.RootFlags    | Should -BeExactly $expectedFlags -Because 'the summary reports the ACTUAL root Flags, which -WhatIf leaves unchanged'
            $elevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
            $elevWarn = @($warnings | Where-Object { $_.Message -like '*requires an elevated session*' })
            if ($elevated) { $elevWarn.Count | Should -Be 0 -Because 'an elevated session gets no preview-only warning' }
            else           { $elevWarn.Count | Should -BeGreaterThan 0 -Because 'a non-elevated session must be told the real run needs elevation' }
        }

        It 'the root-Flags oracle (suite helper) accepts DWORD, a QWORD that fits an [int], and a numeric REG_SZ or REG_EXPAND_SZ, and refuses everything else' {
            $dword  = [Microsoft.Win32.RegistryValueKind]::DWord
            $qword  = [Microsoft.Win32.RegistryValueKind]::QWord
            $string = [Microsoft.Win32.RegistryValueKind]::String
            $expand = [Microsoft.Win32.RegistryValueKind]::ExpandString
            $binary = [Microsoft.Win32.RegistryValueKind]::Binary
            script:Get-ExpectedRootFlags 4 $dword | Should -BeExactly '0x4'
            foreach ($s in '4', '0x4', '+4', '4e0', ' 4 ') {
                script:Get-ExpectedRootFlags $s $string | Should -BeExactly '0x4' -Because "the [int] cast accepts REG_SZ '$s'"
            }
            script:Get-ExpectedRootFlags '4' $expand | Should -BeExactly '0x4' -Because 'GetValue returns a REG_EXPAND_SZ as a string and the script converts it like a REG_SZ'
            script:Get-ExpectedRootFlags '' $string    | Should -BeNullOrEmpty -Because 'an empty REG_SZ is unusable even though [int]"" is 0'
            script:Get-ExpectedRootFlags 'abc' $string | Should -BeNullOrEmpty
            script:Get-ExpectedRootFlags ([byte[]](1, 2, 3)) $binary | Should -BeNullOrEmpty -Because 'REG_BINARY is unusable'
            script:Get-ExpectedRootFlags $null $null | Should -BeExactly '(absent)'
            script:Get-ExpectedRootFlags ([long]4) $qword | Should -BeExactly '0x4'
            script:Get-ExpectedRootFlags ([long][math]::Pow(2, 40)) $qword | Should -BeNullOrEmpty -Because 'a QWORD that does not fit an [int] is unusable'
        }

        It 'ConvertTo-RootFlagsValue converts a DWORD and a numeric REG_SZ, and refuses everything else without throwing (extracted helper)' {
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:Cep, [ref]$null, [ref]$null)
            $def = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'ConvertTo-RootFlagsValue' }, $false)
            $def | Should -Not -BeNullOrEmpty
            . ([scriptblock]::Create($def[0].Extent.Text))
            $dword  = [Microsoft.Win32.RegistryValueKind]::DWord
            $string = [Microsoft.Win32.RegistryValueKind]::String
            $binary = [Microsoft.Win32.RegistryValueKind]::Binary
            $multi  = [Microsoft.Win32.RegistryValueKind]::MultiString

            (ConvertTo-RootFlagsValue -Raw 20 -Kind $dword).Value  | Should -Be 20
            (ConvertTo-RootFlagsValue -Raw -1 -Kind $dword).Value  | Should -Be -1 -Because 'a DWORD 0xFFFFFFFF reads back as Int32 -1 and must stay usable'
            (ConvertTo-RootFlagsValue -Raw 20 -Kind $dword).Reason | Should -BeNullOrEmpty
            (ConvertTo-RootFlagsValue -Raw '20' -Kind $string).Value    | Should -Be 20 -Because 'the 1.0.4 repair of a numeric REG_SZ relies on this'
            (ConvertTo-RootFlagsValue -Raw '0x14' -Kind $string).Value  | Should -Be 20
            (ConvertTo-RootFlagsValue -Raw ' 0x2 ' -Kind $string).Value | Should -Be 2
            # Strings the bare [int] cast accepted before 1.0.8 must still convert (review regression cases).
            (ConvertTo-RootFlagsValue -Raw '+20' -Kind $string).Value          | Should -Be 20
            (ConvertTo-RootFlagsValue -Raw '000000000020' -Kind $string).Value | Should -Be 20
            (ConvertTo-RootFlagsValue -Raw '0x0000000014' -Kind $string).Value | Should -Be 20

            $r = ConvertTo-RootFlagsValue -Raw 'abc' -Kind $string
            $r.Value  | Should -BeNullOrEmpty
            $r.Reason | Should -BeLike "*String*'abc'*"
            $r = ConvertTo-RootFlagsValue -Raw '' -Kind $string
            $r.Value  | Should -BeNullOrEmpty -Because 'an empty REG_SZ is not a number, even though [int]"" is 0'
            $r.Reason | Should -BeLike "*String*''*"
            $r = ConvertTo-RootFlagsValue -Raw '4294967295' -Kind $string
            $r.Value  | Should -BeNullOrEmpty -Because '[int] of that string overflows and must not throw'
            $r = ConvertTo-RootFlagsValue -Raw ([byte[]](1, 2, 3)) -Kind $binary
            $r.Value  | Should -BeNullOrEmpty
            $r.Reason | Should -BeLike "*Binary*'1,2,3'*"
            $r = ConvertTo-RootFlagsValue -Raw ([string[]]('a', 'b')) -Kind $multi
            $r.Value  | Should -BeNullOrEmpty
            $r.Reason | Should -BeLike '*MultiString*'
            $r = ConvertTo-RootFlagsValue -Raw 'abc' -Kind $null
            $r.Reason | Should -BeLike '*String*' -Because 'with no kind the .NET type name is used'

            $r = ConvertTo-RootFlagsValue -Raw $null -Kind $null
            $r.Value  | Should -BeNullOrEmpty
            $r.Reason | Should -BeNullOrEmpty -Because 'an absent value is not an error'
        }

        It 'runs in a session that already holds the released CepRegNative type (three methods, no RegCreateKeyExW): the script loads its own type and completes under -WhatIf' {
            # Regression: the script once guarded its Add-Type with the name CepRegNative. A session
            # that had run the released script (whose type has only RegOpenKeyExW, RegQueryValueExW
            # and RegCloseKey) skipped the Add-Type, and the first Add then failed on the missing
            # RegCreateKeyExW. A type cannot be redefined in a running process, so the script's
            # expanded type must have another name. This stub is the released type: it is loaded
            # (once per process) BEFORE the script runs, and stays for every later case.
            if (-not ('CepRegNative' -as [type])) {
                Add-Type -TypeDefinition @"
using System; using System.Runtime.InteropServices;
public static class CepRegNative {
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int RegOpenKeyExW(IntPtr hKey, string lpSubKey, uint ulOptions, uint samDesired, out IntPtr phkResult);
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int RegQueryValueExW(IntPtr hKey, string lpValueName, IntPtr lpReserved, out uint lpType, byte[] lpData, ref uint lpcbData);
    [DllImport("advapi32.dll")] public static extern int RegCloseKey(IntPtr hKey);
}
"@
            }
            ('CepRegNative' -as [type]) | Should -Not -BeNullOrEmpty
            ([type]'CepRegNative').GetMethod('RegCreateKeyExW') | Should -BeNullOrEmpty -Because 'the stub must be the released shape, without the create method'
            $o = Invoke-Cep @{ Url = $script:KnownUrl; PolicyName = $script:KnownName }
            $o.Mode        | Should -BeExactly 'Add'
            $o.PolicyID    | Should -BeExactly $script:KnownPid
            $o.EntryAction | Should -BeExactly 'Declined' -Because '-WhatIf declines the entry write; the run reached that gate, so the path check ran on the script''s own type'
            $o = & $script:Cep -Url 'https://nothing-here.invalid/ejbca/msae/CEPService?alias' -Location LocalUser -Remove -WhatIf 3>$null
            $o.Mode | Should -BeExactly 'Remove'
        }

        It 'the hive-snapshot encoder has no length limit, escapes every delimiter and keeps multi-string element boundaries (suite helper)' {
            # Regression: [uri]::EscapeDataString throws on Windows PowerShell 5.1 for a string longer
            # than about 65,000 characters, and a REG_SZ under the GP hive can be that long.
            $long = 'x' * 100000
            $script:encodedLong = $null
            { $script:encodedLong = script:ConvertTo-ShapeToken $long } | Should -Not -Throw
            $script:encodedLong.Length | Should -BeGreaterOrEqual 100000
            $delims = script:ConvertTo-ShapeToken '|;,\'
            $delims | Should -BeExactly '\|\;\,\\'
            $delims | Should -Not -Match '(^|[^\\])[|;,]' -Because 'every unescaped delimiter would corrupt the snapshot line'
            $m1 = script:ConvertTo-ShapeToken @('a,b', 'c')
            $m2 = script:ConvertTo-ShapeToken @('a', 'b,c')
            $m1 | Should -BeExactly '[a\,b,c]'
            $m2 | Should -BeExactly '[a,b\,c]'
            $m1 | Should -Not -BeExactly $m2 -Because 'the element boundaries of a REG_MULTI_SZ must survive the encoding'
            # Every case yields a distinct token, and a plain string never collides with a one-element array.
            @($script:encodedLong, $delims, $m1, $m2, (script:ConvertTo-ShapeToken 'a,b'), (script:ConvertTo-ShapeToken @('a,b'))) |
                Select-Object -Unique | Should -HaveCount 6
            (script:ConvertTo-ShapeToken "a`r`nb") | Should -BeExactly 'a\r\nb'
            (script:ConvertTo-ShapeToken $null)    | Should -BeExactly ''
        }
    }

    Context 'Static: parse and help' -Tag 'Static' {

        It 'parses without errors' {
            $errs = $null
            $null = [System.Management.Automation.Language.Parser]::ParseFile($script:Cep, [ref]$null, [ref]$errs)
            $errs | Should -BeNullOrEmpty
        }

        It 'comment-based help binds (Synopsis is present)' {
            (Get-Help $script:Cep).Synopsis.Trim() | Should -Not -BeNullOrEmpty
        }

        It 'carries a PSScriptInfo header (Test-ScriptFileInfo parses it; Version is semver)' {
            $info = Test-ScriptFileInfo -Path $script:Cep -ErrorAction Stop
            $info.Version | Should -Match '^\d+\.\d+\.\d+$'
            $info.Guid    | Should -Not -BeNullOrEmpty
        }
        It 'never references the released native type name CepRegNative (a session that holds that type must still load the expanded one)' {
            # The Add-Type guard and every static call must use the new name. A guard on the old name
            # skips the Add-Type when the released type is present, and the first native create then
            # fails on a missing method. The name of the new type must still CONTAIN the old one,
            # because the Lab cases extract the Add-Type block by that substring.
            $text = Get-Content -LiteralPath $script:Cep -Raw
            $text | Should -Not -Match "'CepRegNative'\s+-as\s+\[type\]" -Because 'the Add-Type guard must not test the released type name'
            $text | Should -Not -Match '\[CepRegNative\]::' -Because 'every native call must go through the expanded type'
            $text | Should -Not -Match 'public static class CepRegNative\s*\{' -Because 'the released type name cannot be redefined in a session that holds it'
            $m = [regex]::Match($text, "'(CepRegNative[A-Za-z0-9_]+)'\s+-as\s+\[type\]")
            $m.Success | Should -BeTrue -Because 'the guard must test a name that contains CepRegNative (the Lab cases find the block by that substring)'
            $typeName = $m.Groups[1].Value
            $text | Should -Match "public static class $typeName\s*\{"
            $text | Should -Match "\[$typeName\]::RegCreateKeyExW\("
            $text | Should -Match "\[$typeName\]::RegOpenKeyExW\("
        }

        It 'documents every non-common parameter' {
            $cmd = Get-Command $script:Cep
            $common = [System.Management.Automation.PSCmdlet]::CommonParameters + [System.Management.Automation.PSCmdlet]::OptionalCommonParameters
            $documented = @((Get-Help $script:Cep).parameters.parameter.name)
            foreach ($p in $cmd.Parameters.Keys | Where-Object { $_ -notin $common }) {
                $documented | Should -Contain $p -Because "parameter -$p should have a .PARAMETER help entry"
            }
        }
    }

    Context 'Guard: parameter-conflict validation (throws before any write)' -Tag 'Guard' {

        It 'rejects -SetAsDefault with -ClearDefault' {
            { & $script:Cep -Url 'https://x/' -PolicyName 'n' -SetAsDefault -ClearDefault } |
                Should -Throw -ExpectedMessage '*mutually exclusive*'
        }

        It 'rejects -DisableUserConfigured with -EnableUserConfigured' {
            { & $script:Cep -Url 'https://x/' -PolicyName 'n' -Location GPMachine -DisableUserConfigured -EnableUserConfigured } |
                Should -Throw -ExpectedMessage '*mutually exclusive*'
        }

        It 'rejects the GP-only root-Flags switches on a non-GP location' {
            { & $script:Cep -Url 'https://x/' -PolicyName 'n' -Location LocalUser -DisableUserConfigured } |
                Should -Throw -ExpectedMessage '*only exist in the Group Policy hive*'
        }

        It 'rejects a non-http(s) URL' {
            { & $script:Cep -Url 'ftp://pki/nope' -PolicyName 'n' -Location LocalUser } |
                Should -Throw -ExpectedMessage '*absolute http/https URI*'
        }

        It 'rejects a URL with a control character (the character sits inside, so Trim cannot hide it)' {
            { & $script:Cep -Url "https://x/`ta" -PolicyName 'n' -Location LocalUser } |
                Should -Throw -ExpectedMessage '*control characters*'
        }

        It 'rejects a URL that is not an absolute URI' {
            { & $script:Cep -Url 'not a url' -PolicyName 'n' -Location LocalUser } |
                Should -Throw -ExpectedMessage '*absolute http/https URI*'
        }

        It 'warns when PolicyName has leading or trailing whitespace (hashed verbatim)' {
            $out = & $script:Cep -Url 'https://x/' -PolicyName ' n ' -Location LocalUser -WhatIf 3>&1 4>$null 5>$null 6>$null
            $w = @($out | Where-Object { $_ -is [System.Management.Automation.WarningRecord] -and $_.Message -like '*whitespace*' })
            $w.Count | Should -BeGreaterThan 0
            $o = @($out | Where-Object { $_ -isnot [System.Management.Automation.WarningRecord] })[-1]
            $o.FriendlyName | Should -BeExactly ' n ' -Because 'the name is written and hashed verbatim'
        }
    }

    # -------------------------------------------------------------------------------------------
    # Lab tier: live registry round-trips in the user-configured stores. Opt-in (-RunLab).
    # Tests are SEQUENTIAL: each builds on the state the previous one verified (add -> update ->
    # default -> replace-sibling -> remove), mirroring how the script is used in real life.
    # -------------------------------------------------------------------------------------------
    Context 'Lab: live registry round-trip (user-configured stores)' -Tag 'Lab' -Skip:(-not $script:LabReady) {

        BeforeAll {
            $script:Prefix  = "PESTER-$([guid]::NewGuid().ToString('N').Substring(0,8))"
            # .invalid is RFC 2606-reserved: this URL can never resolve to a real endpoint. The
            # host carries the run prefix, so the safety-net sweep below is scoped to THIS run.
            $script:UrlBase = "https://$($script:Prefix.ToLower()).lab.invalid/ejbca/msae/CEPService"
            $script:LabUrl  = "$script:UrlBase`?alias"
            $script:LabName = "$script:Prefix Policy"

            # Independent oracles - reference implementations, not the script's code.
            $sha1 = [System.Security.Cryptography.SHA1]::Create()
            $script:ExpectedKey = -join ($sha1.ComputeHash([System.Text.Encoding]::Unicode.GetBytes($script:LabUrl.ToLowerInvariant())) |
                                         ForEach-Object { $_.ToString('x2') })
            $h = [int64]0
            foreach ($c in $script:LabName.ToCharArray()) { $h = ($h * 31 + [int64]$c) -band 4294967295 }
            if ($h -ge 2147483648) { $h -= 4294967296 }
            $script:ExpectedPid = "$h"

            $script:HiveCU = 'HKCU:\Software\Microsoft\Cryptography\PolicyServers'
            $script:HiveLM = 'HKLM:\SOFTWARE\Microsoft\Cryptography\PolicyServers'

            # Snapshot shared state so teardown can restore the (Default) marker EXACTLY. Existed is
            # the oracle for the script's BaseKeyCreated field in the first Add of each location;
            # teardown itself trusts only that field (see BaseKeyCreated below).
            function script:Get-HiveSnapshot([string]$Hive) {
                $existed = Test-Path -LiteralPath $Hive
                @{ Existed = $existed; Marker = if ($existed) { (Get-Item -LiteralPath $Hive).GetValue('') } else { $null } }
            }
            $script:PreCU = script:Get-HiveSnapshot $script:HiveCU
            $script:PreLM = script:Get-HiveSnapshot $script:HiveLM

            # Exact key paths this run creates - the ONLY things teardown deletes outright.
            $script:CreatedKeys = New-Object System.Collections.Generic.List[string]

            # Native registry calls for the fixtures and the teardown. RegCreateKeyExW reports in
            # its disposition whether it CREATED the key (1, REG_CREATED_NEW_KEY) or OPENED one that
            # already existed (2): that disposition is the only creation evidence a fixture accepts
            # (New-Item without -Force checks existence and then calls CreateSubKey, which opens an
            # existing key, so a key another writer created in between would be claimed).
            # RegDeleteKeyW deletes ONE key and fails when the key has subkeys: the provider's
            # Remove-Item calls DeleteSubKeyTree even without -Recurse, so it would also delete a
            # child that appeared after the caller's subkey check.
            if (-not ('PesterRegNative' -as [type])) {
                Add-Type -TypeDefinition @"
using System; using System.Runtime.InteropServices;
public static class PesterRegNative {
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int RegCreateKeyExW(IntPtr hKey, string lpSubKey, int Reserved, string lpClass, uint dwOptions, uint samDesired, IntPtr lpSecurityAttributes, out IntPtr phkResult, out uint lpdwDisposition);
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int RegDeleteKeyW(IntPtr hKey, string lpSubKey);
    [DllImport("advapi32.dll")] public static extern int RegCloseKey(IntPtr hKey);
}
"@
            }
            # Splits HKCU:\... or HKLM:\... into the hive handle and the hive-relative subkey.
            function script:Split-RegistryPath([string]$Path) {
                $isCU = $Path -like 'HKCU:*'
                [pscustomobject]@{
                    Handle = if ($isCU) { [IntPtr]::new(-2147483647) } else { [IntPtr]::new(-2147483646) }   # HKEY_CURRENT_USER / HKEY_LOCAL_MACHINE
                    Rel    = ($Path -replace '^HK(CU|LM):\\?', '').TrimEnd('\')
                }
            }
            # Creates a non-volatile key and returns $true ONLY when this call created it
            # (REG_CREATED_NEW_KEY). $false means the key already existed and was only opened: the
            # caller must not claim it. Any other failure throws with the Win32 code.
            function script:New-LeafKey([string]$Path) {
                $p = script:Split-RegistryPath $Path
                $h = [IntPtr]::Zero; $disp = [uint32]0
                $rc = [PesterRegNative]::RegCreateKeyExW($p.Handle, $p.Rel, 0, $null, 0, 0xF003F, [IntPtr]::Zero, [ref]$h, [ref]$disp)   # KEY_ALL_ACCESS
                if ($rc -ne 0) { throw "Fixture key $Path could not be created (Win32 error $rc)." }
                [void][PesterRegNative]::RegCloseKey($h)
                return ($disp -eq 1)
            }
            # Deletes ONE tracked key, leaf-only, and returns $true when the key is gone afterwards.
            # When the native delete fails, the key is left in place and named in a warning together
            # with the subkeys that remain: a child that appeared after the caller's check is
            # foreign and must survive. This is the ONLY way the suite deletes a real key.
            function script:Remove-LeafKey([string]$Path) {
                $p = script:Split-RegistryPath $Path
                $rc = [PesterRegNative]::RegDeleteKeyW($p.Handle, $p.Rel)
                if ($rc -eq 0 -or $rc -eq 2) { return $true }   # ERROR_SUCCESS, or ERROR_FILE_NOT_FOUND: already gone
                $subs = @()
                try { if (Test-Path -LiteralPath $Path) { $subs = @((Get-Item -LiteralPath $Path).GetSubKeyNames()) } } catch { }
                if ($subs.Count -gt 0) {
                    Write-Warning "Teardown left $Path in place: the leaf-only delete failed (Win32 error $rc) because the key has subkey(s) this run did not create ($($subs -join ', ')). Review and remove it manually by this exact path."
                } else {
                    Write-Warning "Teardown left $Path in place: the leaf-only delete failed (Win32 error $rc). Review and remove it manually by this exact path."
                }
                return $false
            }

            # Native calls for the symbolic-link fixtures: a VOLATILE link key is created with
            # RegCreateKeyExW (REG_OPTION_VOLATILE | REG_OPTION_CREATE_LINK), its REG_LINK value is
            # set with RegSetValueExW, and the LINK OBJECT itself is deleted through a handle opened
            # with REG_OPTION_OPEN_LINK and NtDeleteKey (RegDeleteKeyW and Remove-Item would follow
            # the link and act on its target).
            if (-not ('PesterRegLink' -as [type])) {
                Add-Type -TypeDefinition @"
using System; using System.Runtime.InteropServices;
public static class PesterRegLink {
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int RegCreateKeyExW(IntPtr hKey, string lpSubKey, int Reserved, string lpClass, uint dwOptions, uint samDesired, IntPtr lpSecurityAttributes, out IntPtr phkResult, out uint lpdwDisposition);
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int RegSetValueExW(IntPtr hKey, string lpValueName, int Reserved, uint dwType, byte[] lpData, uint cbData);
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int RegOpenKeyExW(IntPtr hKey, string lpSubKey, uint ulOptions, uint samDesired, out IntPtr phkResult);
    [DllImport("ntdll.dll")] public static extern int NtDeleteKey(IntPtr KeyHandle);
    [DllImport("advapi32.dll")] public static extern int RegCloseKey(IntPtr hKey);
}
"@
            }
            # Deletes the registry symbolic link at hive-relative $LinkRel under HKCU through a handle
            # opened with REG_OPTION_OPEN_LINK (0x8) and DELETE (0x10000): the link object goes, its
            # target is never touched. Every failure is named with the key path: a link that stays
            # (volatile: a reboot removes it at the latest) must be visible to the operator.
            function script:Remove-VolatileLink([string]$LinkRel) {
                $hkcu = [IntPtr]::new(-2147483647)   # HKEY_CURRENT_USER
                $lh = [IntPtr]::Zero
                $rc = [PesterRegLink]::RegOpenKeyExW($hkcu, $LinkRel, 0x8, 0x10000, [ref]$lh)
                if ($rc -ne 0) {
                    Write-Warning "Teardown could not open the link key HKCU:\$LinkRel for deletion (Win32 error $rc). Review and remove the link manually by this exact path (it is volatile: a reboot removes it at the latest)."
                    return $false
                }
                try {
                    $nt = [PesterRegLink]::NtDeleteKey($lh)
                    if ($nt -ne 0) {
                        Write-Warning ("Teardown could not delete the link key HKCU:\$LinkRel (NtDeleteKey returned NTSTATUS 0x{0:X8}). Review and remove the link manually by this exact path (it is volatile: a reboot removes it at the latest)." -f $nt)
                        return $false
                    }
                    return $true
                }
                finally { [void][PesterRegLink]::RegCloseKey($lh) }
            }

            function script:Remove-MarkerValue([string]$Hive) {
                $rel  = $Hive -replace '^HK(CU|LM):\\', ''
                $root = if ($Hive -like 'HKCU:*') { [Microsoft.Win32.Registry]::CurrentUser } else { [Microsoft.Win32.Registry]::LocalMachine }
                $k = $root.OpenSubKey($rel, $true)
                if ($k) { try { $k.DeleteValue('', $false) } finally { $k.Close() } }
            }

            # (Default) marker bookkeeping for AfterAll: the LATEST transition of the marker that
            # this suite can attribute to itself, per hive. A write transition records the value
            # the suite ASKED the script to set (the PolicyId the case passed with -SetAsDefault,
            # given as -Requested) when the summary reports DefaultChanged - never the DefaultMarker
            # the script read back afterwards, which another writer can have replaced between the
            # script's write and that read. A clear transition (DefaultCleared from -Remove, or
            # DefaultChanged from a -ClearDefault run, declared with -ClearDefault) records the
            # value it cleared only when that value was the suite's own latest write. Each
            # transition resets the other field, so LastWrite and LastClear never both hold a
            # value. AfterAll deletes or restores only when the current marker equals LastWrite,
            # or when the marker is absent and LastClear holds the suite's own value; anything
            # else is another writer's and is left in place and named in a warning.
            $script:MarkerState = @{}   # hive path -> @{ LastWrite = value or $null; LastClear = value or $null }
            function script:Get-MarkerState([string]$Hive) {
                if (-not $script:MarkerState.ContainsKey($Hive)) { $script:MarkerState[$Hive] = @{ LastWrite = $null; LastClear = $null } }
                $script:MarkerState[$Hive]
            }
            function script:Register-MarkerChange([string]$Hive, $Summary, [string]$Requested, [switch]$ClearDefault) {
                $st = script:Get-MarkerState $Hive
                foreach ($s in @($Summary)) {
                    if ($null -eq $s) { continue }
                    $changed = [bool]($s.PSObject.Properties['DefaultChanged'] -and $s.DefaultChanged -eq $true)
                    $cleared = [bool]($s.PSObject.Properties['DefaultCleared'] -and $s.DefaultCleared -eq $true)
                    if ($changed -and -not $ClearDefault) {
                        if (-not $Requested) { throw "Register-MarkerChange: a summary with DefaultChanged needs -Requested (the PolicyId the case passed with -SetAsDefault) or -ClearDefault; the marker the script read back is never used as evidence." }
                        $st.LastWrite = $Requested; $st.LastClear = $null
                    }
                    elseif ($cleared -or ($changed -and $ClearDefault)) {
                        # The value cleared is attributable only when it was the suite's own latest write.
                        $st.LastClear = $st.LastWrite; $st.LastWrite = $null
                    }
                }
            }

            # The BaseKeyCreated the script reported in the FIRST Add summary registered under each
            # base key (HiveCU / HiveLM). AfterAll removes a base key only when that summary says
            # this run created it AND the key is empty afterwards. The snapshot's absence check is
            # not enough: another writer can create the base key between that check and the
            # script's write, and the script then only creates its child under a foreign base.
            $script:BaseKeyCreated = @{}

            # A later case that modifies, replaces or removes an entry an earlier case created must
            # first prove that this run owns it: the path is tracked only on the script's own
            # creation evidence (Register-CreatedEntry). When an earlier proof failed, the key is
            # untracked and may be foreign, so the case fails here and touches nothing.
            function script:Assert-OwnedEntry([string]$Path) {
                if (-not $script:CreatedKeys.Contains($Path)) {
                    throw "This case needs the entry at $Path, but this run has not proven that it created it (an earlier case failed, or the key is foreign). The case stops here and does not modify or remove the key."
                }
            }

            # Reconciles this run's ownership of ONE entry path with the registry, with NO assertion:
            # when no key exists at the path, the path is unregistered (a no-op when it was not
            # tracked) and $true is returned; when a key is still there it stays tracked and
            # $false is returned. It is NOT a guard for a call that can delete: a key carries no
            # identity, so a key another writer creates at a deleted path before this check would
            # pass Test-Path and stay tracked. Invoke-RemovalOwned covers those calls; this
            # function serves Confirm-RemovedEntry, where the path is already unregistered.
            function script:Sync-OwnedEntry([string]$Path) {
                $gone = -not (Test-Path -LiteralPath $Path)
                if ($gone) { [void]$script:CreatedKeys.Remove($Path) }
                return $gone
            }

            # The entry key name the script derives from a URL: SHA-1 over the UTF-16LE bytes of
            # the invariant-lowercased URL (the suite's own reference implementation).
            function script:ConvertTo-EntryKey([string]$Url) {
                $sha1 = [System.Security.Cryptography.SHA1]::Create()
                -join ($sha1.ComputeHash([System.Text.Encoding]::Unicode.GetBytes($Url.ToLowerInvariant())) | ForEach-Object { $_.ToString('x2') })
            }

            # Runs ONE script call that can delete tracked entry keys (-Remove for the entry;
            # -ReplaceExisting for the superseded siblings), with this ownership rule: a registry
            # key carries no identity, so the ownership of every path the call can delete is
            # CONSUMED before the script runs (each tracked path in -Paths is unregistered first;
            # an untracked path is ignored throughout). After the call - also when the script
            # threw - a consumed path is registered again ONLY when the script's own summary
            # reports that it did NOT remove it (-Reregister receives the summary object and the
            # path and returns $true then: RemovedEntry false, or the sibling absent from
            # DuplicatesRemoved) AND a key still exists at the path. A path the summary reports
            # as removed stays unregistered whatever exists there now: a key another writer
            # creates at that path between the script's delete and this reconciliation (for
            # example during the script's marker step) is foreign, and a Test-Path check alone
            # would keep it tracked and let AfterAll delete it. When the script returned no
            # summary (it threw), nothing is registered again: a key that still exists at a
            # consumed path is named for manual review. The summary is returned; the original
            # exception passes through untouched.
            function script:Invoke-RemovalOwned([string[]]$Paths, [scriptblock]$Run, [scriptblock]$Reregister) {
                $owned = @($Paths | Where-Object { $script:CreatedKeys.Contains($_) })
                foreach ($p in $owned) { [void]$script:CreatedKeys.Remove($p) }
                $summary = $null
                try { $summary = & $Run }
                finally {
                    $s = @($summary) | Where-Object { $null -ne $_ } | Select-Object -Last 1
                    foreach ($p in $owned) {
                        if ($null -eq $s) {
                            if (Test-Path -LiteralPath $p) { Write-Warning "The script returned no summary for a call that can delete $p (it threw). The key that exists there is NO LONGER tracked for teardown: review it manually and remove it by this exact path only if it is this run's." }
                            continue
                        }
                        if ((& $Reregister $s $p) -eq $true -and (Test-Path -LiteralPath $p)) { $script:CreatedKeys.Add($p) }
                    }
                }
                return $summary
            }

            # The -Reregister rules for Invoke-RemovalOwned. -Remove: the entry is kept when the
            # summary reports RemovedEntry false. -ReplaceExisting: a sibling is kept when no URL
            # in DuplicatesRemoved derives to its key name (the summary carries URLs, not keys).
            $script:KeepUnlessRemovedEntry = { param($Summary, $Path) $null = $Path; ($Summary.RemovedEntry -ne $true) }
            $script:KeepUnlessDuplicateRemoved = {
                param($Summary, $Path)
                $leaf = Split-Path -Path $Path -Leaf
                @(@($Summary.DuplicatesRemoved) | Where-Object { $_ -and (script:ConvertTo-EntryKey $_) -eq $leaf }).Count -eq 0
            }

            # Proves that a consumed entry path is gone, ahead of every assertion on the summary.
            # It runs through Sync-OwnedEntry, so a path that a case did not route through
            # Invoke-RemovalOwned is unregistered here the moment the key is absent, before the
            # assertion; a key that is still there makes the case fail here.
            function script:Confirm-RemovedEntry([string]$Path) {
                (script:Sync-OwnedEntry -Path $Path) | Should -BeTrue -Because "the script reported the removal of the entry at $Path, so the key must be absent"
            }

            # Tracks an entry path for teardown ONLY on creation evidence from the script itself:
            # the Add summary's EntryAction is 'Created' when the script's own RegCreateKeyExW call
            # returned REG_CREATED_NEW_KEY for the entry key and Write-CepEntry then returned.
            # EntryApplied is NOT enough - it is also true for 'Updated', a rewrite of a key that
            # already existed (REG_OPENED_EXISTING_KEY): another writer can create the key between
            # the test's absence check and the script's write, and the script then updates that
            # foreign key. Such a key pre-existed and is left alone (warned, never tracked). When
            # the script threw or reported no write, nothing is tracked either; a
            # key that exists there now (a partial key this run may have left, or a foreign one)
            # is named for manual review instead of being deleted.
            function script:Register-CreatedEntry([string]$Path, $Summary) {
                $action = $null
                foreach ($s in @($Summary)) {
                    if ($null -ne $s -and $s.PSObject.Properties['EntryAction']) { $action = "$($s.EntryAction)" }
                }
                $base = Split-Path -Path $Path -Parent
                foreach ($s in @($Summary)) {
                    if ($null -ne $s -and $s.PSObject.Properties['BaseKeyCreated'] -and -not $script:BaseKeyCreated.ContainsKey($base)) {
                        $script:BaseKeyCreated[$base] = ($s.BaseKeyCreated -eq $true)
                    }
                }
                if ($action -eq 'Created') { $script:CreatedKeys.Add($Path); return }
                if ($action -eq 'Updated') {
                    Write-Warning "The key at $Path pre-existed when the script wrote it (EntryAction = Updated): it is foreign, NOT tracked for teardown, and left alone. Review it manually."
                    return
                }
                if (Test-Path -LiteralPath $Path) {
                    Write-Warning "A key exists at $Path but the script did not report that it created it (it threw, or EntryAction is '$action'). It is NOT tracked for teardown: review it manually and remove it by this exact path only if it is a partial key of this run."
                }
            }
        }

        AfterAll {
            # 1) Surgical: remove ONLY the exact keys this run created (most-recent first). A key is
            #    always registered before any subkey this run creates beneath it, so its tracked
            #    subkeys are removed before it and it is a leaf when its turn comes. A key that
            #    still has a subkey is NOT removed: a subkey this run did not create is foreign
            #    (a CEP entry is a leaf by design), and a tracked subkey that is still there could
            #    not be removed itself. Both are named in a warning and the key is left in place.
            #    The delete itself is leaf-only (Remove-LeafKey, RegDeleteKeyW): a child that
            #    appears between the subkey check and the delete makes the delete fail, so the
            #    key and that child survive. Remove-Item would delete the whole subtree.
            for ($i = $script:CreatedKeys.Count - 1; $i -ge 0; $i--) {
                $path = $script:CreatedKeys[$i]
                try {
                    if (-not (Test-Path -LiteralPath $path)) { continue }
                    $subs = @((Get-Item -LiteralPath $path).GetSubKeyNames())
                    if ($subs.Count -eq 0) { [void](script:Remove-LeafKey -Path $path); continue }
                    $unowned = @($subs | Where-Object { -not $script:CreatedKeys.Contains("$path\$_") })
                    if ($unowned.Count -gt 0) {
                        Write-Warning "AfterAll left $path in place: it has subkey(s) this run did not create ($($unowned -join ', ')). Review and remove it manually by this exact path."
                    } else {
                        Write-Warning "AfterAll left $path in place: its tracked subkey(s) ($($subs -join ', ')) could not be removed. Review and remove it manually by this exact path."
                    }
                } catch { Write-Warning "AfterAll could not remove ${path}: $_" }
            }
            # 2) Safety net, scoped to THIS run: any entry whose URL host carries the run prefix
            #    (a fresh GUID - it cannot match a pre-existing entry). Never a broad wildcard.
            #    STRUCTURAL GUARD: sweep only with a fully-formed run URL base - an unset/empty
            #    one would degenerate the -like pattern to "*" and match real entries.
            if ($script:UrlBase -match '^https://pester-[0-9a-f]{8}\.lab\.invalid/') {
                foreach ($hive in @($script:HiveCU, $script:HiveLM)) {
                    if ($hive -and (Test-Path -LiteralPath $hive)) {
                        foreach ($k in @(Get-ChildItem -LiteralPath $hive)) {
                            $u = $k.GetValue('URL')
                            if ($u -and $u -like "$script:UrlBase*") {
                                # REPORT ONLY: the run URL is not proof that this run created the
                                # entry, so a leftover is named for manual review, never deleted here.
                                Write-Warning "AfterAll safety-net found an untracked entry that carries this run's URL: $($k.PSPath) ($u). Review and remove it manually by this exact path."
                            }
                        }
                    }
                }
            } else {
                Write-Warning "Safety-net sweep skipped: run URL base is unset or malformed ('$script:UrlBase')."
            }
            # 3) Restore each hive's (Default) marker to its snapshotted state ONLY on this suite's
            #    own evidence (Register-MarkerChange): the current value must equal the value the
            #    suite last asked the script to write (LastWrite), or the marker must be absent
            #    because the suite's latest transition cleared that value (LastClear). Any other
            #    current value belongs to another writer: it is left in place and named in a
            #    warning, even on a base key this run did not create. Then remove a base key ONLY if the FIRST Add summary this
            #    run registered under it reports BaseKeyCreated (the script's own RegCreateKeyExW
            #    disposition; a base key that was absent before the run but that the script found
            #    present was created by another writer in between, so it is foreign) and it ends
            #    the run empty - leaf-only, so a subkey that appears in between survives. A base
            #    key this run created that is NOT empty is named in a warning and left in place; a
            #    foreign base key is not mentioned.
            foreach ($pair in @(@($script:HiveCU, $script:PreCU), @($script:HiveLM, $script:PreLM))) {
                $hive = $pair[0]; $pre = $pair[1]
                if (-not (Test-Path -LiteralPath $hive)) { continue }
                $st = script:Get-MarkerState $hive
                $lastWrite = $st.LastWrite; $lastClear = $st.LastClear
                try {
                    # $null-aware on BOTH sides: GetValue('') returns $null only when no (Default)
                    # value exists - a present-but-empty ('') or zero (0) marker must be RESTORED,
                    # not misread as absent (truthiness would conflate the two).
                    $cur = (Get-Item -LiteralPath $hive).GetValue('')
                    $curIsOurs = ($null -ne $cur) -and ($null -ne $lastWrite) -and ("$cur" -ceq $lastWrite)
                    if ($null -ne $pre.Marker) {
                        if ($null -eq $cur) {
                            if ($null -ne $lastClear) { Set-ItemProperty -LiteralPath $hive -Name '(default)' -Value $pre.Marker }
                            else { Write-Warning "AfterAll did not restore the (Default) marker under ${hive}: it was '$($pre.Marker)' before the run and is absent now, but the latest clear this run can attribute to itself is not of its own value. Review it manually." }
                        } elseif ("$cur" -cne "$($pre.Marker)") {
                            if ($curIsOurs) { Set-ItemProperty -LiteralPath $hive -Name '(default)' -Value $pre.Marker }
                            else { Write-Warning "AfterAll did not restore the (Default) marker under ${hive}: it was '$($pre.Marker)' before the run and is '$cur' now, not the value this run last asked the script to write. Review it manually." }
                        }
                    } elseif ($null -ne $cur) {
                        if ($curIsOurs) { script:Remove-MarkerValue $hive }
                        else { Write-Warning "AfterAll left the (Default) marker '$cur' under $hive in place: no marker existed before the run, and it is not the value this run last asked the script to write. Review it manually." }
                    }
                } catch { Write-Warning "AfterAll could not restore the (Default) marker under ${hive}: $_" }
                if ($script:BaseKeyCreated[$hive] -eq $true) {
                    $k = Get-Item -LiteralPath $hive
                    if ($k.SubKeyCount -eq 0 -and $k.ValueCount -eq 0) {
                        try { [void](script:Remove-LeafKey -Path $hive) } catch { Write-Warning "AfterAll could not remove the base key ${hive}: $_" }
                    } else {
                        Write-Warning "AfterAll left the base key $hive in place: this run created it, but it is not empty ($($k.SubKeyCount) subkey(s), $($k.ValueCount) value(s)). Review it manually."
                    }
                }
            }
        }

        It 'Add writes the entry; every value verified against independent oracles' {
            # Ownership: the target path is precomputable, so its absence is proven FIRST. The path
            # is tracked in the finally ONLY when the script's summary reports that it created the
            # key (EntryAction = Created); after a throw nothing is tracked and a key found there is only
            # named for manual review (see Register-CreatedEntry).
            $entryPath = "$script:HiveCU\$script:ExpectedKey"
            Test-Path -LiteralPath $entryPath | Should -BeFalse -Because 'the entry must not pre-exist; it is not this run''s and will not be touched'
            $o = $null
            try {
                $o = & $script:Cep -Url $script:LabUrl -PolicyName $script:LabName -Location LocalUser -Confirm:$false 3>$null
            }
            finally {
                script:Register-CreatedEntry -Path $entryPath -Summary $o
            }
            $o.EntryApplied | Should -BeTrue
            $o.EntryAction  | Should -BeExactly 'Created' -Because 'the entry was proven absent before the run wrote it'
            $o.BaseKeyCreated | Should -Be (-not $script:PreCU.Existed) -Because 'the script must report that it created the base key exactly when the snapshot found none'
            $o.Path | Should -BeExactly "$script:HiveCU\$script:ExpectedKey"
            $k = Get-Item -LiteralPath "$script:HiveCU\$script:ExpectedKey"
            $k.GetValue('URL')          | Should -BeExactly $script:LabUrl
            $k.GetValue('PolicyID')     | Should -BeExactly $script:ExpectedPid
            $k.GetValue('FriendlyName') | Should -BeExactly $script:LabName
            [int]$k.GetValue('Flags')     | Should -Be 0x14
            [int]$k.GetValue('AuthFlags') | Should -Be 0x2
            [int]$k.GetValue('Cost')      | Should -Be 0x7FFFFFFD
        }

        It 'Add repairs a value left with the wrong KIND (a REG_DWORD PolicyID becomes the REG_SZ clients expect); every kind is verified' {
            $k = "$script:HiveCU\$script:ExpectedKey"
            script:Assert-OwnedEntry -Path $k
            New-ItemProperty -LiteralPath $k -Name PolicyID -Value 241064013 -PropertyType DWord -Force | Out-Null
            (Get-Item -LiteralPath $k).GetValueKind('PolicyID') | Should -Be ([Microsoft.Win32.RegistryValueKind]::DWord)
            $o = & $script:Cep -Url $script:LabUrl -PolicyName $script:LabName -Location LocalUser -Confirm:$false 3>$null
            $o.EntryApplied | Should -BeTrue
            $o.EntryAction  | Should -BeExactly 'Updated' -Because 'the key from the first Add already exists; the repair rewrites it'
            $k2 = Get-Item -LiteralPath $k
            $k2.GetValueKind('PolicyID') | Should -Be ([Microsoft.Win32.RegistryValueKind]::String) -Because 'Set-ItemProperty without -Type keeps an existing kind; the script must force REG_SZ'
            $k2.GetValue('PolicyID') | Should -BeExactly $script:ExpectedPid
            foreach ($n in 'URL', 'FriendlyName') { $k2.GetValueKind($n) | Should -Be ([Microsoft.Win32.RegistryValueKind]::String) }
            foreach ($n in 'Flags', 'AuthFlags', 'Cost') { $k2.GetValueKind($n) | Should -Be ([Microsoft.Win32.RegistryValueKind]::DWord) }
        }

        It 'rerunning the same Add is an idempotent update (values unchanged, EntryApplied, EntryAction = Updated)' {
            script:Assert-OwnedEntry -Path "$script:HiveCU\$script:ExpectedKey"
            $o = & $script:Cep -Url $script:LabUrl -PolicyName $script:LabName -Location LocalUser -Confirm:$false 3>$null
            $o.EntryApplied | Should -BeTrue
            $o.EntryAction  | Should -BeExactly 'Updated' -Because 'the key from the first Add already exists, so this is a rewrite, not a creation'
            $k = Get-Item -LiteralPath "$script:HiveCU\$script:ExpectedKey"
            $k.GetValue('PolicyID') | Should -BeExactly $script:ExpectedPid
            [int]$k.GetValue('Flags') | Should -Be 0x14
        }

        It '-SetAsDefault writes the (Default) marker with this PolicyID' {
            script:Assert-OwnedEntry -Path "$script:HiveCU\$script:ExpectedKey"
            $o = & $script:Cep -Url $script:LabUrl -PolicyName $script:LabName -Location LocalUser -SetAsDefault -Confirm:$false 3>$null
            script:Register-MarkerChange -Hive $script:HiveCU -Summary $o -Requested $script:ExpectedPid
            $o.DefaultChanged | Should -BeTrue
            (Get-Item -LiteralPath $script:HiveCU).GetValue('') | Should -BeExactly $script:ExpectedPid
        }

        It '-ReplaceExisting removes a same-PolicyID sibling with a different URL' {
            # The main entry is rewritten by the -ReplaceExisting runs below, so it must be this run's.
            script:Assert-OwnedEntry -Path "$script:HiveCU\$script:ExpectedKey"
            # Pre-flight: no FOREIGN entry may share this run's PolicyID (fresh-GUID name makes a
            # collision essentially impossible, but -ReplaceExisting deletes, so prove it first).
            $foreign = @(Get-ChildItem -LiteralPath $script:HiveCU | Where-Object {
                "$($_.GetValue('PolicyID'))" -eq $script:ExpectedPid -and $_.PSChildName -ne $script:ExpectedKey })
            $foreign | Should -BeNullOrEmpty -Because 'no pre-existing entry may share the run PolicyID before testing -ReplaceExisting'

            # Author a stale sibling: same PolicyID, superseded URL (still under the run prefix).
            $sibUrl = "$script:UrlBase`?stale"
            $sha1 = [System.Security.Cryptography.SHA1]::Create()
            $sibKey = -join ($sha1.ComputeHash([System.Text.Encoding]::Unicode.GetBytes($sibUrl.ToLowerInvariant())) |
                             ForEach-Object { $_.ToString('x2') })
            # Ownership: the sibling key must be absent before the script creates it, and it is
            # tracked in the finally ONLY when the script's summary reports that it created the key (EntryAction = Created).
            $sibPath = "$script:HiveCU\$sibKey"
            Test-Path -LiteralPath $sibPath | Should -BeFalse -Because 'the sibling entry must not pre-exist; it is not this run''s and will not be touched'
            $sib = $null
            try {
                $sib = & $script:Cep -Url $sibUrl -PolicyName $script:LabName -PolicyId $script:ExpectedPid -Location LocalUser -Confirm:$false 3>$null
            }
            finally {
                script:Register-CreatedEntry -Path $sibPath -Summary $sib
            }
            $sib.EntryApplied | Should -BeTrue
            $sib.EntryAction  | Should -BeExactly 'Created'
            Test-Path -LiteralPath $sibPath | Should -BeTrue

            # Leaf guard (1.0.6): a CEP entry is a leaf, so a sibling that has a subkey must NOT be
            # removed recursively (a recursive delete could follow a link beneath a delegated descendant).
            # Fixture discipline: prove the subkey absent, create it with the native call, claim it
            # only on the REG_CREATED_NEW_KEY disposition, and delete it (leaf-only) only then.
            $child = "$script:HiveCU\$sibKey\PESTER-child"
            Test-Path -LiteralPath $child | Should -BeFalse -Because "the fixture subkey $child must not pre-exist; it is not this run's and will not be touched"
            $childCreated = $false
            try {
                $childCreated = script:New-LeafKey -Path $child
                $childCreated | Should -BeTrue -Because 'the fixture subkey must be created by this test, not opened as an existing key'
                # -ReplaceExisting can delete the sibling: its ownership is consumed before the
                # call and restored only when the summary reports it NOT removed and it still exists.
                $o = script:Invoke-RemovalOwned -Paths @($sibPath) -Reregister $script:KeepUnlessDuplicateRemoved -Run {
                    & $script:Cep -Url $script:LabUrl -PolicyName $script:LabName -Location LocalUser -ReplaceExisting -Confirm:$false 3>$null
                }
                @($o.DuplicatesRemoved).Count | Should -Be 0 -Because 'a sibling with a subkey is refused, not removed'
                (@($o.Notes) -join ' ') | Should -BeLike '*subkey*'
                Test-Path -LiteralPath "$script:HiveCU\$sibKey" | Should -BeTrue -Because 'the sibling with a subkey must survive'
                $script:CreatedKeys.Contains($sibPath) | Should -BeTrue -Because 'the summary reports the sibling not removed and the key exists, so this run owns it again'
            }
            finally {
                if ($childCreated -and (Test-Path -LiteralPath $child)) { [void](script:Remove-LeafKey -Path $child) }
            }

            # The sibling's ownership is consumed before this call: the summary reports it removed,
            # so the path stays unregistered whatever exists there afterwards.
            $o = script:Invoke-RemovalOwned -Paths @($sibPath) -Reregister $script:KeepUnlessDuplicateRemoved -Run {
                & $script:Cep -Url $script:LabUrl -PolicyName $script:LabName -Location LocalUser -ReplaceExisting -Confirm:$false 3>$null
            }
            # First: prove the sibling gone, before any assertion on the summary.
            script:Confirm-RemovedEntry -Path $sibPath
            @($o.DuplicatesRemoved) | Should -Contain $sibUrl
            Test-Path -LiteralPath "$script:HiveCU\$script:ExpectedKey" | Should -BeTrue
        }

        It '-Remove deletes the entry and clears the now-orphaned (Default) marker' {
            script:Assert-OwnedEntry -Path "$script:HiveCU\$script:ExpectedKey"
            # The entry's ownership is consumed before the call (also when the script removes the
            # entry and then throws, for example in the default-marker step) and restored only
            # when the summary reports RemovedEntry false and the key still exists.
            $o = script:Invoke-RemovalOwned -Paths @("$script:HiveCU\$script:ExpectedKey") -Reregister $script:KeepUnlessRemovedEntry -Run {
                & $script:Cep -Url $script:LabUrl -Location LocalUser -Remove -Confirm:$false 3>$null
            }
            script:Register-MarkerChange -Hive $script:HiveCU -Summary $o
            # First: prove the entry gone, before any assertion on the summary.
            script:Confirm-RemovedEntry -Path "$script:HiveCU\$script:ExpectedKey"
            $o.RemovedEntry   | Should -BeTrue
            $o.DefaultCleared | Should -BeTrue
            (Get-Item -LiteralPath $script:HiveCU).GetValue('') | Should -BeNullOrEmpty
        }

        It '-Remove a second time reports nothing to remove and changes nothing' {
            # The previous case proved the entry gone and unregistered it. A key at that path now
            # is not this run's, and -Remove must not run against it.
            Test-Path -LiteralPath "$script:HiveCU\$script:ExpectedKey" | Should -BeFalse -Because 'the entry was removed by the previous case; a key at that path now is not this run''s and will not be touched'
            # The path is untracked already: the helper ignores it (nothing to consume, nothing
            # to register again) and every deleting call takes the same route.
            $o = script:Invoke-RemovalOwned -Paths @("$script:HiveCU\$script:ExpectedKey") -Reregister $script:KeepUnlessRemovedEntry -Run {
                & $script:Cep -Url $script:LabUrl -Location LocalUser -Remove -Confirm:$false 3>$null
            }
            # The path is untracked already; the helper keeps it so (a no-op) and proves the key absent.
            script:Confirm-RemovedEntry -Path "$script:HiveCU\$script:ExpectedKey"
            $script:CreatedKeys.Contains("$script:HiveCU\$script:ExpectedKey") | Should -BeFalse -Because 'an untracked path is never registered by a deleting call'
            $o.RemovedEntry   | Should -BeFalse
            $o.DefaultCleared | Should -BeFalse
            (@($o.Notes) -join ' ') | Should -BeLike '*nothing to remove*'
        }

        It 'redundant endpoint: -Remove keeps the (Default) marker while another entry serves the PolicyID, and clears it with the last one' {
            $redUrl = "$script:UrlBase`?redundant"
            $sha1 = [System.Security.Cryptography.SHA1]::Create()
            $redKey = -join ($sha1.ComputeHash([System.Text.Encoding]::Unicode.GetBytes($redUrl.ToLowerInvariant())) |
                             ForEach-Object { $_.ToString('x2') })
            # Ownership: the previous case removed the main entry and unregistered it, so both paths
            # must be absent before the script re-creates them. Each path is tracked in its finally
            # ONLY when the script's summary reports that it created the key (EntryAction = Created), and is unregistered
            # again once its removal is proven.
            $entryPath = "$script:HiveCU\$script:ExpectedKey"
            $redPath   = "$script:HiveCU\$redKey"
            Test-Path -LiteralPath $entryPath | Should -BeFalse -Because 'the main entry was removed by the previous case; a key at that path now is not this run''s and will not be touched'
            Test-Path -LiteralPath $redPath   | Should -BeFalse -Because 'the redundant entry must not pre-exist; it is not this run''s and will not be touched'

            $o = $null
            try {
                $o = & $script:Cep -Url $script:LabUrl -PolicyName $script:LabName -Location LocalUser -SetAsDefault -Confirm:$false 3>$null
            }
            finally {
                script:Register-CreatedEntry -Path $entryPath -Summary $o
                script:Register-MarkerChange -Hive $script:HiveCU -Summary $o -Requested $script:ExpectedPid
            }
            $o.EntryApplied   | Should -BeTrue
            $o.EntryAction    | Should -BeExactly 'Created'
            $o.DefaultChanged | Should -BeTrue
            Test-Path -LiteralPath $entryPath | Should -BeTrue
            $o = $null
            try {
                $o = & $script:Cep -Url $redUrl -PolicyName $script:LabName -PolicyId $script:ExpectedPid -Location LocalUser -Confirm:$false 3>$null
            }
            finally {
                script:Register-CreatedEntry -Path $redPath -Summary $o
            }
            Test-Path -LiteralPath $redPath | Should -BeTrue
            $o.EntryApplied | Should -BeTrue
            $o.EntryAction  | Should -BeExactly 'Created'
            (Get-Item -LiteralPath $script:HiveCU).GetValue('') | Should -BeExactly $script:ExpectedPid

            # Each -Remove consumes the ownership of its entry before the call; the summary
            # reports the entry removed, so the path stays unregistered afterwards.
            $o = script:Invoke-RemovalOwned -Paths @($entryPath) -Reregister $script:KeepUnlessRemovedEntry -Run {
                & $script:Cep -Url $script:LabUrl -Location LocalUser -Remove -Confirm:$false 3>$null
            }
            script:Register-MarkerChange -Hive $script:HiveCU -Summary $o
            # First: prove the entry gone, before any assertion on the summary.
            script:Confirm-RemovedEntry -Path $entryPath
            $o.RemovedEntry   | Should -BeTrue
            $o.DefaultCleared | Should -BeFalse -Because 'the redundant entry still serves the marked PolicyID'
            (@($o.Notes) -join ' ') | Should -BeLike '*marker kept*'
            (Get-Item -LiteralPath $script:HiveCU).GetValue('') | Should -BeExactly $script:ExpectedPid

            # The -Run block reads a local of this case ($redUrl): PowerShell scopes dynamically,
            # so the block sees it through the helper's call chain.
            $o = script:Invoke-RemovalOwned -Paths @($redPath) -Reregister $script:KeepUnlessRemovedEntry -Run {
                & $script:Cep -Url $redUrl -Location LocalUser -Remove -Confirm:$false 3>$null
            }
            script:Register-MarkerChange -Hive $script:HiveCU -Summary $o
            script:Confirm-RemovedEntry -Path $redPath
            $o.RemovedEntry   | Should -BeTrue
            $o.DefaultCleared | Should -BeTrue -Because 'no entry serves the marked PolicyID any more'
            (Get-Item -LiteralPath $script:HiveCU).GetValue('') | Should -BeNullOrEmpty
        }

        It '-ClearDefault after -SetAsDefault removes the (Default) marker (DefaultChanged)' {
            # Ownership: the previous case removed the main entry and unregistered it, so the path
            # must be absent before the script re-creates it. The path is tracked in the finally
            # ONLY when the script's summary reports that it created the key (EntryAction = Created), and is unregistered
            # again once its removal is proven.
            $entryPath = "$script:HiveCU\$script:ExpectedKey"
            Test-Path -LiteralPath $entryPath | Should -BeFalse -Because 'the main entry was removed by the previous case; a key at that path now is not this run''s and will not be touched'
            $o = $null
            try {
                $o = & $script:Cep -Url $script:LabUrl -PolicyName $script:LabName -Location LocalUser -SetAsDefault -Confirm:$false 3>$null
            }
            finally {
                script:Register-CreatedEntry -Path $entryPath -Summary $o
                script:Register-MarkerChange -Hive $script:HiveCU -Summary $o -Requested $script:ExpectedPid
            }
            $o.EntryApplied   | Should -BeTrue
            $o.EntryAction    | Should -BeExactly 'Created'
            $o.DefaultChanged | Should -BeTrue
            Test-Path -LiteralPath $entryPath | Should -BeTrue
            (Get-Item -LiteralPath $script:HiveCU).GetValue('') | Should -BeExactly $script:ExpectedPid

            $o = & $script:Cep -Url $script:LabUrl -PolicyName $script:LabName -Location LocalUser -ClearDefault -Confirm:$false 3>$null
            script:Register-MarkerChange -Hive $script:HiveCU -Summary $o -ClearDefault
            $o.DefaultChanged | Should -BeTrue
            $o.DefaultMarker  | Should -BeExactly ''
            (Get-Item -LiteralPath $script:HiveCU).GetValue('') | Should -BeNullOrEmpty

            # The entry's ownership is consumed before the call; the summary reports the entry
            # removed, so the path stays unregistered afterwards.
            $o = script:Invoke-RemovalOwned -Paths @($entryPath) -Reregister $script:KeepUnlessRemovedEntry -Run {
                & $script:Cep -Url $script:LabUrl -Location LocalUser -Remove -Confirm:$false 3>$null
            }
            script:Register-MarkerChange -Hive $script:HiveCU -Summary $o
            # First: prove the entry gone, before any assertion on the summary.
            script:Confirm-RemovedEntry -Path $entryPath
            $o.RemovedEntry | Should -BeTrue
        }

        It 'Add reports EntryAction = Updated and BaseKeyCreated = $false for an entry key that already exists (the evidence is the disposition of the create call)' {
            # Root cause of the review finding: Test-Path followed by New-Item is two operations, and
            # the provider's own CreateSubKey opens an existing key. The script now derives both
            # fields from the REG_CREATED_NEW_KEY / REG_OPENED_EXISTING_KEY disposition of its own
            # RegCreateKeyExW call. This case pre-creates the (empty) entry key by hand with the
            # native call, claims it only on the REG_CREATED_NEW_KEY disposition, runs Add, and
            # asserts that the script reports an update and no base-key creation. The test deletes
            # the key it created itself (leaf-only) and never relies on the script's summary for that.
            $entryPath = "$script:HiveCU\$script:ExpectedKey"
            Test-Path -LiteralPath $script:HiveCU | Should -BeTrue -Because 'the earlier cases created or found the base key; the entry key must not be the first key created here'
            Test-Path -LiteralPath $entryPath | Should -BeFalse -Because 'the main entry was removed by the previous case; a key at that path now is not this run''s and will not be touched'
            $created = $false
            try {
                $created = script:New-LeafKey -Path $entryPath
                if ($created) { $script:CreatedKeys.Add($entryPath) }
                $created | Should -BeTrue -Because 'REG_CREATED_NEW_KEY (1): the test must have created the key itself, not opened an existing one'
                $o = & $script:Cep -Url $script:LabUrl -PolicyName $script:LabName -Location LocalUser -Confirm:$false 3>$null
                $o.EntryApplied   | Should -BeTrue
                $o.EntryAction    | Should -BeExactly 'Updated' -Because 'the key existed when the script wrote it: RegCreateKeyExW returned REG_OPENED_EXISTING_KEY'
                $o.BaseKeyCreated | Should -BeFalse -Because 'the base key existed before this run of the script'
                $k = Get-Item -LiteralPath $entryPath
                $k.GetValue('URL')      | Should -BeExactly $script:LabUrl
                $k.GetValue('PolicyID') | Should -BeExactly $script:ExpectedPid
                [int]$k.GetValue('Flags') | Should -Be 0x14
            }
            finally {
                if ($created) {
                    [void](script:Remove-LeafKey -Path $entryPath)
                    [void]$script:CreatedKeys.Remove($entryPath)
                }
            }
            Test-Path -LiteralPath $entryPath | Should -BeFalse -Because 'the test removed the key it created'
        }

        It 'the registry path check refuses a key that grants write-class rights to an untrusted principal (extracted checker, throwaway HKCU key)' {
            # The untrusted-OWNER refusal is NOT tested: setting a foreign owner needs SeRestorePrivilege,
            # and a key this user creates is owned by the user or by Administrators (both trusted).
            $ast  = [System.Management.Automation.Language.Parser]::ParseFile($script:Cep, [ref]$null, [ref]$null)
            $shim = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.IfStatementAst] -and $n.Extent.Text -match 'CepRegNative' }, $true)
            $shim | Should -Not -BeNullOrEmpty
            . ([scriptblock]::Create($shim[0].Extent.Text))
            foreach ($name in 'Test-RegistryKeyIsLink', 'Test-RegistryComponentExists', 'Assert-ProtectedRegistryPath') {
                $def = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
                $def | Should -Not -BeNullOrEmpty
                . ([scriptblock]::Create($def[0].Extent.Text))
            }
            $aclRel  = "Software\$script:Prefix-acl"
            $regPath = "Registry::HKEY_CURRENT_USER\$aclRel"
            $users   = [System.Security.Principal.SecurityIdentifier]'S-1-5-32-545'   # BUILTIN\Users: never trusted
            # Fixture discipline: prove the key absent, create it with the native call, register and
            # delete it (leaf-only) only if THIS test created it (REG_CREATED_NEW_KEY disposition).
            Test-Path -LiteralPath "HKCU:\$aclRel" | Should -BeFalse -Because "the fixture key HKCU:\$aclRel must not pre-exist; it is not this run's and will not be touched"
            $aclCreated = $false
            try {
                $aclCreated = script:New-LeafKey -Path "HKCU:\$aclRel"
                if ($aclCreated) { $script:CreatedKeys.Add("HKCU:\$aclRel") }
                $aclCreated | Should -BeTrue -Because 'the fixture key must be created by this test, not opened as an existing key'
                { Assert-ProtectedRegistryPath -Path "HKCU:\$aclRel\PolicyServers" } | Should -Not -Throw -Because 'a fresh key under the user hive has a default ACL'
                $acl = Get-Acl -Path $regPath
                $acl.AddAccessRule((New-Object System.Security.AccessControl.RegistryAccessRule($users, 'SetValue', 'Allow')))
                Set-Acl -Path $regPath -AclObject $acl
                { Assert-ProtectedRegistryPath -Path "HKCU:\$aclRel\PolicyServers" } | Should -Throw -ExpectedMessage '*write-class rights*S-1-5-32-545*'
                # Control: a read-only grant to the same principal is not a write-class right.
                $acl = Get-Acl -Path $regPath
                $null = $acl.RemoveAccessRule((New-Object System.Security.AccessControl.RegistryAccessRule($users, 'SetValue', 'Allow')))
                $acl.AddAccessRule((New-Object System.Security.AccessControl.RegistryAccessRule($users, 'ReadKey', 'Allow')))
                Set-Acl -Path $regPath -AclObject $acl
                { Assert-ProtectedRegistryPath -Path "HKCU:\$aclRel\PolicyServers" } | Should -Not -Throw
            }
            finally {
                # This test never creates a subkey under the fixture, and the delete is leaf-only, so
                # a subkey that something else created there makes the delete fail and warn. The
                # path is unregistered after its own teardown, so AfterAll cannot remove a key that
                # something else creates at that path later.
                if ($aclCreated -and (Test-Path -LiteralPath "HKCU:\$aclRel")) { [void](script:Remove-LeafKey -Path "HKCU:\$aclRel") }
                if ($aclCreated) { [void]$script:CreatedKeys.Remove("HKCU:\$aclRel") }
            }
        }

        It 'Test-RegEntryUsable rejects fragments and wrong kinds, and accepts a complete row (extracted helper, throwaway HKCU keys)' {
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:Cep, [ref]$null, [ref]$null)
            $def = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Test-RegEntryUsable' }, $false)
            $def | Should -Not -BeNullOrEmpty
            . ([scriptblock]::Create($def[0].Extent.Text))
            $base = "HKCU:\Software\$script:Prefix-rows"
            # Fixture discipline: prove the base key absent, create it with the native call, register
            # and delete it only if THIS test created it (REG_CREATED_NEW_KEY disposition). The row
            # subkeys live under that fresh base; each is claimed only on its own REG_CREATED_NEW_KEY
            # disposition and registered then, so the AfterAll rule (remove a tracked key only when
            # every subkey beneath it is tracked) holds for the base key too. The teardown deletes
            # the rows one by one (most recent first) and then the base, every delete leaf-only.
            Test-Path -LiteralPath $base | Should -BeFalse -Because "the fixture key $base must not pre-exist; it is not this run's and will not be touched"
            $baseCreated = $false
            $rows = New-Object System.Collections.Generic.List[string]
            $url = 'https://rows.lab.invalid/cep'; $rowPid = '5'   # not $pid: that automatic variable is read-only
            function New-Row([string]$Name, [hashtable]$Strings, [hashtable]$Dwords) {
                $p = "$base\$Name"
                if (-not (script:New-LeafKey -Path $p)) { throw "The fixture row $p already existed (REG_OPENED_EXISTING_KEY): it is not this test's and will not be touched." }
                $rows.Add($p); $script:CreatedKeys.Add($p)
                foreach ($n in $Strings.Keys) { New-ItemProperty -LiteralPath $p -Name $n -Value $Strings[$n] -PropertyType String | Out-Null }
                foreach ($n in $Dwords.Keys)  { New-ItemProperty -LiteralPath $p -Name $n -Value $Dwords[$n]  -PropertyType DWord  | Out-Null }
                Get-Item -LiteralPath $p
            }
            try {
                $baseCreated = script:New-LeafKey -Path $base
                if ($baseCreated) { $script:CreatedKeys.Add($base) }
                $baseCreated | Should -BeTrue -Because 'the fixture base key must be created by this test, not opened as an existing key'
                $frag     = New-Row -Name 'frag'     -Strings @{ URL = $url; PolicyID = $rowPid } -Dwords @{}
                $pidDword = New-Row -Name 'piddword' -Strings @{ URL = $url; FriendlyName = 'n' } -Dwords @{ PolicyID = 5; Flags = 0x14; AuthFlags = 2; Cost = 1 }
                $noCost   = New-Row -Name 'nocost'   -Strings @{ URL = $url; PolicyID = $rowPid; FriendlyName = 'n' } -Dwords @{ Flags = 0x14; AuthFlags = 2 }
                $costSz   = New-Row -Name 'costsz'   -Strings @{ URL = $url; PolicyID = $rowPid; FriendlyName = 'n'; Cost = '1' } -Dwords @{ Flags = 0x14; AuthFlags = 2 }
                $full     = New-Row -Name 'full'     -Strings @{ URL = $url; PolicyID = $rowPid; FriendlyName = 'n' } -Dwords @{ Flags = 0x14; AuthFlags = 2; Cost = 1 }

                Test-RegEntryUsable -Key $null     -ExpectUrl $url -ExpectPolicyId $rowPid | Should -BeFalse
                Test-RegEntryUsable -Key $frag     -ExpectUrl $url -ExpectPolicyId $rowPid | Should -BeFalse -Because 'URL + PolicyID alone is what an interrupted write leaves'
                Test-RegEntryUsable -Key $pidDword -ExpectUrl $url -ExpectPolicyId $rowPid | Should -BeFalse -Because 'a PolicyID stored as REG_DWORD stringifies equal but is the wrong kind'
                Test-RegEntryUsable -Key $noCost   -ExpectUrl $url -ExpectPolicyId $rowPid | Should -BeFalse -Because 'Cost is missing'
                Test-RegEntryUsable -Key $costSz   -ExpectUrl $url -ExpectPolicyId $rowPid | Should -BeFalse -Because 'Cost must be a DWORD'
                Test-RegEntryUsable -Key $full     -ExpectUrl $url -ExpectPolicyId '6'     | Should -BeFalse -Because 'the PolicyID must match the expected one'
                Test-RegEntryUsable -Key $full     -ExpectUrl $url -ExpectPolicyId $rowPid | Should -BeTrue
            }
            finally {
                # Leaf-only, one key at a time: each registered row (most recent first), then the
                # base. A row is a leaf by construction, so a subkey beneath one, or an unregistered
                # subkey beneath the base, makes that delete fail and warn instead of deleting it.
                # Every path is unregistered after its own teardown, so AfterAll cannot remove a key
                # that something else creates at one of these paths later.
                for ($i = $rows.Count - 1; $i -ge 0; $i--) {
                    if (Test-Path -LiteralPath $rows[$i]) { [void](script:Remove-LeafKey -Path $rows[$i]) }
                    [void]$script:CreatedKeys.Remove($rows[$i])
                }
                if ($baseCreated -and (Test-Path -LiteralPath $base)) { [void](script:Remove-LeafKey -Path $base) }
                if ($baseCreated) { [void]$script:CreatedKeys.Remove($base) }
            }
        }

        It 'GPUser under -WhatIf runs the AD objectGUID lookup and leaves the GP hive unchanged (no writes)' {
            # The GP hive may carry the user's real configuration: this case only previews and NEVER
            # writes there. The hive is snapshotted before and after as one canonical string: for the
            # root key and every subkey (recursively) each value's name, kind and data, plus the subkey
            # names. A regression that rewrites the root Flags or an existing AD row therefore fails
            # the comparison, not only one that adds or removes a name.
            $gp = 'HKCU:\Software\Policies\Microsoft\Cryptography\PolicyServers'
            function Get-GpShape {
                if (-not (Test-Path -LiteralPath $gp)) { return [pscustomobject]@{ Shape = '(absent)'; FlagsRaw = $null; FlagsKind = $null } }
                $lines = New-Object System.Collections.Generic.List[string]
                # Every free-text field goes through the suite's backslash encoder (see
                # ConvertTo-ShapeToken: no length limit, delimiters | ; , escaped, multi-string
                # element boundaries kept), so a name or datum cannot collide with the delimiters.
                function Add-KeyShape([Microsoft.Win32.RegistryKey]$Key, [string]$Rel) {
                    foreach ($n in @($Key.GetValueNames() | Sort-Object)) {
                        $data = $Key.GetValue($n, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                        $lines.Add("$(script:ConvertTo-ShapeToken $Rel)|value|$(script:ConvertTo-ShapeToken $n)|$($Key.GetValueKind($n))|$(script:ConvertTo-ShapeToken $data)")
                    }
                    foreach ($s in @($Key.GetSubKeyNames() | Sort-Object)) {
                        $lines.Add("$(script:ConvertTo-ShapeToken $Rel)|subkey|$(script:ConvertTo-ShapeToken $s)")
                        $child = $Key.OpenSubKey($s)
                        if ($child) { try { Add-KeyShape -Key $child -Rel "$Rel\$s" } finally { $child.Close() } }
                    }
                }
                $k = Get-Item -LiteralPath $gp
                Add-KeyShape -Key $k -Rel ''
                $raw  = $k.GetValue('Flags')
                $kind = if ($null -ne $raw) { $k.GetValueKind('Flags') } else { $null }
                [pscustomobject]@{ Shape = ($lines -join ';'); FlagsRaw = $raw; FlagsKind = $kind }
            }
            $before = Get-GpShape
            # The expected RootFlags is derived from the snapshot FIRST, by the suite's shared
            # oracle (Get-ExpectedRootFlags: the documented conversion contract, independent of the
            # script's helper, and the same one the Unit GPUser case uses). $null means UNUSABLE:
            # the script's step 0a refuses such a value BEFORE its -WhatIf gates, so that run
            # throws instead of returning a summary.
            $expectedFlags = script:Get-ExpectedRootFlags $before.FlagsRaw $before.FlagsKind
            if ($null -ne $expectedFlags) {
                $o = & $script:Cep -Url $script:LabUrl -PolicyName $script:LabName -Location GPUser -WhatIf 3>$null
                (Get-GpShape).Shape | Should -BeExactly $before.Shape -Because '-WhatIf must not change the GP hive'
                Test-Path -LiteralPath "$gp\$script:ExpectedKey" | Should -BeFalse
                $o.EntryApplied | Should -BeFalse
                $o.EntryAction  | Should -BeExactly 'Declined' -Because '-WhatIf declines the entry write at the confirmation gate'
                $o.RootFlags    | Should -BeExactly $expectedFlags
                $joined = $true
                try { $joined = [bool](Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).PartOfDomain } catch { }
                if ($joined) { $o.ADPolicyRow | Should -BeExactly 'not run' -Because 'the lookup succeeded and -WhatIf declined the row write' }
                else         { $o.ADPolicyRow | Should -BeExactly 'skipped (no domain)' }
            }
            else {
                # Refusal path: the run must stop with the step 0a message and write nothing, even
                # under -WhatIf. The hive is compared afterwards exactly as in the usable case.
                { & $script:Cep -Url $script:LabUrl -PolicyName $script:LabName -Location GPUser -WhatIf 3>$null } |
                    Should -Throw -ExpectedMessage '*root Flags*cannot be used*Nothing was written*'
                (Get-GpShape).Shape | Should -BeExactly $before.Shape -Because 'the refusal must leave the GP hive unchanged'
                Test-Path -LiteralPath "$gp\$script:ExpectedKey" | Should -BeFalse
            }
        }

        It 'the registry path check detects a SYMBOLIC LINK and refuses it (extracted checker, throwaway HKCU keys)' {
            # The REAL checker and its native shim, extracted from the script.
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:Cep, [ref]$null, [ref]$null)
            $shim = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.IfStatementAst] -and $n.Extent.Text -match 'CepRegNative' }, $true)
            $shim | Should -Not -BeNullOrEmpty
            . ([scriptblock]::Create($shim[0].Extent.Text))
            foreach ($name in 'Test-RegistryKeyIsLink', 'Test-RegistryComponentExists', 'Assert-ProtectedRegistryPath') {
                $def = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
                $def | Should -Not -BeNullOrEmpty
                . ([scriptblock]::Create($def[0].Extent.Text))
            }
            $hkcu = [IntPtr]::new(-2147483647)   # HKEY_CURRENT_USER
            $sid  = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
            $targetRel = "Software\$script:Prefix-linktarget"; $linkRel = "Software\$script:Prefix-link"
            # Ownership: the test deletes only what it created, so both throwaway names must be absent
            # first. The target is created with the native call and tracked (for the leaf-only
            # AfterAll delete) only on its REG_CREATED_NEW_KEY disposition.
            Test-Path -LiteralPath "HKCU:\$targetRel" | Should -BeFalse -Because 'the throwaway target key must not pre-exist'
            Test-Path -LiteralPath "HKCU:\$linkRel"   | Should -BeFalse -Because 'the throwaway link name must not pre-exist'
            $targetCreated = script:New-LeafKey -Path "HKCU:\$targetRel"
            if ($targetCreated) { $script:CreatedKeys.Add("HKCU:\$targetRel") }
            $targetCreated | Should -BeTrue -Because 'REG_CREATED_NEW_KEY (1): the target key must be new, not an opened existing key'
            # A VOLATILE registry symbolic link (REG_OPTION_VOLATILE 0x1 | REG_OPTION_CREATE_LINK 0x2) - gone at
            # reboot at the latest - whose REG_LINK (6) 'SymbolicLinkValue' names the throwaway target key.
            $h = [IntPtr]::Zero; $disp = [uint32]0
            [PesterRegLink]::RegCreateKeyExW($hkcu, $linkRel, 0, $null, 0x3, 0xF003F, [IntPtr]::Zero, [ref]$h, [ref]$disp) | Should -Be 0
            $disp | Should -Be 1 -Because 'REG_CREATED_NEW_KEY (1): the link key must be new, not an opened existing key'
            # From here the link key exists: ONE try/finally deletes it whatever fails below, including
            # the SymbolicLinkValue write (a link key without a target value would otherwise stay).
            try {
                try {
                    $linkTarget = [System.Text.Encoding]::Unicode.GetBytes("\Registry\User\$sid\$targetRel")
                    [PesterRegLink]::RegSetValueExW($h, 'SymbolicLinkValue', 0, 6, $linkTarget, [uint32]$linkTarget.Length) | Should -Be 0
                } finally { [void][PesterRegLink]::RegCloseKey($h) }
                (Get-Item -LiteralPath "HKCU:\$linkRel").Name | Should -Not -BeNullOrEmpty -Because 'the provider follows the link to the target - which is exactly why the checker must look at the link object'
                Test-RegistryKeyIsLink -HiveHandle $hkcu -SubKey $linkRel   | Should -BeTrue  -Because 'the key IS a link'
                Test-RegistryKeyIsLink -HiveHandle $hkcu -SubKey $targetRel | Should -BeFalse -Because 'a real key is not'
                Test-RegistryKeyIsLink -HiveHandle $hkcu -SubKey 'Software' | Should -BeFalse
                { Assert-ProtectedRegistryPath -Path "HKCU:\$linkRel\PolicyServers" }   | Should -Throw -ExpectedMessage '*SYMBOLIC LINK*'
                { Assert-ProtectedRegistryPath -Path "HKCU:\$targetRel\PolicyServers" } | Should -Not -Throw
            }
            finally {
                # Delete the LINK itself through a handle opened with REG_OPTION_OPEN_LINK - never its
                # target through the link (Remove-Item would follow it). A failure is named with the path.
                [void](script:Remove-VolatileLink -LinkRel $linkRel)
            }
        }

        It 'the path check catches a DANGLING symbolic link that Test-Path reports as absent (extracted checker, throwaway HKCU keys)' {
            # Regression for the round-6 finding: Assert-ProtectedRegistryPath must decide each component's
            # existence with the native REG_OPTION_OPEN_LINK probe, not Test-Path. The provider follows a
            # link to its target, so a link whose target does NOT exist reads as absent to Test-Path - which
            # would let the walk break BEFORE the link check and leave the link to be retargeted at a
            # protected key before the write. The native probe sees the link object itself.
            $ast  = [System.Management.Automation.Language.Parser]::ParseFile($script:Cep, [ref]$null, [ref]$null)
            $shim = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.IfStatementAst] -and $n.Extent.Text -match 'CepRegNative' }, $true)
            . ([scriptblock]::Create($shim[0].Extent.Text))
            foreach ($name in 'Test-RegistryKeyIsLink', 'Test-RegistryComponentExists', 'Assert-ProtectedRegistryPath') {
                $def = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
                $def | Should -Not -BeNullOrEmpty
                . ([scriptblock]::Create($def[0].Extent.Text))
            }
            $hkcu = [IntPtr]::new(-2147483647)   # HKEY_CURRENT_USER
            $sid  = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
            # A link whose SymbolicLinkValue names a target key that is never created: the link DANGLES.
            $danglingRel = "Software\$script:Prefix-dangling"; $absentRel = "Software\$script:Prefix-absent-target"
            Test-RegistryComponentExists -HiveHandle $hkcu -SubKey $danglingRel | Should -BeFalse -Because 'the throwaway link name must not pre-exist (as a key or as a link)'
            $h = [IntPtr]::Zero; $disp = [uint32]0
            [PesterRegLink]::RegCreateKeyExW($hkcu, $danglingRel, 0, $null, 0x3, 0xF003F, [IntPtr]::Zero, [ref]$h, [ref]$disp) | Should -Be 0   # VOLATILE | CREATE_LINK
            $disp | Should -Be 1 -Because 'REG_CREATED_NEW_KEY (1): the link key must be new; the teardown deletes only what this test created'
            # From here the link key exists: ONE try/finally deletes it whatever fails below, including
            # the SymbolicLinkValue write (a link key without a target value would otherwise stay).
            try {
                try {
                    $linkTarget = [System.Text.Encoding]::Unicode.GetBytes("\Registry\User\$sid\$absentRel")
                    [PesterRegLink]::RegSetValueExW($h, 'SymbolicLinkValue', 0, 6, $linkTarget, [uint32]$linkTarget.Length) | Should -Be 0
                } finally { [void][PesterRegLink]::RegCloseKey($h) }
                Test-Path -LiteralPath "HKCU:\$danglingRel" | Should -BeFalse -Because 'the provider follows the link to its ABSENT target - exactly the bypass the native probe closes'
                Test-RegistryComponentExists -HiveHandle $hkcu -SubKey $danglingRel | Should -BeTrue -Because 'the link OBJECT exists even though its target does not'
                Test-RegistryComponentExists -HiveHandle $hkcu -SubKey $absentRel   | Should -BeFalse -Because 'a genuinely absent key returns ERROR_FILE_NOT_FOUND'
                { Assert-ProtectedRegistryPath -Path "HKCU:\$danglingRel\PolicyServers" } | Should -Throw -ExpectedMessage '*SYMBOLIC LINK*'
            }
            finally {
                # OPEN_LINK | DELETE on the link object itself; a failure is named with the path.
                [void](script:Remove-VolatileLink -LinkRel $danglingRel)
            }
        }

        It 'Write-CepEntry updates an existing entry under a base key the caller can read but not write (the native create requests KEY_READ; extracted writer, throwaway HKCU keys)' {
            # Regression: New-RegistryKeyNative once requested KEY_READ | KEY_WRITE (0x20006) on the
            # base key. Before that, an update of an existing entry never opened the base key for
            # writing, so a protected but readable base key was fine; with 0x20006 the open failed
            # with ERROR_ACCESS_DENIED (5). The create now requests KEY_READ (0x20019): opening an
            # existing key needs no write right, and creating a missing key needs KEY_CREATE_SUB_KEY
            # on its parent, exactly as New-Item did. Fixture: a throwaway base key and an entry key
            # beneath it, both created natively and claimed only on REG_CREATED_NEW_KEY; then a
            # non-inheritable DENY of CreateSubKey and SetValue for the running account on the base
            # key alone (the entry key stays writable). The real Write-CepEntry and everything it
            # calls are extracted from the script and pointed at the throwaway base through $hive.
            $ast  = [System.Management.Automation.Language.Parser]::ParseFile($script:Cep, [ref]$null, [ref]$null)
            $shim = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.IfStatementAst] -and $n.Extent.Text -match 'CepRegNative' }, $true)
            $shim | Should -Not -BeNullOrEmpty
            . ([scriptblock]::Create($shim[0].Extent.Text))
            foreach ($name in 'Test-RegistryKeyIsLink', 'Test-RegistryComponentExists', 'Assert-ProtectedRegistryPath', 'New-RegistryKeyNative', 'New-BaseKey', 'Write-CepEntry') {
                $def = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
                $def | Should -Not -BeNullOrEmpty -Because "the script must define $name"
                . ([scriptblock]::Create($def[0].Extent.Text))
            }
            $typeName = [regex]::Match($shim[0].Extent.Text, "'(CepRegNative[A-Za-z0-9_]*)'\s+-as\s+\[type\]").Groups[1].Value
            $typeName | Should -Not -BeNullOrEmpty
            $native = $typeName -as [type]
            $hkcu   = [IntPtr]::new(-2147483647)   # HKEY_CURRENT_USER
            $me     = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
            $baseRel = "Software\$script:Prefix-denybase"
            $hive    = "HKCU:\$baseRel"                     # New-BaseKey reads $hive from the calling scope
            $entryPath = "$hive\entry"
            $regBase = "Registry::HKEY_CURRENT_USER\$baseRel"
            # The extracted New-BaseKey sets $script:baseKeyCreated only when its create call CREATES
            # the base key. The base exists throughout this case, so that assignment never runs and
            # the suite's own $script:BaseKeyCreated table (same name, case-insensitive) is untouched.
            # Asserted at the end.
            Test-Path -LiteralPath $hive | Should -BeFalse -Because "the fixture base key $hive must not pre-exist; it is not this run's and will not be touched"
            $baseCreated = $false; $entryCreated = $false; $denyRule = $null
            try {
                $baseCreated = script:New-LeafKey -Path $hive
                if ($baseCreated) { $script:CreatedKeys.Add($hive) }
                $baseCreated | Should -BeTrue -Because 'REG_CREATED_NEW_KEY (1): the fixture base key must be new'
                $entryCreated = script:New-LeafKey -Path $entryPath
                if ($entryCreated) { $script:CreatedKeys.Add($entryPath) }
                $entryCreated | Should -BeTrue -Because 'REG_CREATED_NEW_KEY (1): the fixture entry key must be new'
                # DENY CreateSubKey | SetValue for the running account on the base key only (no
                # inheritance): the entry key keeps its inherited Allow entries. The path check skips
                # Deny entries, and the owner is the running account, so the check still passes.
                $denyRule = New-Object System.Security.AccessControl.RegistryAccessRule($me, 'CreateSubKey, SetValue', 'None', 'None', 'Deny')
                $acl = Get-Acl -Path $regBase
                $acl.AddAccessRule($denyRule)
                Set-Acl -Path $regBase -AclObject $acl
                # Probe 1: the fixture reproduces the regression - KEY_READ | KEY_WRITE on the base is refused.
                $h = [IntPtr]::Zero; $disp = [uint32]0
                $native::RegCreateKeyExW($hkcu, $baseRel, 0, $null, 0, 0x20006, [IntPtr]::Zero, [ref]$h, [ref]$disp) | Should -Be 5 -Because 'ERROR_ACCESS_DENIED: the base key denies the running account SetValue and CreateSubKey, so a write-class open must fail'
                if ($h -ne [IntPtr]::Zero) { [void]$native::RegCloseKey($h) }
                # Probe 2: KEY_READ opens the same existing key, with the REG_OPENED_EXISTING_KEY disposition.
                $h = [IntPtr]::Zero; $disp = [uint32]0
                $native::RegCreateKeyExW($hkcu, $baseRel, 0, $null, 0, 0x20019, [IntPtr]::Zero, [ref]$h, [ref]$disp) | Should -Be 0 -Because 'opening an existing key with KEY_READ needs no write right'
                $disp | Should -Be 2
                [void]$native::RegCloseKey($h)
                { Assert-ProtectedRegistryPath -Path $entryPath } | Should -Not -Throw -Because 'a Deny entry for the running account is not a write-class grant to an untrusted principal'
                # The real writer: the base key is opened (not created), the entry key is opened (not
                # created), the path is checked again, and the values are written through the provider.
                $created = Write-CepEntry -EntryPath $entryPath -Strings @{ URL = $script:LabUrl; PolicyID = $script:ExpectedPid; FriendlyName = $script:LabName } -Dwords @{ Flags = 0x14; AuthFlags = 2; Cost = 1 }
                $created | Should -BeFalse -Because 'the entry key existed: REG_OPENED_EXISTING_KEY, so the caller reports Updated'
                $k = Get-Item -LiteralPath $entryPath
                $k.GetValue('URL')      | Should -BeExactly $script:LabUrl
                $k.GetValue('PolicyID') | Should -BeExactly $script:ExpectedPid
                $k.GetValueKind('PolicyID') | Should -Be ([Microsoft.Win32.RegistryValueKind]::String)
                [int]$k.GetValue('Flags') | Should -Be 0x14
                $k.GetValueKind('Cost')   | Should -Be ([Microsoft.Win32.RegistryValueKind]::DWord)
                (Get-Item -LiteralPath $hive).ValueCount | Should -Be 0 -Because 'the writer never writes a value on the base key'
                $script:BaseKeyCreated | Should -BeOfType [hashtable] -Because 'the extracted New-BaseKey must not have replaced the suite''s bookkeeping table (the base existed, so it never set its flag)'
            }
            finally {
                # Remove the Deny entry first (a delete needs no denied right, but the fixture must
                # not leave a foreign-looking ACE behind), then the entry, then the base: leaf-only,
                # each only if this test created it, and each unregistered after its own teardown.
                if ($denyRule -and (Test-Path -LiteralPath $hive)) {
                    # The provider's Set-Acl opens the key for WRITE, which the Deny entry blocks, so it
                    # raises a non-terminating error. Open the key for permission changes only, which
                    # the entry does not deny, and remove the rule through that handle.
                    try {
                        $root = if ($hive -like 'HKLM:*') { [Microsoft.Win32.Registry]::LocalMachine } else { [Microsoft.Win32.Registry]::CurrentUser }
                        $rel  = $hive -replace '^HK(CU|LM):\\', ''
                        $rk = $root.OpenSubKey($rel, [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree, [System.Security.AccessControl.RegistryRights]'ChangePermissions, ReadPermissions')
                        if (-not $rk) { throw "OpenSubKey returned no key for '$rel'" }
                        try {
                            $sec = $rk.GetAccessControl()
                            $null = $sec.RemoveAccessRule($denyRule)
                            $rk.SetAccessControl($sec)
                        }
                        finally { $rk.Close() }
                    }
                    catch { Write-Warning "Could not remove the Deny entry from ${hive}: $_" }
                }
                if ($entryCreated -and (Test-Path -LiteralPath $entryPath)) { [void](script:Remove-LeafKey -Path $entryPath) }
                if ($entryCreated) { [void]$script:CreatedKeys.Remove($entryPath) }
                if ($baseCreated -and (Test-Path -LiteralPath $hive)) { [void](script:Remove-LeafKey -Path $hive) }
                if ($baseCreated) { [void]$script:CreatedKeys.Remove($hive) }
            }
            Test-Path -LiteralPath $hive | Should -BeFalse -Because 'the test removed the keys it created'
        }

        It 'LocalMachine add/remove round-trip (elevated session)' -Skip:(-not $script:LabElevated) {
            # Ownership: the entry must be absent before the script creates it. The path is tracked
            # in the finally ONLY when the script's summary reports that it created the key (EntryAction = Created), and is
            # unregistered again once its removal is proven.
            $entryPath = "$script:HiveLM\$script:ExpectedKey"
            Test-Path -LiteralPath $entryPath | Should -BeFalse -Because 'the entry must not pre-exist under the machine hive; it is not this run''s and will not be touched'
            $o = $null
            try {
                $o = & $script:Cep -Url $script:LabUrl -PolicyName $script:LabName -Location LocalMachine -Confirm:$false 3>$null
            }
            finally {
                script:Register-CreatedEntry -Path $entryPath -Summary $o
            }
            $o.EntryApplied | Should -BeTrue
            $o.EntryAction  | Should -BeExactly 'Created' -Because 'the entry was proven absent before the run wrote it'
            $o.BaseKeyCreated | Should -Be (-not $script:PreLM.Existed) -Because 'the script must report that it created the base key exactly when the snapshot found none'
            Test-Path -LiteralPath $entryPath | Should -BeTrue
            (Get-Item -LiteralPath $entryPath).GetValue('PolicyID') | Should -BeExactly $script:ExpectedPid

            # The entry's ownership is consumed before the call (also when the script removes the
            # entry and then throws in the default-marker step); the summary reports the entry
            # removed, so the path stays unregistered afterwards.
            $o = script:Invoke-RemovalOwned -Paths @($entryPath) -Reregister $script:KeepUnlessRemovedEntry -Run {
                & $script:Cep -Url $script:LabUrl -Location LocalMachine -Remove -Confirm:$false 3>$null
            }
            script:Register-MarkerChange -Hive $script:HiveLM -Summary $o
            # First: prove the entry gone, before any assertion on the summary.
            script:Confirm-RemovedEntry -Path $entryPath
            $o.RemovedEntry | Should -BeTrue
        }
    }
}
