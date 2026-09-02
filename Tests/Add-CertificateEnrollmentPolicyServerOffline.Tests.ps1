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
                   URL under the RFC-reserved .invalid TLD (can never reach anything real), every
                   created key is tracked by exact path and removed in teardown, a pre-existing
                   (Default) marker is snapshotted and restored, and a PolicyServers base key is
                   removed only if this run created it and it ends the run empty. The GP-hive
                   locations (GPMachine/GPUser) are deliberately NOT exercised: on a domain
                   member they tattoo pseudo-policy backed by no GPO (the script itself warns) -
                   the GPO suite covers Group Policy delivery against a throwaway unlinked GPO.

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

            # Snapshot shared state so teardown can restore it EXACTLY (marker) and knows whether
            # this run created the base key itself (remove it again only then, and only if empty).
            function script:Get-HiveSnapshot([string]$Hive) {
                $existed = Test-Path -LiteralPath $Hive
                @{ Existed = $existed; Marker = if ($existed) { (Get-Item -LiteralPath $Hive).GetValue('') } else { $null } }
            }
            $script:PreCU = script:Get-HiveSnapshot $script:HiveCU
            $script:PreLM = script:Get-HiveSnapshot $script:HiveLM

            # Exact key paths this run creates - the ONLY things teardown deletes outright.
            $script:CreatedKeys = New-Object System.Collections.Generic.List[string]

            function script:Remove-MarkerValue([string]$Hive) {
                $rel  = $Hive -replace '^HK(CU|LM):\\', ''
                $root = if ($Hive -like 'HKCU:*') { [Microsoft.Win32.Registry]::CurrentUser } else { [Microsoft.Win32.Registry]::LocalMachine }
                $k = $root.OpenSubKey($rel, $true)
                if ($k) { try { $k.DeleteValue('', $false) } finally { $k.Close() } }
            }
        }

        AfterAll {
            # 1) Surgical: remove ONLY the exact keys this run created (most-recent first).
            for ($i = $script:CreatedKeys.Count - 1; $i -ge 0; $i--) {
                try { if (Test-Path -LiteralPath $script:CreatedKeys[$i]) { Remove-Item -LiteralPath $script:CreatedKeys[$i] -Recurse -Force } } catch { }
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
                                try { Remove-Item -LiteralPath $k.PSPath -Recurse -Force } catch { }
                                Write-Warning "AfterAll safety-net removed an untracked test entry: $($k.PSChildName) ($u)"
                            }
                        }
                    }
                }
            } else {
                Write-Warning "Safety-net sweep skipped: run URL base is unset or malformed ('$script:UrlBase')."
            }
            # 3) Restore each hive's (Default) marker to its snapshotted state, and remove a base
            #    key ONLY if this run created it and it ends the run empty.
            foreach ($pair in @(@($script:HiveCU, $script:PreCU), @($script:HiveLM, $script:PreLM))) {
                $hive = $pair[0]; $pre = $pair[1]
                if (-not (Test-Path -LiteralPath $hive)) { continue }
                try {
                    # $null-aware on BOTH sides: GetValue('') returns $null only when no (Default)
                    # value exists - a present-but-empty ('') or zero (0) marker must be RESTORED,
                    # not misread as absent (truthiness would conflate the two).
                    $cur = (Get-Item -LiteralPath $hive).GetValue('')
                    if ($null -ne $pre.Marker) {
                        if ($null -eq $cur -or "$cur" -cne "$($pre.Marker)") { Set-ItemProperty -LiteralPath $hive -Name '(default)' -Value $pre.Marker }
                    } elseif ($null -ne $cur) {
                        script:Remove-MarkerValue $hive
                    }
                } catch { }
                if (-not $pre.Existed) {
                    $k = Get-Item -LiteralPath $hive
                    if ($k.SubKeyCount -eq 0 -and $k.ValueCount -eq 0) {
                        try { Remove-Item -LiteralPath $hive } catch { }
                    }
                }
            }
        }

        It 'Add writes the entry; every value verified against independent oracles' {
            # Track BEFORE invoking: the target path is precomputable, and a mid-write throw would
            # otherwise orphan an untracked partial key (teardown step 1 is Test-Path-guarded, so
            # pre-registering a key that never materializes is harmless).
            $script:CreatedKeys.Add("$script:HiveCU\$script:ExpectedKey")
            $o = & $script:Cep -Url $script:LabUrl -PolicyName $script:LabName -Location LocalUser -Confirm:$false 3>$null
            $o.EntryApplied | Should -BeTrue
            $o.Path | Should -BeExactly "$script:HiveCU\$script:ExpectedKey"
            $k = Get-Item -LiteralPath "$script:HiveCU\$script:ExpectedKey"
            $k.GetValue('URL')          | Should -BeExactly $script:LabUrl
            $k.GetValue('PolicyID')     | Should -BeExactly $script:ExpectedPid
            $k.GetValue('FriendlyName') | Should -BeExactly $script:LabName
            [int]$k.GetValue('Flags')     | Should -Be 0x14
            [int]$k.GetValue('AuthFlags') | Should -Be 0x2
            [int]$k.GetValue('Cost')      | Should -Be 0x7FFFFFFD
        }

        It 'rerunning the same Add is an idempotent update (values unchanged, EntryApplied)' {
            $o = & $script:Cep -Url $script:LabUrl -PolicyName $script:LabName -Location LocalUser -Confirm:$false 3>$null
            $o.EntryApplied | Should -BeTrue
            $k = Get-Item -LiteralPath "$script:HiveCU\$script:ExpectedKey"
            $k.GetValue('PolicyID') | Should -BeExactly $script:ExpectedPid
            [int]$k.GetValue('Flags') | Should -Be 0x14
        }

        It '-SetAsDefault writes the (Default) marker with this PolicyID' {
            $o = & $script:Cep -Url $script:LabUrl -PolicyName $script:LabName -Location LocalUser -SetAsDefault -Confirm:$false 3>$null
            $o.DefaultChanged | Should -BeTrue
            (Get-Item -LiteralPath $script:HiveCU).GetValue('') | Should -BeExactly $script:ExpectedPid
        }

        It '-ReplaceExisting removes a same-PolicyID sibling with a different URL' {
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
            $script:CreatedKeys.Add("$script:HiveCU\$sibKey")   # track BEFORE invoking (see first Add)
            $null = & $script:Cep -Url $sibUrl -PolicyName $script:LabName -PolicyId $script:ExpectedPid -Location LocalUser -Confirm:$false 3>$null
            Test-Path -LiteralPath "$script:HiveCU\$sibKey" | Should -BeTrue

            $o = & $script:Cep -Url $script:LabUrl -PolicyName $script:LabName -Location LocalUser -ReplaceExisting -Confirm:$false 3>$null
            @($o.DuplicatesRemoved) | Should -Contain $sibUrl
            Test-Path -LiteralPath "$script:HiveCU\$sibKey" | Should -BeFalse
            Test-Path -LiteralPath "$script:HiveCU\$script:ExpectedKey" | Should -BeTrue
        }

        It '-Remove deletes the entry and clears the now-orphaned (Default) marker' {
            $o = & $script:Cep -Url $script:LabUrl -Location LocalUser -Remove -Confirm:$false 3>$null
            $o.RemovedEntry   | Should -BeTrue
            $o.DefaultCleared | Should -BeTrue
            Test-Path -LiteralPath "$script:HiveCU\$script:ExpectedKey" | Should -BeFalse
            (Get-Item -LiteralPath $script:HiveCU).GetValue('') | Should -BeNullOrEmpty
        }

        It 'LocalMachine add/remove round-trip (elevated session)' -Skip:(-not $script:LabElevated) {
            $script:CreatedKeys.Add("$script:HiveLM\$script:ExpectedKey")   # track BEFORE invoking (see first Add)
            $o = & $script:Cep -Url $script:LabUrl -PolicyName $script:LabName -Location LocalMachine -Confirm:$false 3>$null
            $o.EntryApplied | Should -BeTrue
            (Get-Item -LiteralPath "$script:HiveLM\$script:ExpectedKey").GetValue('PolicyID') | Should -BeExactly $script:ExpectedPid

            $o = & $script:Cep -Url $script:LabUrl -Location LocalMachine -Remove -Confirm:$false 3>$null
            $o.RemovedEntry | Should -BeTrue
            Test-Path -LiteralPath "$script:HiveLM\$script:ExpectedKey" | Should -BeFalse
        }
    }
}
