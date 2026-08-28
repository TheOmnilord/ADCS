<#
.SYNOPSIS
    Pester suite for Add-CertificateEnrollmentPolicyServerOffline.ps1. Requires Pester 5+.

.DESCRIPTION
    Three always-on tiers (no opt-in Lab tier - the script's only side effects are registry
    writes, and every tier here avoids them):

      -Tag Unit    Exercises the REAL derivations (SHA-1/UTF-16LE subkey, EJBCA String.hashCode()
                   PolicyID, Flags/AuthFlags math) by invoking the script under -WhatIf against
                   the per-user hive (-Location LocalUser, which needs no elevation) and reading
                   the emitted summary object. -WhatIf makes every registry write a no-op, so the
                   tier changes nothing. No AD, no modules, no elevation.
      -Tag Static  The script parses and its comment-based help binds. No AD, no modules.
      -Tag Guard   Parameter-conflict validation. These invocations throw BEFORE the elevation
                   check and before any registry write, so they need neither elevation nor a
                   reachable registry hive and make no changes.

    Oracle constants (independently derived; the ldap: subkey matches the script's own AD_KEY
    documentation, which cross-checks the SHA-1/UTF-16LE method):
      * PolicyID  Java String.hashCode('Example PKI Service')                     = 241064013
      * Subkey    SHA-1(UTF-16LE(lowercased URL)) of the sample EJBCA CEP URL     = dc032f3a...

.EXAMPLE
    Invoke-Pester -Path .\Tests\Add-CertificateEnrollmentPolicyServerOffline.Tests.ps1

.EXAMPLE
    # Parse/help/derivation only - no registry access at all:
    Invoke-Pester -Path .\Tests\Add-CertificateEnrollmentPolicyServerOffline.Tests.ps1 -Tag Unit,Static
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '',
    Justification = 'the container parameter is consumed inside Pester Describe/BeforeAll scriptblocks, which the analyzer cannot see through')]
param(
    [string] $ScriptPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'Add-CertificateEnrollmentPolicyServerOffline.ps1')
)

Describe 'Add-CertificateEnrollmentPolicyServerOffline' {

    BeforeAll {
        $script:Cep = $ScriptPath
        $script:Cep | Should -Exist

        $script:KnownUrl    = 'https://pki.example.net/ejbca/msae/CEPService?alias'
        $script:KnownName   = 'Example PKI Service'
        $script:KnownPid    = '241064013'
        $script:KnownSubkey = 'dc032f3a68521c2445e1e161da81503bddce17a7'

        # Invoke under -WhatIf against the per-user hive (no elevation, no writes) and return the
        # summary object. All non-output streams are silenced so the "What if:" lines stay quiet.
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
}
