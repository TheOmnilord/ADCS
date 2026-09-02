<#
.SYNOPSIS
    Pester suite for Set-ADCSTemplateValidity.ps1. Requires Pester 5+.

.DESCRIPTION
    Three always-on tiers plus one opt-in tier:

      -Tag Unit    Pure helpers extracted from the script by AST (so the REAL code runs, never a
                   copy): the Years/Months/Weeks/Days/Hours day math, the pKIExpirationPeriod
                   byte encoding (negative-FILETIME ticks, little-endian), the human-readable
                   decoder, and the wildcard-preserving LDAP filter escaper. The byte oracles are
                   Windows' own: 1 year and 6 weeks must produce exactly the stock Kerberos
                   Authentication template's pKIExpirationPeriod/pKIOverlapPeriod bytes. No AD.
      -Tag Static  The script parses and its comment-based help binds (non-vacuously). No AD.
      -Tag Guard   Parameter-conflict validation that throws in begin{} BEFORE any AD/LDAP
                   connection is attempted. Runs on CI (no domain needed).
      -Tag Lab     LIVE modification of a throwaway template in AD. Skipped unless -RunLab is
                   passed; needs the RSAT ActiveDirectory module and a writable DC. Surgical by
                   construction: ONE bare pKICertificateTemplate object named PESTER-<hex>-VAL is
                   created, every wildcard the tests use is scoped under the run-unique prefix,
                   and teardown removes the exact tracked DN (with a prefix-scoped, structurally
                   guarded safety net). Pre-existing templates are never touched.

.EXAMPLE
    Invoke-Pester -Path .\Tests\Set-ADCSTemplateValidity.Tests.ps1 -ExcludeTag Lab

.EXAMPLE
    # Full run against the lab AD (uses the current domain's DC unless -LabServer is given):
    $cfg = New-PesterContainer -Path .\Tests\Set-ADCSTemplateValidity.Tests.ps1 -Data @{ RunLab = $true }
    Invoke-Pester -Container $cfg
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '',
    Justification = 'container parameters are consumed inside Pester Describe/BeforeAll scriptblocks, which the analyzer cannot see through')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingEmptyCatchBlock', '',
    Justification = 'best-effort teardown paths (AfterAll object removal) deliberately swallow per-item errors')]
param(
    [bool]   $RunLab     = $false,
    [string] $ScriptPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'Set-ADCSTemplateValidity.ps1'),
    [string] $LabServer  = ''    # Lab: a writable DC; empty = serverless (current domain) binds
)

BeforeDiscovery {
    # -Skip conditions are evaluated during discovery, so the gates must be set here.
    $script:LabReady = $RunLab -and [bool](Get-Module -ListAvailable ActiveDirectory)
}

Describe 'Set-ADCSTemplateValidity' {

    BeforeAll {
        $script:Val = $ScriptPath
        $script:Val | Should -Exist

        # --- AST-extract the pure helpers so the Unit tier exercises the REAL code ------------
        # (The script's begin block connects to AD, so it cannot be dot-sourced wholesale;
        # extracting the function bodies runs them with no side effects.)
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:Val, [ref]$null, [ref]$null)
        foreach ($name in 'ConvertTo-PKIPeriodDays', 'ConvertTo-PKIPeriodBytes', 'ConvertFrom-PKIPeriodBytes', 'ConvertTo-LdapFilterValue') {
            $def = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
            if ($def) { . ([scriptblock]::Create($def[0].Extent.Text)) }
        }

        function script:ToHex([byte[]]$Bytes) { -join ($Bytes | ForEach-Object { $_.ToString('x2') }) }
    }

    Context 'Unit: period math and encoding' -Tag 'Unit' {

        It 'converts every unit to days (AD convention: 365/year, 30/month)' {
            ConvertTo-PKIPeriodDays -Period 2  -PeriodUnit Years  | Should -Be 730
            ConvertTo-PKIPeriodDays -Period 3  -PeriodUnit Months | Should -Be 90
            ConvertTo-PKIPeriodDays -Period 6  -PeriodUnit Weeks  | Should -Be 42
            ConvertTo-PKIPeriodDays -Period 47 -PeriodUnit Days   | Should -Be 47
            ConvertTo-PKIPeriodDays -Period 12 -PeriodUnit Hours  | Should -Be 0.5
        }

        It 'encodes 1 year / 6 weeks to the EXACT bytes Windows puts on stock templates' {
            # Oracle: the built-in Kerberos Authentication template carries these values.
            script:ToHex (ConvertTo-PKIPeriodBytes -Period 1 -PeriodUnit Years) | Should -BeExactly '004039872ee1feff'
            script:ToHex (ConvertTo-PKIPeriodBytes -Period 6 -PeriodUnit Weeks) | Should -BeExactly '0080a60affdeffff'
        }

        It 'encoding is 8 bytes of negative little-endian ticks for any unit' {
            foreach ($case in @(@(200, 'Days'), @(100, 'Days'), @(47, 'Days'), @(12, 'Hours'))) {
                $b = ConvertTo-PKIPeriodBytes -Period $case[0] -PeriodUnit $case[1]
                $b.Length | Should -Be 8
                [System.BitConverter]::ToInt64($b, 0) | Should -BeLessThan 0
            }
        }

        It 'decode(encode(x)) round-trips the SC-081 milestones and common units' {
            ConvertFrom-PKIPeriodBytes -Bytes (ConvertTo-PKIPeriodBytes -Period 200 -PeriodUnit Days)  | Should -BeExactly '200 day(s)'
            ConvertFrom-PKIPeriodBytes -Bytes (ConvertTo-PKIPeriodBytes -Period 100 -PeriodUnit Days)  | Should -BeExactly '100 day(s)'
            ConvertFrom-PKIPeriodBytes -Bytes (ConvertTo-PKIPeriodBytes -Period 47  -PeriodUnit Days)  | Should -BeExactly '47 day(s)'
            ConvertFrom-PKIPeriodBytes -Bytes (ConvertTo-PKIPeriodBytes -Period 2   -PeriodUnit Years) | Should -BeExactly '2 year(s)'
            ConvertFrom-PKIPeriodBytes -Bytes (ConvertTo-PKIPeriodBytes -Period 6   -PeriodUnit Weeks) | Should -BeExactly '6 week(s)'
            ConvertFrom-PKIPeriodBytes -Bytes (ConvertTo-PKIPeriodBytes -Period 12  -PeriodUnit Hours) | Should -BeExactly '12 hour(s)'
        }

        It 'decoder prefers the largest clean unit (90 days reads as 3 months by design)' {
            ConvertFrom-PKIPeriodBytes -Bytes (ConvertTo-PKIPeriodBytes -Period 90 -PeriodUnit Days) | Should -BeExactly '3 month(s)'
        }

        It 'decoder returns N/A for null or non-8-byte input' {
            ConvertFrom-PKIPeriodBytes -Bytes $null | Should -BeExactly 'N/A'
            ConvertFrom-PKIPeriodBytes -Bytes ([byte[]](1, 2, 3)) | Should -BeExactly 'N/A'
        }

        It 'LDAP escaper protects metacharacters but PRESERVES the * and ? wildcards' {
            ConvertTo-LdapFilterValue -Value 'a(b)c' | Should -BeExactly 'a\28b\29c'
            ConvertTo-LdapFilterValue -Value 'a\b'   | Should -BeExactly 'a\5cb'
            ConvertTo-LdapFilterValue -Value 'Web*'  | Should -BeExactly 'Web*'
            ConvertTo-LdapFilterValue -Value 'U?er'  | Should -BeExactly 'U?er'
        }
    }

    Context 'Static: parse and help' -Tag 'Static' {

        It 'parses without errors' {
            $errs = $null
            $null = [System.Management.Automation.Language.Parser]::ParseFile($script:Val, [ref]$null, [ref]$errs)
            $errs | Should -BeNullOrEmpty
        }

        It 'comment-based help binds (Synopsis is real, not auto-generated syntax)' {
            $syn = (Get-Help $script:Val).Synopsis.Trim()
            $syn | Should -Not -BeNullOrEmpty
            $syn | Should -Not -Match '\[\[-|\[<CommonParameters>\]'
        }

        It 'carries a PSScriptInfo header (Test-ScriptFileInfo parses it; Version is semver)' {
            $info = Test-ScriptFileInfo -Path $script:Val -ErrorAction Stop
            $info.Version | Should -Match '^\d+\.\d+\.\d+$'
            $info.Guid    | Should -Not -BeNullOrEmpty
        }
        It 'documents every non-common parameter' {
            $cmd = Get-Command $script:Val
            $common = [System.Management.Automation.PSCmdlet]::CommonParameters + [System.Management.Automation.PSCmdlet]::OptionalCommonParameters
            $documented = @((Get-Help $script:Val).parameters.parameter.name)
            foreach ($p in $cmd.Parameters.Keys | Where-Object { $_ -notin $common }) {
                $documented | Should -Contain $p -Because "parameter -$p should have a .PARAMETER help entry"
            }
        }
    }

    Context 'Guard: validation before any AD connection' -Tag 'Guard' {

        It 'rejects -OverlapPeriod without -OverlapPeriodUnit' {
            { & $script:Val -TemplateName x -ValidityPeriod 1 -ValidityPeriodUnit Years -OverlapPeriod 6 -WhatIf } |
                Should -Throw -ExpectedMessage '*must both be specified together*'
        }

        It 'rejects -OverlapPeriodUnit without -OverlapPeriod' {
            { & $script:Val -TemplateName x -ValidityPeriod 1 -ValidityPeriodUnit Years -OverlapPeriodUnit Weeks -WhatIf } |
                Should -Throw -ExpectedMessage '*must both be specified together*'
        }

        It 'rejects an overlap that is not shorter than the validity' {
            { & $script:Val -TemplateName x -ValidityPeriod 30 -ValidityPeriodUnit Days -OverlapPeriod 1 -OverlapPeriodUnit Months -WhatIf } |
                Should -Throw -ExpectedMessage '*must be shorter than*'
        }
    }

    # -------------------------------------------------------------------------------------------
    # Lab tier: live modification of ONE throwaway template object. Opt-in (-RunLab).
    # Tests are SEQUENTIAL: modify -> idempotent -> -WhatIf -> wildcard dedup, against the same
    # tracked object.
    # -------------------------------------------------------------------------------------------
    Context 'Lab: live template modification' -Tag 'Lab' -Skip:(-not $script:LabReady) {

        BeforeAll {
            # Prefix and teardown-consumed state FIRST - before anything that can throw.
            $script:Prefix  = "PESTER-$([guid]::NewGuid().ToString('N').Substring(0,8))"
            $script:Created = New-Object System.Collections.Generic.List[string]   # exact DNs

            Import-Module ActiveDirectory -ErrorAction Stop
            $script:AP = @{}
            if ($LabServer) { $script:AP.Server = $LabServer }
            $cfgNc = (Get-ADRootDSE @script:AP).configurationNamingContext
            $script:TplBase = "CN=Certificate Templates,CN=Public Key Services,CN=Services,$cfgNc"

            # ONE bare throwaway template: 2-year validity, 6-week overlap, revision 0.
            # Breadcrumb BEFORE creation: teardown never runs on a hard process kill, and later
            # runs' prefix-scoped sweeps can never match this run's prefix - so the exact DN must
            # survive in the console/CI log for manual, exact-name removal after an abort.
            $script:TplName = "$script:Prefix-VAL"
            $script:TplDn   = "CN=$script:TplName,$script:TplBase"
            $script:Created.Add($script:TplDn)
            Write-Host "Lab object about to be created: $script:TplDn (remove manually by this exact DN if the run is killed before teardown)"
            New-ADObject @script:AP -Name $script:TplName -Type pKICertificateTemplate -Path $script:TplBase -OtherAttributes @{
                displayName                     = "$script:TplName (throwaway)"
                revision                        = 100
                'msPKI-Template-Minor-Revision' = 0
                pKIExpirationPeriod             = ([byte[]](ConvertTo-PKIPeriodBytes -Period 2 -PeriodUnit Years))
                pKIOverlapPeriod                = ([byte[]](ConvertTo-PKIPeriodBytes -Period 6 -PeriodUnit Weeks))
            }

            function script:Get-LabTpl {
                Get-ADObject @script:AP -Identity $script:TplDn -Properties pKIExpirationPeriod, pKIOverlapPeriod, 'msPKI-Template-Minor-Revision'
            }
        }

        AfterAll {
            # Surgical: remove ONLY the exact DNs this run created. Report failures truthfully -
            # a swallowed error must never masquerade as a successful removal.
            foreach ($dn in $script:Created) {
                try { Remove-ADObject @script:AP -Identity $dn -Confirm:$false }
                catch { Write-Warning "Teardown could NOT remove tracked object ${dn}: $_" }
            }
            # Safety net, scoped to THIS run's fresh-GUID prefix. STRUCTURAL GUARD: only with a
            # fully-formed prefix - an unset one would widen the LDAP filter. Never widen this.
            if ($script:Prefix -match '^PESTER-[0-9a-f]{8}$' -and $script:TplBase) {
                foreach ($t in @(Get-ADObject @script:AP -SearchBase $script:TplBase -LDAPFilter "(cn=$script:Prefix-*)" -ErrorAction SilentlyContinue)) {
                    try {
                        Remove-ADObject @script:AP -Identity $t.DistinguishedName -Confirm:$false
                        Write-Warning "AfterAll safety-net removed a leftover test object: $($t.DistinguishedName)"
                    }
                    catch { Write-Warning "AfterAll safety-net could NOT remove $($t.DistinguishedName): $_ - remove manually by this exact DN." }
                }
            }
        }

        It 'modifies validity and overlap, bumps the minor revision, and reports the transition' {
            $p = @{ TemplateName = $script:TplName; ValidityPeriod = 200; ValidityPeriodUnit = 'Days'
                    OverlapPeriod = 4; OverlapPeriodUnit = 'Weeks'; Confirm = $false }
            if ($LabServer) { $p.Server = $LabServer }
            $out = @(& $script:Val @p 6>$null)
            $out.Count | Should -Be 1
            $out[0].Status           | Should -BeExactly 'Modified'
            $out[0].PreviousValidity | Should -BeExactly '2 year(s)'
            $out[0].NewValidity      | Should -BeExactly '200 Days'
            $out[0].PreviousOverlap  | Should -BeExactly '6 week(s)'

            # Independent read-back: AD bytes must equal the reference encoding, revision bumped.
            $tpl = script:Get-LabTpl
            script:ToHex ([byte[]]$tpl.pKIExpirationPeriod) | Should -BeExactly (script:ToHex (ConvertTo-PKIPeriodBytes -Period 200 -PeriodUnit Days))
            script:ToHex ([byte[]]$tpl.pKIOverlapPeriod)    | Should -BeExactly (script:ToHex (ConvertTo-PKIPeriodBytes -Period 4 -PeriodUnit Weeks))
            [int]$tpl.'msPKI-Template-Minor-Revision' | Should -Be 1
        }

        It 'rerunning with the same values reports Already set and does NOT bump the revision' {
            $p = @{ TemplateName = $script:TplName; ValidityPeriod = 200; ValidityPeriodUnit = 'Days'
                    OverlapPeriod = 4; OverlapPeriodUnit = 'Weeks'; Confirm = $false }
            if ($LabServer) { $p.Server = $LabServer }
            $out = @(& $script:Val @p 6>$null)
            $out[0].Status | Should -BeExactly 'Already set'
            [int](script:Get-LabTpl).'msPKI-Template-Minor-Revision' | Should -Be 1
        }

        It '-WhatIf reports Skipped and changes nothing in AD' {
            $p = @{ TemplateName = $script:TplName; ValidityPeriod = 47; ValidityPeriodUnit = 'Days'; WhatIf = $true }
            if ($LabServer) { $p.Server = $LabServer }
            $out = @(& $script:Val @p 6>$null)
            $out[0].Status | Should -BeExactly 'Skipped'
            $tpl = script:Get-LabTpl
            script:ToHex ([byte[]]$tpl.pKIExpirationPeriod) | Should -BeExactly (script:ToHex (ConvertTo-PKIPeriodBytes -Period 200 -PeriodUnit Days))
            [int]$tpl.'msPKI-Template-Minor-Revision' | Should -Be 1
        }

        It 'two overlapping wildcard patterns match the template once (deduplication), with no false not-found warning' {
            $p = @{ TemplateName = @("$script:Prefix-V*", $script:TplName); ValidityPeriod = 100; ValidityPeriodUnit = 'Days'; Confirm = $false }
            if ($LabServer) { $p.Server = $LabServer }
            $mixed = @(& $script:Val @p 3>&1 6>$null)
            $warnings = @($mixed | Where-Object { $_ -is [System.Management.Automation.WarningRecord] })
            $out = @($mixed | Where-Object { $_ -isnot [System.Management.Automation.WarningRecord] })
            $out.Count | Should -Be 1
            $out[0].Status | Should -BeExactly 'Modified'
            # The deduplicated second pattern DID match - it must not warn "no templates found".
            $warnings | Should -BeNullOrEmpty
            [int](script:Get-LabTpl).'msPKI-Template-Minor-Revision' | Should -Be 2
        }

        It 'a run-scoped pattern with no matches warns and returns nothing' {
            $p = @{ TemplateName = "$script:Prefix-NOMATCH*"; ValidityPeriod = 1; ValidityPeriodUnit = 'Years'; Confirm = $false }
            if ($LabServer) { $p.Server = $LabServer }
            $out = @(& $script:Val @p 3>&1 6>$null)
            @($out | Where-Object { $_ -is [System.Management.Automation.WarningRecord] }) | Should -Not -BeNullOrEmpty
            @($out | Where-Object { $_ -isnot [System.Management.Automation.WarningRecord] }).Count | Should -Be 0
        }
    }
}
