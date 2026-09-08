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
      -Tag Guard   Parameter-conflict validation that throws BEFORE any AD/LDAP connection is
                   attempted, plus the unreachable -Server refusal (a dead host name must fail
                   fast with a message that names the server). Runs on CI (no domain needed).
      -Tag Lab     LIVE modification of throwaway templates in AD. Skipped unless -RunLab is
                   passed; needs the RSAT ActiveDirectory module and a writable DC. Surgical by
                   construction: TWO bare pKICertificateTemplate objects are created, named
                   PESTER-<hex>-VAL (fully populated) and PESTER-<hex>-BARE (no displayName, no
                   pKIExpirationPeriod, no pKIOverlapPeriod, no minor revision). Every wildcard
                   the tests use is scoped under the run-unique prefix. Each fixture proves its DN
                   absent before creation and records the created object's objectGUID; teardown
                   removes by that GUID only (with a prefix-scoped, structurally guarded safety
                   net). Pre-existing templates are never touched. Two cases inject failures: one
                   spawns powershell.exe (and pwsh.exe when on PATH) with -NonInteractive so that
                   ShouldProcess throws, and one places a temporary Deny WriteProperty ACE for
                   the current identity on the VAL object (removed in a finally).

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
    # Engines for the Lab non-interactive-host case: Windows PowerShell always, pwsh when on PATH.
    $script:LabEngines = @(@{ Exe = 'powershell.exe' })
    if (Get-Command pwsh.exe -ErrorAction SilentlyContinue) { $script:LabEngines += @{ Exe = 'pwsh.exe' } }
}

Describe 'Set-ADCSTemplateValidity' {

    BeforeAll {
        $script:Val = $ScriptPath
        $script:Val | Should -Exist

        # --- AST-extract the pure helpers so the Unit tier exercises the REAL code ------------
        # (The script's begin block connects to AD, so it cannot be dot-sourced wholesale;
        # extracting the function bodies runs them with no side effects.)
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:Val, [ref]$null, [ref]$null)
        foreach ($name in 'ConvertTo-PKIPeriodDays', 'ConvertTo-PKIPeriodBytes', 'ConvertFrom-PKIPeriodBytes', 'ConvertFrom-PKIPeriodBytesToDays', 'ConvertTo-LdapFilterValue') {
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

        It 'ConvertFrom-PKIPeriodBytesToDays returns exact days for comparisons (6 weeks = 42, 1 year = 365, 12 hours = 0.5) and $null for bad input' {
            ConvertFrom-PKIPeriodBytesToDays -Bytes (ConvertTo-PKIPeriodBytes -Period 6 -PeriodUnit Weeks)  | Should -Be 42
            ConvertFrom-PKIPeriodBytesToDays -Bytes (ConvertTo-PKIPeriodBytes -Period 1 -PeriodUnit Years)  | Should -Be 365
            ConvertFrom-PKIPeriodBytesToDays -Bytes (ConvertTo-PKIPeriodBytes -Period 12 -PeriodUnit Hours) | Should -Be 0.5
            ConvertFrom-PKIPeriodBytesToDays -Bytes $null | Should -BeNullOrEmpty
            ConvertFrom-PKIPeriodBytesToDays -Bytes ([byte[]](1, 2, 3)) | Should -BeNullOrEmpty
            # the case the script now refuses: a stock 6-week overlap under a 30-day validity
            (ConvertFrom-PKIPeriodBytesToDays -Bytes (ConvertTo-PKIPeriodBytes -Period 6 -PeriodUnit Weeks)) -ge (ConvertTo-PKIPeriodDays -Period 30 -PeriodUnit Days) | Should -BeTrue
        }

        It 'LDAP escaper protects metacharacters but PRESERVES the * wildcard' {
            ConvertTo-LdapFilterValue -Value 'a(b)c' | Should -BeExactly 'a\28b\29c'
            ConvertTo-LdapFilterValue -Value 'a\b'   | Should -BeExactly 'a\5cb'
            ConvertTo-LdapFilterValue -Value 'Web*'  | Should -BeExactly 'Web*'
            # ? is NOT an LDAP wildcard (RFC 4515 has none for a single character): it passes
            # through unescaped and the server matches it literally - so the help must not
            # advertise it as one.
            ConvertTo-LdapFilterValue -Value 'U?er'  | Should -BeExactly 'U?er'
            (Get-Help $script:Val -Parameter TemplateName | Out-String) | Should -Not -Match '\* and \?'
        }

        It 'LDAP escaper turns a NUL character into \00 (RFC 4515)' {
            ConvertTo-LdapFilterValue -Value "a`0b" | Should -BeExactly 'a\00b'
            ConvertTo-LdapFilterValue -Value "`0"   | Should -BeExactly '\00'
        }
    }

    Context 'Static: parse and help' -Tag 'Static' {

        It 'parses without errors' {
            $errs = $null
            $null = [System.Management.Automation.Language.Parser]::ParseFile($script:Val, [ref]$null, [ref]$errs)
            $errs | Should -BeNullOrEmpty
        }

        It 'is a FLAT script with no begin/process/end blocks (the 5.1 -File + non-console-stdin silent no-op fix must not regress)' {
            # Under Windows PowerShell 5.1, powershell.exe -File with a non-console stdin (a scheduler,
            # CI, WinRM/psexec, or the `< NUL` idiom) never runs a process{} block - the script would
            # search nothing, print nothing and exit 0. The script takes no pipeline input, so it must
            # run as a flat body; a reintroduced process{} would silently break unattended runs.
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:Val, [ref]$null, [ref]$null)
            $ast.BeginBlock   | Should -BeNullOrEmpty -Because 'a begin{} implies a process{} that 5.1 -File skips on non-console stdin'
            $ast.ProcessBlock | Should -BeNullOrEmpty -Because 'the process{} block silently no-ops under 5.1 -File with a redirected/EOF stdin'
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

        It 'refuses an unreachable -Server fast, with an error that names the server (v1.0.6 fix)' {
            # Before 1.0.6 the RootDSE bind of a dead host did not throw; the empty
            # configurationNamingContext produced a base DN ending in "CN=Services," and the error
            # blamed the templates container. The refusal must now name the server, and it must
            # not wait for an LDAP connect timeout (measured well under 1 s on both engines).
            $sw = [Diagnostics.Stopwatch]::StartNew()
            { & $script:Val -TemplateName x -ValidityPeriod 1 -ValidityPeriodUnit Years -Server nonexistent.invalid -WhatIf } |
                Should -Throw -ExpectedMessage "*RootDSE on the server 'nonexistent.invalid' returned no configurationNamingContext*"
            $sw.Stop()
            $sw.Elapsed.TotalSeconds | Should -BeLessThan 15 -Because 'the refusal happens right after the RootDSE read, not after an LDAP timeout'
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
            $script:Created = New-Object System.Collections.Generic.List[hashtable]   # @{ Dn; Guid } per object this run PROVED it created

            Import-Module ActiveDirectory -ErrorAction Stop
            $script:AP = @{}
            if ($LabServer) { $script:AP.Server = $LabServer }
            $cfgNc = (Get-ADRootDSE @script:AP).configurationNamingContext
            $script:TplBase = "CN=Certificate Templates,CN=Public Key Services,CN=Services,$cfgNc"

            # Fixture discipline: prove the intended DN is ABSENT (throw and register nothing if it
            # is not - that object is somebody else's), then create with -PassThru -ErrorAction Stop
            # and record the objectGUID of the object New-ADObject itself returned. The GUID never
            # comes from a later lookup by DN: under the default Continue preference a concurrent
            # name collision makes New-ADObject emit a NON-terminating error, and a DN lookup would
            # then register the OTHER object's GUID for teardown to delete. The absence check is
            # terminating for the same reason: a failed read must never pass as "absent".
            # Teardown removes by the recorded GUID only.
            # Breadcrumb BEFORE creation: teardown never runs on a hard process kill, and later
            # runs' prefix-scoped sweeps can never match this run's prefix - so the exact DN must
            # survive in the console/CI log for manual, exact-name removal after an abort.
            function script:New-LabTemplate([string]$Name, [hashtable]$Attributes) {
                $dn = "CN=$Name,$script:TplBase"
                $existing = @(Get-ADObject @script:AP -SearchBase $script:TplBase -SearchScope OneLevel -LDAPFilter "(cn=$Name)" -ErrorAction Stop)
                if ($existing.Count -gt 0) {
                    throw "Lab fixture refused: an object already exists at $dn. It is not this run's; nothing was registered and it will not be touched."
                }
                Write-Host "Lab object about to be created: $dn (remove manually by this exact DN if the run is killed before teardown)"
                $obj = New-ADObject @script:AP -Name $Name -Type pKICertificateTemplate -Path $script:TplBase -OtherAttributes $Attributes -PassThru -ErrorAction Stop
                if (-not $obj -or -not $obj.ObjectGUID) {
                    throw "Lab fixture refused: New-ADObject returned no object for $dn, so this run cannot prove it created one. Nothing was registered; review the DN manually."
                }
                $script:Created.Add(@{ Dn = $dn; Guid = [guid]$obj.ObjectGUID })
                $dn
            }

            # ONE bare throwaway template: 2-year validity, 6-week overlap, revision 0.
            $script:TplName = "$script:Prefix-VAL"
            $script:TplDn   = script:New-LabTemplate -Name $script:TplName -Attributes @{
                displayName                     = "$script:TplName (throwaway)"
                revision                        = 100
                'msPKI-Template-Minor-Revision' = 0
                pKIExpirationPeriod             = ([byte[]](ConvertTo-PKIPeriodBytes -Period 2 -PeriodUnit Years))
                pKIOverlapPeriod                = ([byte[]](ConvertTo-PKIPeriodBytes -Period 6 -PeriodUnit Weeks))
            }

            function script:Get-LabTpl {
                Get-ADObject @script:AP -Identity $script:TplDn -Properties pKIExpirationPeriod, pKIOverlapPeriod, 'msPKI-Template-Minor-Revision'
            }

            # The DC's DNS host name, so that one case passes -Server explicitly and exercises the
            # script's LDAP://<server>/... read AND write paths even in a serverless Lab run.
            $script:LabDc = if ($LabServer) { $LabServer } else { [string](Get-ADRootDSE @script:AP).dnsHostName }
            $script:LabDc | Should -Not -BeNullOrEmpty -Because 'RootDSE must expose dnsHostName'

            function script:Get-LabAdsiPath([string]$Dn) {
                if ($LabServer) { "LDAP://$LabServer/$Dn" } else { "LDAP://$Dn" }
            }

            # A SECOND throwaway with the optional attributes ABSENT: no displayName, no
            # pKIExpirationPeriod, no pKIOverlapPeriod, no msPKI-Template-Minor-Revision. It drives
            # the script's absent-attribute else branches (DisplayName falls back to cn, periods
            # read as N/A, revision starts at 0, no retained-overlap refusal).
            $script:BareName = "$script:Prefix-BARE"
            $script:BareDn   = script:New-LabTemplate -Name $script:BareName -Attributes @{ revision = 100 }
        }

        AfterAll {
            # Surgical: remove ONLY the objects this run proved it created, by objectGUID, and only
            # while the object at the tracked DN still carries that GUID. Report failures
            # truthfully - a swallowed error must never masquerade as a successful removal.
            foreach ($entry in $script:Created) {
                try {
                    $cur = Get-ADObject @script:AP -Identity $entry.Dn
                    if ([guid]$cur.ObjectGUID -ne $entry.Guid) {
                        Write-Warning "Teardown left $($entry.Dn) alone: its objectGUID $($cur.ObjectGUID) is not the one this run created ($($entry.Guid))."
                        continue
                    }
                    Remove-ADObject @script:AP -Identity $entry.Guid -Confirm:$false
                }
                catch { Write-Warning "Teardown could NOT remove tracked object $($entry.Dn): $_" }
            }
            # Safety net, scoped to THIS run's fresh-GUID prefix: REPORT ONLY. A prefix is not proof
            # that this run created an object, so a leftover is named for manual review, never
            # deleted here. STRUCTURAL GUARD: only with a fully-formed prefix - an unset one would
            # widen the LDAP filter. Never widen this.
            if ($script:Prefix -match '^PESTER-[0-9a-f]{8}$' -and $script:TplBase) {
                $tracked = @($script:Created | ForEach-Object { $_.Dn })
                foreach ($t in @(Get-ADObject @script:AP -SearchBase $script:TplBase -LDAPFilter "(cn=$script:Prefix-*)" -ErrorAction SilentlyContinue)) {
                    $why = if ($tracked -contains $t.DistinguishedName) { 'a tracked object the GUID check above left alone' } else { 'an object this run never registered' }
                    Write-Warning "AfterAll safety-net found a leftover object with this run's prefix ($why): $($t.DistinguishedName) (objectGUID $($t.ObjectGUID)). Review and remove it manually by this exact DN."
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

        It 'refuses to shorten validity below the RETAINED overlap: Error row, template unchanged, non-zero exit (v1.0.2 fix, previously untested)' {
            # Prior test left this template at validity 100 Days, overlap 4 Weeks (28 days). A new
            # validity of 20 Days with NO -OverlapPeriod would leave the 28-day overlap longer than the
            # validity - an invalid pair - so the run must report an Error, change nothing, and exit 1.
            $before = script:Get-LabTpl
            $p = @{ TemplateName = $script:TplName; ValidityPeriod = 20; ValidityPeriodUnit = 'Days'; Confirm = $false }
            if ($LabServer) { $p.Server = $LabServer }
            $global:LASTEXITCODE = 0
            $out = @(& $script:Val @p 3>&1 6>$null 2>$null)
            $rows = @($out | Where-Object { $_.PSObject.Properties['Status'] })
            $rows.Count | Should -Be 1
            $rows[0].Status | Should -BeLike 'Error*' -Because 'the retained overlap is not shorter than the new validity'
            $after = script:Get-LabTpl
            script:ToHex ([byte[]]$after.pKIExpirationPeriod) | Should -BeExactly (script:ToHex ([byte[]]$before.pKIExpirationPeriod)) -Because 'the template must be left untouched'
            [int]$after.'msPKI-Template-Minor-Revision' | Should -Be ([int]$before.'msPKI-Template-Minor-Revision')
            $LASTEXITCODE | Should -Be 1 -Because 'a run with an error exits non-zero (Write-Error + exit 1) while the structured report survives'
        }

        It 'rerunning the same validity WITHOUT -OverlapPeriod reports Already set and does NOT bump the revision (the common operator path)' {
            # State: validity 100 Days, overlap 4 Weeks, revision 2. With no -OverlapPeriod the
            # overlap comparison short-circuits to "equal", so an equal validity is a no-op.
            $p = @{ TemplateName = $script:TplName; ValidityPeriod = 100; ValidityPeriodUnit = 'Days'; Confirm = $false }
            if ($LabServer) { $p.Server = $LabServer }
            $out = @(& $script:Val @p 6>$null)
            $out.Count | Should -Be 1
            $out[0].Status     | Should -BeExactly 'Already set'
            $out[0].NewOverlap | Should -BeExactly '(unchanged)'
            [int](script:Get-LabTpl).'msPKI-Template-Minor-Revision' | Should -Be 2
        }

        It 'the same validity with a DIFFERENT overlap reports Modified and bumps the revision, through an explicit -Server' {
            # Validity stays 100 Days (equal), overlap 4 -> 2 Weeks (differs): the run must write.
            # -Server is passed explicitly here so the LDAP://<server>/<dn> write path is exercised.
            $p = @{ TemplateName = $script:TplName; ValidityPeriod = 100; ValidityPeriodUnit = 'Days'
                    OverlapPeriod = 2; OverlapPeriodUnit = 'Weeks'; Confirm = $false; Server = $script:LabDc }
            $out = @(& $script:Val @p 6>$null)
            $out.Count | Should -Be 1
            $out[0].Status          | Should -BeExactly 'Modified'
            $out[0].PreviousOverlap | Should -BeExactly '4 week(s)'
            $out[0].NewOverlap      | Should -BeExactly '2 Weeks'
            $tpl = script:Get-LabTpl
            script:ToHex ([byte[]]$tpl.pKIExpirationPeriod) | Should -BeExactly (script:ToHex (ConvertTo-PKIPeriodBytes -Period 100 -PeriodUnit Days))
            script:ToHex ([byte[]]$tpl.pKIOverlapPeriod)    | Should -BeExactly (script:ToHex (ConvertTo-PKIPeriodBytes -Period 2 -PeriodUnit Weeks))
            [int]$tpl.'msPKI-Template-Minor-Revision' | Should -Be 3
        }

        It 'a non-interactive host (<Exe> -NonInteractive -File, no -Confirm:$false) makes ShouldProcess throw: exit 1, one Error row, template unchanged (v1.0.4 fix)' -ForEach $script:LabEngines {
            # ConfirmImpact High with no -Confirm:$false asks for confirmation; under -NonInteractive
            # the prompt is impossible and ShouldProcess THROWS. The script's try wraps that call so
            # the failure is counted and reported as an Error row instead of escaping uncounted.
            $before = script:Get-LabTpl
            $stdout = Join-Path $TestDrive "ni-$Exe.out.txt"
            $stderr = Join-Path $TestDrive "ni-$Exe.err.txt"
            $argList = @('-NonInteractive', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$script:Val`"",
                         '-TemplateName', $script:TplName, '-ValidityPeriod', '150', '-ValidityPeriodUnit', 'Days')
            if ($LabServer) { $argList += @('-Server', $LabServer) }
            $proc = Start-Process -FilePath $Exe -ArgumentList $argList -Wait -PassThru -NoNewWindow -RedirectStandardOutput $stdout -RedirectStandardError $stderr
            $proc.ExitCode | Should -Be 1 -Because 'the counted confirmation failure must end the run with exit 1'

            # The 7-property rows print as Format-List. pwsh wraps property names in ANSI colour
            # codes even with a redirected stdout (strip them), and both engines wrap a long value
            # at the console width (compare the exception text with whitespace removed).
            $text = (Get-Content -Path $stdout -Raw) -replace '\x1b\[[0-9;]*[A-Za-z]', ''
            $statusLines = @([regex]::Matches($text, '(?m)^Status\s*:\s*(.*)$') | ForEach-Object { $_.Groups[1].Value.Trim() })
            $statusLines.Count | Should -Be 1 -Because 'exactly one row (the throwaway) is emitted'
            $statusLines[0] | Should -BeLike 'Error:*'
            ($text -replace '\s+', '') | Should -Match 'NonInteractive' -Because 'the Error row carries the ShouldProcess exception text'
            $text | Should -Match ('(?m)^TemplateName\s*:\s*' + [regex]::Escape($script:TplName))

            $after = script:Get-LabTpl
            script:ToHex ([byte[]]$after.pKIExpirationPeriod) | Should -BeExactly (script:ToHex ([byte[]]$before.pKIExpirationPeriod))
            [int]$after.'msPKI-Template-Minor-Revision' | Should -Be ([int]$before.'msPKI-Template-Minor-Revision')
        }

        It 'a failed SetInfo (Deny WriteProperty ACE for the current identity) yields an Error row, exit 1, and no change' {
            $before = script:Get-LabTpl
            $sid  = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
            $rule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule(
                $sid, [System.DirectoryServices.ActiveDirectoryRights]::WriteProperty, [System.Security.AccessControl.AccessControlType]::Deny)
            $aclPath = script:Get-LabAdsiPath $script:TplDn
            try {
                $entry = [ADSI]$aclPath
                $sec = $entry.ObjectSecurity
                $sec.AddAccessRule($rule)
                $entry.ObjectSecurity = $sec
                $entry.CommitChanges()

                $p = @{ TemplateName = $script:TplName; ValidityPeriod = 150; ValidityPeriodUnit = 'Days'; Confirm = $false }
                if ($LabServer) { $p.Server = $LabServer }
                $global:LASTEXITCODE = 0
                $out = @(& $script:Val @p 3>&1 6>$null 2>$null)
                $rows = @($out | Where-Object { $_.PSObject.Properties['Status'] })
                $rows.Count | Should -Be 1
                $rows[0].Status | Should -BeLike 'Error:*' -Because 'SetInfo is refused by the Deny ACE'
                $LASTEXITCODE | Should -Be 1
            }
            finally {
                # Remove the Deny ACE so later cases (and teardown) keep their rights. WriteDacl is
                # not denied, so this always succeeds while the object exists.
                $entry = [ADSI]$aclPath
                $sec = $entry.ObjectSecurity
                $null = $sec.RemoveAccessRule($rule)
                $entry.ObjectSecurity = $sec
                $entry.CommitChanges()
            }
            $after = script:Get-LabTpl
            script:ToHex ([byte[]]$after.pKIExpirationPeriod) | Should -BeExactly (script:ToHex ([byte[]]$before.pKIExpirationPeriod)) -Because 'the refused write must change nothing'
            [int]$after.'msPKI-Template-Minor-Revision' | Should -Be ([int]$before.'msPKI-Template-Minor-Revision')
        }

        It 'a template with NO displayName, pKIExpirationPeriod, pKIOverlapPeriod or minor revision: DisplayName = cn, periods N/A, no refusal, revision 1' {
            $p = @{ TemplateName = $script:BareName; ValidityPeriod = 1; ValidityPeriodUnit = 'Years'; Confirm = $false }
            if ($LabServer) { $p.Server = $LabServer }
            $mixed = @(& $script:Val @p 2>&1 3>&1 6>$null)
            @($mixed | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] })   | Should -BeNullOrEmpty -Because 'an absent overlap cannot trigger the retained-overlap refusal'
            @($mixed | Where-Object { $_ -is [System.Management.Automation.WarningRecord] }) | Should -BeNullOrEmpty
            $out = @($mixed | Where-Object { $_.PSObject.Properties['Status'] })
            $out.Count | Should -Be 1
            $out[0].DisplayName      | Should -BeExactly $script:BareName -Because 'with no displayName the script falls back to cn'
            $out[0].PreviousValidity | Should -BeExactly 'N/A'
            $out[0].PreviousOverlap  | Should -BeExactly 'N/A'
            $out[0].NewOverlap       | Should -BeExactly '(unchanged)'
            $out[0].Status           | Should -BeExactly 'Modified'
            $tpl = Get-ADObject @script:AP -Identity $script:BareDn -Properties displayName, pKIExpirationPeriod, pKIOverlapPeriod, 'msPKI-Template-Minor-Revision'
            script:ToHex ([byte[]]$tpl.pKIExpirationPeriod) | Should -BeExactly (script:ToHex (ConvertTo-PKIPeriodBytes -Period 1 -PeriodUnit Years))
            $tpl.pKIOverlapPeriod | Should -BeNullOrEmpty -Because 'the script does not invent an overlap'
            $tpl.displayName      | Should -BeNullOrEmpty -Because 'the script does not invent a displayName'
            [int]$tpl.'msPKI-Template-Minor-Revision' | Should -Be 1 -Because 'an absent revision counts as 0 and is bumped to 1'
        }
    }
}
