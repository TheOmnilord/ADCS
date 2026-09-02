<#
.SYNOPSIS
    Pester suite for Sync-ADCSTemplate.ps1. Requires Pester 5+ (tested on 6.x).

.DESCRIPTION
    Three always-on tiers plus one opt-in tier:

      -Tag Unit    Pure helper functions, extracted from the script by AST (so the REAL code
                   is exercised, never a copy) and run in-process. No AD, no DC.
      -Tag Static  The script parses without errors and its comment-based help binds.
                   No AD, no DC.
      -Tag Guard   Parameter/mode validation. These invocations throw BEFORE any DC is
                   contacted, so they need the ActiveDirectory module present but no reachable
                   DC and make no changes.
      -Tag Lab     Live end-to-end tests that CREATE and REMOVE AD objects. Skipped unless
                   -RunLab is passed. Cleanup is surgical (exact DN of each object this suite
                   created, tracked per run under a unique PESTER-<hex> prefix); pre-existing
                   objects are never touched.

.EXAMPLE
    # Safe tiers only (Unit + Static + Guard) - no changes, needs the AD module:
    Invoke-Pester -Path .\Tests\Sync-ADCSTemplate.Tests.ps1 -ExcludeTag Lab

.EXAMPLE
    # Static/Unit only - runs even without the AD module:
    Invoke-Pester -Path .\Tests\Sync-ADCSTemplate.Tests.ps1 -Tag Unit,Static

.EXAMPLE
    # Full run against a lab. Configure servers via -Data on a container:
    $cfg = New-PesterContainer -Path .\Tests\Sync-ADCSTemplate.Tests.ps1 -Data @{
        RunLab          = $true
        AronsServer     = 'ARONS-DC1.arons.local'
        ChildServer     = 'WIN-1UP9S490HDR.child.arons.local'   # optional: child-domain root-SID path
        NorefjellServer = '192.168.1.101'                        # optional: cross-forest, no-AD CS target
    }
    Invoke-Pester -Container $cfg
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingConvertToSecureStringWithPlainText', '',
    Justification = 'a dummy credential built solely to trigger the -Credential-requires-Server guard; never used to authenticate')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingEmptyCatchBlock', '',
    Justification = 'best-effort polling and cleanup paths (retry loop, AfterEach/AfterAll teardown) deliberately swallow per-attempt errors')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '',
    Justification = 'container parameters are consumed inside Pester Describe/BeforeDiscovery scriptblocks, which the analyzer cannot see through')]
param(
    [bool]   $RunLab          = $false,
    [string] $ScriptPath      = (Join-Path (Split-Path $PSScriptRoot -Parent) 'Sync-ADCSTemplate.ps1'),
    [string] $SourceTemplate  = 'KerberosAuthentication',   # a template present in every forest with the PK Services structure
    [string] $AronsServer     = '',                         # target DC for the Lab tier (a single DC). Required for -RunLab.
    [string] $ChildServer     = '',                         # optional: a child-domain DC (child-domain root-SID path)
    [string] $NorefjellServer = ''                          # optional: a DC/IP in a SEPARATE forest without AD CS (cross-forest)
)

BeforeDiscovery {
    # -Skip conditions are evaluated during discovery, so anything they reference must be set here.
    $script:LabReady     = $RunLab -and $AronsServer
    $script:ChildReady   = $RunLab -and $ChildServer
    $script:XForestReady = $RunLab -and $AronsServer -and $NorefjellServer
    $script:HasAD        = [bool](Get-Module -ListAvailable ActiveDirectory)
}

Describe 'Sync-ADCSTemplate' {

    BeforeAll {
        $script:Sync = $ScriptPath
        $script:Sync | Should -Exist

        # --- AST-extract the pure helper functions so the Unit tier exercises the real code ---
        # (The script has a mandatory -Mode and runs main logic on load, so it cannot be dot-sourced
        # wholesale; extracting the function definitions gives their real bodies with no side effects.)
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:Sync, [ref]$null, [ref]$null)
        foreach ($name in 'Get-RandomHex', 'ConvertTo-LdapFilterValue', 'New-SyntheticOidBase', 'Get-AttrCanonical', 'Compare-TemplateAttributes', 'Convert-ToLatestCompatibility') {
            $def = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
            if ($def) { . ([scriptblock]::Create($def[0].Extent.Text)) }
        }
        # Get-AttrCanonical references $script:ByteAttributes - mirror the script's definition.
        $script:ByteAttributes = @('pKIExpirationPeriod', 'pKIKeyUsage', 'pKIOverlapPeriod')

    }

    Context 'Unit: pure helpers' -Tag 'Unit' {

        It 'ConvertTo-LdapFilterValue escapes RFC 4515 metacharacters' {
            ConvertTo-LdapFilterValue 'a*b'   | Should -Be 'a\2ab'
            ConvertTo-LdapFilterValue 'a(b)c' | Should -Be 'a\28b\29c'
            ConvertTo-LdapFilterValue 'a\b'   | Should -Be 'a\5cb'
            ConvertTo-LdapFilterValue 'plain' | Should -Be 'plain'
        }

        It 'Get-RandomHex returns N uppercase hex characters' {
            $h = Get-RandomHex -Length 32
            $h | Should -Match '^[0-9A-F]{32}$'
        }

        It 'New-SyntheticOidBase produces a well-formed base under the MS template arc' {
            $oid = New-SyntheticOidBase
            $oid | Should -Match '^1\.3\.6\.1\.4\.1\.311\.21\.8(\.[0-9]+){5}$'
        }

        It 'Get-AttrCanonical: a literal pipe cannot collide two different multi-values' {
            # The '|' join must escape embedded pipes, else @("a|b") and @("a","b") look identical.
            (Get-AttrCanonical -Name 'x' -Value @('a|b')) |
                Should -Not -Be (Get-AttrCanonical -Name 'x' -Value @('a', 'b'))
        }

        It 'Compare-TemplateAttributes: a case-only difference is a mismatch (case-sensitive -ceq)' {
            # Get-AttrCanonical preserves case; the case-sensitive guarantee is the -ceq in
            # Compare-TemplateAttributes, so assert through it (Pester's -Be is case-insensitive).
            $src = [pscustomobject]@{ 'msPKI-Foo' = 'Value' }
            $tgt = [pscustomobject]@{ 'msPKI-Foo' = 'VALUE' }
            $diff = Compare-TemplateAttributes -Source $src -Target $tgt
            ($diff | Where-Object Attribute -eq 'msPKI-Foo').Match | Should -BeFalse
        }

        It 'Get-AttrCanonical: $null and an empty collection canonicalize identically' {
            (Get-AttrCanonical -Name 'x' -Value $null) |
                Should -Be (Get-AttrCanonical -Name 'x' -Value @())
        }

        It 'Get-AttrCanonical: byte attributes compare in order (byte order is significant)' {
            (Get-AttrCanonical -Name 'pKIKeyUsage' -Value ([byte[]](1, 16))) |
                Should -Not -Be (Get-AttrCanonical -Name 'pKIKeyUsage' -Value ([byte[]](16, 1)))
        }

        It 'Get-AttrCanonical: multi-value set comparison is order-insensitive' {
            (Get-AttrCanonical -Name 'x' -Value @('a', 'b')) |
                Should -Be (Get-AttrCanonical -Name 'x' -Value @('b', 'a'))
        }

        It 'Compare-TemplateAttributes: flags a value that differs' {
            $src = [pscustomobject]@{ 'pKIKeyUsage' = [byte[]](160, 0) }
            $tgt = [pscustomobject]@{ 'pKIKeyUsage' = [byte[]](0, 0) }
            $diff = Compare-TemplateAttributes -Source $src -Target $tgt
            ($diff | Where-Object Attribute -eq 'pKIKeyUsage').Match | Should -BeFalse
        }

        It 'Compare-TemplateAttributes: a TARGET-only PKI attribute appears in the diff' {
            $src = [pscustomobject]@{ 'flags' = 1 }
            $tgt = [pscustomobject]@{ 'flags' = 1; 'msPKI-Extra' = 'surprise' }
            $diff = Compare-TemplateAttributes -Source $src -Target $tgt
            ($diff | Where-Object Attribute -eq 'msPKI-Extra') | Should -Not -BeNullOrEmpty
        }

        It 'Convert-ToLatestCompatibility: CSP-based v2 -> v4 with the exact stock v4 bytes (0x06060100)' {
            # Oracle: real MMC-made v4 Kerberos Authentication templates carry 0x06060100.
            $a = @{ 'msPKI-Template-Schema-Version' = [int]2; 'msPKI-Private-Key-Flag' = [int]0
                    'flags' = [BitConverter]::ToInt32([BitConverter]::GetBytes([uint32]0x10060), 0)
                    'msPKI-Template-Minor-Revision' = [int]0
                    'pKIDefaultCSPs' = @('1,Microsoft RSA SChannel Cryptographic Provider') }
            $r = Convert-ToLatestCompatibility -Attributes $a
            $r.Upgraded | Should -BeTrue
            $a['msPKI-Template-Schema-Version'] | Should -Be 4
            [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$a['msPKI-Private-Key-Flag']), 0) | Should -Be ([uint32]0x06060100)
            [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$a['flags']), 0) | Should -Be ([uint32]0x20060)   # IS_DEFAULT -> IS_MODIFIED
            $a['msPKI-Template-Minor-Revision'] | Should -Be 1
        }

        It 'Convert-ToLatestCompatibility: CNG/KSP template (no CSP list) gets 0x06060000, NOT the legacy-provider bit' {
            $a = @{ 'msPKI-Template-Schema-Version' = [int]3; 'msPKI-Private-Key-Flag' = [int]0 }
            $r = Convert-ToLatestCompatibility -Attributes $a
            $r.Upgraded | Should -BeTrue
            $r.LegacyProvider | Should -BeFalse
            [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$a['msPKI-Private-Key-Flag']), 0) | Should -Be ([uint32]0x06060000)
        }

        It 'Convert-ToLatestCompatibility: REPLACES an existing compatibility level, never OR-accumulates the version nibbles' {
            # A v3 source already at "Windows Server 2008 R2 / Windows 7" (both nibbles = 3, 0x03030000).
            # A plain -bor would give 3|6 = 7 (an invalid level); the transform must land on exactly 6/6.
            $cng = @{ 'msPKI-Template-Schema-Version' = [int]3
                      'msPKI-Private-Key-Flag' = [BitConverter]::ToInt32([BitConverter]::GetBytes([uint32]0x03030000), 0) }
            $null = Convert-ToLatestCompatibility -Attributes $cng
            [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$cng['msPKI-Private-Key-Flag']), 0) | Should -Be ([uint32]0x06060000)

            # Same, CSP-based and with an unrelated high-nibble flag (0x00200000) that must survive.
            $csp = @{ 'msPKI-Template-Schema-Version' = [int]2
                      'msPKI-Private-Key-Flag' = [BitConverter]::ToInt32([BitConverter]::GetBytes([uint32]0x05250010), 0)
                      'pKIDefaultCSPs' = @('1,Microsoft RSA SChannel Cryptographic Provider') }
            $null = Convert-ToLatestCompatibility -Attributes $csp
            # nibbles 5/5 -> 6/6; 0x00200000 and 0x10 preserved; CSP adds 0x100.
            [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$csp['msPKI-Private-Key-Flag']), 0) | Should -Be ([uint32]0x06260110)
        }

        It 'Convert-ToLatestCompatibility: an empty/absent CSP list does not fool the @($null).Count trap' {
            foreach ($csp in @($null, @())) {
                $a = @{ 'msPKI-Template-Schema-Version' = [int]2; 'msPKI-Private-Key-Flag' = [int]0; 'pKIDefaultCSPs' = $csp }
                $r = Convert-ToLatestCompatibility -Attributes $a
                $r.LegacyProvider | Should -BeFalse -Because 'no CSP entries means CNG/KSP - no 0x100'
                [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$a['msPKI-Private-Key-Flag']), 0) | Should -Be ([uint32]0x06060000)
            }
        }

        It 'Convert-ToLatestCompatibility: v1 is left untouched (not upgradable in place)' {
            $a = @{ 'msPKI-Template-Schema-Version' = [int]1; 'msPKI-Private-Key-Flag' = [int]0 }
            $r = Convert-ToLatestCompatibility -Attributes $a
            $r.Upgraded | Should -BeFalse
            $a['msPKI-Template-Schema-Version'] | Should -Be 1
            $a['msPKI-Private-Key-Flag'] | Should -Be 0
        }

        It 'Convert-ToLatestCompatibility: a template already at v4 is a no-op' {
            $a = @{ 'msPKI-Template-Schema-Version' = [int]4; 'msPKI-Private-Key-Flag' = [int]0x06060100 }
            $r = Convert-ToLatestCompatibility -Attributes $a
            $r.Upgraded | Should -BeFalse
            $a['msPKI-Private-Key-Flag'] | Should -Be 0x06060100
        }
    }

    Context 'Static: parse and help' -Tag 'Static' {

        It 'parses without errors' {
            $errs = $null
            [System.Management.Automation.Language.Parser]::ParseFile($script:Sync, [ref]$null, [ref]([ref]$errs).Value) | Out-Null
            $errs2 = $null
            $null = [System.Management.Automation.Language.Parser]::ParseFile($script:Sync, [ref]$null, [ref]$errs2)
            $errs2 | Should -BeNullOrEmpty
        }

        It 'comment-based help binds (Synopsis is present)' {
            (Get-Help $script:Sync).Synopsis.Trim() | Should -Not -BeNullOrEmpty
        }

        It 'carries a PSScriptInfo header (Test-ScriptFileInfo parses it; Version is semver)' {
            $info = Test-ScriptFileInfo -Path $script:Sync -ErrorAction Stop
            $info.Version | Should -Match '^\d+\.\d+\.\d+$'
            $info.Guid    | Should -Not -BeNullOrEmpty
        }
        It 'documents every non-common parameter' {
            $cmd = Get-Command $script:Sync
            $common = [System.Management.Automation.PSCmdlet]::CommonParameters + [System.Management.Automation.PSCmdlet]::OptionalCommonParameters
            $help = Get-Help $script:Sync
            $documented = @($help.parameters.parameter.name)
            foreach ($p in $cmd.Parameters.Keys | Where-Object { $_ -notin $common }) {
                $documented | Should -Contain $p -Because "parameter -$p should have a .PARAMETER help entry"
            }
        }

        It 'provides runnable examples for every mode' {
            $ex = (Get-Help $script:Sync -Examples | Out-String)
            foreach ($mode in 'Export', 'Import', 'Sync', 'Validate') {
                $ex | Should -Match "-Mode $mode"
            }
        }
    }

    Context 'Guard: parameter and mode validation' -Tag 'Guard' -Skip:(-not $script:HasAD) {

        It 'rejects a parameter the mode does not consume (-Mode Import -StripOid)' {
            { & $script:Sync -Mode Import -Path x.json -StripOid } |
                Should -Throw -ExpectedMessage '*not applicable to -Mode Import*'
        }

        It 'requires -Path for -Mode Export' {
            { & $script:Sync -Mode Export } | Should -Throw -ExpectedMessage '*-Path is required*'
        }

        It 'requires -SourceServer for -Mode Sync' {
            { & $script:Sync -Mode Sync } | Should -Throw -ExpectedMessage '*-SourceServer is required*'
        }

        It 'requires -Server when -Credential is given' {
            $cred = [pscredential]::new('x\y', (ConvertTo-SecureString 'z' -AsPlainText -Force))
            { & $script:Sync -Mode Import -Path x.json -Credential $cred } |
                Should -Throw -ExpectedMessage '*-Credential requires -Server*'
        }

        It 'rejects -SkipAcl with -EnrollPrincipals' {
            { & $script:Sync -Mode Import -Path x.json -SkipAcl -EnrollPrincipals @{ a = 'Read' } } |
                Should -Throw -ExpectedMessage '*-SkipAcl and -EnrollPrincipals are mutually exclusive*'
        }

        It 'rejects -SkipAcl with an explicit -AclBase' {
            { & $script:Sync -Mode Import -Path x.json -SkipAcl -AclBase Schema } |
                Should -Throw -ExpectedMessage '*-SkipAcl and -AclBase are mutually exclusive*'
        }

        It 'rejects -OidRoot without -OidHandling GenerateFromRoot' {
            { & $script:Sync -Mode Import -Path x.json -OidRoot 1.2.3 } |
                Should -Throw -ExpectedMessage '*-OidRoot is only consumed by -OidHandling GenerateFromRoot*'
        }

        It 'rejects -SourceServer outside -Mode Sync' {
            { & $script:Sync -Mode Export -Path x.json -SourceServer dc1 } |
                Should -Throw -ExpectedMessage '*not applicable to -Mode Export*'
        }

        It 'rejects -UpgradeCompatibility for -Mode Export (Import/Sync only)' {
            { & $script:Sync -Mode Export -Path x.json -UpgradeCompatibility } |
                Should -Throw -ExpectedMessage '*not applicable to -Mode Export*'
        }

        It 'rejects -UpgradeCompatibility for -Mode Validate' {
            { & $script:Sync -Mode Validate -TemplateName K -UpgradeCompatibility } |
                Should -Throw -ExpectedMessage '*not applicable to -Mode Validate*'
        }
    }

    # -------------------------------------------------------------------------------------------
    # Lab tier: live create/verify/remove. Opt-in (-RunLab) and per-capability server config.
    # -------------------------------------------------------------------------------------------
    Context 'Lab: live operations' -Tag 'Lab' -Skip:(-not $script:LabReady) {

        BeforeAll {
            # Prefix FIRST - before anything that can throw. AfterAll runs even when BeforeAll
            # dies, and its safety-net sweep must never see an unset (= unscoped) prefix.
            $script:Prefix   = "PESTER-$([guid]::NewGuid().ToString('N').Substring(0,8))"
            Import-Module ActiveDirectory -ErrorAction Stop
            $script:TmpDir   = Join-Path $env:TEMP $script:Prefix
            New-Item -ItemType Directory -Force $script:TmpDir | Out-Null
            $script:AP       = @{ Server = $AronsServer }
            $script:ConfigNC = (Get-ADRootDSE @script:AP).configurationNamingContext
            $script:TplBase  = "CN=Certificate Templates,CN=Public Key Services,CN=Services,$script:ConfigNC"
            $script:OidBase  = "CN=OID,CN=Public Key Services,CN=Services,$script:ConfigNC"
            $script:ForestRoot = (Get-ADObject @script:AP -Identity $script:OidBase -Properties 'msPKI-Cert-Template-OID').'msPKI-Cert-Template-OID'
            $script:Created  = New-Object System.Collections.Generic.List[object]  # @{ Server; DN }

            # Export once - the input file every Import/Sync test reuses.
            $script:ExportFile = Join-Path $script:TmpDir 'src.json'
            & $script:Sync -Mode Export -Path $script:ExportFile -TemplateName $SourceTemplate -Server $AronsServer *> $null

            function script:New-LabName { "$script:Prefix-$([guid]::NewGuid().ToString('N').Substring(0,6))" }

            # Read an object with retry - a remote forest (cross-forest target) can lag briefly over
            # ADWS after a write, so poll until it appears (or the timeout elapses).
            function script:Get-ADObjectRetry {
                param([hashtable]$AdParams, [string]$Identity, [string]$SearchBase, [string]$Filter, [string[]]$Properties, [int]$TimeoutSec = 25)
                $q = @{} + $AdParams
                if ($Properties) { $q['Properties'] = $Properties }
                $deadline = (Get-Date).AddSeconds($TimeoutSec)
                do {
                    $o = $null
                    try {
                        if ($Identity) { $o = Get-ADObject @q -Identity $Identity -ErrorAction Stop }
                        else { $o = Get-ADObject @q -SearchBase $SearchBase -LDAPFilter $Filter -ErrorAction Stop }
                    }
                    catch { }
                    if ($o) { return $o }
                    Start-Sleep -Milliseconds 750
                } while ((Get-Date) -lt $deadline)
                return $null
            }

            # Run the script and TRACK the created object from its own "Created template: <DN>" output -
            # so cleanup never depends on a read-back that could lag. Returns @{ Output; DN; Oid }.
            function script:Invoke-LabCreate {
                param([hashtable]$SyncParams)
                # -Width 4096: stop Out-String wrapping the "Created template: <DN>" line, which would
                # truncate the DN mid-string. '.' never spans newlines, so (.+) grabs the whole DN line.
                $out = & $script:Sync @SyncParams *>&1 | Out-String -Width 4096
                $ap  = @{ Server = $SyncParams.Server }
                $dn  = if ($out -match 'Created template:\s*(.+)') { $Matches[1].Trim() } else { $null }
                $oid = if ($out -match 'Template OID:\s*([0-9.]+)') { $Matches[1] } else { $null }
                if ($dn) {
                    $script:Created.Add(@{ Server = $ap.Server; DN = $dn })
                    if ($oid) {
                        $oidC = "CN=OID,CN=Public Key Services,CN=Services,$((Get-ADRootDSE @ap).configurationNamingContext)"
                        $c = script:Get-ADObjectRetry -AdParams $ap -SearchBase $oidC -Filter "(msPKI-Cert-Template-OID=$oid)"
                        if ($c) { $script:Created.Add(@{ Server = $ap.Server; DN = $c.DistinguishedName }) }
                    }
                }
                [pscustomobject]@{ Output = $out; DN = $dn; Oid = $oid }
            }
        }

        AfterEach {
            # Surgical: remove ONLY the exact DNs this suite created, most-recent first.
            for ($i = $script:Created.Count - 1; $i -ge 0; $i--) {
                $o = $script:Created[$i]
                try { Remove-ADObject -Server $o.Server -Identity $o.DN -Confirm:$false -ErrorAction Stop } catch { }
            }
            $script:Created.Clear()
        }

        AfterAll {
            # Safety net: sweep every configured server for anything carrying THIS run's unique
            # prefix (a fresh GUID - it cannot match a pre-existing object), in case a test threw
            # before its object was tracked. Prefix-scoped, never a broad wildcard.
            # STRUCTURAL GUARD: the sweep runs only when the prefix has its full PESTER-<hex8>
            # shape - an unset/empty prefix would otherwise widen the LDAP filters to unscoped
            # patterns like (cn=-*) against the very container that holds real templates. AfterAll
            # runs even when BeforeAll throws, so never assume the prefix is set. Never widen this.
            if ($script:Prefix -match '^PESTER-[0-9a-f]{8}$') {
                foreach ($srv in @($AronsServer, $ChildServer, $NorefjellServer | Where-Object { $_ } | Select-Object -Unique)) {
                    try {
                        $cfg = (Get-ADRootDSE -Server $srv).configurationNamingContext
                        $tpls = "CN=Certificate Templates,CN=Public Key Services,CN=Services,$cfg"
                        $oidC = "CN=OID,CN=Public Key Services,CN=Services,$cfg"
                        foreach ($t in @(Get-ADObject -Server $srv -SearchBase $tpls -LDAPFilter "(cn=$script:Prefix-*)" -Properties 'msPKI-Cert-Template-OID' -ErrorAction SilentlyContinue)) {
                            $oid = $t.'msPKI-Cert-Template-OID'
                            Remove-ADObject -Server $srv -Identity $t.DistinguishedName -Confirm:$false -ErrorAction SilentlyContinue
                            if ($oid) {
                                $c = Get-ADObject -Server $srv -SearchBase $oidC -LDAPFilter "(msPKI-Cert-Template-OID=$oid)" -ErrorAction SilentlyContinue
                                if ($c) { Remove-ADObject -Server $srv -Identity $c.DistinguishedName -Confirm:$false -ErrorAction SilentlyContinue }
                            }
                            Write-Warning "AfterAll safety-net removed an untracked test object: $($t.DistinguishedName)"
                        }
                        # Companion OID objects whose display name carries the prefix but whose template was already removed.
                        foreach ($c in @(Get-ADObject -Server $srv -SearchBase $oidC -LDAPFilter "(DisplayName=$script:Prefix-*)" -ErrorAction SilentlyContinue)) {
                            Remove-ADObject -Server $srv -Identity $c.DistinguishedName -Confirm:$false -ErrorAction SilentlyContinue
                        }
                    }
                    catch { }
                }
            } else {
                Write-Warning "Safety-net sweep skipped: run prefix is unset or malformed ('$script:Prefix')."
            }
            if ($script:TmpDir) { Remove-Item -LiteralPath $script:TmpDir -Recurse -Force -ErrorAction SilentlyContinue }
        }

        It 'Export writes a BOM-marked JSON with identity, OID, and no ACL attribute' {
            $bytes = [System.IO.File]::ReadAllBytes($script:ExportFile)
            ($bytes[0..2] -join ',') | Should -Be '239,187,191'   # UTF-8 BOM
            $j = Get-Content -LiteralPath $script:ExportFile -Raw -Encoding UTF8 | ConvertFrom-Json
            $j.name | Should -Be $SourceTemplate
            $j.'msPKI-Cert-Template-OID' | Should -Not -BeNullOrEmpty
            $j.PSObject.Properties.Name | Should -Not -Contain 'pKIEnrollmentAccess'
        }

        It 'Export refuses a wildcard (metacharacter treated literally)' {
            { & $script:Sync -Mode Export -Path (Join-Path $script:TmpDir 'w.json') -TemplateName 'Kerb*' -Server $AronsServer } |
                Should -Throw -ExpectedMessage '*was not found*'
        }

        It 'Import -OidHandling <Mode> creates the template and a companion OID object' -ForEach @(
            @{ Mode = 'GenerateRandom'; Extra = @{} }
            @{ Mode = 'Generate'; Extra = @{} }
            @{ Mode = 'GenerateFromRoot'; Extra = @{ OidRoot = '1.3.6.1.4.1.311.21.8.90000001.90000002.90000003.90000004.90000005' } }
        ) {
            $cn = script:New-LabName
            $p = @{ Mode = 'Import'; Path = $script:ExportFile; Server = $AronsServer; NewTemplateName = $cn; NewDisplayName = $cn; OidHandling = $Mode; AclBase = 'Schema' }
            $Extra.GetEnumerator() | ForEach-Object { $p[$_.Key] = $_.Value }
            $r = script:Invoke-LabCreate -SyncParams $p
            $r.DN | Should -Not -BeNullOrEmpty -Because 'the script should report a created template'
            $comp = script:Get-ADObjectRetry -AdParams $script:AP -SearchBase $script:OidBase -Filter "(msPKI-Cert-Template-OID=$($r.Oid))"
            $comp | Should -Not -BeNullOrEmpty -Because "$Mode should register a companion OID display object"
            if ($Mode -eq 'Generate') {
                $r.Oid | Should -BeLike "$script:ForestRoot.*" -Because 'Generate mints under the real forest OID root'
            }
        }

        It '-UpgradeCompatibility raises the imported copy to v4 with the stock v4 private-key-flag' {
            # Fixture is the built-in Kerberos Authentication template (schema v2, CSP-based), so the
            # upgrade must land it at v4 / 0x06060100 - the exact value real MMC-made v4 copies carry.
            $cn = script:New-LabName
            $r = script:Invoke-LabCreate -SyncParams @{ Mode = 'Import'; Path = $script:ExportFile; Server = $AronsServer
                NewTemplateName = $cn; NewDisplayName = $cn; OidHandling = 'GenerateRandom'; SkipAcl = $true; UpgradeCompatibility = $true }
            $r.DN | Should -Not -BeNullOrEmpty
            $t = script:Get-ADObjectRetry -AdParams $script:AP -Identity $r.DN -Properties 'msPKI-Template-Schema-Version', 'msPKI-Private-Key-Flag'
            [int]$t.'msPKI-Template-Schema-Version' | Should -Be 4
            [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$t.'msPKI-Private-Key-Flag'), 0) | Should -Be ([uint32]0x06060100)
        }

        It 'a plain import (no -UpgradeCompatibility) preserves the source schema version' {
            $cn = script:New-LabName
            $r = script:Invoke-LabCreate -SyncParams @{ Mode = 'Import'; Path = $script:ExportFile; Server = $AronsServer
                NewTemplateName = $cn; NewDisplayName = $cn; OidHandling = 'GenerateRandom'; SkipAcl = $true }
            $t = script:Get-ADObjectRetry -AdParams $script:AP -Identity $r.DN -Properties 'msPKI-Template-Schema-Version'
            [int]$t.'msPKI-Template-Schema-Version' | Should -Be 2 -Because 'the Kerberos Authentication fixture is schema v2 and must be copied as-is without the switch'
        }

        It 'Import -AclBase Standard writes a protected DACL; SkipAcl leaves the schema default' {
            $cn1 = script:New-LabName
            $r1 = script:Invoke-LabCreate -SyncParams @{ Mode = 'Import'; Path = $script:ExportFile; Server = $AronsServer; NewTemplateName = $cn1; NewDisplayName = $cn1; OidHandling = 'GenerateRandom'; AclBase = 'Standard' }
            $t1 = script:Get-ADObjectRetry -AdParams $script:AP -Identity $r1.DN -Properties nTSecurityDescriptor
            $t1.nTSecurityDescriptor.AreAccessRulesProtected | Should -BeTrue

            $cn2 = script:New-LabName
            $r2 = script:Invoke-LabCreate -SyncParams @{ Mode = 'Import'; Path = $script:ExportFile; Server = $AronsServer; NewTemplateName = $cn2; NewDisplayName = $cn2; OidHandling = 'GenerateRandom'; SkipAcl = $true }
            $t2 = script:Get-ADObjectRetry -AdParams $script:AP -Identity $r2.DN -Properties nTSecurityDescriptor
            $t2.nTSecurityDescriptor.AreAccessRulesProtected | Should -BeFalse
        }

        It 'Import -AclBase PrincipalsOnly writes exactly the requested grants' {
            $cn = script:New-LabName
            $r = script:Invoke-LabCreate -SyncParams @{ Mode = 'Import'; Path = $script:ExportFile; Server = $AronsServer; NewTemplateName = $cn; NewDisplayName = $cn; OidHandling = 'GenerateRandom'; AclBase = 'PrincipalsOnly'; EnrollPrincipals = @{ 'AuthenticatedUsers' = 'Read'; 'DomainControllers' = 'Enroll', 'Autoenroll' } }
            $t = script:Get-ADObjectRetry -AdParams $script:AP -Identity $r.DN -Properties nTSecurityDescriptor
            $rules = @($t.nTSecurityDescriptor.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]))
            ($rules | Where-Object { $_.IdentityReference.Value -eq 'S-1-5-11' }) | Should -Not -BeNullOrEmpty
            $enrollGuid = [Guid]'0e10c968-78fb-11d2-90d4-00c04f79dc55'
            $dcSid = (Get-ADDomain @script:AP).DomainSID.Value + '-516'
            ($rules | Where-Object { $_.IdentityReference.Value -eq $dcSid -and $_.ObjectType -eq $enrollGuid }) | Should -Not -BeNullOrEmpty
        }

        It 'Import refuses a duplicate cn and a duplicate OID' {
            { & $script:Sync -Mode Import -Path $script:ExportFile -Server $AronsServer `
                    -NewTemplateName $SourceTemplate -NewDisplayName x -OidHandling GenerateRandom } |
                Should -Throw -ExpectedMessage '*already exists*'
            { & $script:Sync -Mode Import -Path $script:ExportFile -Server $AronsServer `
                    -NewTemplateName (script:New-LabName) -NewDisplayName x -OidHandling Preserve } |
                Should -Throw -ExpectedMessage '*already carries OID*'
        }

        It 'Import rejects a tampered OID and an injected cn from the file' {
            $bad = Join-Path $script:TmpDir 'bad.json'
            $j = Get-Content -LiteralPath $script:ExportFile -Raw -Encoding UTF8 | ConvertFrom-Json
            $j.'msPKI-Cert-Template-OID' = '1.2.3.*)(cn=*'
            $j | ConvertTo-Json -Depth 10 | Out-File -Encoding utf8 $bad
            { & $script:Sync -Mode Import -Path $bad -Server $AronsServer -NewTemplateName (script:New-LabName) -NewDisplayName x } |
                Should -Throw -ExpectedMessage '*not a valid dotted OID*'

            $j2 = Get-Content -LiteralPath $script:ExportFile -Raw -Encoding UTF8 | ConvertFrom-Json
            $j2.name = 'evil,cn=x'
            $j2 | ConvertTo-Json -Depth 10 | Out-File -Encoding utf8 $bad
            { & $script:Sync -Mode Import -Path $bad -Server $AronsServer -OidHandling GenerateRandom } |
                Should -Throw -ExpectedMessage '*may only contain*'
        }

        It 'Validate passes both pipelines for the source template' {
            $out = & $script:Sync -Mode Validate -TemplateName $SourceTemplate -Server $AronsServer *>&1 | Out-String
            $out | Should -Match 'OVERALL PASS'
            $out | Should -Match 'Cleaned up'
            # Validate cleans up its own throwaways; nothing to track.
        }

        It 'Sync refuses a same-forest copy unless -Server is explicit' {
            { & $script:Sync -Mode Sync -SourceServer $AronsServer -NewTemplateName (script:New-LabName) `
                    -NewDisplayName x -OidHandling GenerateRandom } |
                Should -Throw -ExpectedMessage '*SAME forest*'
        }

        Context 'Child domain' -Skip:(-not $script:ChildReady) {
            It 'root-SID ACEs (RID 498/519) resolve to the forest-root domain via the child DC' {
                $cn = script:New-LabName
                $cp = @{ Server = $ChildServer }
                $r = script:Invoke-LabCreate -SyncParams @{ Mode = 'Import'; Path = $script:ExportFile; Server = $ChildServer; NewTemplateName = $cn; NewDisplayName = $cn; OidHandling = 'GenerateRandom'; AclBase = 'Standard' }
                $r.DN | Should -Not -BeNullOrEmpty
                $t = script:Get-ADObjectRetry -AdParams $cp -Identity $r.DN -Properties nTSecurityDescriptor
                $rootSid = (Get-ADDomain -Server (Get-ADForest @cp).RootDomain).DomainSID.Value
                $rules = @($t.nTSecurityDescriptor.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]))
                ($rules | Where-Object { $_.IdentityReference.Value -eq "$rootSid-519" }) | Should -Not -BeNullOrEmpty
                ($rules | Where-Object { $_.IdentityReference.Value -eq "$rootSid-498" }) | Should -Not -BeNullOrEmpty
            }
        }

        Context 'Cross-forest (no-AD CS target)' -Skip:(-not $script:XForestReady) {
            It 'Sync creates the template in the target forest and resolves principals there' {
                $cn = script:New-LabName
                $tp = @{ Server = $NorefjellServer }
                $r = script:Invoke-LabCreate -SyncParams @{ Mode = 'Sync'; SourceServer = $AronsServer; Server = $NorefjellServer; TemplateName = $SourceTemplate; NewTemplateName = $cn; NewDisplayName = $cn; OidHandling = 'GenerateRandom'; AclBase = 'PrincipalsOnly'; EnrollPrincipals = @{ 'DomainControllers' = 'Enroll' } }
                $r.DN | Should -Not -BeNullOrEmpty -Because 'the template must be created in the target forest'
                $r.DN | Should -BeLike '*DC=norefjell,DC=local' -Because 'creation must land in the target forest'
                $t = script:Get-ADObjectRetry -AdParams $tp -Identity $r.DN -Properties nTSecurityDescriptor
                $tgtDcSid = (Get-ADDomain @tp).DomainSID.Value + '-516'
                $rules = @($t.nTSecurityDescriptor.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]))
                ($rules | Where-Object { $_.IdentityReference.Value -eq $tgtDcSid }) |
                    Should -Not -BeNullOrEmpty -Because 'DomainControllers must resolve in the TARGET forest'
            }

            It 'Generate fails cleanly against a forest with no PKI OID root' {
                { & $script:Sync -Mode Sync -SourceServer $AronsServer -Server $NorefjellServer -TemplateName $SourceTemplate `
                        -NewTemplateName (script:New-LabName) -NewDisplayName x -OidHandling Generate -AclBase Schema -WhatIf } |
                    Should -Throw -ExpectedMessage '*no PKI OID root*'
            }
        }
    }
}
