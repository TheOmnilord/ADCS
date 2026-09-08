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
                   -RunLab is passed. Cleanup is surgical: an object is tracked only by the DN and
                   objectGUID the script printed at creation (or a -PassThru returned to the
                   test), never by a later lookup, and it is removed only when the object at its
                   DN still carries the recorded objectGUID. The
                   objects carry a unique PESTER-<hex> prefix per run; the AfterAll prefix sweep
                   reports an untracked prefixed object but never deletes it. Pre-existing
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
        # Resolve-PrincipalSid, Import-Template and Export-Template reach the directory only through
        # Get-ADObject / New-ADObject / Remove-ADObject / Get-ADObjectIfPresent / Get-SchemaAttributeType,
        # which their Unit Contexts shadow with stub functions - no AD module is needed for them.
        foreach ($name in 'Get-RandomHex', 'ConvertTo-LdapFilterValue', 'New-SyntheticOidBase', 'Get-AttrCanonical', 'Compare-TemplateAttributes', 'Convert-ToLatestCompatibility', 'ConvertTo-ImportAttributeValue', 'Get-LinkedIssuancePolicy',
                         'ConvertTo-SchemaTypedValue', 'Get-RootDomainSid', 'Get-WellKnownTokenSid', 'Resolve-PrincipalSid', 'Export-Template', 'Resolve-OidDisplay', 'Resolve-TemplateOid', 'Import-Template') {
            $def = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
            $def | Should -Not -BeNullOrEmpty -Because "function $name must be defined in the script"
            . ([scriptblock]::Create($def[0].Extent.Text))
        }
        # Get-AttrCanonical and ConvertTo-ImportAttributeValue reference the script's attribute-type
        # lists: extract the very assignments, so the tests can never drift from the script.
        foreach ($listName in '$script:IntAttributes', '$script:MultiValueAttributes', '$script:ByteAttributes', '$script:OidListAttributes') {
            $asg = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq $listName }, $true)
            $asg | Should -Not -BeNullOrEmpty -Because "$listName must be defined in the script"
            . ([scriptblock]::Create($asg[0].Extent.Text))
        }

        # Asserts that a DACL is EXACTLY the expected grants: one Allow ACE per expected
        # (SID, right, object GUID), nothing inherited, no deny, no ACE for any other principal, and
        # no ACE carrying bits beyond its expected right - the DS may expand a generic right into
        # its specific bits (GenericRead -> ReadProperty|ListChildren|ListObject|ReadControl, and
        # GenericWrite -> WriteProperty|Self|ReadControl), which is tolerated; WriteDacl, WriteOwner,
        # Delete, CreateChild, an unexpected ExtendedRight or anything else is not.
        function script:Assert-ExactDacl {
            param([object[]]$Rules, [object[]]$Expected)   # Expected: @{ Sid; Right ('read'|'write'|'enroll'|'autoenroll'|'fullcontrol') }
            $enrollGuid     = [Guid]'0e10c968-78fb-11d2-90d4-00c04f79dc55'
            $autoenrollGuid = [Guid]'a05b8cc2-17bc-4802-a710-e7c15ab866a2'
            $spec = @{
                read       = @{ Bits = [int][System.DirectoryServices.ActiveDirectoryRights]::GenericRead;    Allow = 0x80000000 -bor 0x20094; Guid = [Guid]::Empty }
                write      = @{ Bits = [int][System.DirectoryServices.ActiveDirectoryRights]::GenericWrite;   Allow = 0x40000000 -bor 0x20028; Guid = [Guid]::Empty }
                enroll     = @{ Bits = [int][System.DirectoryServices.ActiveDirectoryRights]::ExtendedRight;  Allow = 0x100;                  Guid = $enrollGuid }
                autoenroll = @{ Bits = [int][System.DirectoryServices.ActiveDirectoryRights]::ExtendedRight;  Allow = 0x100;                  Guid = $autoenrollGuid }
            }
            $rules = @($Rules)
            @($rules | Where-Object { $_.IsInherited }).Count | Should -Be 0 -Because 'the DACL is protected; nothing may be inherited'
            @($rules | Where-Object { $_.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow }).Count | Should -Be 0
            $expectedSids = @($Expected | ForEach-Object { $_.Sid } | Sort-Object -Unique)
            @($rules | ForEach-Object { $_.IdentityReference.Value } | Sort-Object -Unique) | Should -Be $expectedSids -Because 'no principal beyond the expected ones may hold anything'
            # The DS coalesces same-typed grants for one principal into ONE ACE (Domain Admins'
            # read + write become a single GenericRead|GenericWrite entry), so the unit of
            # comparison is (principal, object GUID): one ACE each, carrying exactly the union of
            # the expected rights and nothing beyond their tolerated expansion.
            $groups = @{}
            foreach ($e in $Expected) {
                $s = $spec[$e.Right]; $k = "$($e.Sid)|$($s.Guid)"
                if (-not $groups.ContainsKey($k)) { $groups[$k] = @{ Sid = $e.Sid; Guid = $s.Guid; Bits = [int64]0; Allow = [int64]0; Rights = @() } }
                $groups[$k].Bits   = $groups[$k].Bits  -bor (([int64]$s.Bits) -band 0xFFFFFFFF)
                $groups[$k].Allow  = $groups[$k].Allow -bor ([int64]$s.Allow)
                $groups[$k].Rights += $e.Right
            }
            $rules.Count | Should -Be $groups.Count -Because 'exactly one ACE per (principal, object type)'
            foreach ($g in $groups.Values) {
                $match = @($rules | Where-Object { $_.IdentityReference.Value -eq $g.Sid -and $_.ObjectType -eq $g.Guid })
                $match.Count | Should -Be 1 -Because "exactly one ACE for $($g.Sid) / $($g.Guid)"
                $bits = ([int64][int]$match[0].ActiveDirectoryRights) -band 0xFFFFFFFF
                ($bits -band $g.Bits) | Should -Be $g.Bits -Because "the ACE for $($g.Sid) must carry $($g.Rights -join '+')"
                $extra = $bits -band (-bnot $g.Allow)
                $extra | Should -Be 0 -Because "the ACE for $($g.Sid) must carry no right beyond $($g.Rights -join '+') (extra bits 0x$('{0:X}' -f $extra))"
            }
        }

        # A REAL PSCmdlet for the helpers that take -CallerCmdlet: its ShouldProcess answers $true
        # here (or $false under -WhatIf), exactly as the script's own $PSCmdlet does. The body
        # receives the cmdlet as its first argument.
        function script:Invoke-WithCmdlet {
            [CmdletBinding(SupportsShouldProcess)]
            param([scriptblock]$Body)
            & $Body $PSCmdlet
        }

        # Parse ONE "<Prefix> <DN> (objectGUID <guid>)" line the script prints when it creates an
        # object ("Created template:" / "Created OID object:"). The DN is non-greedy and the GUID
        # suffix is anchored at the end of the line, so a DN that itself contains parentheses (the
        # cn allowlist permits them) still parses. Returns @{ DN; Guid } - Guid is $null when the
        # line carries no objectGUID - or $null when no such line exists. The Lab tier tracks the
        # objects it must remove ONLY through this parser, and the Unit tier checks it against the
        # very lines the script prints, so the producer and the consumer cannot drift apart.
        # (?m): ^ and $ work per line; \s* before $ also eats the \r of a CRLF line.
        function script:Read-LabCreatedLine {
            param([Parameter(Mandatory)][AllowEmptyString()][string]$Output, [Parameter(Mandatory)][string]$Prefix)
            $p = [regex]::Escape($Prefix)
            $g = '[0-9A-Fa-f]{8}(?:-[0-9A-Fa-f]{4}){3}-[0-9A-Fa-f]{12}'
            if ($Output -match "(?m)^$p\s*(.+?)\s*\(objectGUID ($g)\)\s*`$") { return @{ DN = $Matches[1]; Guid = [guid]$Matches[2] } }
            if ($Output -match "(?m)^$p\s*(.+?)\s*`$") { return @{ DN = $Matches[1]; Guid = $null } }
            return $null
        }

        # Exclusive-creation evidence for a folder. New-Item -ItemType Directory and
        # [IO.Directory]::CreateDirectory both succeed when the folder already exists, and a
        # Test-Path before them is a separate operation - another process can create the name in
        # between, and the caller would then "own" (and later remove recursively) a foreign folder.
        # Win32 CreateDirectoryW is atomic: it returns false with error 183 (ERROR_ALREADY_EXISTS)
        # when the name is taken, so a true result proves THIS call made the folder. The last-error
        # read happens in the same C# frame as the call: a PowerShell frame could run other interop
        # (with SetLastError) between the two and clobber the value.
        if (-not ('PesterSyncFsNative' -as [type])) {
            Add-Type -TypeDefinition @"
using System; using System.Runtime.InteropServices;
public static class PesterSyncFsNative {
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool CreateDirectoryW(string lpPathName, IntPtr lpSecurityAttributes);
    // 0 when this call created the folder; otherwise the Win32 error (183 = ERROR_ALREADY_EXISTS).
    public static int CreateDirectory(string path) {
        return CreateDirectoryW(path, IntPtr.Zero) ? 0 : Marshal.GetLastWin32Error();
    }
}
"@
        }
        # Returns 'Created' when THIS call made the folder and 'AlreadyExists' when the name was
        # already taken (folder or file). Any other failure throws with the Win32 error. The Lab
        # tier records ownership of its scratch folder ONLY from a 'Created' result.
        function script:New-LabExclusiveDirectory {
            param([Parameter(Mandatory)][string]$Path)
            $err = [PesterSyncFsNative]::CreateDirectory($Path)
            if ($err -eq 0)   { return 'Created' }
            if ($err -eq 183) { return 'AlreadyExists' }
            throw "CreateDirectoryW failed for '$Path' (Win32 error $err)."
        }

        # Removes a Lab scratch folder by proof, not by name. Exclusive creation of the folder
        # proves ownership of the folder only - not of what is inside it later. So: only the
        # files in -TrackedFiles go (the exact paths the suite recorded when it named them), one
        # file at a time, with Remove-Item on that path and never a folder. Then the folder
        # itself goes NON-recursively: Directory.Delete without the recursive flag fails when
        # anything remains, and anything that remains is not this run's - it stays and is
        # reported by path. A tracked path that is not directly under -Path, or that is now a
        # folder (File.Exists is false for a folder), is never removed. Never Remove-Item -Recurse.
        function script:Remove-LabScratchFolder {
            param([Parameter(Mandatory)][string]$Path, [string[]]$TrackedFiles)
            foreach ($f in @($TrackedFiles | Where-Object { $_ })) {
                $parent = [System.IO.Path]::GetDirectoryName($f)
                if (-not [string]::Equals($parent, $Path, [System.StringComparison]::OrdinalIgnoreCase)) {
                    Write-Warning "Scratch cleanup left '$f' alone: it is not directly under '$Path'."
                    continue
                }
                if (-not [System.IO.File]::Exists($f)) { continue }   # never written, already gone, or a folder now
                try { Remove-Item -LiteralPath $f -Force -ErrorAction Stop }
                catch { Write-Warning "Scratch cleanup could not remove tracked file '$f': $($_.Exception.Message)" }
            }
            if (-not [System.IO.Directory]::Exists($Path)) { return }
            $left = @([System.IO.Directory]::GetFileSystemEntries($Path))
            if ($left.Count -gt 0) {
                Write-Warning "Scratch folder '$Path' holds $($left.Count) entries this run did not create, so it was NOT removed - review it manually."
                foreach ($e in $left) { Write-Warning "  untracked entry left in place: $e" }
                return
            }
            try { [System.IO.Directory]::Delete($Path) }   # non-recursive: fails when anything remains
            catch { Write-Warning "Scratch cleanup could not remove the empty folder '$Path': $($_.Exception.Message)" }
        }

        # Evidence that THIS run's -Mode Export call wrote the file at -Path. Existence is not
        # evidence: the script can fail before it writes (source lookup, ambiguous name) while
        # another process puts a file of that name in the scratch folder. Two checks, both must
        # hold: (a) the script's own "Export completed: <path>" line (Export-Template prints it
        # right after WriteAllText, with the resolved full path) names this exact path in -Output
        # (un-wrapped, Out-String -Width 4096), and (b) the file parses as JSON and every property
        # in -Expect carries the value this run exported (the source name, or its OID and schema
        # version when the name was stripped). Returns $null when both hold, else the reason.
        # The Unit tier checks this against the very line the script prints.
        function script:Test-LabExportEvidence {
            param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][AllowEmptyString()][string]$Output, [Parameter(Mandatory)][hashtable]$Expect)
            $full = [System.IO.Path]::GetFullPath($Path)
            $named = $false
            foreach ($m in [regex]::Matches($Output, '(?m)^Export completed:\s*(.+?)\s*$')) {
                $printed = $m.Groups[1].Value
                try { $printed = [System.IO.Path]::GetFullPath($printed) } catch { continue }
                if ([string]::Equals($printed, $full, [System.StringComparison]::OrdinalIgnoreCase)) { $named = $true; break }
            }
            if (-not $named) { return "the script printed no 'Export completed:' line for '$Path'" }
            if (-not [System.IO.File]::Exists($Path)) { return "no file exists at '$Path'" }
            try { $j = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8) | ConvertFrom-Json -ErrorAction Stop }
            catch { return "the file at '$Path' is not JSON: $($_.Exception.Message)" }
            if ($null -eq $j) { return "the file at '$Path' is empty" }
            foreach ($k in $Expect.Keys) {
                $prop = $j.PSObject.Properties[$k]
                if ($null -eq $prop) { return "the file at '$Path' has no '$k' property" }
                if ("$($prop.Value)" -ne "$($Expect[$k])") { return "the file at '$Path' has '$k' = '$($prop.Value)', expected '$($Expect[$k])'" }
            }
            return $null
        }

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

        It 'Get-AttrCanonical: a trailing backslash cannot forge the escaped separator' {
            # The escape character itself is escaped first (\ -> \\), so a value ending in '\'
            # followed by the joiner ('\\|') can never read as an escaped pipe ('\|'): @('a\','b')
            # and @('a|b') must stay distinct. (A review once claimed -replace '\\','\\' was a
            # no-op; .NET replacement strings do not process backslashes, so it does double them.)
            (Get-AttrCanonical -Name 'x' -Value @('a\', 'b')) |
                Should -Not -Be (Get-AttrCanonical -Name 'x' -Value @('a|b'))
            (Get-AttrCanonical -Name 'x' -Value @('a\')) | Should -BeExactly 'a\\'
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

        It 'ConvertTo-ImportAttributeValue: integer attributes accept only integral values in the Int32 range - anything else is REFUSED, never dropped or coerced' {
            ConvertTo-ImportAttributeValue -Name 'msPKI-RA-Signature' -Value 1          | Should -Be 1
            (ConvertTo-ImportAttributeValue -Name 'msPKI-RA-Signature' -Value 1) -is [int] | Should -BeTrue
            ConvertTo-ImportAttributeValue -Name 'msPKI-RA-Signature' -Value ([long]2)   | Should -Be 2
            ConvertTo-ImportAttributeValue -Name 'msPKI-RA-Signature' -Value 3.0        | Should -Be 3
            ConvertTo-ImportAttributeValue -Name 'msPKI-Private-Key-Flag' -Value (-1)   | Should -Be -1
            ConvertTo-ImportAttributeValue -Name 'msPKI-RA-Signature' -Value @(4)       | Should -Be 4
            # A giant or non-finite JSON number overflows the [decimal] cast; it must still be
            # refused with the ATTRIBUTE-NAMED message (a bare cast threw a raw ".NET cannot convert"
            # error that did not name the attribute), on both 5.1 and 7.
            # 1e-30 is the DECIMAL-UNDERFLOW case: [decimal]1e-30 is 0, so a post-cast integer check
            # would silently coerce this malformed value to 0. It must be refused, not coerced.
            foreach ($bad in @(0.4, '5', 'abc', $true, @(1, 2), 4294967296, 1e40, 1e-30, 5e-28, ([double]::PositiveInfinity), ([double]::NaN))) {
                { ConvertTo-ImportAttributeValue -Name 'msPKI-RA-Signature' -Value $bad } |
                    Should -Throw -ExpectedMessage '*Import refused*msPKI-RA-Signature*' -Because "'$bad' must be refused with the attribute named (a bare cast dropped, coerced, underflowed, or threw an un-attributed error)"
            }
            { ConvertTo-ImportAttributeValue -Name 'msPKI-RA-Signature' -Value @() }   | Should -Throw -ExpectedMessage '*Import refused*'
            { ConvertTo-ImportAttributeValue -Name 'msPKI-RA-Signature' -Value $null } | Should -Throw -ExpectedMessage '*Import refused*'
            { ConvertTo-ImportAttributeValue -Name 'msPKI-RA-Signature' -Value '' }    | Should -Throw -ExpectedMessage '*Import refused*'
        }

        It 'ConvertTo-ImportAttributeValue: period attributes are exactly 8 bytes, key usage 1-2, every element a byte' {
            $p = ConvertTo-ImportAttributeValue -Name 'pKIExpirationPeriod' -Value @(0, 64, 57, 135, 46, 225, 254, 255)
            $p -is [byte[]] | Should -BeTrue
            $p.Count | Should -Be 8
            (ConvertTo-ImportAttributeValue -Name 'pKIKeyUsage' -Value @(160, 0)).Count | Should -Be 2
            (ConvertTo-ImportAttributeValue -Name 'pKIKeyUsage' -Value @(128)).Count    | Should -Be 1
            { ConvertTo-ImportAttributeValue -Name 'pKIOverlapPeriod' -Value @(1, 2, 3) } | Should -Throw -ExpectedMessage '*Import refused*pKIOverlapPeriod*' -Because 'a bare [byte[]] cast accepted a three-byte "period"'
            { ConvertTo-ImportAttributeValue -Name 'pKIKeyUsage' -Value @(1, 2, 3) }     | Should -Throw -ExpectedMessage '*Import refused*'
            { ConvertTo-ImportAttributeValue -Name 'pKIExpirationPeriod' -Value @(0, 64, 57, 135, 46, 225, 254, 256) }  | Should -Throw -ExpectedMessage '*not a byte*'
            { ConvertTo-ImportAttributeValue -Name 'pKIExpirationPeriod' -Value @(0, 64, 57, 135, 46, 225, 254, 'ff') } | Should -Throw -ExpectedMessage '*not a byte*'
            # decimal-underflow element: [decimal]1e-30 is 0, so this must be refused, not coerced to a 0 byte
            { ConvertTo-ImportAttributeValue -Name 'pKIExpirationPeriod' -Value @(0, 64, 57, 135, 46, 225, 254, 1e-30) } | Should -Throw -ExpectedMessage '*not a byte*'
            { ConvertTo-ImportAttributeValue -Name 'pKIExpirationPeriod' -Value 'AEA5hy7h/v8=' } | Should -Throw -ExpectedMessage '*got a string*'
        }

        It 'ConvertTo-ImportAttributeValue: multi-value attributes take non-empty strings only; OID lists must be dotted OIDs' {
            # the function returns ONE array of strings (the caller casts it to the AD collection type)
            $one = ConvertTo-ImportAttributeValue -Name 'pKIExtendedKeyUsage' -Value '1.3.6.1.5.5.7.3.1'
            $one -is [array] | Should -BeTrue
            $one.Count | Should -Be 1
            $one[0] | Should -BeExactly '1.3.6.1.5.5.7.3.1'
            (ConvertTo-ImportAttributeValue -Name 'msPKI-Certificate-Policy' -Value @('1.3.6.1.4.1.311.21.8.1.2', '1.3.6.1.4.1.311.21.8.1.3')).Count | Should -Be 2
            (ConvertTo-ImportAttributeValue -Name 'pKIDefaultCSPs' -Value @('1,Microsoft RSA SChannel Cryptographic Provider')).Count | Should -Be 1
            { ConvertTo-ImportAttributeValue -Name 'msPKI-Certificate-Policy' -Value 'not-an-oid' }   | Should -Throw -ExpectedMessage '*not a dotted OID*'
            { ConvertTo-ImportAttributeValue -Name 'msPKI-Certificate-Policy' -Value @('1.2.3', 5) }  | Should -Throw -ExpectedMessage '*expected a string*'
            { ConvertTo-ImportAttributeValue -Name 'pKIExtendedKeyUsage' -Value @() }                | Should -Throw -ExpectedMessage '*no values*'
            { ConvertTo-ImportAttributeValue -Name 'pKIDefaultCSPs' -Value "1,Provider`r`nX" }        | Should -Throw -ExpectedMessage '*control characters*'
            { ConvertTo-ImportAttributeValue -Name 'msPKI-Supersede-Templates' -Value @('') }        | Should -Throw -ExpectedMessage '*empty element*'
        }

        It 'the dotted-OID validation pattern is anchored with \z, so a trailing newline cannot slip past the OID uniqueness guard' {
            # $ matches before a trailing LF, so a tampered msPKI-Cert-Template-OID ending in "\n"
            # passed validation and then missed the (msPKI-Cert-Template-OID=<oid>) uniqueness search
            # (AD does not fold the newline). \z anchors at the true end of the string.
            $pat = '^(0|[1-9]\d*)(\.(0|[1-9]\d*))+\z'
            '1.3.6.1.4.1.311.21.8.1.2'            | Should -Match $pat
            ("1.3.6.1.4.1.311.21.8.1.2" + "`n")   | Should -Not -Match $pat -Because 'a trailing LF must not pass ($ would have allowed it)'
            ("1.3.6.1.4.1.311.21.8.1.2" + "`r`n") | Should -Not -Match $pat
            # tie the assertion to the real code: Resolve-TemplateOid must use \z, not $
            (Get-Content -LiteralPath $script:Sync -Raw) | Should -Match ([regex]::Escape('(\.(0|[1-9]\d*))+\z')) -Because 'the OID validations must be anchored with \z'
            (Get-Content -LiteralPath $script:Sync -Raw) | Should -Not -Match ([regex]::Escape('(\.(0|[1-9]\d*))+$')) -Because 'no OID validation may still use the newline-permissive $'
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

        It 'Convert-ToLatestCompatibility: a schema-3 source with a KSP provider list keeps CNG semantics (no 0x100); its own legacy bit is preserved either way' {
            # A v3 template legitimately lists KSPs in pKIDefaultCSPs ("1,Microsoft Software Key
            # Storage Provider"). A populated list is NOT evidence of a CSP template - only schema-2
            # (2003) sources are CSP-only. The source's own CT_FLAG_USE_LEGACY_PROVIDER bit decides.
            $ksp = @{ 'msPKI-Template-Schema-Version' = [int]3; 'msPKI-Private-Key-Flag' = [int]0
                      'pKIDefaultCSPs' = @('1,Microsoft Software Key Storage Provider') }
            $r = Convert-ToLatestCompatibility -Attributes $ksp
            $r.LegacyProvider | Should -BeFalse -Because 'setting 0x100 would switch the v4 copy to legacy CryptoAPI key handling'
            [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$ksp['msPKI-Private-Key-Flag']), 0) | Should -Be ([uint32]0x06060000)
            $v3csp = @{ 'msPKI-Template-Schema-Version' = [int]3; 'msPKI-Private-Key-Flag' = [int]0x100
                        'pKIDefaultCSPs' = @('1,Microsoft RSA SChannel Cryptographic Provider') }
            $r = Convert-ToLatestCompatibility -Attributes $v3csp
            $r.LegacyProvider | Should -BeTrue -Because 'the source itself carries the legacy-provider bit'
            [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$v3csp['msPKI-Private-Key-Flag']), 0) | Should -Be ([uint32]0x06060100)
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

        It 'Get-LinkedIssuancePolicy scans only msPKI-Certificate-Policy (issued-cert policy), not msPKI-RA-Policies (signing-cert requirement)' {
            # An OID only in msPKI-RA-Policies yields NO AMA candidates, so the function returns before
            # touching AD (the empty-OID early return) - a stub -ADParams is never dereferenced. This
            # proves msPKI-RA-Policies is excluded: those policies constrain the enrollment-agent
            # signing certificate, they are not stamped into the issued cert, so an AMA link on one
            # grants the enrollee nothing and must not refuse the import.
            Get-LinkedIssuancePolicy -Attributes @{ 'msPKI-RA-Policies' = @('1.3.6.1.4.1.311.10.3.10') } -ConfigNC 'DC=x' -ADParams @{} | Should -BeNullOrEmpty
            Get-LinkedIssuancePolicy -Attributes @{} -ConfigNC 'DC=x' -ADParams @{} | Should -BeNullOrEmpty
            Get-LinkedIssuancePolicy -Attributes @{ 'msPKI-Certificate-Policy' = $null } -ConfigNC 'DC=x' -ADParams @{} | Should -BeNullOrEmpty
        }

        It 'Convert-ToLatestCompatibility: a minor revision at Int32.MaxValue is refused ATOMICALLY (no half-upgrade)' {
            # Incrementing Int32.MaxValue overflowed the [Int32] cast; the old order set schema/flags
            # first, so the template was left half-upgraded (v4 schema, un-bumped revision) and still
            # reported Upgraded. Now the overflow throws before ANY attribute is mutated.
            $a = @{ 'msPKI-Template-Schema-Version' = [int]2; 'msPKI-Private-Key-Flag' = [int]0
                    'flags' = [BitConverter]::ToInt32([BitConverter]::GetBytes([uint32]0x10000), 0)
                    'msPKI-Template-Minor-Revision' = [int]::MaxValue }
            { Convert-ToLatestCompatibility -Attributes $a } | Should -Throw -ExpectedMessage '*Int32.MaxValue*'
            $a['msPKI-Template-Schema-Version']   | Should -Be 2      -Because 'nothing may be mutated when the revision cannot be incremented'
            $a['msPKI-Private-Key-Flag']          | Should -Be 0
            $a['msPKI-Template-Minor-Revision']   | Should -Be ([int]::MaxValue)
        }

        It 'Convert-ToLatestCompatibility: a schema-2 source with msPKI-RA-Application-Policies is NOT upgraded and the hashtable stays untouched' {
            # That attribute's encoding differs at v3/v4; stamping v4 would drop the RA-signature
            # application-policy requirement (the 1.0.5 fix). Refused before any mutation.
            $a = @{ 'msPKI-Template-Schema-Version' = [int]2; 'msPKI-Private-Key-Flag' = [int]0; 'flags' = [int]0x10000
                    'msPKI-Template-Minor-Revision' = [int]3; 'msPKI-RA-Application-Policies' = @('1.3.6.1.4.1.311.10.3.10') }
            $r = Convert-ToLatestCompatibility -Attributes $a
            $r.Upgraded    | Should -BeFalse
            $r.FromVersion | Should -Be 2
            $r.Reason      | Should -Match 'msPKI-RA-Application-Policies'
            $a['msPKI-Template-Schema-Version'] | Should -Be 2
            $a['msPKI-Private-Key-Flag']        | Should -Be 0
            $a['flags']                         | Should -Be 0x10000
            $a['msPKI-Template-Minor-Revision'] | Should -Be 3
        }

        It 'Convert-ToLatestCompatibility: an absent schema version is treated as v1 (left alone); an absent private-key flag starts from 0' {
            $noVer = @{ 'msPKI-Private-Key-Flag' = [int]0 }
            $r = Convert-ToLatestCompatibility -Attributes $noVer
            $r.Upgraded    | Should -BeFalse
            $r.FromVersion | Should -Be 1
            $noVer.ContainsKey('msPKI-Template-Schema-Version') | Should -BeFalse -Because 'a v1 source is not upgradable in place, so nothing is added'
            $noPkf = @{ 'msPKI-Template-Schema-Version' = [int]3 }
            $r = Convert-ToLatestCompatibility -Attributes $noPkf
            $r.Upgraded | Should -BeTrue
            $noPkf['msPKI-Template-Schema-Version'] | Should -Be 4
            [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$noPkf['msPKI-Private-Key-Flag']), 0) | Should -Be ([uint32]0x06060000) -Because 'an absent flag is 0 plus the Server 2016 / Windows 10 nibbles'
        }

        It 'ConvertTo-ImportAttributeValue: JSON numbers arrive in the engine''s own type (5.1: Decimal for a fraction, 7: Double) and are judged by value' {
            # ConvertFrom-Json yields [decimal] for 0.4 on Windows PowerShell 5.1 and [double] on
            # PowerShell 7, so each engine drives its own branch of the integer validation here.
            $frac  = ('{"v":0.4}'        | ConvertFrom-Json).v
            $big   = ('{"v":4294967296}' | ConvertFrom-Json).v
            $seven = ('{"v":7}'          | ConvertFrom-Json).v
            if ($PSVersionTable.PSVersion.Major -le 5) { $frac -is [decimal] | Should -BeTrue -Because 'Windows PowerShell parses a JSON fraction as Decimal' }
            else                                        { $frac -is [double]  | Should -BeTrue -Because 'PowerShell 7 parses a JSON fraction as Double' }
            { ConvertTo-ImportAttributeValue -Name 'msPKI-RA-Signature' -Value $frac } | Should -Throw -ExpectedMessage '*Import refused*msPKI-RA-Signature*expected an integer*'
            { ConvertTo-ImportAttributeValue -Name 'msPKI-RA-Signature' -Value $big }  | Should -Throw -ExpectedMessage '*Import refused*msPKI-RA-Signature*outside the Int32 range*'
            $r = ConvertTo-ImportAttributeValue -Name 'msPKI-RA-Signature' -Value $seven
            $r | Should -Be 7
            $r -is [int] | Should -BeTrue
            # the [decimal] branch itself, on both engines: fraction and out-of-range refused, integral accepted
            { ConvertTo-ImportAttributeValue -Name 'msPKI-RA-Signature' -Value ([decimal]0.4) }        | Should -Throw -ExpectedMessage '*expected an integer*'
            { ConvertTo-ImportAttributeValue -Name 'msPKI-RA-Signature' -Value ([decimal]4294967296) } | Should -Throw -ExpectedMessage '*outside the Int32 range*'
            ConvertTo-ImportAttributeValue -Name 'msPKI-RA-Signature' -Value ([decimal]7) | Should -Be 7
        }

        It 'ConvertTo-ImportAttributeValue: an unsupported value type is refused by type name; a byte element accepts an integral double and refuses a fraction or $null' {
            { ConvertTo-ImportAttributeValue -Name 'msPKI-RA-Signature' -Value (Get-Date) } | Should -Throw -ExpectedMessage '*Import refused*expected an integer, got DateTime*'
            $p = ConvertTo-ImportAttributeValue -Name 'pKIKeyUsage' -Value @([double]2.0, [double]0.0)
            $p -is [byte[]] | Should -BeTrue
            $p[0] | Should -Be 2
            { ConvertTo-ImportAttributeValue -Name 'pKIKeyUsage' -Value @([double]2.5, 0) } | Should -Throw -ExpectedMessage '*element 0 (2.5) is not a byte*'
            { ConvertTo-ImportAttributeValue -Name 'pKIKeyUsage' -Value @(1, $null) }       | Should -Throw -ExpectedMessage '*element 1 is not a byte*'
        }

        It 'ConvertTo-LdapFilterValue escapes NUL as \00' {
            ConvertTo-LdapFilterValue ('a' + [char]0 + 'b') | Should -Be 'a\00b'
        }

        It 'ConvertTo-SchemaTypedValue: Int and String take exactly one element, MultiString and Bytes cast, anything else or a failed cast is $null (not consumed)' {
            $i = ConvertTo-SchemaTypedValue -SchemaType 'Int' -Value 7
            $i | Should -Be 7
            $i -is [int] | Should -BeTrue
            ConvertTo-SchemaTypedValue -SchemaType 'Int' -Value @(7)      | Should -Be 7
            ConvertTo-SchemaTypedValue -SchemaType 'Int' -Value ([long]7) | Should -Be 7
            ConvertTo-SchemaTypedValue -SchemaType 'Int' -Value @(1, 2)   | Should -BeNullOrEmpty -Because 'a single-valued target attribute cannot take two elements'
            ConvertTo-SchemaTypedValue -SchemaType 'Int' -Value @()       | Should -BeNullOrEmpty -Because 'an empty array must not fabricate 0'
            ConvertTo-SchemaTypedValue -SchemaType 'Int' -Value 'abc'     | Should -BeNullOrEmpty -Because 'a failed cast drops the attribute instead of throwing'
            $s = ConvertTo-SchemaTypedValue -SchemaType 'String' -Value 'x'
            $s | Should -BeExactly 'x'
            $s -is [string] | Should -BeTrue
            ConvertTo-SchemaTypedValue -SchemaType 'String' -Value @('x', 'y') | Should -BeNullOrEmpty
            $m = ConvertTo-SchemaTypedValue -SchemaType 'MultiString' -Value @('a', 'b')
            $m -is [object[]] | Should -BeTrue
            $m.Count | Should -Be 2
            $m[1] | Should -BeExactly 'b'
            (ConvertTo-SchemaTypedValue -SchemaType 'MultiString' -Value 'solo').Count | Should -Be 1
            $b = ConvertTo-SchemaTypedValue -SchemaType 'Bytes' -Value @(1, 2, 255)
            $b -is [byte[]] | Should -BeTrue
            $b.Count | Should -Be 3
            ConvertTo-SchemaTypedValue -SchemaType 'Bytes' -Value @(1, 'zz') | Should -BeNullOrEmpty -Because 'a string where the target expects octets is a failed cast'
            ConvertTo-SchemaTypedValue -SchemaType $null -Value 'x'         | Should -BeNullOrEmpty -Because 'an attribute the target schema cannot type is not consumed'
            ConvertTo-SchemaTypedValue -SchemaType 'Other' -Value 'x'       | Should -BeNullOrEmpty
        }

        It 'Get-WellKnownTokenSid resolves the domain-relative, universal and forest-root tokens from the domain objects alone; an unknown token is $null' {
            $dom = [pscustomobject]@{ DomainSID = [pscustomobject]@{ Value = 'S-1-5-21-1-2-3' }; NetBIOSName = 'ARONS'; DNSRoot = 'arons.local' }
            $for = [pscustomobject]@{ RootDomain = 'arons.local' }
            $script:RootDomainSidCache = $null   # the script's per-run cache, filled by the first enterprise token below
            $sid = { param($n) (Get-WellKnownTokenSid -Norm $n -TargetDomain $dom -TargetForest $for -ADParams @{}).Value }
            & $sid 'everyone'                    | Should -Be 'S-1-1-0'
            & $sid 'authenticatedusers'          | Should -Be 'S-1-5-11'
            & $sid 'enterprisedomaincontrollers' | Should -Be 'S-1-5-9'
            & $sid 'domainusers'                 | Should -Be 'S-1-5-21-1-2-3-513'
            & $sid 'domaincomputers'             | Should -Be 'S-1-5-21-1-2-3-515'
            & $sid 'domainadmins'                | Should -Be 'S-1-5-21-1-2-3-512'
            & $sid 'domaincontrollers'           | Should -Be 'S-1-5-21-1-2-3-516'
            # the target domain IS the forest root: the root SID needs no directory read
            & $sid 'enterpriseadmins'            | Should -Be 'S-1-5-21-1-2-3-519'
            & $sid 'enterpriserodcs'             | Should -Be 'S-1-5-21-1-2-3-498'
            Get-WellKnownTokenSid -Norm 'nosuchtoken' -TargetDomain $dom -TargetForest $for -ADParams @{} | Should -BeNullOrEmpty
            $script:RootDomainSidCache = $null
        }

        It 'Export-Template -StripOid -StripIdentity writes a BOM-marked JSON without the OID, name and displayName; without the switches they stay' {
            $view = [pscustomobject]@{ name = 'Src'; displayName = 'Source'; objectClass = 'pKICertificateTemplate'; flags = 1; revision = 100
                                       'msPKI-Cert-Template-OID' = '1.2.3.4'; 'msPKI-Template-Schema-Version' = 2 }
            $path = Join-Path $TestDrive 'stripped.json'
            $written = script:Invoke-WithCmdlet -Body { param($c) Export-Template -InputObject $view -Path $path -StripOid -StripIdentity -CallerCmdlet $c 6>$null }
            $written | Should -BeTrue
            $bytes = [System.IO.File]::ReadAllBytes($path)
            ($bytes[0..2] -join ',') | Should -Be '239,187,191'
            $j = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
            $j.PSObject.Properties.Name | Should -Not -Contain 'msPKI-Cert-Template-OID'
            $j.PSObject.Properties.Name | Should -Not -Contain 'name'
            $j.PSObject.Properties.Name | Should -Not -Contain 'displayName'
            $j.flags | Should -Be 1
            $j.'msPKI-Template-Schema-Version' | Should -Be 2
            $view2 = [pscustomobject]@{ name = 'Src'; displayName = 'Source'; flags = 1; 'msPKI-Cert-Template-OID' = '1.2.3.4' }
            $path2 = Join-Path $TestDrive 'full.json'
            $null = script:Invoke-WithCmdlet -Body { param($c) Export-Template -InputObject $view2 -Path $path2 -CallerCmdlet $c 6>$null }
            $j2 = [System.IO.File]::ReadAllText($path2, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
            $j2.name | Should -Be 'Src'
            $j2.displayName | Should -Be 'Source'
            $j2.'msPKI-Cert-Template-OID' | Should -Be '1.2.3.4'
        }

        It 'Export-Template under -WhatIf writes no file and returns $false' {
            $path = Join-Path $TestDrive 'whatif.json'
            $r = script:Invoke-WithCmdlet -WhatIf -Body { param($c) Export-Template -InputObject ([pscustomobject]@{ name = 'Src'; flags = 1 }) -Path $path -CallerCmdlet $c 6>$null }
            $r | Should -BeFalse
            Test-Path -LiteralPath $path | Should -BeFalse
        }

        It 'New-LabExclusiveDirectory reports Created only for the call that made the folder, AlreadyExists after that' {
            $path = Join-Path $TestDrive 'exclusive-dir'
            Test-Path -LiteralPath $path | Should -BeFalse
            script:New-LabExclusiveDirectory -Path $path | Should -Be 'Created'
            Test-Path -LiteralPath $path -PathType Container | Should -BeTrue
            script:New-LabExclusiveDirectory -Path $path | Should -Be 'AlreadyExists'
            # A file at the name is "taken" as well - never reported as created.
            $file = Join-Path $TestDrive 'exclusive-file'
            Set-Content -LiteralPath $file -Value 'x'
            script:New-LabExclusiveDirectory -Path $file | Should -Be 'AlreadyExists'
            # Any other Win32 failure (here: a missing parent) throws instead of returning a verdict.
            { script:New-LabExclusiveDirectory -Path (Join-Path $TestDrive 'no-such-parent\leaf') } | Should -Throw '*Win32 error 3*'
        }

        It 'Remove-LabScratchFolder removes only the tracked files and leaves an untracked entry, a tracked path that is now a folder, and the folder itself' {
            $dir = Join-Path $TestDrive 'scratch-mixed'
            $null = New-Item -ItemType Directory -Path $dir
            $tracked   = Join-Path $dir 'tracked.json';   Set-Content -LiteralPath $tracked -Value 'a'
            $unwritten = Join-Path $dir 'unwritten.json'                                   # tracked, never written
            $stranger  = Join-Path $dir 'stranger.txt';   Set-Content -LiteralPath $stranger -Value 'b'
            $sub       = Join-Path $dir 'sub';            $null = New-Item -ItemType Directory -Path $sub
            $asFolder  = Join-Path $dir 'now-a-folder';   $null = New-Item -ItemType Directory -Path $asFolder
            $outside   = Join-Path $TestDrive 'outside.txt'; Set-Content -LiteralPath $outside -Value 'c'
            $w = script:Remove-LabScratchFolder -Path $dir -TrackedFiles @($tracked, $unwritten, $asFolder, $outside) 3>&1
            Test-Path -LiteralPath $tracked  | Should -BeFalse
            Test-Path -LiteralPath $stranger | Should -BeTrue  -Because 'an untracked file is never deleted'
            Test-Path -LiteralPath $sub      | Should -BeTrue  -Because 'an untracked folder is never deleted'
            Test-Path -LiteralPath $asFolder | Should -BeTrue  -Because 'a tracked path that is a folder now is never deleted'
            Test-Path -LiteralPath $outside  | Should -BeTrue  -Because 'a tracked path outside the scratch folder is never deleted'
            Test-Path -LiteralPath $dir      | Should -BeTrue  -Because 'the folder stays while anything untracked remains'
            $text = @($w | ForEach-Object { "$_" }) -join "`n"
            $text | Should -Match 'NOT removed'
            $text | Should -Match ([regex]::Escape($stranger))
            $text | Should -Match ([regex]::Escape($sub))
            $text | Should -Match ([regex]::Escape($outside))
        }

        It 'Remove-LabScratchFolder removes the folder non-recursively once only tracked files were inside' {
            $dir = Join-Path $TestDrive 'scratch-owned'
            $null = New-Item -ItemType Directory -Path $dir
            $a = Join-Path $dir 'a.json'; Set-Content -LiteralPath $a -Value 'a'
            $b = Join-Path $dir 'b.json'; Set-Content -LiteralPath $b -Value 'b'
            $w = script:Remove-LabScratchFolder -Path $dir -TrackedFiles @($a, $b, (Join-Path $dir 'never-written.json')) 3>&1
            @($w).Count | Should -Be 0
            Test-Path -LiteralPath $dir | Should -BeFalse
            # An absent folder is not an error either (BeforeAll may have died before the create).
            $w2 = script:Remove-LabScratchFolder -Path (Join-Path $TestDrive 'absent') -TrackedFiles @() 3>&1
            @($w2).Count | Should -Be 0
        }

        It 'Test-LabExportEvidence accepts only the "Export completed:" line the script prints for the path plus matching content' {
            # The evidence line comes from the script's own Export-Template, so the parser and the
            # producer cannot drift apart. Out-String -Width 4096 renders it exactly as the Lab tier reads it.
            $view = [pscustomobject]@{ name = 'Src'; displayName = 'Source'; flags = 1; 'msPKI-Cert-Template-OID' = '1.2.3.4'; 'msPKI-Template-Schema-Version' = 2 }
            $path = Join-Path $TestDrive 'evidence.json'
            $out = script:Invoke-WithCmdlet -Body { param($c) Export-Template -InputObject $view -Path $path -CallerCmdlet $c } *>&1 | Out-String -Width 4096
            $out | Should -Match ([regex]::Escape("Export completed: $path"))
            script:Test-LabExportEvidence -Path $path -Output $out -Expect @{ name = 'Src' } | Should -BeNullOrEmpty
            script:Test-LabExportEvidence -Path $path -Output $out -Expect @{ 'msPKI-Cert-Template-OID' = '1.2.3.4'; 'msPKI-Template-Schema-Version' = 2 } | Should -BeNullOrEmpty
            # The file exists, but the output names another path, no path, or the content is not this run's: refused with a reason.
            script:Test-LabExportEvidence -Path $path -Output "Export completed: $(Join-Path $TestDrive 'other.json')`r`n" -Expect @{ name = 'Src' } | Should -Match 'no .Export completed:. line'
            script:Test-LabExportEvidence -Path $path -Output '' -Expect @{ name = 'Src' } | Should -Match 'no .Export completed:. line'
            script:Test-LabExportEvidence -Path $path -Output $out -Expect @{ name = 'Other' } | Should -Match "'name' = 'Src', expected 'Other'"
            script:Test-LabExportEvidence -Path $path -Output $out -Expect @{ displayName = 'Source'; revision = 100 } | Should -Match "no 'revision' property"
            # A file another process wrote at the named path is refused: not JSON, or absent.
            $foreign = Join-Path $TestDrive 'foreign.json'
            Set-Content -LiteralPath $foreign -Value 'not json'
            script:Test-LabExportEvidence -Path $foreign -Output "Export completed: $foreign`r`n" -Expect @{ name = 'Src' } | Should -Match 'is not JSON'
            $gone = Join-Path $TestDrive 'gone.json'
            script:Test-LabExportEvidence -Path $gone -Output "Export completed: $gone`r`n" -Expect @{ name = 'Src' } | Should -Match 'no file exists'
        }
    }

    Context 'Unit: Resolve-PrincipalSid against a stubbed directory' -Tag 'Unit' {

        BeforeAll {
            # A canned directory keyed by the exact LDAP filter Resolve-PrincipalSid builds. The stub
            # shadows the AD cmdlet for this Context only (a function outranks a cmdlet of the same
            # name, and the Context scope ends with the Context), so no AD module is needed.
            $script:FakeDir   = @{}
            $script:FakeCalls = New-Object System.Collections.Generic.List[string]
            function Get-ADObject {
                [CmdletBinding()]
                param([string]$LDAPFilter, [string]$Identity, [string]$SearchBase, [string[]]$Properties, [string]$Server, [int]$ResultSetSize, [pscredential]$Credential)
                $script:FakeCalls.Add($LDAPFilter)
                if ($script:FakeDir.ContainsKey($LDAPFilter)) { return $script:FakeDir[$LDAPFilter] }
                return   # a real cmdlet emits nothing on a miss - never an explicit $null
            }
            function script:New-FakePrincipal {
                param([string]$Dn, [string]$Sid)
                [pscustomobject]@{ DistinguishedName = $Dn; objectSid = [System.Security.Principal.SecurityIdentifier]$Sid }
            }
            $script:FakeDom = [pscustomobject]@{ DomainSID = [pscustomobject]@{ Value = 'S-1-5-21-1-2-3' }; NetBIOSName = 'ARONS'; DNSRoot = 'arons.local' }
            $script:FakeFor = [pscustomobject]@{ RootDomain = 'arons.local' }
            function script:Resolve-Fake {
                param([string]$Id)
                Resolve-PrincipalSid -Identity $Id -TargetDomain $script:FakeDom -TargetForest $script:FakeFor -ADParams @{}
            }
        }

        BeforeEach {
            $script:FakeDir.Clear()
            $script:FakeCalls.Clear()
            $script:RootDomainSidCache = $null
        }

        It 'a raw SID string resolves without a directory read' {
            (script:Resolve-Fake 'S-1-5-21-1-2-3-1105').Value | Should -Be 'S-1-5-21-1-2-3-1105'
            $script:FakeCalls.Count | Should -Be 0
        }

        It 'a DOMAIN\name with a foreign prefix is refused before any directory read' {
            { script:Resolve-Fake 'OTHER\bob' } | Should -Throw -ExpectedMessage "*names domain 'OTHER'*Use a SID*"
            $script:FakeCalls.Count | Should -Be 0
        }

        It 'a DOMAIN\name with the target prefix resolves by sAMAccountName' {
            $script:FakeDir['(sAMAccountName=bob)'] = script:New-FakePrincipal 'CN=bob,CN=Users,DC=arons,DC=local' 'S-1-5-21-1-2-3-1105'
            (script:Resolve-Fake 'ARONS\bob').Value | Should -Be 'S-1-5-21-1-2-3-1105'
            $script:FakeCalls | Should -Contain '(sAMAccountName=bob)'
        }

        It 'a UPN is looked up as a UPN and checked for a sAMAccountName shadow; the same object as shadow is accepted' {
            $bob = script:New-FakePrincipal 'CN=bob,CN=Users,DC=arons,DC=local' 'S-1-5-21-1-2-3-1105'
            $script:FakeDir['(userPrincipalName=bob@arons.local)'] = $bob
            (script:Resolve-Fake 'bob@arons.local').Value | Should -Be 'S-1-5-21-1-2-3-1105'
            $script:FakeCalls | Should -Contain '(userPrincipalName=bob@arons.local)'
            $script:FakeCalls | Should -Contain '(sAMAccountName=bob@arons.local)' -Because 'the shadow check must always run for a UPN'
            $script:FakeDir['(sAMAccountName=bob@arons.local)'] = $bob   # the SAME object: same SID, no ambiguity
            (script:Resolve-Fake 'bob@arons.local').Value | Should -Be 'S-1-5-21-1-2-3-1105'
        }

        It 'a UPN that a DIFFERENT object carries as its sAMAccountName is refused, also in the DOMAIN\user@domain form' {
            $script:FakeDir['(userPrincipalName=bob@arons.local)'] = script:New-FakePrincipal 'CN=bob,CN=Users,DC=arons,DC=local' 'S-1-5-21-1-2-3-1105'
            $script:FakeDir['(sAMAccountName=bob@arons.local)']    = script:New-FakePrincipal 'CN=evil,CN=Users,DC=arons,DC=local' 'S-1-5-21-1-2-3-2222'
            { script:Resolve-Fake 'bob@arons.local' } | Should -Throw -ExpectedMessage '*different directory object carries it as its sAMAccountName*CN=evil*'
            { script:Resolve-Fake 'ARONS\bob@arons.local' } | Should -Throw -ExpectedMessage '*different directory object carries it as its sAMAccountName*' -Because 'the prefixed form takes the UPN-only path too (the 1.0.5 fix)'
        }

        It 'a UPN with no UPN match but a sAMAccountName match is refused rather than resolved to the shadow' {
            $script:FakeDir['(sAMAccountName=bob@arons.local)'] = script:New-FakePrincipal 'CN=evil,CN=Users,DC=arons,DC=local' 'S-1-5-21-1-2-3-2222'
            { script:Resolve-Fake 'bob@arons.local' } | Should -Throw -ExpectedMessage '*different directory object carries it as its sAMAccountName*'
        }

        It 'a name that matches more than one object is refused as ambiguous' {
            $script:FakeDir['(sAMAccountName=dup)'] = @(
                (script:New-FakePrincipal 'CN=dup1,CN=Users,DC=arons,DC=local' 'S-1-5-21-1-2-3-3001'),
                (script:New-FakePrincipal 'CN=dup2,CN=Users,DC=arons,DC=local' 'S-1-5-21-1-2-3-3002')
            )
            { script:Resolve-Fake 'dup' } | Should -Throw -ExpectedMessage '*ambiguous: 2 objects match*CN=dup1*CN=dup2*'
        }

        It 'a well-known token that also matches a directory object with a DIFFERENT SID is refused; the same SID is accepted' {
            $script:FakeDir['(sAMAccountName=Domain Admins)'] = script:New-FakePrincipal 'CN=Fake Admins,CN=Users,DC=arons,DC=local' 'S-1-5-21-1-2-3-9999'
            { script:Resolve-Fake 'Domain Admins' } | Should -Throw -ExpectedMessage '*matches BOTH the well-known token and a DIFFERENT directory object*'
            $script:FakeDir['(sAMAccountName=Domain Admins)'] = script:New-FakePrincipal 'CN=Domain Admins,CN=Users,DC=arons,DC=local' 'S-1-5-21-1-2-3-512'
            (script:Resolve-Fake 'Domain Admins').Value | Should -Be 'S-1-5-21-1-2-3-512'
        }

        It 'a well-known token with no directory object resolves to the token SID; an unknown name fails closed after the directory was consulted' {
            (script:Resolve-Fake 'Everyone').Value                      | Should -Be 'S-1-1-0'
            (script:Resolve-Fake 'Domain Computers').Value              | Should -Be 'S-1-5-21-1-2-3-515'
            (script:Resolve-Fake 'enterprise-domain-controllers').Value | Should -Be 'S-1-5-9'
            { script:Resolve-Fake 'nobody' } | Should -Throw -ExpectedMessage '*Could not resolve principal*by sAMAccountName or UPN*'
            $script:FakeCalls | Should -Contain '(sAMAccountName=nobody)'
        }
    }

    Context 'Unit: Import-Template against a stubbed directory' -Tag 'Unit' {

        BeforeAll {
            # The directory as Import-Template sees it: the Certificate Templates and OID containers
            # exist, no template or display object does, New-ADObject records what it was asked to
            # create, Remove-ADObject records what it was asked to remove. Set Carrier to make a
            # second template appear with the new OID right after the create (the concurrent-claim
            # case); set RemoveFails to make the template rollback fail.
            $script:Imp = @{ Creates = $null; Removes = $null; TemplateCreated = $false; TemplateGuid = [guid]::NewGuid(); CompanionGuid = [guid]::NewGuid(); Carrier = $false; RemoveFails = $false; SchemaTypes = @{} }
            function Get-ADObjectIfPresent {
                param([Parameter(Mandatory)][string]$Identity, [hashtable]$ADParams = @{}, [string[]]$Properties)
                if ($Identity -like 'CN=Certificate Templates,*' -or $Identity -like 'CN=OID,*') { return [pscustomobject]@{ DistinguishedName = $Identity } }
                return   # nothing on a miss - never an explicit $null
            }
            function Get-ADObject {
                [CmdletBinding()]
                param([string]$LDAPFilter, [string]$Identity, [string]$SearchBase, [string[]]$Properties, [string]$Server, [int]$ResultSetSize, [pscredential]$Credential)
                if ($script:Imp.Carrier -and $script:Imp.TemplateCreated -and $SearchBase -like 'CN=Certificate Templates,*' -and $LDAPFilter -like '(msPKI-Cert-Template-OID=*') {
                    return [pscustomobject]@{ DistinguishedName = "CN=Intruder,$SearchBase"; ObjectGUID = [guid]::NewGuid() }
                }
                return   # a real cmdlet emits nothing on a miss - an explicit $null would count as one "carrier"
            }
            function New-ADObject {
                [CmdletBinding(SupportsShouldProcess)]
                param([string]$Path, [string]$Name, [string]$DisplayName, [string]$Type, [hashtable]$OtherAttributes, [switch]$PassThru, [string]$Server, [pscredential]$Credential)
                $script:Imp.Creates.Add(@{ Path = $Path; Name = $Name; DisplayName = $DisplayName; Type = $Type; OtherAttributes = $OtherAttributes })
                if ($Type -eq 'pKICertificateTemplate') {
                    $script:Imp.TemplateCreated = $true
                    if ($PassThru) { return [pscustomobject]@{ DistinguishedName = "CN=$Name,$Path"; ObjectGUID = $script:Imp.TemplateGuid } }
                }
                elseif ($Type -eq 'msPKI-Enterprise-Oid' -and $PassThru) {
                    return [pscustomobject]@{ DistinguishedName = "CN=$Name,$Path"; ObjectGUID = $script:Imp.CompanionGuid }
                }
            }
            function Remove-ADObject {
                [CmdletBinding(SupportsShouldProcess)]
                param($Identity, [string]$Server, [pscredential]$Credential)
                $script:Imp.Removes.Add("$Identity")
                if ($script:Imp.RemoveFails) { throw 'access denied (stub)' }
            }
            function Get-SchemaAttributeType {
                param([string]$AttributeName, [string]$ConfigNC, [hashtable]$ADParams)
                $script:Imp.SchemaTypes[$AttributeName]
            }
            function script:Invoke-FakeImport {
                # Output: the returned DN plus every WarningRecord (3>&1); the Write-Host narration is dropped.
                param([psobject]$View)
                script:Invoke-WithCmdlet -Body { param($c) Import-Template -InputObject $View -ConfigNC 'CN=Configuration,DC=x,DC=test' -ADParams @{} -CallerCmdlet $c 6>$null 3>&1 }
            }
        }

        BeforeEach {
            $script:Imp.Creates = New-Object System.Collections.Generic.List[object]
            $script:Imp.Removes = New-Object System.Collections.Generic.List[string]
            $script:Imp.TemplateCreated = $false
            $script:Imp.Carrier = $false
            $script:Imp.RemoveFails = $false
            $script:Imp.SchemaTypes = @{}
        }

        It 'types an unknown PKI attribute from the target schema: Int/String need one element, Bytes cast, an untyped or multi-element single-valued one is dropped with a warning' {
            $script:Imp.SchemaTypes = @{ 'msPKI-Foo' = 'Int'; 'msPKI-Bar' = 'Int'; 'msPKI-Baz' = 'String'; 'msPKI-Qux' = 'Bytes'; 'msPKI-Unknown' = $null }
            $view = [pscustomobject]@{ name = 'Src'; displayName = 'Src'; flags = 1; revision = 100; 'msPKI-Cert-Template-OID' = '1.2.3.4'
                                       'msPKI-Foo' = 7; 'msPKI-Bar' = @(1, 2); 'msPKI-Baz' = 'text'; 'msPKI-Qux' = @(1, 2, 3); 'msPKI-Unknown' = 'x' }
            $out = @(script:Invoke-FakeImport -View $view)
            $dn = @($out | Where-Object { $_ -is [string] })
            $dn.Count | Should -Be 1
            $dn[0] | Should -Be 'CN=Src,CN=Certificate Templates,CN=Public Key Services,CN=Services,CN=Configuration,DC=x,DC=test'
            $tpl = @($script:Imp.Creates | Where-Object { $_.Type -eq 'pKICertificateTemplate' })
            $tpl.Count | Should -Be 1
            $oa = $tpl[0].OtherAttributes
            $oa['msPKI-Foo'] | Should -Be 7
            $oa['msPKI-Foo'] -is [int] | Should -BeTrue
            $oa['msPKI-Baz'] | Should -BeExactly 'text'
            $oa['msPKI-Baz'] -is [string] | Should -BeTrue
            $oa['msPKI-Qux'] -is [byte[]] | Should -BeTrue
            $oa['msPKI-Qux'].Count | Should -Be 3
            $oa.ContainsKey('msPKI-Bar')     | Should -BeFalse -Because 'two elements cannot go into a single-valued Int attribute'
            $oa.ContainsKey('msPKI-Unknown') | Should -BeFalse -Because 'an attribute the target schema cannot type is not copied'
            $oa['msPKI-Cert-Template-OID'] | Should -Be '1.2.3.4'
            $oa['flags'] | Should -Be 1
            $warn = @($out | Where-Object { $_ -is [System.Management.Automation.WarningRecord] } | ForEach-Object { $_.Message })
            @($warn | Where-Object { $_ -match 'will NOT be written' -and $_ -match 'msPKI-Bar' -and $_ -match 'msPKI-Unknown' }).Count | Should -Be 1
            $script:Imp.Removes.Count | Should -Be 0
        }

        It 'prints a "Created OID object:" and a "Created template:" line carrying the DN and objectGUID New-ADObject returned at creation, and the Lab parser reads them back' {
            $view = [pscustomobject]@{ name = 'Src'; displayName = 'Src'; flags = 1; 'msPKI-Cert-Template-OID' = '1.2.3.4' }
            # 6>&1: the Write-Host narration is what this test is about; warnings are dropped.
            $lines = @(script:Invoke-WithCmdlet -Body { param($c) Import-Template -InputObject $view -ConfigNC 'CN=Configuration,DC=x,DC=test' -ADParams @{} -CallerCmdlet $c 6>&1 3>$null } |
                Where-Object { $_ -is [System.Management.Automation.InformationRecord] } | ForEach-Object { "$_" })
            $compCreate = @($script:Imp.Creates | Where-Object { $_.Type -eq 'msPKI-Enterprise-Oid' })
            $compCreate.Count | Should -Be 1 -Because 'Preserve registers a display object for an OID the target does not know'
            $compDN = "CN=$($compCreate[0].Name),$($compCreate[0].Path)"
            $tplDN  = 'CN=Src,CN=Certificate Templates,CN=Public Key Services,CN=Services,CN=Configuration,DC=x,DC=test'
            $lines | Should -Contain "Created OID object: $compDN (objectGUID $($script:Imp.CompanionGuid))"
            $lines | Should -Contain "Created template: $tplDN (objectGUID $($script:Imp.TemplateGuid))"
            # the Lab tracker parses exactly these lines (CRLF-joined, as Out-String renders them)
            $text = ($lines -join "`r`n") + "`r`n"
            $tpl = script:Read-LabCreatedLine -Output $text -Prefix 'Created template:'
            $tpl.DN   | Should -BeExactly $tplDN
            $tpl.Guid | Should -Be $script:Imp.TemplateGuid
            $comp = script:Read-LabCreatedLine -Output $text -Prefix 'Created OID object:'
            $comp.DN   | Should -BeExactly $compDN
            $comp.Guid | Should -Be $script:Imp.CompanionGuid
            # a DN with parentheses, and a line without a GUID, still parse as intended
            $odd = "Created template: CN=A (b),CN=Certificate Templates,DC=x (objectGUID $($script:Imp.TemplateGuid))`r`nCreated OID object: CN=1.2,CN=OID,DC=x`r`n"
            (script:Read-LabCreatedLine -Output $odd -Prefix 'Created template:').DN | Should -BeExactly 'CN=A (b),CN=Certificate Templates,DC=x'
            $noGuid = script:Read-LabCreatedLine -Output $odd -Prefix 'Created OID object:'
            $noGuid.DN   | Should -BeExactly 'CN=1.2,CN=OID,DC=x'
            $noGuid.Guid | Should -BeNullOrEmpty -Because 'a line without an objectGUID must not be mistaken for one with it'
            script:Read-LabCreatedLine -Output 'nothing here' -Prefix 'Created template:' | Should -BeNullOrEmpty
        }

        It 'rolls back the new template AND its companion OID object when another template claimed the OID concurrently, then rethrows' {
            $script:Imp.Carrier = $true
            $view = [pscustomobject]@{ name = 'Src'; displayName = 'Src'; flags = 1; 'msPKI-Cert-Template-OID' = '1.2.3.4' }
            $warn = New-Object System.Collections.Generic.List[string]
            $err = $null
            try { script:Invoke-FakeImport -View $view | ForEach-Object { if ($_ -is [System.Management.Automation.WarningRecord]) { $warn.Add($_.Message) } } }
            catch { $err = $_ }
            $err | Should -Not -BeNullOrEmpty -Because 'the failure must propagate after the rollback'
            $err.Exception.Message | Should -BeLike '*claimed concurrently by another template*CN=Intruder*'
            $script:Imp.Creates.Count | Should -Be 2
            $script:Imp.Creates[0].Type | Should -Be 'msPKI-Enterprise-Oid' -Because 'Preserve registers a display object for an OID the target does not know'
            $script:Imp.Creates[1].Type | Should -Be 'pKICertificateTemplate'
            $script:Imp.Removes.Count | Should -Be 2
            $script:Imp.Removes[0] | Should -Be $script:Imp.TemplateGuid.ToString() -Because 'the template is removed first, by the GUID captured at creation'
            $script:Imp.Removes[1] | Should -Be "CN=$($script:Imp.Creates[0].Name),$($script:Imp.Creates[0].Path)" -Because 'the companion is removed once the template is gone'
            @($warn | Where-Object { $_ -match 'rolled back the new template' }).Count | Should -Be 1
            @($warn | Where-Object { $_ -match 'Rolled back the companion OID object' }).Count | Should -Be 1
        }

        It 'keeps the companion OID object when the new template cannot be removed, and names both for manual cleanup' {
            $script:Imp.Carrier = $true
            $script:Imp.RemoveFails = $true
            $view = [pscustomobject]@{ name = 'Src'; displayName = 'Src'; flags = 1; 'msPKI-Cert-Template-OID' = '1.2.3.4' }
            $warn = New-Object System.Collections.Generic.List[string]
            $err = $null
            try { script:Invoke-FakeImport -View $view | ForEach-Object { if ($_ -is [System.Management.Automation.WarningRecord]) { $warn.Add($_.Message) } } }
            catch { $err = $_ }
            $err.Exception.Message | Should -BeLike '*claimed concurrently*'
            $script:Imp.Removes.Count | Should -Be 1 -Because 'the companion removal is not attempted while the template survives'
            @($warn | Where-Object { $_ -match 'could not be removed' -and $_ -match 'clean up manually' -and $_ -match 'CN=Src,' }).Count | Should -Be 1
            @($warn | Where-Object { $_ -match 'companion OID object was KEPT' }).Count | Should -Be 1
        }

        # The MultiString arm is the one conversion Import-Template does itself (the cast to the AD
        # collection type after ConvertTo-SchemaTypedValue), so it needs the ActiveDirectory module
        # for that type. The stubs above stay in force: a function outranks the module's cmdlet of
        # the same name. Discovered only where the module is available - the Unit tier stays
        # module-free elsewhere, and nothing is ever skipped.
        if ($script:HasAD) {
            Context 'MultiString arm (ActiveDirectory module present)' {

                BeforeAll { Import-Module ActiveDirectory -ErrorAction Stop }

                It 'copies a MultiString attribute as an ADPropertyValueCollection of strings, one element per source value, a scalar as one element' {
                    $script:Imp.SchemaTypes = @{ 'msPKI-Multi' = 'MultiString'; 'msPKI-Solo' = 'MultiString'; 'msPKI-Foo' = 'Int' }
                    $view = [pscustomobject]@{ name = 'Src'; displayName = 'Src'; flags = 1; revision = 100; 'msPKI-Cert-Template-OID' = '1.2.3.4'
                                               'msPKI-Multi' = @('a', 7, 'c'); 'msPKI-Solo' = 'solo'; 'msPKI-Foo' = 3 }
                    $out = @(script:Invoke-FakeImport -View $view)
                    @($out | Where-Object { $_ -is [string] }).Count | Should -Be 1 -Because 'the import must complete and return the DN'
                    $tpl = @($script:Imp.Creates | Where-Object { $_.Type -eq 'pKICertificateTemplate' })
                    $tpl.Count | Should -Be 1
                    $oa = $tpl[0].OtherAttributes
                    $oa.ContainsKey('msPKI-Multi') | Should -BeTrue
                    $oa['msPKI-Multi'] -is [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection] | Should -BeTrue -Because 'New-ADObject must receive the AD collection type, not an object[]'
                    $oa['msPKI-Multi'].Count | Should -Be 3 -Because 'the cast must spread the elements, not wrap the array as one value'
                    $multi = @($oa['msPKI-Multi'])
                    $multi[0] | Should -BeExactly 'a'
                    $multi[1] | Should -BeExactly '7'
                    $multi[1] -is [string] | Should -BeTrue -Because 'every element is cast to string before the collection is built'
                    $multi[2] | Should -BeExactly 'c'
                    $oa['msPKI-Solo'] -is [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection] | Should -BeTrue
                    $oa['msPKI-Solo'].Count | Should -Be 1
                    @($oa['msPKI-Solo'])[0] | Should -BeExactly 'solo'
                    $oa['msPKI-Foo'] | Should -Be 3
                    $warn = @($out | Where-Object { $_ -is [System.Management.Automation.WarningRecord] } | ForEach-Object { $_.Message })
                    @($warn | Where-Object { $_ -match 'will NOT be written' }).Count | Should -Be 0 -Because 'a typed MultiString attribute is consumed, not dropped'
                    $script:Imp.Removes.Count | Should -Be 0
                }
            }
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

    # A discovery condition, not -Skip: CI fails a run that reports a skipped test, so without the
    # ActiveDirectory module the Guard Context must not exist at all.
    if ($script:HasAD) {
    Context 'Guard: parameter and mode validation' -Tag 'Guard' {

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

        It 'rejects -OidHandling GenerateFromRoot without -OidRoot before any DC is contacted (the file need not even exist)' {
            { & $script:Sync -Mode Import -Path x.json -OidHandling GenerateFromRoot } |
                Should -Throw -ExpectedMessage "*'GenerateFromRoot' requires -OidRoot*"
            { & $script:Sync -Mode Sync -SourceServer dc1 -OidHandling GenerateFromRoot } |
                Should -Throw -ExpectedMessage "*'GenerateFromRoot' requires -OidRoot*"
        }

        It 'rejects -SourceServer outside -Mode Sync' {
            { & $script:Sync -Mode Export -Path x.json -SourceServer dc1 } |
                Should -Throw -ExpectedMessage '*not applicable to -Mode Export*'
        }

        It 'rejects -UpgradeCompatibility for -Mode Export (Import/Sync only)' {
            { & $script:Sync -Mode Export -Path x.json -UpgradeCompatibility } |
                Should -Throw -ExpectedMessage '*not applicable to -Mode Export*'
        }

        It 'rejects -AllowLinkedIssuancePolicy for -Mode Export and -Mode Validate (Import/Sync only)' {
            { & $script:Sync -Mode Export -Path x.json -TemplateName K -AllowLinkedIssuancePolicy } |
                Should -Throw -ExpectedMessage '*AllowLinkedIssuancePolicy*'
            { & $script:Sync -Mode Validate -TemplateName K -AllowLinkedIssuancePolicy } |
                Should -Throw -ExpectedMessage '*AllowLinkedIssuancePolicy*'
        }

        It 'rejects -UpgradeCompatibility for -Mode Validate' {
            { & $script:Sync -Mode Validate -TemplateName K -UpgradeCompatibility } |
                Should -Throw -ExpectedMessage '*not applicable to -Mode Validate*'
        }
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
            # Scratch folder: EXCLUSIVE creation only. A name is not proof of ownership, so the
            # folder must be absent first (throw if not, and never adopt it). The pre-check only
            # gives a clear message: New-Item would also succeed on a folder another process
            # created between the check and the create, so the create goes through
            # CreateDirectoryW (New-LabExclusiveDirectory), which reports whether THIS call made
            # the folder. $script:TmpDirCreated is set ONLY from that 'Created' verdict.
            # Owning the folder does not own its later contents: another process can add a file
            # or a subfolder. So every file the suite puts inside is recorded BY EXACT PATH in
            # $script:TmpFiles - only AFTER the write, once Confirm-LabScratchFile has proved the
            # file exists - and AfterAll removes exactly those files one at a time, then the
            # folder non-recursively - never Remove-Item -Recurse.
            $script:TmpDir        = Join-Path $env:TEMP $script:Prefix
            $script:TmpDirCreated = $false
            $script:TmpFiles      = New-Object System.Collections.Generic.List[string]
            if (Test-Path -LiteralPath $script:TmpDir) {
                throw "Lab fixture refused: '$script:TmpDir' already exists. It is not this run's folder, so it will not be reused or removed."
            }
            $verdict = script:New-LabExclusiveDirectory -Path $script:TmpDir
            if ($verdict -ne 'Created') {
                throw "Lab fixture refused: '$script:TmpDir' appeared between the existence check and the create ($verdict). It is not this run's folder, so it will not be reused or removed."
            }
            $script:TmpDirCreated = $true
            $script:AP       = @{ Server = $AronsServer }
            $script:ConfigNC = (Get-ADRootDSE @script:AP).configurationNamingContext
            $script:TplBase  = "CN=Certificate Templates,CN=Public Key Services,CN=Services,$script:ConfigNC"
            $script:OidBase  = "CN=OID,CN=Public Key Services,CN=Services,$script:ConfigNC"
            $script:ForestRoot = (Get-ADObject @script:AP -Identity $script:OidBase -Properties 'msPKI-Cert-Template-OID').'msPKI-Cert-Template-OID'
            $script:Created  = New-Object System.Collections.Generic.List[object]  # @{ Server; DN; Guid; DisplayName; Oid } - Guid is the objectGUID; AfterEach clears it
            $script:Ledger   = New-Object System.Collections.Generic.List[object]  # every record ever added to Created, for the whole run - never cleared; AfterAll reads it
            $script:Removed  = New-Object System.Collections.Generic.List[string]  # lowercased DN of every tracked object this run removed

            # The ONE way to name a file in the scratch folder. The suite names every file it puts
            # there - also the ones the script writes, because Export takes the path from -Path.
            # A name is a RESERVATION only: an intended write is not evidence of a write, so
            # nothing is tracked here. The caller proves the write afterwards with
            # Confirm-LabScratchFile, which is what records the path. AfterAll removes exactly
            # the recorded paths and nothing else.
            function script:New-LabScratchPath {
                param([Parameter(Mandatory)][string]$Name)
                if ($Name -ne [System.IO.Path]::GetFileName($Name)) { throw "New-LabScratchPath: '$Name' must be a bare file name." }
                return (Join-Path $script:TmpDir $Name)
            }

            # Called right after every write (the suite's own, or the script's through -Path):
            # asserts the file now exists as a leaf directly under the scratch folder and records
            # its exact path. Only a path that passed this check is ever removed by AfterAll. The
            # suite's own writes call this directly. A file the SCRIPT writes reaches this only
            # through Invoke-LabExport, after Test-LabExportEvidence attributed the write to this
            # run: existence alone is not evidence (the script can fail before it writes while
            # another process creates the name).
            function script:Confirm-LabScratchFile {
                param([Parameter(Mandatory)][string]$Path)
                $parent = [System.IO.Path]::GetDirectoryName($Path)
                if (-not [string]::Equals($parent, $script:TmpDir, [System.StringComparison]::OrdinalIgnoreCase)) {
                    throw "Confirm-LabScratchFile: '$Path' is not directly under the scratch folder '$script:TmpDir'."
                }
                if (-not [System.IO.File]::Exists($Path)) {
                    throw "Confirm-LabScratchFile: expected a file at '$Path' after the write, but none exists."
                }
                if (-not $script:TmpFiles.Contains($Path)) { $script:TmpFiles.Add($Path) }
            }

            # The counterpart for a call that must write nothing: asserts that nothing exists at
            # the path and makes sure it is not tracked, so a file another process puts there
            # later is never removed by AfterAll.
            function script:Confirm-LabScratchAbsent {
                param([Parameter(Mandatory)][string]$Path)
                if ([System.IO.File]::Exists($Path) -or [System.IO.Directory]::Exists($Path)) {
                    throw "Confirm-LabScratchAbsent: '$Path' exists, but the call was expected to write nothing."
                }
                $null = $script:TmpFiles.Remove($Path)
            }

            # The ONE way the suite runs the script's -Mode Export into the scratch folder. A file
            # the SCRIPT writes is registered only with evidence attributable to this run
            # (Test-LabExportEvidence: the script's own "Export completed: <path>" line AND file
            # content that carries what this run exported), and only through Confirm-LabScratchFile.
            # A failed call never registers: a file that exists after a failure is not this run's
            # and is reported by path. Returns the script's un-wrapped output text.
            function script:Invoke-LabExport {
                param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][hashtable]$Expect, [hashtable]$ExportParams = @{})
                $errs = New-Object System.Collections.Generic.List[object]
                $p = @{ Mode = 'Export'; Path = $Path; Server = $AronsServer } + $ExportParams
                $out = & { try { & $script:Sync @p } catch { $errs.Add($_) } } *>&1 | Out-String -Width 4096
                if ($errs.Count -gt 0) {
                    if ([System.IO.File]::Exists($Path)) { Write-Warning "Export failed but '$Path' exists; it is NOT tracked (not this run's write) - review it manually." }
                    throw $errs[0]
                }
                $why = script:Test-LabExportEvidence -Path $Path -Output $out -Expect $Expect
                if ($null -eq $why) { script:Confirm-LabScratchFile -Path $Path }
                else { Write-Warning "Export of '$Path' is NOT tracked: $why - review it manually." }
                return $out
            }

            # Export once - the input file every Import/Sync test reuses. The fixture is refused
            # unless the export was tracked: an untracked file is not proof of a write.
            $script:ExportFile = script:New-LabScratchPath -Name 'src.json'
            $null = script:Invoke-LabExport -Path $script:ExportFile -Expect @{ name = $SourceTemplate } -ExportParams @{ TemplateName = $SourceTemplate }
            if (-not $script:TmpFiles.Contains($script:ExportFile)) { throw "Lab fixture refused: Export did not prove a write of '$script:ExportFile'." }
            # The source's OID and schema version, read from the tracked export: they identify a
            # -StripIdentity export of the same template, which carries no name.
            $script:SourceView = [System.IO.File]::ReadAllText($script:ExportFile, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
            if (-not $script:SourceView.'msPKI-Cert-Template-OID') { throw "Lab fixture refused: '$script:ExportFile' carries no msPKI-Cert-Template-OID." }

            function script:New-LabName { "$script:Prefix-$([guid]::NewGuid().ToString('N').Substring(0,6))" }

            # Track an object this suite CREATED by its exact DN and its objectGUID on the server that
            # holds it. AfterEach removes tracked objects most-recent first, by GUID, and only while
            # the DN still carries that GUID (Remove-LabObjectIfOwned). An intended creation is never
            # registered: an object is tracked only with the DN and objectGUID captured AT CREATION -
            # the ones the script printed from the object New-ADObject -PassThru returned, or the
            # ones a -PassThru returned to the test itself. A later lookup by name, DN, OID or
            # display name is not creation evidence and never registers anything. The GUID is
            # mandatory: a name, a prefix or a missing GUID never establishes ownership, so a record
            # without a GUID cannot exist. -DisplayName and -Oid record what the script was asked to
            # give a companion OID object; the ownership check compares them as well.
            # Every record also goes to the run-wide Ledger, which AfterEach never clears: the
            # AfterAll sweep uses it to apply the same ownership rule to a tracked object it finds.
            function script:Register-LabObject {
                param([Parameter(Mandatory)][string]$Server, [Parameter(Mandatory)][string]$DN, [Parameter(Mandatory)][guid]$Guid, [string]$DisplayName, [string]$Oid)
                $rec = @{ Server = $Server; DN = $DN; Guid = $Guid; DisplayName = $DisplayName; Oid = $Oid }
                $script:Created.Add($rec)
                $script:Ledger.Add($rec)
            }

            # One direct read by DN: the object, or $null when no object has that DN. Any other
            # failure (transport, permissions) is rethrown - "absent" is never inferred from a read
            # that failed for another reason.
            function script:Get-LabObjectByDN {
                param([hashtable]$AdParams, [string]$DN, [string[]]$Properties)
                $q = @{} + $AdParams
                if ($Properties) { $q['Properties'] = $Properties }
                try { return (Get-ADObject @q -Identity $DN -ErrorAction Stop) }
                catch {
                    if ($_.Exception.GetType().Name -eq 'ADIdentityNotFoundException') { return $null }
                    throw
                }
            }

            # THE ownership check. AfterEach and AfterAll both remove a tracked object only through
            # this function, so the two cannot drift. It reads the object at the record's DN on
            # -Server (the record's own server by default) and removes it only when it is the
            # tracked object: it must carry the recorded objectGUID and, when the record has them,
            # the recorded displayName and msPKI-Cert-Template-OID. It is then removed BY GUID. Any
            # other GUID means somebody else's object: left alone and reported. A record without a
            # GUID (Register-LabObject refuses one, so this is a guard) is never deleted, only
            # reported: an unreadable or missing GUID is not identity, and a DN is not ownership.
            # Returns @{ Outcome = 'Removed' | 'Absent' | 'Foreign' | 'Unverifiable'; Reason }. A DN
            # with no object is 'Absent' (already gone, or a replica lags). Any other failure is
            # thrown to the caller.
            function script:Remove-LabObjectIfOwned {
                param([Parameter(Mandatory)]$Record, [string]$Server)
                if (-not $Server) { $Server = $Record.Server }
                if ($null -eq $Record.Guid) {
                    return [pscustomobject]@{ Outcome = 'Unverifiable'; Reason = 'the record carries no objectGUID, so ownership cannot be checked' }
                }
                $props = @()
                if ($Record.DisplayName) { $props += 'displayName' }
                if ($Record.Oid)         { $props += 'msPKI-Cert-Template-OID' }
                $cur = script:Get-LabObjectByDN -AdParams @{ Server = $Server } -DN $Record.DN -Properties $props
                if (-not $cur) { return [pscustomobject]@{ Outcome = 'Absent'; Reason = $null } }
                $foreign = $null
                if (-not $cur.ObjectGUID) { return [pscustomobject]@{ Outcome = 'Unverifiable'; Reason = 'the object hides its objectGUID from this caller, so it cannot be matched to the recorded one' } }
                if ($cur.ObjectGUID -ne $Record.Guid) { $foreign = "objectGUID $($cur.ObjectGUID), not the recorded $($Record.Guid)" }
                elseif ($Record.DisplayName -and "$($cur.displayName)" -ne $Record.DisplayName) { $foreign = "displayName '$($cur.displayName)', not the recorded '$($Record.DisplayName)'" }
                elseif ($Record.Oid -and "$($cur.'msPKI-Cert-Template-OID')" -ne $Record.Oid) { $foreign = "OID '$($cur.'msPKI-Cert-Template-OID')', not the recorded '$($Record.Oid)'" }
                if ($foreign) { return [pscustomobject]@{ Outcome = 'Foreign'; Reason = $foreign } }
                Remove-ADObject -Server $Server -Identity $Record.Guid -Confirm:$false -ErrorAction Stop
                [pscustomobject]@{ Outcome = 'Removed'; Reason = $null }
            }

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

            # Run the script and TRACK what it reported creating, by the objectGUID it printed.
            # BEFORE the run: the deterministic template DN (the cn and the container are known) must
            # be absent. An object that already carries it is not this run's: the run is failed with
            # a message instead, and nothing is registered - tracking an intended creation once
            # risked deleting an object the test never created.
            # AFTER the run: the script prints "Created template: <DN> (objectGUID <guid>)" only after
            # New-ADObject -PassThru returned that object, and "Created OID object: <DN> (objectGUID
            # <guid>)" the same way for the companion it creates. The DN and the GUID on those lines
            # come from the returned object itself - creation-time identity, never a later lookup.
            # Each object is registered by THAT DN and THAT GUID, and only by them: no read-back by
            # DN (a replacement put at the DN after the create would be adopted) and no search by
            # OID or display name (a matching object another writer created would be adopted). A
            # line that carries no objectGUID registers nothing and is reported with its DN. A
            # refused run prints no such line, so nothing is tracked. The template line is printed
            # before the ACL step, so a template the script left behind after an ACL failure is
            # still registered. The companion line is printed before the template create, so a
            # companion the script could not roll back is registered as well; one it did roll back
            # is simply absent at cleanup. A captured script error is rethrown once tracking is done.
            # Returns @{ Output; DN; Oid }.
            function script:Invoke-LabCreate {
                param([hashtable]$SyncParams)
                if (-not $SyncParams.NewTemplateName -or -not $SyncParams.NewDisplayName) {
                    throw 'Invoke-LabCreate requires NewTemplateName and NewDisplayName in -SyncParams: without them the expected DN and the companion display name cannot be known, and nothing could be checked.'
                }
                $ap = @{ Server = $SyncParams.Server }
                $cfgNC = (Get-ADRootDSE @ap).configurationNamingContext
                $oidC  = "CN=OID,CN=Public Key Services,CN=Services,$cfgNC"
                $expectedDN = "CN=$($SyncParams.NewTemplateName),CN=Certificate Templates,CN=Public Key Services,CN=Services,$cfgNC"
                $pre = script:Get-LabObjectByDN -AdParams $ap -DN $expectedDN
                if ($pre) {
                    throw "Lab pre-check: '$expectedDN' already exists on $($ap.Server) (objectGUID $($pre.ObjectGUID)). It is not this run's object, so it is neither tracked nor removed; the test cannot run under that name."
                }
                # The script's error is captured, not propagated, so that the objects it reported
                # creating are tracked first; the partial output still carries the lines parsed below.
                # -Width 4096: stop Out-String wrapping the "Created ...: <DN> (objectGUID <guid>)"
                # lines, which would break the DN or the GUID across lines.
                $errs = New-Object System.Collections.Generic.List[object]
                $out = & { try { & $script:Sync @SyncParams } catch { $errs.Add($_) } } *>&1 | Out-String -Width 4096
                $tpl  = script:Read-LabCreatedLine -Output $out -Prefix 'Created template:'
                $comp = script:Read-LabCreatedLine -Output $out -Prefix 'Created OID object:'
                $dn   = if ($tpl) { $tpl.DN } else { $null }
                $oid  = if ($out -match 'Template OID:\s*([0-9.]+)') { $Matches[1] } else { $null }
                # Tracking must not hide the script's own error: a failure in here is rethrown only
                # after a captured script error.
                $trackErr = $null
                try {
                    if ($dn -and $dn -ne $expectedDN) {
                        Write-Warning "Lab tracking: the script reported creating '$dn' on $($ap.Server), not the expected '$expectedDN' - not tracked; review it manually."
                    }
                    elseif ($dn -and -not $tpl.Guid) {
                        Write-Warning "Lab tracking: the script reported creating '$dn' on $($ap.Server) without an objectGUID - not tracked; review it manually (remove it by DN only after you confirmed it is this run's object)."
                    }
                    elseif ($dn) {
                        script:Register-LabObject -Server $ap.Server -DN $dn -Guid $tpl.Guid
                    }
                    if ($comp -and $comp.DN -notlike "CN=*,$oidC") {
                        Write-Warning "Lab tracking: the script reported creating the OID object '$($comp.DN)' on $($ap.Server), outside the expected container '$oidC' - not tracked; review it manually."
                    }
                    elseif ($comp -and -not $comp.Guid) {
                        Write-Warning "Lab tracking: the script reported creating the OID object '$($comp.DN)' on $($ap.Server) without an objectGUID - not tracked; review it manually (remove it by DN only after you confirmed it is this run's object)."
                    }
                    elseif ($comp) {
                        # The record keeps the display name the script was asked to give the companion
                        # and the OID it printed, so the ownership check can compare them as well.
                        script:Register-LabObject -Server $ap.Server -DN $comp.DN -Guid $comp.Guid -DisplayName "$($SyncParams.NewDisplayName)" -Oid "$oid"
                    }
                }
                catch { $trackErr = $_ }
                if ($errs.Count) { throw $errs[0] }
                if ($trackErr) { throw $trackErr }
                [pscustomobject]@{ Output = $out; DN = $dn; Oid = $oid }
            }
        }

        AfterEach {
            # Surgical: remove ONLY the objects this suite created, most-recent first, each through
            # Remove-LabObjectIfOwned - the one ownership check the AfterAll sweep uses as well. A DN
            # with no object is fine (already gone - the script rolled it back, or a replica lags);
            # an object that is not the tracked one (another objectGUID), or one whose objectGUID
            # cannot be checked, is left alone and reported. Any other failure is retried and then
            # REPORTED - a swallowed failure once left an object for the AfterAll sweep to find. The
            # retry also covers a target that lags after a write.
            for ($i = $script:Created.Count - 1; $i -ge 0; $i--) {
                $o = $script:Created[$i]
                $last = $null
                $result = $null
                for ($attempt = 1; $attempt -le 4; $attempt++) {
                    try {
                        $result = script:Remove-LabObjectIfOwned -Record $o
                        $last = $null
                        break
                    }
                    catch {
                        $last = $_
                        if ($attempt -lt 4) { Start-Sleep -Milliseconds 750 }
                    }
                }
                if ($last) { Write-Warning "Cleanup could not remove '$($o.DN)' (objectGUID $($o.Guid)) on $($o.Server): $($last.Exception.Message) - review it manually and remove it by that objectGUID only." }
                elseif ($result.Outcome -eq 'Foreign') { Write-Warning "Cleanup left '$($o.DN)' on $($o.Server) alone: it now carries $($result.Reason), so it is not this run's object." }
                elseif ($result.Outcome -eq 'Unverifiable') { Write-Warning "Cleanup left '$($o.DN)' on $($o.Server) alone: $($result.Reason) - review it manually." }
                else { $script:Removed.Add($o.DN.ToLowerInvariant()) }
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
            # OWNERSHIP: the sweep is not a way around the GUID check. An object it finds at a DN the
            # run-wide Ledger knows is a TRACKED object and goes through Remove-LabObjectIfOwned
            # exactly as in AfterEach: a foreign objectGUID at a tracked DN stays, and is reported.
            # An object at a DN the Ledger does not know is UNTRACKED: the prefix is not proof that
            # this run created it (a pre-check that found an existing prefixed object throws without
            # registering it), and a GUID the sweep reads now proves nothing either. The sweep
            # REPORTS such an object - DN and objectGUID - for manual review and never deletes it.
            # A companion OID object is never chased by its OID alone: the displayName sweep finds
            # it, and it goes through the same rules. The Ledger and Removed lists can be unset when
            # BeforeAll threw.
            if ($script:Prefix -match '^PESTER-[0-9a-f]{8}$') {
                foreach ($srv in @($AronsServer, $ChildServer, $NorefjellServer | Where-Object { $_ } | Select-Object -Unique)) {
                    try {
                        $cfg = (Get-ADRootDSE -Server $srv).configurationNamingContext
                        $tpls = "CN=Certificate Templates,CN=Public Key Services,CN=Services,$cfg"
                        $oidC = "CN=OID,CN=Public Key Services,CN=Services,$cfg"
                        $found = @(Get-ADObject -Server $srv -SearchBase $tpls -LDAPFilter "(cn=$script:Prefix-*)" -ErrorAction SilentlyContinue) +
                                 @(Get-ADObject -Server $srv -SearchBase $oidC -LDAPFilter "(displayName=$script:Prefix-*)" -ErrorAction SilentlyContinue)
                        foreach ($t in @($found | Where-Object { $_ })) {
                            $dnKey = $t.DistinguishedName.ToLowerInvariant()
                            $rec = @($script:Ledger | Where-Object { $_.DN.ToLowerInvariant() -eq $dnKey }) | Select-Object -Last 1
                            try {
                                if ($rec) {
                                    $r = script:Remove-LabObjectIfOwned -Record $rec -Server $srv
                                    if ($r.Outcome -eq 'Foreign') {
                                        Write-Warning "AfterAll safety-net left '$($t.DistinguishedName)' on $srv alone: it carries $($r.Reason), so it is not this run's object."
                                    }
                                    elseif ($r.Outcome -eq 'Unverifiable') {
                                        Write-Warning "AfterAll safety-net left '$($t.DistinguishedName)' on $srv alone: $($r.Reason) - review it manually."
                                    }
                                    elseif ($r.Outcome -eq 'Removed' -and $script:Removed -and $script:Removed.Contains($dnKey)) {
                                        # This run removed that exact object already: this server is a replica that has
                                        # not received the deletion yet (a child DC lags the root by ~18 s in this lab).
                                        Write-Warning "AfterAll safety-net found a lagging replica of an object this run had already removed: $($t.DistinguishedName) on $srv"
                                    }
                                    elseif ($r.Outcome -eq 'Removed') {
                                        Write-Warning "AfterAll safety-net removed a tracked object that AfterEach had not removed: $($t.DistinguishedName) on $srv"
                                    }
                                }
                                else {
                                    # Report only. The Ledger does not own this object, so nothing
                                    # proves this run created it; the prefix and the GUID read here
                                    # are not proof. Never delete it from here.
                                    $guidText = if ($t.ObjectGUID) { "objectGUID $($t.ObjectGUID)" } else { 'objectGUID not readable' }
                                    Write-Warning "AfterAll safety-net found an UNTRACKED object carrying this run's prefix: '$($t.DistinguishedName)' ($guidText) on $srv. This run cannot prove it created it, so it was NOT removed - review it manually."
                                }
                            }
                            catch { Write-Warning "AfterAll safety-net could not check or remove '$($t.DistinguishedName)' on $($srv): $($_.Exception.Message)" }
                        }
                    }
                    catch { }
                }
            } else {
                Write-Warning "Safety-net sweep skipped: run prefix is unset or malformed ('$script:Prefix')."
            }
            # Remove the scratch folder only when THIS run created it (the CreateDirectoryW
            # 'Created' verdict in BeforeAll set the flag). A folder that merely carries the name
            # is never removed: it is reported for manual review and left in place. Even an owned
            # folder is not emptied by name: Remove-LabScratchFolder removes only the files this
            # run recorded in $script:TmpFiles (one file at a time, never a folder), then the folder
            # non-recursively; anything untracked that remains stays and is reported by path.
            if ($script:TmpDirCreated -and $script:TmpDir) {
                script:Remove-LabScratchFolder -Path $script:TmpDir -TrackedFiles @($script:TmpFiles)
            }
            elseif ($script:TmpDir -and (Test-Path -LiteralPath $script:TmpDir)) {
                Write-Warning "Scratch folder '$script:TmpDir' exists but this run has no proof it created it, so it was NOT removed - review it manually."
            }
        }

        It 'Export writes a BOM-marked JSON with identity, OID, and no ACL attribute' {
            $bytes = [System.IO.File]::ReadAllBytes($script:ExportFile)
            ($bytes[0..2] -join ',') | Should -Be '239,187,191'   # UTF-8 BOM
            $j = Get-Content -LiteralPath $script:ExportFile -Raw -Encoding UTF8 | ConvertFrom-Json
            $j.name | Should -Be $SourceTemplate
            $j.'msPKI-Cert-Template-OID' | Should -Not -BeNullOrEmpty
            $j.PSObject.Properties.Name | Should -Not -Contain 'pKIEnrollmentAccess'
        }

        It 'Export refuses a wildcard (metacharacter treated literally) and writes no file' {
            $w = script:New-LabScratchPath -Name 'w.json'   # a name only: a refused export must not become tracked
            { & $script:Sync -Mode Export -Path $w -TemplateName 'Kerb*' -Server $AronsServer } |
                Should -Throw -ExpectedMessage '*was not found*'
            script:Confirm-LabScratchAbsent -Path $w   # a refused export writes nothing, so nothing is tracked
            Test-Path -LiteralPath $w | Should -BeFalse -Because 'a refused export writes nothing'
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

        It 'Import REFUSES a malformed export before creating anything (a bad msPKI-RA-Signature is not dropped)' {
            $tampered = script:New-LabScratchPath -Name 'tampered.json'
            $obj = [System.IO.File]::ReadAllText($script:ExportFile, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
            $obj | Add-Member -NotePropertyName 'msPKI-RA-Signature' -NotePropertyValue 'abc' -Force
            [System.IO.File]::WriteAllText($tampered, ($obj | ConvertTo-Json -Depth 8), [System.Text.UTF8Encoding]::new($true))
            script:Confirm-LabScratchFile -Path $tampered
            $cn = script:New-LabName
            { & $script:Sync -Mode Import -Path $tampered -Server $AronsServer -NewTemplateName $cn -NewDisplayName $cn -OidHandling GenerateRandom -SkipAcl *>$null } |
                Should -Throw -ExpectedMessage '*Import refused*msPKI-RA-Signature*'
            (Get-ADObject @script:AP -SearchBase $script:TplBase -LDAPFilter "(cn=$cn)") | Should -BeNullOrEmpty -Because 'nothing may be created from a malformed export'
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
            # The Standard set is the WHOLE DACL: exactly these six principals with exactly these
            # rights, nothing inherited, nothing from the schema default, no extra bit anywhere (a
            # stray Everyone:FullControl, or WriteProperty added to Authenticated Users, would pass
            # a test that only looks for the expected entries).
            $domSid  = (Get-ADDomain @script:AP).DomainSID.Value
            $rootSid = (Get-ADDomain @script:AP -Identity (Get-ADForest @script:AP).RootDomain).DomainSID.Value
            $rules1  = @($t1.nTSecurityDescriptor.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]))
            script:Assert-ExactDacl -Rules $rules1 -Expected @(
                @{ Sid = 'S-1-5-11';        Right = 'read' }
                @{ Sid = "$rootSid-498";    Right = 'enroll' }, @{ Sid = "$rootSid-498"; Right = 'autoenroll' }
                @{ Sid = "$domSid-512";     Right = 'read' },   @{ Sid = "$domSid-512";  Right = 'write' }, @{ Sid = "$domSid-512"; Right = 'enroll' }
                @{ Sid = "$domSid-516";     Right = 'enroll' }, @{ Sid = "$domSid-516";  Right = 'autoenroll' }
                @{ Sid = "$rootSid-519";    Right = 'read' },   @{ Sid = "$rootSid-519"; Right = 'write' }, @{ Sid = "$rootSid-519"; Right = 'enroll' }
                @{ Sid = 'S-1-5-9';         Right = 'enroll' }, @{ Sid = 'S-1-5-9';      Right = 'autoenroll' }
            )

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
            $dcSid = (Get-ADDomain @script:AP).DomainSID.Value + '-516'
            # "exactly the requested grants": Authenticated Users read, Domain Controllers enroll +
            # autoenroll - three ACEs, no other principal, no other bit, nothing inherited, no deny
            $t.nTSecurityDescriptor.AreAccessRulesProtected | Should -BeTrue
            script:Assert-ExactDacl -Rules $rules -Expected @(
                @{ Sid = 'S-1-5-11'; Right = 'read' }
                @{ Sid = $dcSid;     Right = 'enroll' }, @{ Sid = $dcSid; Right = 'autoenroll' }
            )
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
            $bad = script:New-LabScratchPath -Name 'bad.json'
            $j = Get-Content -LiteralPath $script:ExportFile -Raw -Encoding UTF8 | ConvertFrom-Json
            $j.'msPKI-Cert-Template-OID' = '1.2.3.*)(cn=*'
            $j | ConvertTo-Json -Depth 10 | Out-File -Encoding utf8 -LiteralPath $bad
            script:Confirm-LabScratchFile -Path $bad
            { & $script:Sync -Mode Import -Path $bad -Server $AronsServer -NewTemplateName (script:New-LabName) -NewDisplayName x } |
                Should -Throw -ExpectedMessage '*not a valid dotted OID*'

            $j2 = Get-Content -LiteralPath $script:ExportFile -Raw -Encoding UTF8 | ConvertFrom-Json
            $j2.name = 'evil,cn=x'
            $j2 | ConvertTo-Json -Depth 10 | Out-File -Encoding utf8 -LiteralPath $bad
            script:Confirm-LabScratchFile -Path $bad
            { & $script:Sync -Mode Import -Path $bad -Server $AronsServer -OidHandling GenerateRandom } |
                Should -Throw -ExpectedMessage '*may only contain*'
        }

        It 'Import -AclBase SchemaPlusStandard keeps the schema default (unprotected, SYSTEM Full Control) and adds the Standard grants' {
            $cn = script:New-LabName
            $r = script:Invoke-LabCreate -SyncParams @{ Mode = 'Import'; Path = $script:ExportFile; Server = $AronsServer; NewTemplateName = $cn; NewDisplayName = $cn; OidHandling = 'GenerateRandom'; AclBase = 'SchemaPlusStandard' }
            $r.DN | Should -Not -BeNullOrEmpty
            $r.Output | Should -Match 'added on top of the schema-default ACL'
            $t = script:Get-ADObjectRetry -AdParams $script:AP -Identity $r.DN -Properties nTSecurityDescriptor
            $sd = $t.nTSecurityDescriptor
            $sd.AreAccessRulesProtected | Should -BeFalse -Because 'SchemaPlusStandard removes nothing and keeps inheritance'
            $own = @($sd.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]) |
                Where-Object { -not $_.IsInherited -and $_.AccessControlType -eq [System.Security.AccessControl.AccessControlType]::Allow })
            # the schema default is still there: SYSTEM keeps Full Control
            $sys = @($own | Where-Object { $_.IdentityReference.Value -eq 'S-1-5-18' })
            $sys.Count | Should -Be 1
            (([int64][int]$sys[0].ActiveDirectoryRights) -band 0xF01FF) | Should -Be 0xF01FF -Because 'the schema-default SYSTEM Full Control ACE is kept'
            # every Standard grant is present, compared by (SID, object GUID): the DS coalesces
            # same-typed grants for one principal into ONE ACE (Domain Admins read+write join the
            # schema-default Full Control entry), so each pair must have exactly one ACE that
            # carries AT LEAST the expected bits - the schema default may add more.
            $domSid  = (Get-ADDomain @script:AP).DomainSID.Value
            $rootSid = (Get-ADDomain @script:AP -Identity (Get-ADForest @script:AP).RootDomain).DomainSID.Value
            $enrollGuid = [Guid]'0e10c968-78fb-11d2-90d4-00c04f79dc55'
            $autoGuid   = [Guid]'a05b8cc2-17bc-4802-a710-e7c15ab866a2'
            $read = 0x20094; $write = 0x20028; $ext = 0x100   # the DS expansions of GenericRead / GenericWrite; ExtendedRight
            $expected = @(
                @{ Sid = 'S-1-5-11';     Guid = [Guid]::Empty; Bits = $read }
                @{ Sid = "$rootSid-498"; Guid = $enrollGuid;   Bits = $ext }, @{ Sid = "$rootSid-498"; Guid = $autoGuid; Bits = $ext }
                @{ Sid = "$domSid-512";  Guid = [Guid]::Empty; Bits = ($read -bor $write) }, @{ Sid = "$domSid-512"; Guid = $enrollGuid; Bits = $ext }
                @{ Sid = "$domSid-516";  Guid = $enrollGuid;   Bits = $ext }, @{ Sid = "$domSid-516"; Guid = $autoGuid; Bits = $ext }
                @{ Sid = "$rootSid-519"; Guid = [Guid]::Empty; Bits = ($read -bor $write) }, @{ Sid = "$rootSid-519"; Guid = $enrollGuid; Bits = $ext }
                @{ Sid = 'S-1-5-9';      Guid = $enrollGuid;   Bits = $ext }, @{ Sid = 'S-1-5-9'; Guid = $autoGuid; Bits = $ext }
            )
            foreach ($e in $expected) {
                $m = @($own | Where-Object { $_.IdentityReference.Value -eq $e.Sid -and $_.ObjectType -eq $e.Guid })
                $m.Count | Should -Be 1 -Because "one ACE for $($e.Sid) / $($e.Guid) (the DS coalesces same-typed grants)"
                $bits = ([int64][int]$m[0].ActiveDirectoryRights) -band 0xFFFFFFFF
                ($bits -band $e.Bits) | Should -Be $e.Bits -Because "the ACE for $($e.Sid) / $($e.Guid) must carry the Standard grant (0x$('{0:X}' -f $e.Bits))"
            }
        }

        It 'Import -WhatIf creates nothing and previews the ACL decision' {
            $cn = script:New-LabName
            $out = & $script:Sync -Mode Import -Path $script:ExportFile -Server $AronsServer -NewTemplateName $cn -NewDisplayName $cn -OidHandling GenerateRandom -AclBase Schema -WhatIf *>&1 | Out-String -Width 4096
            $out | Should -Match 'What if: Would leave the schema-default ACL'
            $out | Should -Not -Match 'Created template:'
            (Get-ADObject @script:AP -SearchBase $script:TplBase -LDAPFilter "(cn=$cn)")          | Should -BeNullOrEmpty -Because '-WhatIf must create no template'
            (Get-ADObject @script:AP -SearchBase $script:OidBase -LDAPFilter "(displayName=$cn)") | Should -BeNullOrEmpty -Because '-WhatIf must register no companion OID object'
        }

        It 'Import with -Server set to the DOMAIN name warns that the steps may reach different replicas (-WhatIf: nothing created)' {
            $domainName = (Get-ADDomain @script:AP).DNSRoot
            $cn = script:New-LabName
            $out = & $script:Sync -Mode Import -Path $script:ExportFile -Server $domainName -NewTemplateName $cn -NewDisplayName $cn -OidHandling GenerateRandom -AclBase Schema -WhatIf *>&1 | Out-String -Width 4096
            $out | Should -Match 'is a DOMAIN name'
            (Get-ADObject @script:AP -SearchBase $script:TplBase -LDAPFilter "(cn=$cn)") | Should -BeNullOrEmpty
        }

        It 'Import refuses a -StripIdentity export without -NewTemplateName, and a -StripOid export with -OidHandling Preserve' {
            $noId = script:New-LabScratchPath -Name 'strip-identity.json'
            # no name to match after -StripIdentity: the source's OID and schema version identify this run's export
            $null = script:Invoke-LabExport -Path $noId -ExportParams @{ TemplateName = $SourceTemplate; StripIdentity = $true } `
                -Expect @{ 'msPKI-Cert-Template-OID' = $script:SourceView.'msPKI-Cert-Template-OID'; 'msPKI-Template-Schema-Version' = $script:SourceView.'msPKI-Template-Schema-Version' }
            Test-Path -LiteralPath $noId -PathType Leaf | Should -BeTrue -Because 'the suite named the file the script wrote'
            $script:TmpFiles | Should -Contain $noId -Because 'the confirmed write is what makes the file removable'
            $j = [System.IO.File]::ReadAllText($noId, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
            $j.PSObject.Properties.Name | Should -Not -Contain 'name'
            $j.PSObject.Properties.Name | Should -Not -Contain 'displayName'
            $j.'msPKI-Cert-Template-OID' | Should -Not -BeNullOrEmpty -Because '-StripIdentity keeps the OID'
            { & $script:Sync -Mode Import -Path $noId -Server $AronsServer -OidHandling GenerateRandom -AclBase Schema } |
                Should -Throw -ExpectedMessage '*supply -NewTemplateName*'
            $noOid = script:New-LabScratchPath -Name 'strip-oid.json'
            $null = script:Invoke-LabExport -Path $noOid -Expect @{ name = $SourceTemplate } -ExportParams @{ TemplateName = $SourceTemplate; StripOid = $true }
            Test-Path -LiteralPath $noOid -PathType Leaf | Should -BeTrue -Because 'the suite named the file the script wrote'
            $script:TmpFiles | Should -Contain $noOid -Because 'the confirmed write is what makes the file removable'
            $j2 = [System.IO.File]::ReadAllText($noOid, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
            $j2.PSObject.Properties.Name | Should -Not -Contain 'msPKI-Cert-Template-OID'
            $j2.name | Should -Be $SourceTemplate -Because '-StripOid keeps the identity'
            $cn = script:New-LabName
            { & $script:Sync -Mode Import -Path $noOid -Server $AronsServer -NewTemplateName $cn -NewDisplayName $cn -AclBase Schema } |
                Should -Throw -ExpectedMessage '*Re-export without -StripOid*'
            (Get-ADObject @script:AP -SearchBase $script:TplBase -LDAPFilter "(cn=$cn)") | Should -BeNullOrEmpty
        }

        It 'Import refuses an issuance policy the target forest links to a group (AMA) and accepts it only with -AllowLinkedIssuancePolicy' {
            # A prefixed universal group and a prefixed issuance-policy OID object linked to it
            # (msDS-OIDToGroupLink) - the Authentication Mechanism Assurance shape. Both tracked by
            # DN and objectGUID from the -PassThru objects.
            $grp = New-ADGroup @script:AP -Name "$script:Prefix-AMA" -GroupScope Universal -GroupCategory Security -Path (Get-ADDomain @script:AP).UsersContainer -PassThru -ErrorAction Stop
            script:Register-LabObject -Server $AronsServer -DN $grp.DistinguishedName -Guid $grp.ObjectGUID
            $policyOid = "$(New-SyntheticOidBase).$(Get-Random -Minimum 10000000 -Maximum 99999999).400"
            $oidObj = New-ADObject @script:AP -Path $script:OidBase -Name "$script:Prefix-AMA-POLICY" -Type 'msPKI-Enterprise-Oid' -PassThru -ErrorAction Stop `
                -OtherAttributes @{ DisplayName = "$script:Prefix-AMA-Policy"; flags = [int]2; 'msPKI-Cert-Template-OID' = $policyOid; 'msDS-OIDToGroupLink' = $grp.DistinguishedName }
            script:Register-LabObject -Server $AronsServer -DN $oidObj.DistinguishedName -Guid $oidObj.ObjectGUID
            # an export whose issued certificates would carry that policy
            $tampered = script:New-LabScratchPath -Name 'ama.json'
            $obj = [System.IO.File]::ReadAllText($script:ExportFile, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
            $obj | Add-Member -NotePropertyName 'msPKI-Certificate-Policy' -NotePropertyValue @($policyOid) -Force
            [System.IO.File]::WriteAllText($tampered, ($obj | ConvertTo-Json -Depth 8), [System.Text.UTF8Encoding]::new($true))
            script:Confirm-LabScratchFile -Path $tampered
            $cn = script:New-LabName
            { & $script:Sync -Mode Import -Path $tampered -Server $AronsServer -NewTemplateName $cn -NewDisplayName $cn -OidHandling GenerateRandom -AclBase Schema } |
                Should -Throw -ExpectedMessage "*Import refused*Authentication Mechanism Assurance*$policyOid*$($grp.DistinguishedName)*"
            (Get-ADObject @script:AP -SearchBase $script:TplBase -LDAPFilter "(cn=$cn)") | Should -BeNullOrEmpty -Because 'the refusal happens before anything is created'
            $r = script:Invoke-LabCreate -SyncParams @{ Mode = 'Import'; Path = $tampered; Server = $AronsServer; NewTemplateName = $cn; NewDisplayName = $cn; OidHandling = 'GenerateRandom'; AclBase = 'Schema'; AllowLinkedIssuancePolicy = $true }
            $r.DN | Should -Not -BeNullOrEmpty -Because '-AllowLinkedIssuancePolicy accepts the mapping deliberately'
            $r.Output | Should -Match 'Authentication Mechanism Assurance'
            $r.Output | Should -Match 'AllowLinkedIssuancePolicy given; proceeding'
            $t = script:Get-ADObjectRetry -AdParams $script:AP -Identity $r.DN -Properties 'msPKI-Certificate-Policy'
            @($t.'msPKI-Certificate-Policy') | Should -Contain $policyOid
        }

        It 'Validate passes both pipelines for the source template' {
            $out = & $script:Sync -Mode Validate -TemplateName $SourceTemplate -Server $AronsServer *>&1 | Out-String
            $out | Should -Match 'OVERALL PASS'
            $out | Should -Match 'Cleaned up'
            # Validate cleans up its own throwaways; nothing to track.
        }

        It 'Validate -KeepArtifacts leaves the throwaway templates and the export file in place; the suite tracks them only from the run''s own output' {
            # The throwaways carry the script's own RoundtripTest-/DirectPathTest- names, not this
            # run's prefix, so the AfterAll sweep cannot find them: they are removed ONLY through
            # the records this test registers. Every record comes from the run's own output - the
            # "Created template:" / "Created OID object:" lines (DN and objectGUID from the object
            # New-ADObject -PassThru returned) and the "Export completed:" line plus matching
            # JSON for the file. A failed run registers nothing and reports what it kept.
            $keep = script:New-LabScratchPath -Name 'keep.json'
            $errs = New-Object System.Collections.Generic.List[object]
            $out = & { try { & $script:Sync -Mode Validate -TemplateName $SourceTemplate -Path $keep -KeepArtifacts -Server $AronsServer } catch { $errs.Add($_) } } *>&1 | Out-String -Width 4096
            # Validate creates one throwaway per pipeline, so there are several "Created" lines;
            # the parser reads the first match only, so feed it one line at a time.
            $lines = $out -split "`r?`n"
            $tpls  = @(foreach ($l in $lines) { if ($l -like 'Created template:*')   { script:Read-LabCreatedLine -Output $l -Prefix 'Created template:' } })
            $comps = @(foreach ($l in $lines) { if ($l -like 'Created OID object:*') { script:Read-LabCreatedLine -Output $l -Prefix 'Created OID object:' } })
            if ($errs.Count) {
                foreach ($k in @($tpls + $comps)) { Write-Warning "Validate -KeepArtifacts failed; the kept object '$($k.DN)' (objectGUID $($k.Guid)) on $AronsServer is NOT tracked - review it manually." }
                if ([System.IO.File]::Exists($keep)) { Write-Warning "Validate -KeepArtifacts failed but '$keep' exists; it is NOT tracked (not this run's write) - review it manually." }
                throw $errs[0]
            }
            # Register BEFORE any assertion, so a failed assertion cannot leak a kept object. Same
            # rules as Invoke-LabCreate: the printed DN and objectGUID only, and only inside the
            # expected container.
            foreach ($t in $tpls) {
                if (-not $t.Guid) { Write-Warning "Lab tracking: the script reported creating '$($t.DN)' on $AronsServer without an objectGUID - not tracked; review it manually."; continue }
                if ($t.DN -notlike "CN=*,$script:TplBase") { Write-Warning "Lab tracking: the script reported creating '$($t.DN)' on $AronsServer outside '$script:TplBase' - not tracked; review it manually."; continue }
                script:Register-LabObject -Server $AronsServer -DN $t.DN -Guid $t.Guid
            }
            foreach ($c in $comps) {
                if (-not $c.Guid) { Write-Warning "Lab tracking: the script reported creating the OID object '$($c.DN)' on $AronsServer without an objectGUID - not tracked; review it manually."; continue }
                if ($c.DN -notlike "CN=*,$script:OidBase") { Write-Warning "Lab tracking: the script reported creating the OID object '$($c.DN)' on $AronsServer outside '$script:OidBase' - not tracked; review it manually."; continue }
                script:Register-LabObject -Server $AronsServer -DN $c.DN -Guid $c.Guid
            }
            $why = script:Test-LabExportEvidence -Path $keep -Output $out -Expect @{ name = $SourceTemplate }
            if ($null -eq $why) { script:Confirm-LabScratchFile -Path $keep }
            else { Write-Warning "Validate -KeepArtifacts export '$keep' is NOT tracked: $why - review it manually." }

            $out | Should -Match 'OVERALL PASS'
            $out | Should -Not -Match 'Cleaned up'
            $tpls.Count | Should -Be 2   # one throwaway per pipeline (file and direct)
            $keptDNs = @([regex]::Matches($out, "(?m)^-KeepArtifacts: left throwaway template '(.+?)' in place") | ForEach-Object { $_.Groups[1].Value })
            $keptDNs.Count | Should -Be $tpls.Count
            foreach ($t in $tpls) { $keptDNs | Should -Contain $t.DN }
            $out | Should -Match ("(?m)^-KeepArtifacts: left file '" + [regex]::Escape($keep) + "' in place")
            [System.IO.File]::Exists($keep) | Should -BeTrue
            foreach ($k in @($tpls + $comps)) {
                # Still there after the run, and still the object the script created.
                $cur = script:Get-LabObjectByDN -AdParams $script:AP -DN $k.DN
                $cur | Should -Not -BeNullOrEmpty
                $cur.ObjectGUID | Should -Be $k.Guid
                # Tracked: a regression in the evidence parsing fails here instead of leaking.
                @($script:Created | Where-Object { $_.DN -eq $k.DN -and $_.Guid -eq $k.Guid }).Count | Should -Be 1
            }
            $script:TmpFiles.Contains($keep) | Should -BeTrue
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
                # PrincipalsOnly replaces the whole DACL, so the caller must grant itself a Read or
                # the cleanup read-back cannot see the objectGUID of the created object, and the
                # ownership check then leaves it in place as unverifiable. The target forest
                # has no trust to the caller's forest: it authenticates the caller as its own account,
                # so the caller's SID here is not the one the source forest knows. Authenticated Users
                # (S-1-5-11) is in every authenticated token on every forest, so it is the identity
                # this test can grant. The DomainControllers grant is the one the test is about.
                $r = script:Invoke-LabCreate -SyncParams @{ Mode = 'Sync'; SourceServer = $AronsServer; Server = $NorefjellServer; TemplateName = $SourceTemplate; NewTemplateName = $cn; NewDisplayName = $cn; OidHandling = 'GenerateRandom'; AclBase = 'PrincipalsOnly'; EnrollPrincipals = @{ 'AuthenticatedUsers' = 'Read'; 'DomainControllers' = 'Enroll' } }
                $r.DN | Should -Not -BeNullOrEmpty -Because 'the template must be created in the target forest'
                $r.DN | Should -BeLike '*DC=norefjell,DC=local' -Because 'creation must land in the target forest'
                $t = script:Get-ADObjectRetry -AdParams $tp -Identity $r.DN -Properties nTSecurityDescriptor
                $tgtDcSid = (Get-ADDomain @tp).DomainSID.Value + '-516'
                $rules = @($t.nTSecurityDescriptor.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]))
                ($rules | Where-Object { $_.IdentityReference.Value -eq $tgtDcSid }) |
                    Should -Not -BeNullOrEmpty -Because 'DomainControllers must resolve in the TARGET forest'
                # exactly the requested grants: the target's Domain Controllers enroll, Authenticated
                # Users read - two ACEs, no other principal, no other bit, nothing inherited, no deny
                $t.nTSecurityDescriptor.AreAccessRulesProtected | Should -BeTrue
                script:Assert-ExactDacl -Rules $rules -Expected @(
                    @{ Sid = 'S-1-5-11'; Right = 'read' }
                    @{ Sid = $tgtDcSid;  Right = 'enroll' }
                )
                $t.ObjectGUID | Should -Not -BeNullOrEmpty -Because 'the Read grant keeps the object readable, so the ownership check can compare its objectGUID'
                @($script:Created | Where-Object { $_.DN -eq $r.DN -and $_.Guid -eq $t.ObjectGUID -and $_.Server -eq $NorefjellServer }).Count |
                    Should -Be 1 -Because 'the template must be tracked by the objectGUID the script printed at creation, which the target must report back'
            }

            It 'Sync with the default -OidHandling Preserve carries a fresh source OID into the target forest and registers its companion there' {
                # A prefixed throwaway in the SOURCE forest with a fresh OID: the seeded default
                # templates of the target share the source forest's OIDs, so the stock source
                # template would be refused there as a duplicate OID.
                $srcCn = script:New-LabName
                $src = script:Invoke-LabCreate -SyncParams @{ Mode = 'Import'; Path = $script:ExportFile; Server = $AronsServer; NewTemplateName = $srcCn; NewDisplayName = $srcCn; OidHandling = 'GenerateRandom'; SkipAcl = $true }
                $src.DN  | Should -Not -BeNullOrEmpty
                $src.Oid | Should -Not -BeNullOrEmpty
                $cn = script:New-LabName
                $tp = @{ Server = $NorefjellServer }
                $r = script:Invoke-LabCreate -SyncParams @{ Mode = 'Sync'; SourceServer = $AronsServer; Server = $NorefjellServer; TemplateName = $srcCn; NewTemplateName = $cn; NewDisplayName = $cn; AclBase = 'Schema' }
                $r.DN | Should -Not -BeNullOrEmpty -Because 'the template must be created in the target forest'
                $r.DN | Should -BeLike '*DC=norefjell,DC=local'
                $r.Oid | Should -Be $src.Oid -Because 'Preserve keeps the source OID'
                $r.Output | Should -Match '\(Preserve\)'
                $t = script:Get-ADObjectRetry -AdParams $tp -Identity $r.DN -Properties 'msPKI-Cert-Template-OID'
                $t.'msPKI-Cert-Template-OID' | Should -Be $src.Oid
                $tgtOidC = "CN=OID,CN=Public Key Services,CN=Services,$((Get-ADRootDSE @tp).configurationNamingContext)"
                $comp = script:Get-ADObjectRetry -AdParams $tp -SearchBase $tgtOidC -Filter "(msPKI-Cert-Template-OID=$($src.Oid))" -Properties displayName
                $comp | Should -Not -BeNullOrEmpty -Because 'Preserve registers a display object for a carried OID the target does not know'
                $comp.displayName | Should -Be $cn
                $comp.DistinguishedName | Should -BeLike '*DC=norefjell,DC=local'
                @($script:Created | Where-Object { $_.DN -eq $comp.DistinguishedName -and $_.Guid -eq $comp.ObjectGUID -and $_.Server -eq $NorefjellServer }).Count |
                    Should -Be 1 -Because 'the companion must be tracked by the exact DN and objectGUID the script printed at creation'
                @($script:Created | Where-Object { $_.DN -eq $r.DN -and $_.Guid -eq $t.ObjectGUID -and $_.Server -eq $NorefjellServer }).Count |
                    Should -Be 1 -Because 'the template must be tracked by the objectGUID the script printed at creation, which the target must report back'
            }

            It 'Generate fails cleanly against a forest with no PKI OID root' {
                { & $script:Sync -Mode Sync -SourceServer $AronsServer -Server $NorefjellServer -TemplateName $SourceTemplate `
                        -NewTemplateName (script:New-LabName) -NewDisplayName x -OidHandling Generate -AclBase Schema -WhatIf } |
                    Should -Throw -ExpectedMessage '*no PKI OID root*'
            }
        }
    }
}
