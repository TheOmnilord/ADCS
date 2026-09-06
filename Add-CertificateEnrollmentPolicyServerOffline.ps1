<#PSScriptInfo
.VERSION 1.0.5
.GUID 61adf5d1-6eb5-4f41-8670-e9da72134570
.AUTHOR Sveinung Svea
.PROJECTURI https://github.com/TheOmnilord/ADCS
.LICENSEURI https://github.com/TheOmnilord/ADCS/blob/main/LICENSE
.TAGS ADCS PKI CertificateServices
.RELEASENOTES
1.0.4 - Every existing key on the target path is checked before any write: a registry symbolic link, an untrusted owner or write-class rights for an untrusted principal refuse the run (a link planted where PolicyServers did not exist yet would have carried an elevated first-time write to whatever key it pointed at); string values and the (Default) marker are written as REG_SZ explicitly and every value's KIND is verified (Set-ItemProperty kept an existing wrong kind and the string compare accepted it; a REG_SZ root Flags with the right number is now repaired to a DWORD); the cmdlets inside an approved action pass -Confirm:$false and a removal is verified before it is reported (a declined nested Remove-Item prompt reported a removal that never happened and could clear the marker of an entry that still existed)
1.0.3 - Help text only: -ReplaceExisting documents the complete-row requirement; no code change
1.0.2 - A pre-existing AD-policy row or CEP entry counts as complete only with URL, PolicyID, FriendlyName and DWORD Flags/AuthFlags/Cost present (URL + PolicyID alone is what an interrupted write leaves), for the prerequisite, the (Default) marker and -ReplaceExisting alike
1.0.1 - For the GP locations the domain objectGUID for the AD Enrollment Policy row is resolved BEFORE the CEP entry is written; on a domain-joined machine a lookup failure now aborts the run with nothing written (previously the entry was written first and the failure became a warning, leaving a GP configuration that removes the AD enrollment policy); on a workgroup machine the row is still skipped with a warning
1.0.0 - Initial release
#>

<#
.SYNOPSIS
    Registers (or removes) a certificate enrollment policy (CEP) server in the registry
    completely OFFLINE - no "Validate Server" round-trip, no contact with the policy server.

.DESCRIPTION
    Writes the same registry values the "Certificate Services Client - Certificate Enrollment
    Policy" dialog produces, computing everything locally instead of calling the server's
    MS-XCEP GetPolicies endpoint:

      * Subkey name = SHA-1 over the UTF-16LE bytes of the invariant-lowercased URL
        (same derivation X509Enrollment.CX509PolicyServerUrl::UpdateRegistry uses)
      * PolicyID    = Java String.hashCode() of the policy name - EJBCA MSAE behavior, where
        the alias's "Policy Name" is both the friendly name and the hashCode input. For other
        CEP products (e.g. Microsoft's CEP web service, which returns a GUID) pass -PolicyId.

    Locations:
      LocalMachine (default) / LocalUser - the "user configured" store that certlm.msc /
        certmgr.msc "Manage enrollment policies" manages (Software\Microsoft\...).
      GPMachine / GPUser - the Group Policy hives (Software\Policies\...).
        WARNING (tattooing): on a domain member, values written directly into the GP hives are
        pseudo-policy backed by no GPO: they appear in no GPMC/RSoP/gpresult report, gpupdate
        neither restores nor removes them, the certificate MMC shows them read-only, and a real
        GPO that later manages the same key will silently overwrite shared values while the
        entry subkey lingers as a phantom. On domain members use the companion
        Add-CertificateEnrollmentPolicyServerToGpo.ps1 instead; the GP locations here are
        intended for standalone/workgroup machines and lab work.

    For the GP locations the script also maintains the PolicyServers root Flags value. The
    root Flags are DISABLE bits (EnrollmentPolicyFlags): 0x2 makes clients ignore the whole
    GP-provided list (never set; the script clears it with a warning if found), 0x4 makes
    clients ignore user-configured servers. Existing bits are preserved: rerunning without
    -DisableUserConfigured does NOT clear a previously set 0x4 (use -EnableUserConfigured).

    For the GP locations the script by default also ensures the built-in AD enrollment policy
    row ("LDAP:", subkey 37c9dc30f207f27f61a2f7c3aed598a6e2920b54, PolicyID = the domain
    object's objectGUID, Cost 0xFFFFFFFF), because enabling GP-based CEP configuration
    suppresses the client-side synthesized AD default policy: without this row, machines lose
    the AD enrollment policy and autoenrollment against AD-published templates stops. Opt out
    with -SkipADPolicy (e.g. on workgroup machines, where the lookup is skipped with a warning
    anyway, or when AD enrollment is deliberately being removed).

    Robustness: all registry writes run with ErrorActionPreference Stop; the CEP entry and the
    AD policy row are verified by read-back (including detection of missing values); the
    summary object reports ACTUAL registry state for RootFlags and DefaultMarker, and per-gate
    outcome fields (EntryApplied, ADPolicyRow, DefaultChanged) show what each confirmation
    gate really did. Supports -WhatIf / -Confirm.

.PARAMETER Url
    Full CEP URI, e.g. https://pki.example.net/ejbca/msae/CEPService?alias - must match the
    server exactly (the SHA-1 subkey name and the clients' GetPolicies calls use it verbatim).
    The value is trimmed, must be an absolute http/https URI, and must not contain control
    characters. NOTE: if you rerun with a DIFFERENT URL, the old entry is not removed
    automatically - the script warns about same-PolicyID siblings; pass -ReplaceExisting to
    delete them.

.PARAMETER PolicyName
    The EJBCA MSAE alias "Policy Name". Becomes FriendlyName and (unless -PolicyId is given)
    the input to the PolicyID hash - it is hashed VERBATIM, so keep it identical to the EJBCA
    configuration; renaming it in EJBCA changes the PolicyID and orphans deployed entries.

.PARAMETER PolicyId
    Explicit PolicyID for non-EJBCA servers (must match the server's GetPolicies response).

.PARAMETER Location
    LocalMachine (default) | LocalUser | GPMachine | GPUser. See DESCRIPTION for the GP-hive
    tattooing warning. GPMachine, GPUser and LocalMachine require an elevated session
    (HKCU\Software\Policies is writable only by administrators).

.PARAMETER Authentication
    Client authentication type for the CEP endpoint: Anonymous (1), Kerberos (2, default,
    = "Windows integrated"), UsernamePassword (4), Certificate (8).

.PARAMETER Cost
    Priority; lower = preferred among endpoints sharing a PolicyID. Full DWORD range 1 to
    4294967295 (0xFFFFFFFF); default 0x7FFFFFFD = the dialog's "Priority: Default".
    Pass large values in decimal (a 0xFFFFFFFF literal is parsed by PowerShell as Int32 -1).

.PARAMETER NoAutoEnroll
    Leave "Enable for automatic enrollment and renewal" off (clears entry Flags bit 0x10).

.PARAMETER AllowUntrustedIssuer
    Equivalent to UNchecking "Require strong validation during enrollment" (sets entry Flags
    bit 0x20, PsfAllowUnTrustedCA).

.PARAMETER NoClientId
    Do not send the ClientId attribute in requests (clears entry Flags bit 0x4). Default is
    to include it, matching what the GPO editor writes (Flags = 0x14).

.PARAMETER SetAsDefault
    Mark this policy as the default enrollment policy: writes its PolicyID into the unnamed
    "(Default)" REG_SZ value on the PolicyServers key (what the dialog's Default checkbox
    does). Gets its own confirmation gate and -WhatIf line. Affects interactive enrollment
    preselection only; autoenrollment ignores it.

.PARAMETER ClearDefault
    Remove the unnamed "(Default)" marker so no policy is marked default.

.PARAMETER SkipADPolicy
    GP locations only: do not write the AD enrollment policy row. Background: once GP-based
    CEP configuration exists, the client stops auto-generating the built-in "Active
    Directory Enrollment Policy" and uses only the configured entries - without the LDAP:
    row the machine loses the AD enrollment policy and autoenrollment against AD-published
    templates stops. The script writes the row by default to prevent that; pass this switch
    only when that removal is intended.

.PARAMETER ReplaceExisting
    Remove sibling entries that share this PolicyID but have a different URL (typically stale
    entries from an earlier run with a typo'd or superseded URL). The AD policy row is never
    treated as a removable sibling. Without this switch the script only warns - multiple URLs
    per PolicyID is also the legitimate redundant-endpoint pattern. The cleanup (like the
    (Default) marker and the AD-row prerequisite) acts only when the entry that replaces the
    siblings is COMPLETE at that moment - written and verified by this run, or already present
    with URL and PolicyID as requested AND FriendlyName plus DWORD-typed Flags, AuthFlags and
    Cost; URL + PolicyID alone (what an interrupted write leaves) does not count.

.PARAMETER DisableUserConfigured
    GP locations only: set root Flags bit 0x4 so clients ignore user-configured policy
    servers. Preserved on later runs; clear again with -EnableUserConfigured.

.PARAMETER EnableUserConfigured
    GP locations only: clear root Flags bit 0x4.

.PARAMETER Remove
    Removal mode: delete the entry for -Url from the chosen location, clear the (Default)
    marker if it pointed at that entry's PolicyID and no remaining entry still serves that
    PolicyID, and list remaining entries. The summary's RootFlags field shows the actual root
    Flags value that remains. Other entries and the AD policy row are left alone.

.EXAMPLE
    .\Add-CertificateEnrollmentPolicyServerOffline.ps1 -Url 'https://pki.example.net/ejbca/msae/CEPService?alias' -PolicyName 'Example PKI Service' -WhatIf

    Preview every operation, including the computed subkey hash and PolicyID.

.EXAMPLE
    .\Add-CertificateEnrollmentPolicyServerOffline.ps1 -Url 'https://pki.example.net/ejbca/msae/CEPService?alias' -PolicyName 'Example PKI Service' -Location LocalUser -SetAsDefault

.EXAMPLE
    .\Add-CertificateEnrollmentPolicyServerOffline.ps1 -Url 'https://pki.example.net/ejbca/msae/CEPService?alias' -Location LocalUser -Remove

.NOTES
    Complete manual teardown of a GP-location deployment (per hive H = HKLM for GPMachine,
    HKCU for GPUser) requires removing: every entry subkey under
    H:\SOFTWARE\Policies\Microsoft\Cryptography\PolicyServers (including the AD policy row),
    the root Flags value, the unnamed (Default) value, and - if autoenrollment was configured
    separately - the H:\SOFTWARE\Policies\Microsoft\Cryptography\AutoEnrollment key. -Remove
    handles one entry and its (Default) marker; shared/root configuration is always left in
    place (the summary and warnings tell you what remains).
#>
[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Add')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'Remove',
    Justification = 'Mandatory discriminator for the Remove parameter set; consumed via $PSCmdlet.ParameterSetName, not by reading $Remove.')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Add')]
    [Parameter(Mandatory, ParameterSetName = 'Remove')]
    [string]$Url,

    [Parameter(Mandatory, ParameterSetName = 'Add')]
    [string]$PolicyName,

    [Parameter(ParameterSetName = 'Add')]
    [string]$PolicyId,

    [ValidateSet('LocalMachine','LocalUser','GPMachine','GPUser')]
    [string]$Location = 'LocalMachine',

    [Parameter(ParameterSetName = 'Add')]
    [ValidateSet('Anonymous','Kerberos','UsernamePassword','Certificate')]
    [string]$Authentication = 'Kerberos',

    [Parameter(ParameterSetName = 'Add')]
    [ValidateRange(1, 4294967295)]
    [long]$Cost = 0x7FFFFFFD,

    [Parameter(ParameterSetName = 'Add')] [switch]$NoAutoEnroll,
    [Parameter(ParameterSetName = 'Add')] [switch]$AllowUntrustedIssuer,
    [Parameter(ParameterSetName = 'Add')] [switch]$NoClientId,
    [Parameter(ParameterSetName = 'Add')] [switch]$SetAsDefault,
    [switch]$ClearDefault,
    [Parameter(ParameterSetName = 'Add')] [switch]$SkipADPolicy,
    [Parameter(ParameterSetName = 'Add')] [switch]$ReplaceExisting,
    [Parameter(ParameterSetName = 'Add')] [switch]$DisableUserConfigured,
    [Parameter(ParameterSetName = 'Add')] [switch]$EnableUserConfigured,

    [Parameter(Mandatory, ParameterSetName = 'Remove')]
    [switch]$Remove
)

$ErrorActionPreference = 'Stop'
$notes = New-Object System.Collections.Generic.List[string]

$hive = switch ($Location) {
    'GPMachine'    { 'HKLM:\SOFTWARE\Policies\Microsoft\Cryptography\PolicyServers' }
    'GPUser'       { 'HKCU:\Software\Policies\Microsoft\Cryptography\PolicyServers' }
    'LocalMachine' { 'HKLM:\SOFTWARE\Microsoft\Cryptography\PolicyServers' }
    'LocalUser'    { 'HKCU:\Software\Microsoft\Cryptography\PolicyServers' }
}
$isGP = $Location -like 'GP*'
$AD_KEY = '37c9dc30f207f27f61a2f7c3aed598a6e2920b54'   # SHA-1 of utf16le "ldap:"

# ---- parameter conflicts -------------------------------------------------------------------
if ($SetAsDefault -and $ClearDefault) { throw '-SetAsDefault and -ClearDefault are mutually exclusive.' }
if ($DisableUserConfigured -and $EnableUserConfigured) { throw '-DisableUserConfigured and -EnableUserConfigured are mutually exclusive.' }
if (($DisableUserConfigured -or $EnableUserConfigured) -and -not $isGP) {
    throw "-DisableUserConfigured/-EnableUserConfigured only exist in the Group Policy hive (root Flags); they have no effect for -Location $Location. Use a GP location."
}

# ---- privilege check (before any gate, so -WhatIf previews are honest) ---------------------
$needsElevation = $Location -in 'GPMachine','GPUser','LocalMachine'   # HKCU\Software\Policies is admin-writable only
if ($needsElevation) {
    $elevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
                ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $elevated) {
        if ($WhatIfPreference) { Write-Warning "Preview only: a real run with -Location $Location requires an elevated session." }
        else { throw "-Location $Location writes $hive and requires an elevated session." }
    }
}

# ---- input validation ----------------------------------------------------------------------
$Url = $Url.Trim()
if ($Url -match '[\x00-\x1F]') { throw 'Url contains control characters.' }
if ($PSCmdlet.ParameterSetName -eq 'Add') {
    $parsed = $null
    if (-not [System.Uri]::TryCreate($Url, [System.UriKind]::Absolute, [ref]$parsed) -or $parsed.Scheme -notin 'http','https') {
        throw "Url must be an absolute http/https URI. Got: '$Url'"
    }
    if ($PolicyName -match '[\x00-\x1F]') { throw 'PolicyName contains control characters.' }
    if ($PolicyName -ne $PolicyName.Trim()) {
        Write-Warning "PolicyName has leading/trailing whitespace. It is hashed VERBATIM - make sure this exactly matches the EJBCA Policy Name."
    }
}

# ---- derivations ---------------------------------------------------------------------------
$sha1 = [System.Security.Cryptography.SHA1]::Create()
$key  = -join ($sha1.ComputeHash([System.Text.Encoding]::Unicode.GetBytes($Url.ToLowerInvariant())) |
               ForEach-Object { $_.ToString('x2') })
$target = "$hive\$key"

# ---- registry path protection ---------------------------------------------------------------
# The writes below are privileged (an elevated session for three of the four locations) and land
# at a path the script builds, so every EXISTING key from the hive root down to the deepest
# existing component of the target must be (a) a real key, not a registry SYMBOLIC LINK - a link
# planted where 'PolicyServers' does not exist yet would carry an elevated first-time write to
# whatever protected key it points at (Get-Item/New-Item follow links silently; -LiteralPath only
# stops wildcards) - (b) owned by a trusted principal, and (c) free of write-class grants (create
# subkey, set value, delete, WRITE_DAC, WRITE_OWNER) to untrusted principals, who could otherwise
# plant such a link or swap the key between this check and the write. Trusted: SYSTEM,
# Administrators, TrustedInstaller, the running account and its Domain/Enterprise Admins, and the
# CREATOR OWNER / OWNER RIGHTS placeholders (they resolve to the running account for keys this
# script creates). The default ACLs of HKLM\SOFTWARE\Microsoft\Cryptography, HKLM\SOFTWARE\Policies
# and the user's own HKCU pass; a misdelegated parent is refused - fail closed, no override.
if (-not ('CepRegNative' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class CepRegNative {
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int RegOpenKeyExW(IntPtr hKey, string lpSubKey, uint ulOptions, uint samDesired, out IntPtr phkResult);
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int RegQueryValueExW(IntPtr hKey, string lpValueName, IntPtr lpReserved, out uint lpType, byte[] lpData, ref uint lpcbData);
    [DllImport("advapi32.dll")]
    public static extern int RegCloseKey(IntPtr hKey);
}
"@
}
function Test-RegistryKeyIsLink {
    # $true when the key at hive-relative $SubKey is a registry symbolic link. Opened with
    # REG_OPTION_OPEN_LINK (0x8) the handle is the LINK object itself, which carries the REG_LINK
    # value 'SymbolicLinkValue'; a real key opened the same way has no such value.
    param([IntPtr]$HiveHandle, [string]$SubKey)
    $h = [IntPtr]::Zero
    $rc = [CepRegNative]::RegOpenKeyExW($HiveHandle, $SubKey, 0x8, 0x1, [ref]$h)   # KEY_QUERY_VALUE
    if ($rc -ne 0) { throw "Registry key '$SubKey' could not be opened for the link check (Win32 error $rc)." }
    try {
        $type = [uint32]0; $cb = [uint32]0
        $q = [CepRegNative]::RegQueryValueExW($h, 'SymbolicLinkValue', [IntPtr]::Zero, [ref]$type, $null, [ref]$cb)
        return ($q -eq 0 -or $q -eq 234)   # ERROR_SUCCESS / ERROR_MORE_DATA: the value exists -> a link
    }
    finally { [void][CepRegNative]::RegCloseKey($h) }
}
function Assert-ProtectedRegistryPath {
    param([string]$Path)   # HKLM:\... or HKCU:\...; every EXISTING component is checked, root first
    $isCU = $Path -like 'HKCU:*'
    $hiveHandle = if ($isCU) { [IntPtr]::new(-2147483647) } else { [IntPtr]::new(-2147483646) }   # HKEY_CURRENT_USER / HKEY_LOCAL_MACHINE
    $hivePrefix = if ($isCU) { 'HKCU:' } else { 'HKLM:' }
    $rootName   = if ($isCU) { 'HKEY_CURRENT_USER' } else { 'HKEY_LOCAL_MACHINE' }
    $me = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $trusted = @{ 'S-1-5-18' = 1; 'S-1-5-32-544' = 1; 'S-1-3-0' = 1; 'S-1-3-4' = 1
                  'S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464' = 1 }
    $trusted[$me.User.Value] = 1
    if ($me.User.AccountDomainSid) {
        $trusted["$($me.User.AccountDomainSid.Value)-512"] = 1
        $trusted["$($me.User.AccountDomainSid.Value)-519"] = 1
    }
    $writeMask = ([int64][System.Security.AccessControl.RegistryRights]'CreateSubKey, SetValue, Delete, ChangePermissions, TakeOwnership') -bor 0x10000000 -bor 0x40000000   # + GENERIC_ALL, GENERIC_WRITE
    $sub = ''
    foreach ($p in @(($Path -replace '^HK(CU|LM):\\?', '') -split '\\' | Where-Object { $_ })) {
        $sub = if ($sub) { "$sub\$p" } else { $p }
        $full = "$hivePrefix\$sub"
        if (-not (Test-Path -LiteralPath $full)) { break }   # the rest does not exist yet: this run creates it under the parent just checked
        if (Test-RegistryKeyIsLink -HiveHandle $hiveHandle -SubKey $sub) {
            throw "Refusing to write: registry key '$full' is a SYMBOLIC LINK. A link on the path would carry this run's privileged writes to whatever key it points at. Remove the link (and find out who planted it) before rerunning."
        }
        # Get-Acl -LiteralPath is broken for registry keys on Windows PowerShell 5.1 (it reports the
        # key as missing); the provider-qualified -Path form works on both engines, and escaping the
        # key name keeps -Path literal in effect.
        $acl = Get-Acl -Path ("Registry::$rootName\" + [System.Management.Automation.WildcardPattern]::Escape($sub)) -ErrorAction Stop
        $owner = $acl.GetOwner([System.Security.Principal.SecurityIdentifier])
        if (-not $owner) { throw "The owner of registry key '$full' could not be read; refusing to write." }
        if (-not $trusted.ContainsKey($owner.Value)) {
            throw "Refusing to write: registry key '$full' is owned by untrusted principal $($owner.Value) (an owner can always re-permission and replace a key). Fix the key's ownership before rerunning."
        }
        $bad = @()
        foreach ($rule in $acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier])) {
            if ($rule.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow) { continue }
            $sid = $rule.IdentityReference.Value
            if ($trusted.ContainsKey($sid)) { continue }
            if ((([int64][int]$rule.RegistryRights) -band 0xFFFFFFFF) -band $writeMask) { $bad += $sid }
        }
        if ($bad.Count) {
            throw "Refusing to write: registry key '$full' grants write-class rights (create subkey / set value / delete / change permissions / take ownership) to untrusted principal(s) $(@($bad | Sort-Object -Unique) -join ', '), who could plant a symbolic link or swap the key while this run writes. Restrict the key's ACL before rerunning."
        }
    }
}
Assert-ProtectedRegistryPath -Path $target

function ConvertTo-DwordInt([long]$v) { [BitConverter]::ToInt32([BitConverter]::GetBytes([uint32]$v), 0) }
function Get-CepEntries {
    if (Test-Path -LiteralPath $hive) {
        Get-ChildItem -LiteralPath $hive | ForEach-Object {
            [pscustomobject]@{ Key = $_.PSChildName; URL = $_.GetValue('URL'); PolicyID = $_.GetValue('PolicyID') }
        }
    }
}
function Get-DefaultMarker { if (Test-Path -LiteralPath $hive) { (Get-Item -LiteralPath $hive).GetValue('') } }
function Remove-DefaultMarker {
    $rel  = $hive -replace '^HK(CU|LM):\\', ''
    $root = if ($hive -like 'HKCU:*') { [Microsoft.Win32.Registry]::CurrentUser } else { [Microsoft.Win32.Registry]::LocalMachine }
    $k = $root.OpenSubKey($rel, $true)
    if ($k) { try { $k.DeleteValue('', $false) } finally { $k.Close() } }
}
function Get-RootFlagsDisplay {
    if (-not $isGP) { return 'n/a (local location)' }
    if (Test-Path -LiteralPath $hive) {
        $v = (Get-Item -LiteralPath $hive).GetValue('Flags')
        if ($null -ne $v) { return '0x{0:X}' -f [int]$v }
    }
    return '(absent)'
}
function Write-CepEntry {
    param([string]$EntryPath, [hashtable]$Strings, [hashtable]$Dwords)
    if (-not (Test-Path -LiteralPath $EntryPath)) {
        $parent = Split-Path -Path $EntryPath -Parent
        if (-not (Test-Path -LiteralPath $parent)) { New-Item -Path $parent -Force -Confirm:$false | Out-Null }  # -Force only creates missing parents here (guarded)
        New-Item -Path $EntryPath -Confirm:$false | Out-Null   # leaf WITHOUT -Force: never recreate/wipe an existing key
    }
    # -Type String, always: Set-ItemProperty without -Type KEEPS an existing value's kind when the
    # conversion succeeds, so a PolicyID some earlier tool left as REG_DWORD would stay a DWORD while
    # the string compare below read it back as equal. -Confirm:$false on every cmdlet: inside an
    # already-approved action they must not raise their own prompts (a declined inner prompt would
    # return normally and leave the entry half-written).
    foreach ($n in $Strings.Keys) { Set-ItemProperty -LiteralPath $EntryPath -Name $n -Value ([string]$Strings[$n]) -Type String -Confirm:$false }
    foreach ($n in $Dwords.Keys)  { New-ItemProperty -LiteralPath $EntryPath -Name $n -Value $Dwords[$n] -PropertyType DWord -Force -Confirm:$false | Out-Null }
    $chk = Get-Item -LiteralPath $EntryPath
    $bad = @()
    foreach ($n in $Strings.Keys) {
        $got = $chk.GetValue($n)
        if ($null -eq $got -or [string]$got -cne [string]$Strings[$n] -or $chk.GetValueKind($n) -ne [Microsoft.Win32.RegistryValueKind]::String) { $bad += "$n (value or kind)" }
    }
    foreach ($n in $Dwords.Keys) {
        $got = $chk.GetValue($n)
        if ($null -eq $got -or [int]$got -ne [int]$Dwords[$n] -or $chk.GetValueKind($n) -ne [Microsoft.Win32.RegistryValueKind]::DWord) { $bad += "$n (value or kind)" }
    }
    if ($bad) { throw "Post-write verification failed for value(s): $($bad -join ', ') under $EntryPath" }
}

# ============================ REMOVE MODE ===================================================
if ($PSCmdlet.ParameterSetName -eq 'Remove') {
    $removedEntry = $false; $defaultCleared = $false
    $entries0 = @(Get-CepEntries)
    $mine = $entries0 | Where-Object { $_.Key -eq $key } | Select-Object -First 1
    $marker0 = Get-DefaultMarker
    if (-not $mine) {
        $notes.Add("No entry for this URL (key $key) under $hive - nothing to remove.")
    } else {
        $entryPid = $mine.PolicyID
        $markerMatches = $marker0 -and ("$marker0" -eq "$entryPid")
        $survivorServes = @($entries0 | Where-Object { $_.Key -ne $key -and "$($_.PolicyID)" -eq "$marker0" }).Count -gt 0
        if ($PSCmdlet.ShouldProcess($target, "Remove CEP entry (PolicyID=$entryPid)")) {
            # -Confirm:$false: the script's own prompt above IS the approval. A nested Remove-Item
            # prompt that the operator declined returned normally, and this run then reported a
            # removal that never happened - and cleared the marker of an entry that still existed.
            try { Remove-Item -LiteralPath $target -Recurse -Force -Confirm:$false } catch { throw "Failed to remove ${target}: $_" }
            if (Test-Path -LiteralPath $target) { throw "Removal of $target did not take effect (the key still exists)." }
            $removedEntry = $true
        }
        if ($markerMatches -and $survivorServes) {
            $notes.Add("(Default) marker kept: another entry still serves PolicyID $marker0 (redundant endpoint).")
        }
        if ($markerMatches -and -not $survivorServes -and ($removedEntry -or $WhatIfPreference)) {
            if ($PSCmdlet.ShouldProcess($hive, "Clear (Default) marker (would point at removed PolicyID $entryPid)")) {
                Remove-DefaultMarker; $defaultCleared = $true
            }
        }
    }
    if ($ClearDefault -and -not $defaultCleared -and (Get-DefaultMarker)) {
        if ($PSCmdlet.ShouldProcess($hive, 'Clear (Default) marker')) { Remove-DefaultMarker; $defaultCleared = $true }
    }
    $left = @(Get-CepEntries)
    if ($left.Count -gt 0) { $notes.Add("Remaining entries under ${hive}: $(($left | ForEach-Object { $_.URL }) -join ' | ')") }
    elseif ($isGP -and (Test-Path -LiteralPath $hive)) {
        $notes.Add('No entries remain, but the PolicyServers key (root Flags / (Default)) still exists - clients still treat GP CEP configuration as present. Delete the whole key to fully revert (see NOTES).')
    }
    foreach ($n in $notes) { Write-Warning $n }
    return [pscustomobject]@{
        Mode = 'Remove'; Location = $Location; Path = $target
        RemovedEntry = $removedEntry; DefaultCleared = $defaultCleared
        RootFlags = Get-RootFlagsDisplay; DefaultMarker = "$(Get-DefaultMarker)"; Notes = @($notes)
    }
}

# ============================ ADD MODE ======================================================
if (-not $PolicyId) {
    # EJBCA MSAE: PolicyID = Java String.hashCode() of the Policy Name (32-bit signed wrap).
    # NB: mask with the decimal literal - 0xFFFFFFFF parses as Int32 -1 in PowerShell.
    $h = [int64]0
    foreach ($c in $PolicyName.ToCharArray()) { $h = ($h * 31 + [int64]$c) -band 4294967295 }
    if ($h -ge 2147483648) { $h -= 4294967296 }
    $PolicyId = "$h"
}
$authFlags = @{ Anonymous = 1; Kerberos = 2; UsernamePassword = 4; Certificate = 8 }[$Authentication]
$flags = 0
if (-not $NoAutoEnroll)    { $flags = $flags -bor 0x10 }  # PsfAutoEnrollmentEnabled
if (-not $NoClientId)      { $flags = $flags -bor 0x4  }  # PsfUseClientId (GPO-editor default -> 0x14)
if ($AllowUntrustedIssuer) { $flags = $flags -bor 0x20 }  # PsfAllowUnTrustedCA

$entryApplied = $false; $adRow = 'n/a'; $defaultChanged = $false; $dupRemoved = @()

# ---- 0. prerequisite for the AD enrollment policy row (GP locations) - resolved BEFORE any write
# Once GP-based CEP configuration exists, the client stops synthesizing the AD enrollment policy;
# without the LDAP: row it is lost. So the domain objectGUID the row needs is resolved first: on a
# domain-joined machine a failure aborts the run with nothing written (it must not become a
# warning after the entry exists); on a workgroup machine there is no AD policy to lose, so the
# row is skipped with a warning as documented.
$adPid = $null
if ($isGP -and -not $SkipADPolicy) {
    try {
        $dn  = ([ADSI]'LDAP://RootDSE').defaultNamingContext.Value
        if (-not $dn) { throw 'RootDSE returned no defaultNamingContext' }
        $dom = [ADSI]("LDAP://$dn")
        $adPid = '{' + (New-Object Guid (, ([byte[]]$dom.Properties['objectGUID'][0]))).ToString().ToUpper() + '}'
    } catch {
        $lookupError = $_
        # Fail closed on the membership probe too: if it cannot be determined, assume joined.
        $joined = $true
        try { $joined = [bool](Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).PartOfDomain } catch { Write-Verbose "Domain membership probe failed ($_); assuming domain-joined." }
        if ($joined) {
            throw "Could not resolve the domain objectGUID needed for the AD Enrollment Policy row on this domain-joined machine ($lookupError). Nothing was written. Without that row the machine would LOSE the AD enrollment policy (autoenrollment against AD-published templates stops), so the run stops here: fix the lookup (connectivity, permissions), or pass -SkipADPolicy to omit the row deliberately."
        }
        Write-Warning "Workgroup machine (no domain objectGUID available): skipping the AD policy row - there is no AD enrollment policy to preserve here. ($lookupError)"
        $adRow = 'skipped (no domain)'
    }
}

# ---- 1. AD enrollment policy row FIRST (GP locations; prevents losing the AD default policy)
# Written before the CEP entry so the hive is never left in the state "CEP entry present, LDAP:
# row absent" - neither through a write failure (aborts before the entry exists) nor through a
# declined confirmation (the CEP step below refuses to run then).
# "Already present" means a COMPLETE, correct row: URL 'LDAP:' AND this domain's PolicyID - not
# merely the key (Write-CepEntry creates the key before its values, so an interrupted earlier run
# can leave an empty one that would not restore the AD enrollment policy).
function Test-RegEntryUsable {
    # A policy-server row a client can USE (the GPO script's Test-PolEntryUsable, against a live
    # key): URL and PolicyID as expected, FriendlyName present, and Flags / AuthFlags / Cost present
    # as DWORDs. Write-CepEntry writes the strings first, so an interrupted run leaves URL +
    # PolicyID and nothing else - not a row that may satisfy the AD-row prerequisite or the gate
    # that lets the (Default) marker be set and -ReplaceExisting delete the working siblings.
    param($Key, [string]$ExpectUrl, [string]$ExpectPolicyId)
    if (-not $Key) { return $false }
    if ("$($Key.GetValue('URL'))" -ne $ExpectUrl -or "$($Key.GetValue('PolicyID'))" -ne $ExpectPolicyId) { return $false }
    $names = @($Key.GetValueNames())
    if ($names -notcontains 'FriendlyName') { return $false }
    foreach ($n in 'Flags', 'AuthFlags', 'Cost') {
        if ($names -notcontains $n) { return $false }
        if ($Key.GetValueKind($n) -ne [Microsoft.Win32.RegistryValueKind]::DWord) { return $false }
    }
    return $true
}
$adRowPresent = $false
if ($isGP -and $adPid -and (Test-Path -LiteralPath "$hive\$AD_KEY")) {
    $adRowPresent = Test-RegEntryUsable -Key (Get-Item -LiteralPath "$hive\$AD_KEY") -ExpectUrl 'LDAP:' -ExpectPolicyId "$adPid"
}
if ($isGP) {
    if ($SkipADPolicy) { $adRow = 'skipped (-SkipADPolicy)' }
    elseif ($adPid) {
        $adTarget = "$hive\$AD_KEY"
        if ($PSCmdlet.ShouldProcess($adTarget, "Ensure AD Enrollment Policy row (URL=LDAP:, PolicyID=$adPid, Flags=0x14, Cost=0xFFFFFFFF)")) {
            try {
                Write-CepEntry -EntryPath $adTarget `
                    -Strings @{ URL = 'LDAP:'; PolicyID = $adPid; FriendlyName = 'Active Directory Enrollment Policy' } `
                    -Dwords  @{ Flags = 0x14; AuthFlags = 2; Cost = (ConvertTo-DwordInt 4294967295) }
            } catch { throw "AD policy row write to $adTarget failed: $_ (the CEP entry was NOT written)" }
            $adRow = 'applied'
        } else { $adRow = 'not run' }
    }
}

function New-AddSummary([string[]]$Removed) {
    # The Add-mode result object, from the ACTUAL registry state at the time it is built.
    [pscustomobject]@{
        Mode           = 'Add'
        Location       = $Location
        Path           = $target
        Url            = $Url
        PolicyID       = $PolicyId
        FriendlyName   = $PolicyName
        Flags          = '0x{0:X}' -f $flags
        Authentication = '{0} (0x{1:X})' -f $Authentication, $authFlags
        Cost           = '0x{0:X}' -f $Cost
        EntryApplied   = $entryApplied
        ADPolicyRow    = $adRow
        RootFlags      = Get-RootFlagsDisplay
        DefaultMarker  = "$(Get-DefaultMarker)"
        DefaultChanged = $defaultChanged
        DuplicatesRemoved = @($Removed)
        Notes          = @($notes)
    }
}

# ---- 2. the CEP entry ----------------------------------------------------------------------
# For the GP locations this depends on step 1: without an LDAP: row (just written, already
# present, deliberately omitted with -SkipADPolicy, or moot on a workgroup machine) the entry is
# NOT written and the Add workflow STOPS here - a declined AD-row prompt declines the CEP entry
# and everything that builds on it (root Flags, (Default) marker, sibling cleanup). -WhatIf
# previews all steps regardless.
$entryPreexisting = Test-Path -LiteralPath $target
$adRowSatisfied = (-not $isGP) -or $SkipADPolicy -or $adRow -in 'applied', 'skipped (no domain)' -or $adRowPresent
if (-not $adRowSatisfied -and -not $WhatIfPreference) {
    $notes.Add("CEP entry NOT written and the run stopped here: the AD Enrollment Policy row was declined and $hive carries none. Writing the entry without it would make this machine LOSE the AD enrollment policy (autoenrollment against AD-published templates stops). Nothing else was changed. Re-run and accept the AD row, or pass -SkipADPolicy to omit it deliberately.")
    foreach ($n in $notes) { Write-Warning $n }
    return New-AddSummary -Removed @()
}
$verb = if ($entryPreexisting) { 'Update' } else { 'Create' }
$action = "$verb CEP entry '{0}' (URL={1}, PolicyID={2}, Flags=0x{3:X}, AuthFlags=0x{4:X} {5}, Cost=0x{6:X})" -f `
          $PolicyName, $Url, $PolicyId, $flags, $authFlags, $Authentication, $Cost
if ($PSCmdlet.ShouldProcess($target, $action)) {
    try {
        Write-CepEntry -EntryPath $target `
            -Strings @{ URL = $Url; PolicyID = $PolicyId; FriendlyName = $PolicyName } `
            -Dwords  @{ Flags = $flags; AuthFlags = $authFlags; Cost = (ConvertTo-DwordInt $Cost) }
    } catch { throw "CEP entry write to $target failed (the key may be partially written - inspect it): $_" }
    $entryApplied = $true
}
# The (Default) marker and the removal of stale siblings point AT the entry, so they run only
# when an entry for THIS URL serving THIS PolicyID exists (written now, or already present from
# an earlier run with the same PolicyID). A declined CEP prompt must not leave a marker pointing
# at nothing, and an existing entry that still serves a DIFFERENT PolicyID (the declined update
# would have changed it) must not let the marker or the sibling cleanup act for the requested one.
# A pre-existing entry counts only when COMPLETE for this request - URL and PolicyID both as
# requested (an interrupted earlier write can leave a key with a PolicyID but no URL, which no
# client can use). Re-evaluated from the live registry right before the sibling cleanup.
function Test-EntryComplete {
    if (-not (Test-Path -LiteralPath $target)) { return $false }
    Test-RegEntryUsable -Key (Get-Item -LiteralPath $target) -ExpectUrl $Url -ExpectPolicyId "$PolicyId"
}
$entryExists = $entryApplied -or $WhatIfPreference -or (Test-EntryComplete)
# Root Flags live under the PolicyServers key ITSELF; writing them with no usable policy-server
# entry in this location (the CEP entry declined AND no AD Enrollment Policy row) activates GP CEP
# configuration with no server, so clients lose the AD enrollment policy. Gate on a real entry.
# (-SkipADPolicy alone does not count: it means the AD row was deliberately omitted, not that a
# server exists.)
$policyServerPresent = $entryExists -or $adRowPresent -or ($adRow -eq 'applied')

# ---- 3. root Flags (GP locations; DISABLE bits, preserved across runs) ---------------------
if ($isGP) {
    $existing = if (Test-Path -LiteralPath $hive) { (Get-Item -LiteralPath $hive).GetValue('Flags') } else { $null }
    # The kind matters as much as the number: a Flags left as REG_SZ by some other tool is not a
    # value the client reads, so it is rewritten as a DWORD even when its number is already right.
    $existingKind = if ($null -ne $existing) { try { (Get-Item -LiteralPath $hive).GetValueKind('Flags') } catch { $null } } else { $null }
    $newFlags = if ($null -ne $existing) { [int]$existing } else { 0 }
    if ($newFlags -band 0x2) {
        Write-Warning 'Existing root Flags had bit 0x2 set (clients IGNORE the GP-provided policy list). Clearing it.'
        $newFlags = $newFlags -band (-bnot 0x2)
    }
    if ($DisableUserConfigured) { $newFlags = $newFlags -bor 0x4 }
    if ($EnableUserConfigured)  { $newFlags = $newFlags -band (-bnot 0x4) }
    if (-not $policyServerPresent -and -not $WhatIfPreference) {
        $notes.Add("Root Flags NOT written: no usable policy-server entry exists (the CEP entry was declined and there is no AD Enrollment Policy row). Writing PolicyServers root values would activate GP CEP configuration with no server, and clients would lose the AD enrollment policy.")
    }
    elseif (($null -eq $existing) -or ([int]$existing -ne $newFlags) -or ($existingKind -ne [Microsoft.Win32.RegistryValueKind]::DWord)) {
        $from = if ($null -ne $existing) { '0x{0:X}' -f [int]$existing } else { '(absent)' }
        if ($PSCmdlet.ShouldProcess($hive, ('Set root Flags {0} -> 0x{1:X} (disable bits: 0x2 ignore GP list, 0x4 ignore user-configured)' -f $from, $newFlags))) {
            if (-not (Test-Path -LiteralPath $hive)) { New-Item -Path $hive -Force -Confirm:$false | Out-Null }
            New-ItemProperty -LiteralPath $hive -Name Flags -Value $newFlags -PropertyType DWord -Force -Confirm:$false | Out-Null
            if ((Get-Item -LiteralPath $hive).GetValueKind('Flags') -ne [Microsoft.Win32.RegistryValueKind]::DWord) { throw "Root Flags under $hive did not end up as a DWORD." }
        }
    }
}

# ---- 4. (Default) marker -------------------------------------------------------------------
if ($SetAsDefault) {
    if (-not $entryExists) {
        $notes.Add("(Default) marker NOT set: the CEP entry was declined and does not exist under $hive, so the marker would point at nothing.")
    }
    elseif ($PSCmdlet.ShouldProcess($hive, "Set (Default) marker = $PolicyId (default enrollment policy = '$PolicyName')")) {
        if (-not (Test-Path -LiteralPath $hive)) { New-Item -Path $hive -Force -Confirm:$false | Out-Null }
        Set-ItemProperty -LiteralPath $hive -Name '(default)' -Value ([string]$PolicyId) -Type String -Confirm:$false
        if ((Get-Item -LiteralPath $hive).GetValueKind('') -ne [Microsoft.Win32.RegistryValueKind]::String) { throw "The (Default) marker under $hive did not end up as a REG_SZ." }
        $defaultChanged = $true
    }
}
if ($ClearDefault -and (Get-DefaultMarker)) {
    if ($PSCmdlet.ShouldProcess($hive, 'Clear (Default) marker')) { Remove-DefaultMarker; $defaultChanged = $true }
}

# ---- 5. consistency checks: stale duplicates and orphaned default --------------------------
$entries = @(Get-CepEntries)
# Sibling removal is judged on the registry as it is NOW: the replacing entry must be complete at
# this moment - "written earlier this run" does not count, another writer may have removed or
# changed it since - or the "superseded" siblings are the only working endpoints and must stay.
$entryExists = $WhatIfPreference -or (Test-EntryComplete)
$dups = @($entries | Where-Object { "$($_.PolicyID)" -eq "$PolicyId" -and $_.Key -ne $key -and $_.Key -ne $AD_KEY })
foreach ($d in $dups) {
    if ($ReplaceExisting -and -not $entryExists) {
        $notes.Add("Superseded entry '$($d.URL)' (key $($d.Key)) NOT removed: the replacing CEP entry was declined and does not exist, so removing it would leave no endpoint for PolicyID $PolicyId.")
    }
    elseif ($ReplaceExisting) {
        if ($PSCmdlet.ShouldProcess("$hive\$($d.Key)", "Remove superseded entry with same PolicyID (URL=$($d.URL))")) {
            # Re-validate AFTER the approval (a -Confirm prompt can stay open indefinitely): the
            # replacement must still be complete, and this sibling must still serve the requested
            # PolicyID under a different URL, in the live registry right now.
            if (-not (Test-EntryComplete)) {
                $notes.Add("Superseded entry '$($d.URL)' (key $($d.Key)) NOT removed: the replacing CEP entry is no longer complete (changed or removed while the prompt was open).")
                continue
            }
            $sib = Get-Item -LiteralPath "$hive\$($d.Key)" -ErrorAction SilentlyContinue
            if (-not $sib -or "$($sib.GetValue('PolicyID'))" -ne "$PolicyId" -or "$($sib.GetValue('URL'))" -eq $Url) {
                $notes.Add("Entry at key $($d.Key) NOT removed: it no longer serves PolicyID $PolicyId under a different URL (changed or removed while the prompt was open).")
                continue
            }
            Remove-Item -LiteralPath "$hive\$($d.Key)" -Recurse -Force -Confirm:$false
            if (Test-Path -LiteralPath "$hive\$($d.Key)") { throw "Removal of superseded entry $($d.Key) did not take effect (the key still exists)." }
            $dupRemoved += $d.URL
        }
    } else {
        $notes.Add("Another entry shares PolicyID $PolicyId with a different URL: '$($d.URL)' (key $($d.Key)). If that is a stale/typo entry rerun with -ReplaceExisting; if it is an intended redundant endpoint, ignore this.")
    }
}
$marker = Get-DefaultMarker
if ($marker) {
    $entries = @(Get-CepEntries)
    if (-not ($entries | Where-Object { "$($_.PolicyID)" -eq "$marker" })) {
        $notes.Add("The (Default) marker points at PolicyID '$marker', which matches NO entry - interactive enrollment has no valid default. Fix with -SetAsDefault on the right policy or -ClearDefault.")
    }
}
foreach ($n in $notes) { Write-Warning $n }

New-AddSummary -Removed $dupRemoved
