<#PSScriptInfo
.VERSION 1.0.8
.GUID 61adf5d1-6eb5-4f41-8670-e9da72134570
.AUTHOR Sveinung Svea
.PROJECTURI https://github.com/TheOmnilord/ADCS
.LICENSEURI https://github.com/TheOmnilord/ADCS/blob/main/LICENSE
.TAGS ADCS PKI CertificateServices
.RELEASENOTES
1.0.8 - The root Flags of a GP location are converted by a new helper (ConvertTo-RootFlagsValue) instead of a bare [int] cast. A DWORD is used as it is, and a non-empty REG_SZ is converted with the same [int] cast as before (a decimal, a 0x hex number, a sign, leading zeros), so every numeric string the script accepted is still repaired to a DWORD. A value that cannot be converted (text such as 'abc', an empty string, REG_BINARY, REG_MULTI_SZ) is refused BEFORE the AD Enrollment Policy row and the CEP entry are written, with a message that names the hive, the kind and the value; previously the cast ran after both writes and killed the run with a raw conversion error and no summary. A value that turns unusable while the run is in progress is skipped with a note instead of an error. The RootFlags field of the summary shows '(unusable: ...)' for such a value instead of throwing. A code comment at the RegQueryValueExW call explains why $null is a real null for its byte[] parameter and why ERROR_MORE_DATA cannot occur with a null buffer. The Add summary gains an EntryAction field (Created, Updated, Declined or None) that states what happened to the CEP entry key, because EntryApplied is true for a new key AND for a rewrite of a key that already existed; a caller that cleans up only keys the run created (the Lab tests) needs the distinction. The Add summary also gains a BaseKeyCreated field that is true only when this run itself created the PolicyServers key of the location (the base key is created by one helper, New-BaseKey, so a key that another writer created between the existence check and the write is never reported as created); the same caller needs this evidence because a base key that was absent before the run may have been created by another writer since; both EntryAction and BaseKeyCreated take their evidence from the disposition (REG_CREATED_NEW_KEY or REG_OPENED_EXISTING_KEY) that RegCreateKeyExW returns for the one call that creates or opens the key - the base key, the CEP entry and the AD row are all created that way now, never with New-Item - because Test-Path followed by New-Item is two operations and the provider's own CreateSubKey opens an existing key, so neither proved creation when another writer created the key in between. The native create requests KEY_READ (0x20019), not KEY_READ | KEY_WRITE: opening an existing key needs no write right, and creating a missing key needs KEY_CREATE_SUB_KEY on its parent (which New-Item needed too), so an update of an existing entry under a base key that the caller can read but not write succeeds as it did before 1.0.8. The native helper type is named CepRegNative2, so a session that already loaded the released CepRegNative (which has no RegCreateKeyExW) still loads the expanded type instead of failing on a missing method. After the native creates and BEFORE any value write, Write-CepEntry runs Assert-ProtectedRegistryPath on the entry path again (every component exists now, so the link, owner and ACL of the base key and the entry are all checked), and New-BaseKey runs it on the base key whenever the disposition says the key already existed; a base key that another writer created after the preflight, or an entry name that an untrusted principal created as a symbolic link beneath it, is therefore refused before the values are written through it.
1.0.7 - Help text only: the comment-based help is rewritten to the repository writing style (STYLE.md, derived from ASD-STE100 Simplified Technical English) - short sentences, active voice, no figurative language, acronyms defined, a CAUTION line on -ReplaceExisting and -Remove; every fact, condition and default is kept; no code change
1.0.6 - Write-CepEntry validates the specific key it writes (Assert-ProtectedRegistryPath), so the AD Enrollment Policy row - a sibling leaf of the CEP entry - is no longer written to a delegated or symlinked key that escaped the up-front preflight of the CEP entry alone; the root-Flags gate recognises any pre-existing USABLE policy-server entry (a complete row, not an incomplete fragment) in the location, not only the requested entry and the AD row; the -ReplaceExisting sibling cleanup validates each sibling's registry path before reading or removing it, so a sibling that is a symbolic link cannot make the recursive delete destroy a key in another location; -Remove and -ReplaceExisting refuse to recursively delete a CEP entry that has subkeys (a CEP entry is a leaf by design), closing a path where a recursive delete could follow a registry symbolic link planted beneath a delegated descendant; Assert-ProtectedRegistryPath decides each component's existence with a native REG_OPTION_OPEN_LINK probe instead of Test-Path, so a DANGLING symbolic link (target absent, which Test-Path reports as not-found) can no longer let the walk break before the link check and leave the link to be retargeted at a protected key before the write
1.0.5 - Root Flags are no longer written when no usable policy-server entry exists in a GP location (the CEP entry declined AND no AD Enrollment Policy row): writing PolicyServers root values there activated GP CEP configuration with no server, so clients lost the AD enrollment policy
1.0.4 - Every existing key on the target path is checked before any write: a registry symbolic link, an untrusted owner or write-class rights for an untrusted principal refuse the run (a link planted where PolicyServers did not exist yet would have carried an elevated first-time write to whatever key it pointed at); string values and the (Default) marker are written as REG_SZ explicitly and every value's KIND is verified (Set-ItemProperty kept an existing wrong kind and the string compare accepted it; a REG_SZ root Flags with the right number is now repaired to a DWORD); the cmdlets inside an approved action pass -Confirm:$false and a removal is verified before it is reported (a declined nested Remove-Item prompt reported a removal that never happened and could clear the marker of an entry that still existed)
1.0.3 - Help text only: -ReplaceExisting documents the complete-row requirement; no code change
1.0.2 - A pre-existing AD-policy row or CEP entry counts as complete only with URL, PolicyID, FriendlyName and DWORD Flags/AuthFlags/Cost present (URL + PolicyID alone is what an interrupted write leaves), for the prerequisite, the (Default) marker and -ReplaceExisting alike
1.0.1 - For the GP locations the domain objectGUID for the AD Enrollment Policy row is resolved BEFORE the CEP entry is written; on a domain-joined machine a lookup failure now aborts the run with nothing written (previously the entry was written first and the failure became a warning, leaving a GP configuration that removes the AD enrollment policy); on a workgroup machine the row is still skipped with a warning
1.0.0 - Initial release
#>

<#
.SYNOPSIS
    Registers or removes a Certificate Enrollment Policy (CEP) server in the registry offline,
    without contact with the policy server.

.DESCRIPTION
    The script writes the same registry values that the "Certificate Services Client -
    Certificate Enrollment Policy" dialog writes. The dialog calls the MS-XCEP GetPolicies
    endpoint of the server. The script computes everything locally instead and never calls
    that endpoint. It makes no "Validate Server" round-trip and no other contact with the
    policy server. The script derives two values:
      * Subkey name: the SHA-1 hash of the UTF-16LE bytes of the URL, after the script converts
        the URL to lowercase with the invariant culture.
        X509Enrollment.CX509PolicyServerUrl::UpdateRegistry uses the same derivation.
      * PolicyID: the Java String.hashCode() of the policy name. This is the EJBCA MSAE
        behavior: the "Policy Name" of the alias is both the friendly name and the input to
        the hash. Other CEP products use other values. For example, the CEP web service of
        Microsoft returns a GUID. For those products, pass -PolicyId.

    Locations:
      * LocalMachine and LocalUser are the "user configured" stores under Software\Microsoft.
        LocalMachine is the default. The "Manage enrollment policies" dialog of certlm.msc and
        certmgr.msc manages these stores.
      * GPMachine and GPUser are the Group Policy (GP) hives under Software\Policies. Read the
        warning below before you use them on a domain member.

    WARNING: on a domain member, values that the script writes directly into the GP hives are
    pseudo-policy that no Group Policy Object (GPO) backs. The name of this effect is
    "tattooing". The values appear in no Group Policy Management Console (GPMC), Resultant Set
    of Policy (RSoP) or gpresult report. The gpupdate command neither restores nor removes
    them. The certificate Microsoft Management Console (MMC) snap-in shows them read-only. A
    real GPO that later manages the same key overwrites the shared values without a message,
    and the entry subkey remains unmanaged.

    On a domain member, use the companion script Add-CertificateEnrollmentPolicyServerToGpo.ps1
    instead. The GP locations of this script are intended for standalone or workgroup machines
    and for lab work.

    For the GP locations, the script also maintains the Flags value on the PolicyServers root
    key. The root Flags are DISABLE bits, named EnrollmentPolicyFlags. Bit 0x2 makes clients
    ignore the whole list that GP provides. The script never sets bit 0x2. When the script
    finds bit 0x2 set, it clears the bit and writes a warning. Bit 0x4 makes clients ignore
    user-configured servers.

    The script preserves the existing bits. A later run without -DisableUserConfigured does
    NOT clear a bit 0x4 that an earlier run set. Use -EnableUserConfigured to clear it.

    For the GP locations, the script by default also makes sure that the built-in Active
    Directory (AD) enrollment policy row exists. The row has the URL "LDAP:" and the subkey
    name 37c9dc30f207f27f61a2f7c3aed598a6e2920b54. Its PolicyID is the objectGUID of the
    domain object, and its Cost is 0xFFFFFFFF. The row is required because a GP-based CEP
    configuration suppresses the AD default policy that the client synthesizes on its own.
    Without this row, the machine loses the AD enrollment policy, and Auto-Enrollment against
    AD-published certificate templates stops.

    Pass -SkipADPolicy to opt out. For example, pass it on a workgroup machine, or when you
    intend to remove AD enrollment. On a workgroup machine, the lookup fails and the script
    skips the row with a warning in any case.

    Robustness: the script runs every registry write with ErrorActionPreference Stop. It reads
    back the CEP entry and the AD policy row to verify them, and this read-back detects missing
    values. The summary object reports the ACTUAL registry state in RootFlags and DefaultMarker.
    The outcome fields EntryApplied, ADPolicyRow and DefaultChanged show what each confirmation
    gate did. The script supports -WhatIf and -Confirm.

    Before the first write, the script checks every existing key from the hive root down to the
    target key. A registry symbolic link, an untrusted owner, or write-class rights for an
    untrusted principal make the script refuse the run. After the script creates the
    PolicyServers key and the entry key, and before it writes a value, the script checks the
    entry path again. The script refuses a link or a delegated key that another writer created
    after the first check, before a write goes through it. The script also checks the
    PolicyServers key again when its create call reports that the key already existed.

    The summary field EntryAction states what the script did to the key of the CEP entry. It
    has one of four values:
      * Created: the script created the key and wrote its values.
      * Updated: the key existed when the script wrote it, and the script rewrote its values.
      * Declined: the confirmation gate or -WhatIf declined the write.
      * None: the run stopped before the entry step.

    The summary field BaseKeyCreated is $true only when this run created the PolicyServers key
    of the location. In every other case it is $false. A caller that removes only the keys it
    created can use this field together with EntryAction.

    The evidence for both fields is the disposition that the Windows function RegCreateKeyExW
    returns. The script creates each key with one such call. The call reports whether it created
    the key (REG_CREATED_NEW_KEY) or opened a key that already existed (REG_OPENED_EXISTING_KEY).
    The script does not use an existence check followed by a create as evidence. Another writer
    can create the key between those two operations, and the script would then report a foreign
    key as created.

.PARAMETER Url
    The full CEP URI (for example https://pki.example.net/ejbca/msae/CEPService?alias). The
    value must match the server exactly. The SHA-1 subkey name and the GetPolicies calls of
    the clients use it verbatim. The script trims the value. The value must be an absolute
    http or https URI and must not contain control characters.

    NOTE: when you rerun with a DIFFERENT URL, the script does not remove the old entry
    automatically. The script warns about siblings with the same PolicyID. Pass the switch
    -ReplaceExisting to delete them.

.PARAMETER PolicyName
    The "Policy Name" of the EJBCA MSAE alias. The script writes it as the FriendlyName. Unless
    you pass -PolicyId, the script also hashes it to get the PolicyID. The script hashes the
    name VERBATIM, so keep it identical to the EJBCA configuration. When you rename the alias
    in EJBCA, the PolicyID changes and the deployed entries become orphans.

.PARAMETER PolicyId
    An explicit PolicyID for servers that are not EJBCA. The value must match the GetPolicies
    response of the server.

.PARAMETER Location
    One of LocalMachine, LocalUser, GPMachine and GPUser. The default is LocalMachine. See the
    DESCRIPTION for the tattooing warning about the Group Policy (GP) hives. GPMachine, GPUser
    and LocalMachine require an elevated session. Only administrators can write
    HKCU\Software\Policies.

.PARAMETER Authentication
    The client authentication type for the CEP endpoint: Anonymous (1), Kerberos (2),
    UsernamePassword (4) or Certificate (8). The default is Kerberos, which the dialog names
    "Windows integrated".

.PARAMETER Cost
    The priority of the endpoint. Among endpoints that share a PolicyID, clients prefer the
    lower Cost. The full DWORD range 1 to 4294967295 (0xFFFFFFFF) is valid. The default is
    0x7FFFFFFD, which the dialog names "Priority: Default". Pass a large value in decimal:
    PowerShell parses a 0xFFFFFFFF literal as the Int32 value -1.

.PARAMETER NoAutoEnroll
    Leaves the dialog option "Enable for automatic enrollment and renewal" off. The script
    clears bit 0x10 of the entry Flags value.

.PARAMETER AllowUntrustedIssuer
    Equivalent to clearing the dialog option "Require strong validation during enrollment".
    The script sets bit 0x20 of the entry Flags value, PsfAllowUnTrustedCA.

.PARAMETER NoClientId
    Makes clients omit the ClientId attribute from their requests. The script clears bit 0x4
    of the entry Flags value. By default the script includes the attribute, which matches the
    Flags = 0x14 that the Group Policy Object (GPO) editor writes.

.PARAMETER SetAsDefault
    Marks this policy as the default enrollment policy. The script writes its PolicyID into
    the unnamed "(Default)" REG_SZ value on the PolicyServers key. This is what the Default
    checkbox of the dialog does. The switch gets its own confirmation gate and its own -WhatIf
    line. It affects only the preselection for interactive enrollment. Auto-Enrollment ignores
    it.

.PARAMETER ClearDefault
    Removes the unnamed "(Default)" marker, so that no policy is marked as the default.

.PARAMETER SkipADPolicy
    Only for the GP locations GPMachine and GPUser: the script does not write the Active
    Directory (AD) enrollment policy row. Background: once a GP-based CEP configuration
    exists, the client stops generating the built-in "Active Directory Enrollment Policy" and
    uses only the configured entries. Without the LDAP: row, the machine loses the AD
    enrollment policy, and Auto-Enrollment against AD-published certificate templates stops.
    The script writes the row by default to prevent that loss. Pass this switch only when you
    intend that removal.

.PARAMETER ReplaceExisting
    Removes sibling entries that share this PolicyID but have a different URL. Such siblings
    are typically stale entries from an earlier run with a mistyped or superseded URL. The
    script never treats the AD policy row as a removable sibling. Without this switch, the
    script only warns, because several URLs per PolicyID is also the legitimate pattern for
    redundant endpoints.

    CAUTION: the script deletes every sibling entry with this PolicyID and a different URL,
    including a working redundant endpoint.

    The cleanup acts only when the entry that replaces the siblings is COMPLETE at that
    moment. The (Default) marker and the AD-row prerequisite apply the same completeness rule.
    The entry is complete in one of two cases:
      * This run wrote and verified the entry.
      * The entry was already present with the requested URL and PolicyID, and also with
        FriendlyName and DWORD-typed Flags, AuthFlags and Cost.
    URL and PolicyID alone, which is what an interrupted write leaves, does not count.

.PARAMETER DisableUserConfigured
    Only for the GP locations GPMachine and GPUser: sets bit 0x4 of the root Flags, so that
    clients ignore user-configured policy servers. Later runs preserve the bit. Clear it again
    with -EnableUserConfigured.

.PARAMETER EnableUserConfigured
    Only for the GP locations GPMachine and GPUser: clears bit 0x4 of the root Flags.

.PARAMETER Remove
    Removal mode. The script deletes the entry for -Url from the chosen location. It clears
    the (Default) marker when the marker pointed at the PolicyID of that entry and no
    remaining entry still serves that PolicyID. It lists the remaining entries. The RootFlags
    field of the summary shows the actual root Flags value that remains. The script leaves
    other entries and the AD policy row alone.

    CAUTION: the script deletes the entry for -Url and can clear the (Default) marker of the
    chosen location.

.EXAMPLE
    .\Add-CertificateEnrollmentPolicyServerOffline.ps1 -Url 'https://pki.example.net/ejbca/msae/CEPService?alias' -PolicyName 'Example PKI Service' -WhatIf

    Previews every operation, including the computed subkey hash and the PolicyID. The script
    writes nothing.

.EXAMPLE
    .\Add-CertificateEnrollmentPolicyServerOffline.ps1 -Url 'https://pki.example.net/ejbca/msae/CEPService?alias' -PolicyName 'Example PKI Service' -Location LocalUser -SetAsDefault

    Registers the server in the LocalUser store and marks it as the default enrollment policy.

.EXAMPLE
    .\Add-CertificateEnrollmentPolicyServerOffline.ps1 -Url 'https://pki.example.net/ejbca/msae/CEPService?alias' -Location LocalUser -Remove

    Removes the entry for this URL from the LocalUser store.

.NOTES
    A complete manual teardown of a deployment in a Group Policy (GP) location requires that
    you remove the items below. H is the hive: HKLM for GPMachine, HKCU for GPUser.
      * Every entry subkey under H:\SOFTWARE\Policies\Microsoft\Cryptography\PolicyServers,
        including the Active Directory (AD) policy row.
      * The root Flags value.
      * The unnamed (Default) value.
      * The H:\SOFTWARE\Policies\Microsoft\Cryptography\AutoEnrollment key, when
        Auto-Enrollment was configured separately.
    The -Remove switch handles one entry and its (Default) marker. It always leaves the shared
    root configuration in place. The summary and the warnings tell you what remains.
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
# The type is named CepRegNative2: a session that ran the released script already holds a type
# named CepRegNative without RegCreateKeyExW and RegDeleteKeyW, and a type cannot be redefined in
# a running process. The guard on the new name loads this definition next to the old one.
if (-not ('CepRegNative2' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class CepRegNative2 {
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int RegOpenKeyExW(IntPtr hKey, string lpSubKey, uint ulOptions, uint samDesired, out IntPtr phkResult);
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int RegQueryValueExW(IntPtr hKey, string lpValueName, IntPtr lpReserved, out uint lpType, byte[] lpData, ref uint lpcbData);
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int RegCreateKeyExW(IntPtr hKey, string lpSubKey, int Reserved, string lpClass, uint dwOptions, uint samDesired, IntPtr lpSecurityAttributes, out IntPtr phkResult, out uint lpdwDisposition);
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int RegDeleteKeyW(IntPtr hKey, string lpSubKey);
    [DllImport("advapi32.dll")]
    public static extern int RegCloseKey(IntPtr hKey);
}
"@
}
function New-RegistryKeyNative {
    # Creates the key at $Path (HKCU:\... or HKLM:\...) with RegCreateKeyExW and returns $true
    # only when THIS call created it. The API creates the missing ancestors and reports in its
    # disposition whether it created the key (1, REG_CREATED_NEW_KEY) or opened one that already
    # existed (2, REG_OPENED_EXISTING_KEY). That disposition is the creation evidence the script
    # reports: Test-Path followed by New-Item is two operations, and the provider itself checks
    # existence and then calls CreateSubKey, which opens an existing key. Another writer can create
    # the key between those two steps, and neither form can tell. dwOptions 0 = a non-volatile key;
    # samDesired 0x20019 = KEY_READ, never a write right: the handle is closed at once and the
    # provider cmdlets write the values. Opening an existing key with KEY_READ needs no write right
    # on it, so an update of an entry under a base key the caller can read but not write succeeds
    # (as it did with New-Item). Creating a missing key needs KEY_CREATE_SUB_KEY on its parent,
    # which New-Item needed too. A non-zero return code throws with the Win32 code.
    param([string]$Path)
    $isCU = $Path -like 'HKCU:*'
    $hiveHandle = if ($isCU) { [IntPtr]::new(-2147483647) } else { [IntPtr]::new(-2147483646) }   # HKEY_CURRENT_USER / HKEY_LOCAL_MACHINE
    $rel = ($Path -replace '^HK(CU|LM):\\?', '').TrimEnd('\')
    $h = [IntPtr]::Zero; $disp = [uint32]0
    $rc = [CepRegNative2]::RegCreateKeyExW($hiveHandle, $rel, 0, $null, 0, 0x20019, [IntPtr]::Zero, [ref]$h, [ref]$disp)
    if ($rc -ne 0) { throw "Registry key '$Path' could not be created or opened (Win32 error $rc)." }
    [void][CepRegNative2]::RegCloseKey($h)
    return ($disp -eq 1)
}
function Test-RegistryKeyIsLink {
    # $true when the key at hive-relative $SubKey is a registry symbolic link. Opened with
    # REG_OPTION_OPEN_LINK (0x8) the handle is the LINK object itself, which carries the REG_LINK
    # value 'SymbolicLinkValue'; a real key opened the same way has no such value.
    param([IntPtr]$HiveHandle, [string]$SubKey)
    $h = [IntPtr]::Zero
    $rc = [CepRegNative2]::RegOpenKeyExW($HiveHandle, $SubKey, 0x8, 0x1, [ref]$h)   # KEY_QUERY_VALUE
    if ($rc -ne 0) { throw "Registry key '$SubKey' could not be opened for the link check (Win32 error $rc)." }
    try {
        $type = [uint32]0; $cb = [uint32]0
        # $null for the byte[] lpData parameter is a real null on both engines (a [string]
        # parameter would get "" instead). With a null buffer the API only reports the size, so
        # ERROR_MORE_DATA (234) cannot occur here. The check stays for a caller that passes a buffer.
        $q = [CepRegNative2]::RegQueryValueExW($h, 'SymbolicLinkValue', [IntPtr]::Zero, [ref]$type, $null, [ref]$cb)
        return ($q -eq 0 -or $q -eq 234)   # ERROR_SUCCESS / ERROR_MORE_DATA: the value exists -> a link
    }
    finally { [void][CepRegNative2]::RegCloseKey($h) }
}
function Test-RegistryComponentExists {
    # $true when a key OR a symbolic link exists at hive-relative $SubKey. Opened with
    # REG_OPTION_OPEN_LINK (0x8) the call targets the component ITSELF, so a link is seen as present
    # regardless of whether its target exists - unlike Test-Path, which follows the link and reports a
    # DANGLING link (target absent) as 'not found'. Fails closed on any error other than a genuine
    # ERROR_FILE_NOT_FOUND so an unreadable component is never mistaken for an absent one.
    param([IntPtr]$HiveHandle, [string]$SubKey)
    $h = [IntPtr]::Zero
    $rc = [CepRegNative2]::RegOpenKeyExW($HiveHandle, $SubKey, 0x8, 0x1, [ref]$h)   # REG_OPTION_OPEN_LINK | KEY_QUERY_VALUE
    if ($rc -eq 0) { [void][CepRegNative2]::RegCloseKey($h); return $true }
    if ($rc -eq 2) { return $false }   # ERROR_FILE_NOT_FOUND: truly absent
    throw "Registry key '$SubKey' could not be probed for existence (Win32 error $rc)."
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
        # Decide existence with the native link-aware open, NOT Test-Path: the provider follows a
        # symbolic link to its target, so a DANGLING link (target absent) reads as 'not found' and would
        # let the loop break before the link/owner/ACL checks below - after which the link could be
        # retargeted at a protected key and take this run's privileged write. The native probe opens the
        # component itself, so a link is seen as present and falls through to the link check that rejects it.
        if (-not (Test-RegistryComponentExists -HiveHandle $hiveHandle -SubKey $sub)) { break }   # the rest does not exist yet: this run creates it under the parent just checked
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
function ConvertTo-RootFlagsValue {
    # Converts the raw root Flags value of a PolicyServers key to the [int] the script works with.
    # Returns an object with Value and Reason. Value is the [int], or $null when the raw value is
    # unusable; Reason then names the kind and the value. A missing value (Raw $null) gives Value
    # $null and Reason $null. A DWORD is taken as it is. A non-empty REG_SZ is converted with the
    # same [int] cast the script always used (a decimal, a 0x hex number, a sign, leading zeros),
    # so the 1.0.4 repair of a numeric REG_SZ to a DWORD still runs for every string it accepted.
    # Text, an empty string, REG_BINARY and REG_MULTI_SZ are unusable. [int]'abc' throws on both
    # engines, and a raw cast after the entry writes killed the run with no summary. This helper
    # never throws.
    param($Raw, $Kind)
    if ($null -eq $Raw) { return [pscustomobject]@{ Value = $null; Reason = $null } }
    $kindName = if ($null -ne $Kind) { "$Kind" } else { $Raw.GetType().Name }
    $value = $null
    if ($Raw -is [string]) {
        $s = $Raw.Trim()
        if ($s.Length -gt 0) { try { $value = [int]$s } catch { $value = $null } }
    }
    elseif ($Raw -is [System.ValueType]) { try { $value = [int]$Raw } catch { $value = $null } }
    if ($null -ne $value) { return [pscustomobject]@{ Value = $value; Reason = $null } }
    $display = if ($Raw -is [Array]) { @($Raw | ForEach-Object { "$_" }) -join ',' } else { "$Raw" }
    if ($display.Length -gt 64) { $display = $display.Substring(0, 64) + '...' }
    [pscustomobject]@{ Value = $null; Reason = "the value is kind $kindName ('$display'), not a number that the script can convert to a DWORD" }
}
function Get-RootFlagsValue {
    # Reads the raw root Flags value and its kind from the PolicyServers key of this location and
    # returns the ConvertTo-RootFlagsValue result with the raw value added. Never throws on a bad value.
    $raw = $null; $kind = $null
    if (Test-Path -LiteralPath $hive) {
        $k = Get-Item -LiteralPath $hive
        $raw = $k.GetValue('Flags')
        if ($null -ne $raw) { try { $kind = $k.GetValueKind('Flags') } catch { $kind = $null } }
    }
    $conv = ConvertTo-RootFlagsValue -Raw $raw -Kind $kind
    [pscustomobject]@{ Raw = $raw; Kind = $kind; Value = $conv.Value; Reason = $conv.Reason }
}
function Get-RootFlagsDisplay {
    if (-not $isGP) { return 'n/a (local location)' }
    $rf = Get-RootFlagsValue
    if ($null -ne $rf.Value) { return '0x{0:X}' -f $rf.Value }
    if ($null -ne $rf.Raw) { return "(unusable: $($rf.Reason))" }
    return '(absent)'
}
# $true only when THIS run created the PolicyServers key of the location (New-BaseKey). It is
# reported as BaseKeyCreated in the Add summary. A caller that removes only the keys it created
# (the Lab tests) needs the script's own evidence: a base key that was absent before the run may
# have been created by another writer since, and such a key is not the caller's to remove.
$baseKeyCreated = $false
function New-BaseKey {
    # Creates the PolicyServers key of this location (and its missing ancestors) with
    # RegCreateKeyExW and records the creation. The evidence is the disposition of that one call:
    # only REG_CREATED_NEW_KEY sets the flag. A key that another writer created first is opened,
    # not created, and the flag stays $false. The flag is never reset: a later call in the same run
    # opens the key this run created. A key that already existed is checked again right here:
    # the preflight stopped at the deepest existing component, so a base key that another writer
    # created since (with a delegated DACL, or as a symbolic link) was never checked. The check
    # throws, so nothing is written to such a key. A key this run created earlier passes again.
    if (New-RegistryKeyNative -Path $hive) { $script:baseKeyCreated = $true }
    else { Assert-ProtectedRegistryPath -Path $hive }
}
function Write-CepEntry {
    param([string]$EntryPath, [hashtable]$Strings, [hashtable]$Dwords)
    # Validate the specific key being written here, not just the shared ancestor chain checked once
    # up front for the HTTP CEP entry. The AD Enrollment Policy row is a SIBLING leaf ($hive\<AD_KEY>)
    # of that entry, so a delegated or symlinked AD_KEY key would otherwise take this privileged write
    # unchecked. Assert-ProtectedRegistryPath walks from the hive root to the deepest existing
    # component of $EntryPath, so it covers both the entry and the AD row.
    Assert-ProtectedRegistryPath -Path $EntryPath
    # Both callers write a direct child of $hive (the CEP entry and the AD row). New-BaseKey
    # creates the base key first, so the run records whether the base key is its own; the entry
    # key is then created (or opened, when it exists) by one RegCreateKeyExW call. The function
    # returns $true only when that call CREATED the entry key (REG_CREATED_NEW_KEY): the caller
    # reports it as EntryAction. Neither call ever recreates or wipes an existing key.
    New-BaseKey
    $created = New-RegistryKeyNative -Path $EntryPath
    # Check the path AGAIN, after the creates and before any value write. The check above stopped
    # at the deepest existing component. When the base key was absent then, a concurrent writer
    # can have created it with a delegated DACL, and an untrusted principal can then have created
    # the entry name as a symbolic link; the create above opens such a link (dwOptions 0 follows
    # it) and the value writes below would go through it. Every component exists now, so this
    # second check covers the link status, the owner and the ACL of the base key and the entry.
    # It throws on a failure, so no value is written.
    Assert-ProtectedRegistryPath -Path $EntryPath
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
    return $created
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
            # A CEP entry is a LEAF (only values). Refuse a recursive delete of one that has subkeys:
            # Assert-ProtectedRegistryPath validated the entry and its ancestors, but a recursive
            # Remove-Item descends into subkeys too, and on 5.1 .NET opens descendants WITHOUT
            # REG_OPTION_OPEN_LINK - so a registry link planted beneath a delegated descendant would be
            # followed and the link target's subkeys deleted. An entry with subkeys is anomalous.
            $tk = Get-Item -LiteralPath $target -ErrorAction SilentlyContinue
            if ($tk -and $tk.SubKeyCount -gt 0) { throw "Refusing to remove ${target}: a CEP entry is a leaf key, but this one has $($tk.SubKeyCount) subkey(s) - a recursive delete could follow a registry symbolic link planted beneath a delegated descendant. Investigate and remove it manually." }
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
# What happened to the entry key: None (the run stopped before step 2), Declined (the gate or
# -WhatIf said no), Created (the key was absent) or Updated (the key existed and was rewritten).
# EntryApplied alone cannot tell Created from Updated, and a caller that removes only the keys
# it created needs that difference.
$entryAction = 'None'

# ---- 0a. the existing root Flags must be usable (GP locations) - checked BEFORE any write
# Step 3 rewrites the root Flags from the existing value. A value that another tool left as text
# (REG_SZ 'abc', REG_BINARY, an empty string) cannot be converted. Refuse it here, before the AD
# row and the CEP entry are written, so the operator sees one clear message and nothing changes.
if ($isGP) {
    $rf0 = Get-RootFlagsValue
    if ($null -ne $rf0.Raw -and $null -eq $rf0.Value) {
        throw "Refusing to write: the root Flags value under $hive cannot be used ($($rf0.Reason)). Fix or delete that value before rerunning. Nothing was written."
    }
}

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
    # A usable row must actually HAVE a URL and PolicyID: an empty one is not a server a client can
    # reach. Checked independently of the expected values (a caller passing the row's own empty
    # URL/PolicyID as expected would otherwise pass the empty-equals-empty comparison).
    if (-not "$($Key.GetValue('URL'))" -or -not "$($Key.GetValue('PolicyID'))") { return $false }
    if ("$($Key.GetValue('URL'))" -ne $ExpectUrl -or "$($Key.GetValue('PolicyID'))" -ne $ExpectPolicyId) { return $false }
    $names = @($Key.GetValueNames())
    # URL, PolicyID and FriendlyName must be REG_SZ (a PolicyID stored as REG_DWORD stringifies to a
    # matching value but is not the string a client reads); Flags/AuthFlags/Cost must be REG_DWORD.
    foreach ($s in 'URL', 'PolicyID', 'FriendlyName') {
        if ($names -notcontains $s) { return $false }
        if ($Key.GetValueKind($s) -ne [Microsoft.Win32.RegistryValueKind]::String) { return $false }
    }
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
                $null = Write-CepEntry -EntryPath $adTarget `
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
        EntryAction    = $entryAction
        BaseKeyCreated = $baseKeyCreated
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
    # EntryAction is derived from the disposition of the RegCreateKeyExW call inside Write-CepEntry
    # (the one call that creates or opens the entry key), not from the existence check that named
    # the prompt: a key that another writer created while the prompt waited is opened and updated,
    # and the summary must say so, or a caller that cleans up only created keys would adopt it.
    $createdAtWrite = $false
    try {
        $createdAtWrite = Write-CepEntry -EntryPath $target `
            -Strings @{ URL = $Url; PolicyID = $PolicyId; FriendlyName = $PolicyName } `
            -Dwords  @{ Flags = $flags; AuthFlags = $authFlags; Cost = (ConvertTo-DwordInt $Cost) }
    } catch { throw "CEP entry write to $target failed (the key may be partially written - inspect it): $_" }
    $entryApplied = $true
    $entryAction = if ($createdAtWrite) { 'Created' } else { 'Updated' }
} else {
    $entryAction = 'Declined'
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
# entry in this location (the CEP entry declined AND no AD row) activates GP CEP configuration with
# no server, so clients lose the AD enrollment policy. Gate on a real entry: the requested CEP entry,
# the AD row, OR any OTHER pre-existing USABLE entry in this location - so a declined requested entry
# does not wrongly block a root-Flags change when another working server exists (that would over-
# restrict), while an incomplete fragment (e.g. a subkey carrying only FriendlyName) does NOT count,
# so clearing 0x2 cannot activate a list of zero usable servers. Usable = a complete row (URL,
# PolicyID, FriendlyName and DWORD Flags/AuthFlags/Cost), the same bar Test-RegEntryUsable enforces.
$anyUsableServer = $false
foreach ($e in @(Get-CepEntries)) {
    if (-not "$($e.URL)") { continue }
    $k = Get-Item -LiteralPath "$hive\$($e.Key)" -ErrorAction SilentlyContinue
    if ($k -and (Test-RegEntryUsable -Key $k -ExpectUrl "$($e.URL)" -ExpectPolicyId "$($e.PolicyID)")) { $anyUsableServer = $true; break }
}
$policyServerPresent = $entryExists -or ($adRow -eq 'applied') -or $anyUsableServer

# ---- 3. root Flags (GP locations; DISABLE bits, preserved across runs) ---------------------
if ($isGP) {
    # Read the value again: steps 1 and 2 do not touch it, but another writer may have. Step 0a
    # refused an unusable value before any write; one that turned unusable since is skipped with
    # a note, never a raw conversion error after the entry writes.
    $rf = Get-RootFlagsValue
    $existing = $rf.Value
    # The kind matters as much as the number: a Flags left as REG_SZ by some other tool is not a
    # value the client reads, so it is rewritten as a DWORD even when its number is already right.
    $existingKind = $rf.Kind
    $newFlags = if ($null -ne $existing) { $existing } else { 0 }
    if ($newFlags -band 0x2) {
        Write-Warning 'Existing root Flags had bit 0x2 set (clients IGNORE the GP-provided policy list). Clearing it.'
        $newFlags = $newFlags -band (-bnot 0x2)
    }
    if ($DisableUserConfigured) { $newFlags = $newFlags -bor 0x4 }
    if ($EnableUserConfigured)  { $newFlags = $newFlags -band (-bnot 0x4) }
    if ($null -ne $rf.Raw -and $null -eq $existing) {
        $notes.Add("Root Flags NOT written: the existing value under $hive became unusable while this run was in progress ($($rf.Reason)). Fix or delete that value and rerun.")
    }
    elseif (-not $policyServerPresent -and -not $WhatIfPreference) {
        $notes.Add("Root Flags NOT written: no usable policy-server entry exists (the CEP entry was declined and there is no AD Enrollment Policy row). Writing PolicyServers root values would activate GP CEP configuration with no server, and clients would lose the AD enrollment policy.")
    }
    elseif (($null -eq $existing) -or ($existing -ne $newFlags) -or ($existingKind -ne [Microsoft.Win32.RegistryValueKind]::DWord)) {
        $from = if ($null -ne $existing) { '0x{0:X}' -f $existing } else { '(absent)' }
        if ($PSCmdlet.ShouldProcess($hive, ('Set root Flags {0} -> 0x{1:X} (disable bits: 0x2 ignore GP list, 0x4 ignore user-configured)' -f $from, $newFlags))) {
            New-BaseKey
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
        New-BaseKey
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
            # Validate the sibling's registry path BEFORE reading or removing it: a sibling that is a
            # registry SYMBOLIC LINK to a key in another location would otherwise be read through the
            # link and then removed by Remove-Item -Recurse, which on PowerShell 7 opens the link
            # target and deletes THROUGH it - destroying an unrelated key. Assert-ProtectedRegistryPath
            # refuses a link (and an untrusted owner/writer); a refusal skips this sibling with a note
            # rather than failing the whole run.
            try { Assert-ProtectedRegistryPath -Path "$hive\$($d.Key)" }
            catch {
                $notes.Add("Superseded entry '$($d.URL)' (key $($d.Key)) NOT removed: its registry path failed the protection check ($($_.Exception.Message)) - it may be a symbolic link or a delegated key. Investigate and remove it manually.")
                continue
            }
            $sib = Get-Item -LiteralPath "$hive\$($d.Key)" -ErrorAction SilentlyContinue
            if (-not $sib -or "$($sib.GetValue('PolicyID'))" -ne "$PolicyId" -or "$($sib.GetValue('URL'))" -eq $Url) {
                $notes.Add("Entry at key $($d.Key) NOT removed: it no longer serves PolicyID $PolicyId under a different URL (changed or removed while the prompt was open).")
                continue
            }
            # As in -Remove: a CEP entry is a leaf. Refuse a recursive delete of a sibling with subkeys
            # (a recursive Remove-Item can follow a registry link beneath a delegated descendant).
            if ($sib.SubKeyCount -gt 0) {
                $notes.Add("Superseded entry '$($d.URL)' (key $($d.Key)) NOT removed: it has $($sib.SubKeyCount) subkey(s) - a CEP entry is a leaf, and a recursive delete could follow a symbolic link planted beneath a delegated descendant. Investigate and remove it manually.")
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
