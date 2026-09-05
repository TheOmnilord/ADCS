<#PSScriptInfo
.VERSION 1.0.1
.GUID 54763db6-2359-401f-8960-ef0de5911aaf
.AUTHOR Sveinung Svea
.PROJECTURI https://github.com/TheOmnilord/ADCS
.LICENSEURI https://github.com/TheOmnilord/ADCS/blob/main/LICENSE
.TAGS ADCS PKI CertificateServices
.RELEASENOTES
1.0.1 - The domain objectGUID for the AD Enrollment Policy row is resolved BEFORE any GPO write and a lookup failure now aborts the run (previously the CEP entry was written first and the failure became a warning, leaving a GPO that removes the AD enrollment policy from every client in scope); a GUID-shaped -GpoName falls back to the GPO ID only after an independent listing proves no GPO carries that display name (any other name-lookup failure is rethrown); registry.pol reads validate the PReg header, refuse oversized files and trailing junk, and return the EFFECTIVE value (records replayed in file order: last write wins, **del./**delvals./**DeleteValues honoured) instead of the first matching record; the deletion-order damage check is judged per value, so a deletion between an obsolete record and its replacement no longer raises a false DAMAGED warning
1.0.0 - Initial release
#>

<#
.SYNOPSIS
    Writes (or removes) a certificate enrollment policy (CEP) server directly in a domain GPO -
    offline with respect to the policy server: no "Validate Server" round-trip, ever.

.DESCRIPTION
    The "Certificate Services Client - Certificate Enrollment Policy" setting is plain
    registry-based policy ([MS-GPREG]) stored in the GPO's Registry.pol. This script authors
    those values with Set-GPRegistryValue, which also handles what hand-editing SYSVOL would
    get wrong: AD + GPT.INI version bumps and registering the Registry CSE on the GPO. Values
    authored here appear in the GPME Public Key Policies dialog exactly as if clicked in.

    Derivations (computed locally, no server contact):
      * Subkey name = SHA-1 over the UTF-16LE bytes of the invariant-lowercased URL
      * PolicyID    = Java String.hashCode() of the policy name (EJBCA MSAE behavior).
                      For CEP servers that return a GUID (e.g. Microsoft CEP) pass -PolicyId.

    By default the script ALSO ensures the built-in AD enrollment policy row ("LDAP:", subkey
    37c9dc30f207f27f61a2f7c3aed598a6e2920b54, PolicyID = the domain object's objectGUID,
    Cost 0xFFFFFFFF): enabling GP-based CEP configuration suppresses the client-synthesized
    AD default policy, so a GPO carrying only your CEP entry would silently REMOVE the AD
    enrollment policy from every client in scope and stop autoenrollment against AD-published
    templates. Opt out only deliberately, with -SkipADPolicy. The AD row is never treated as
    a removable duplicate by -ReplaceExisting.

    The PolicyServers root Flags are DISABLE bits (EnrollmentPolicyFlags): 0x2 makes clients
    ignore the whole GP-provided list (never set; cleared with a warning if found), 0x4 makes
    clients ignore user-configured servers. Existing bits are preserved across runs: rerunning
    without -DisableUserConfigured does NOT clear a previously set 0x4 (use
    -EnableUserConfigured to clear it explicitly).

    Robustness: all writes run with ErrorActionPreference Stop and each value is written with
    its own Set-GPRegistryValue call - the cmdlet's multi-value "list" form is deliberately
    avoided because it plants a **delVals. deletion record that, ordered after
    individually-authored values, makes clients DELETE the whole entry. The script detects
    and warns about such mis-ordered deletion records in the GPO. Entries are verified by
    read-back (including detection of missing values); transient SYSVOL/registry.pol
    contention is retried. State inspection reads registry.pol from the PDC emulator (or
    -Server when given) so reads and writes see the same replica. Still avoid editing the
    same GPO concurrently from two sessions or GPME - Set-GPRegistryValue is an unlocked
    read-modify-write and concurrent writers can silently lose values; the read-back makes
    this script detect such loss for its own entry.

    Requires the GroupPolicy module (GPMC / RSAT) and permission to edit the GPO.
    Supports -WhatIf / -Confirm.

.PARAMETER GpoName
    Display name of an existing GPO, or its GUID. The name lookup is tried FIRST; only if no
    GPO has that display name is the value treated as a GUID. Create and link a GPO first if
    needed:  New-GPO -Name 'PKI - Enrollment Policy' | New-GPLink -Target 'OU=...,DC=...'

.PARAMETER Url
    Full CEP URI - must match the server exactly (the SHA-1 subkey name and the clients'
    GetPolicies calls use it verbatim). Trimmed; must be absolute http/https; control
    characters rejected. Rerunning with a DIFFERENT URL does not remove the old entry - the
    script warns about same-PolicyID siblings; pass -ReplaceExisting to delete them.

.PARAMETER PolicyName
    The EJBCA MSAE alias "Policy Name". Becomes FriendlyName and (unless -PolicyId is given)
    the PolicyID hash input - hashed VERBATIM; renaming it in EJBCA changes the PolicyID and
    orphans deployed entries.

.PARAMETER PolicyId
    Explicit PolicyID for non-EJBCA servers (must match the server's GetPolicies response).

.PARAMETER Scope
    Machine (default) = Computer Configuration (HKLM), User = User Configuration (HKCU).

.PARAMETER Authentication
    Client authentication type for the CEP endpoint: Anonymous (1), Kerberos (2, default,
    = "Windows integrated"), UsernamePassword (4), Certificate (8).

.PARAMETER Cost
    Priority; lower = preferred among endpoints sharing a PolicyID. Full DWORD range
    1..4294967295; default 0x7FFFFFFD ("Priority: Default"). Pass large values in decimal
    (a 0xFFFFFFFF literal is parsed by PowerShell as Int32 -1).

.PARAMETER NoAutoEnroll
    Leave "Enable for automatic enrollment and renewal" off (clears entry Flags bit 0x10).

.PARAMETER AllowUntrustedIssuer
    Equivalent to UNchecking "Require strong validation during enrollment" (sets Flags 0x20).

.PARAMETER NoClientId
    Do not send the ClientId attribute (clears Flags bit 0x4; default matches GPME's 0x14).

.PARAMETER SetAsDefault
    Write this policy's PolicyID into the unnamed "(Default)" REG_SZ value on the
    PolicyServers key (the dialog's Default checkbox). Interactive preselection only.

.PARAMETER ClearDefault
    Remove the unnamed "(Default)" marker.

.PARAMETER SkipADPolicy
    Do not write the AD enrollment policy row into the GPO. Background: as soon as a GPO
    delivers ANY CEP configuration, clients in scope stop auto-generating the built-in
    "Active Directory Enrollment Policy" and use only the entries the GPO lists. A GPO
    without the LDAP: row therefore TAKES AWAY the AD enrollment policy from its clients,
    and autoenrollment against AD-published templates (the internal enterprise CA) stops.
    The script writes the row by default to prevent that; pass this switch only when that
    removal is intended (no internal ADCS, or a deliberate migration off AD enrollment).

.PARAMETER ReplaceExisting
    Remove sibling entries sharing this PolicyID under a different URL (stale/typo entries).
    The AD policy row is never removed by this. Without the switch the script only warns,
    since multiple URLs per PolicyID is also the legitimate redundant-endpoint pattern.

.PARAMETER DisableUserConfigured
    Set root Flags bit 0x4 (clients ignore user-configured policy servers). Preserved on
    later runs; clear with -EnableUserConfigured.

.PARAMETER EnableUserConfigured
    Clear root Flags bit 0x4.

.PARAMETER EnableAutoEnrollmentPolicy
    Also write the "Certificate Services Client - Auto-Enrollment" setting into the same GPO
    scope: AEPolicy (default 7 = enabled + renew/update + template update), expiration
    notification percent (default 10) and store (default MY). If the GPO already carries
    different Auto-Enrollment values the script WARNS with old -> new before its confirmation
    gate. Without autoenrollment enabled somewhere, the policy list alone enrolls nothing.

.PARAMETER AEPolicy
    AEPolicy value to write with -EnableAutoEnrollmentPolicy (default 7; 0x8000=32768 writes
    it disabled).

.PARAMETER AEExpirationPercent
    OfflineExpirationPercent to write with -EnableAutoEnrollmentPolicy (default 10).

.PARAMETER AEStore
    OfflineExpirationStoreNames to write with -EnableAutoEnrollmentPolicy (default 'MY').

.PARAMETER Domain
    Optional: domain for the GroupPolicy cmdlets and the AD/SYSVOL lookups.

.PARAMETER Server
    Optional: domain controller to target - used for BOTH the GroupPolicy cmdlets and the
    registry.pol state reads, keeping them on the same replica.

.PARAMETER Remove
    Removal mode: delete the entry for -Url from the GPO scope, clear the (Default) marker if
    it pointed at that entry's PolicyID and no remaining entry still serves that PolicyID,
    and report what remains (remaining entries, actual root Flags, Auto-Enrollment). Other
    entries and shared configuration are left alone.

.EXAMPLE
    .\Add-CertificateEnrollmentPolicyServerToGpo.ps1 -GpoName 'PKI - Enrollment Policy' -Url 'https://pki.example.net/ejbca/msae/CEPService?alias' -PolicyName 'Example PKI Service' -EnableAutoEnrollmentPolicy -WhatIf

.EXAMPLE
    .\Add-CertificateEnrollmentPolicyServerToGpo.ps1 -GpoName 'PKI - Enrollment Policy' -Url 'https://pki.example.net/ejbca/msae/CEPService?alias' -Scope User -Remove

.NOTES
    Complete teardown of everything this script can write, per scope prefix H (HKLM for
    -Scope Machine, HKCU for User), all via Remove-GPRegistryValue -Name <GPO>:
      -Key 'H\SOFTWARE\Policies\Microsoft\Cryptography\PolicyServers\<entry-hash>'   (each entry)
      -Key 'H\SOFTWARE\Policies\Microsoft\Cryptography\PolicyServers\37c9dc30f207f27f61a2f7c3aed598a6e2920b54'  (AD row)
      -Key 'H\SOFTWARE\Policies\Microsoft\Cryptography\PolicyServers' -ValueName 'Flags'
      -Key 'H\SOFTWARE\Policies\Microsoft\Cryptography\PolicyServers' -ValueName ''    ((Default) marker)
      -Key 'H\SOFTWARE\Policies\Microsoft\Cryptography\AutoEnrollment'                 (if AE was enabled)
    Clients pick changes up at the next GP refresh; force with gpupdate, then trigger
    enrollment with "certutil -pulse" (machine) / "certutil -user -pulse" (user).
#>
# NOTE: deliberately NO '#Requires -Modules GroupPolicy'. On PowerShell 7 the GroupPolicy
# module is not visible to Get-Module -ListAvailable (it lives only in the Windows PowerShell
# module path), so #Requires would refuse to run the script even though the module loads fine
# through the WinPSCompatSession shim. The explicit Import-Module below runs under
# $ErrorActionPreference = 'Stop' and fails just as cleanly when the module truly is absent.
[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Add')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPositionalParameters', '',
    Justification = 'The script''s own small internal helpers (Get-PolValue, Test-PolDeletionOrder) are intentionally called positionally for readability; no external cmdlet is called positionally.')]
param(
    [Parameter(Mandatory, Position = 0, ParameterSetName = 'Add')]
    [Parameter(Mandatory, Position = 0, ParameterSetName = 'Remove')]
    [string]$GpoName,

    [Parameter(Mandatory, ParameterSetName = 'Add')]
    [Parameter(Mandatory, ParameterSetName = 'Remove')]
    [string]$Url,

    [Parameter(Mandatory, ParameterSetName = 'Add')]
    [string]$PolicyName,

    [Parameter(ParameterSetName = 'Add')]
    [string]$PolicyId,

    [ValidateSet('Machine','User')]
    [string]$Scope = 'Machine',

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
    [Parameter(ParameterSetName = 'Add')] [switch]$EnableAutoEnrollmentPolicy,
    [Parameter(ParameterSetName = 'Add')] [ValidateRange(0, 65535)] [int]$AEPolicy = 7,
    [Parameter(ParameterSetName = 'Add')] [ValidateRange(1, 99)]    [int]$AEExpirationPercent = 10,
    [Parameter(ParameterSetName = 'Add')] [string]$AEStore = 'MY',

    [string]$Domain,
    [string]$Server,

    [Parameter(Mandatory, ParameterSetName = 'Remove')]
    [switch]$Remove
)

$ErrorActionPreference = 'Stop'
Import-Module GroupPolicy
$notes = New-Object System.Collections.Generic.List[string]

# ---- parameter conflicts -------------------------------------------------------------------
if ($SetAsDefault -and $ClearDefault) { throw '-SetAsDefault and -ClearDefault are mutually exclusive.' }
if ($DisableUserConfigured -and $EnableUserConfigured) { throw '-DisableUserConfigured and -EnableUserConfigured are mutually exclusive.' }
foreach ($p in 'AEPolicy','AEExpirationPercent','AEStore') {
    if ($PSBoundParameters.ContainsKey($p) -and -not $EnableAutoEnrollmentPolicy) {
        throw "-$p only has effect together with -EnableAutoEnrollmentPolicy."
    }
}

# ---- resolve the GPO: display name FIRST, GUID only as fallback ----------------------------
$domParams = @{}
if ($Domain) { $domParams.Domain = $Domain }
if ($Server) { $domParams.Server = $Server }
$gpo = $null
try { $gpo = Get-GPO @domParams -Name $GpoName } catch {
    $nameError = $_
    $guid = [Guid]::Empty
    if (-not [Guid]::TryParse($GpoName, [ref]$guid)) { throw }
    # "Not found by name" is the ONLY failure that makes the GUID reading legitimate - not an
    # authorization error, a transient RPC/SYSVOL failure or a PS7 remoting quirk, any of which
    # would otherwise silently retarget the run at whatever GPO happens to carry that ID. Prove
    # the absence with an independent listing: a GPO that DOES carry this display name means the
    # name lookup failed for another reason, and the GUID must not paper over it.
    $sameName = @(Get-GPO @domParams -All | Where-Object { $_.DisplayName -eq $GpoName })
    if ($sameName.Count) {
        throw "A GPO named '$GpoName' exists but could not be read by name ($($nameError.Exception.Message)); refusing to reinterpret the value as a GPO ID."
    }
    try { $gpo = Get-GPO @domParams -Guid $guid } catch {
        throw "No GPO found with display name '$GpoName', and no GPO has that ID either."
    }
}
$wr = @{ Guid = $gpo.Id } + $domParams          # every write targets the resolved GPO by ID
$gpoLabel = "GPO '$($gpo.DisplayName)' ($($gpo.Id))"

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
        Write-Warning 'PolicyName has leading/trailing whitespace. It is hashed VERBATIM - make sure this exactly matches the EJBCA Policy Name.'
    }
    if ($AEStore -match '[\x00-\x1F]') { throw 'AEStore contains control characters.' }
}

# ---- derivations ---------------------------------------------------------------------------
$hivePrefix = if ($Scope -eq 'Machine') { 'HKLM' } else { 'HKCU' }
$sideFolder = if ($Scope -eq 'Machine') { 'Machine' } else { 'User' }
$baseKey  = "$hivePrefix\SOFTWARE\Policies\Microsoft\Cryptography\PolicyServers"
$aeKey    = "$hivePrefix\SOFTWARE\Policies\Microsoft\Cryptography\AutoEnrollment"
$AD_KEY   = '37c9dc30f207f27f61a2f7c3aed598a6e2920b54'      # SHA-1 of utf16le "ldap:"
$sha1 = [System.Security.Cryptography.SHA1]::Create()
$hash = -join ($sha1.ComputeHash([System.Text.Encoding]::Unicode.GetBytes($Url.ToLowerInvariant())) |
               ForEach-Object { $_.ToString('x2') })
$entryKey = "$baseKey\$hash"

# State reads must see the same replica the GroupPolicy cmdlets write (PDC emulator by
# default, or -Server). \\domain\SYSVOL would hit the site-nearest DC and could be stale.
$polHost = $gpo.DomainName
if ($Server) { $polHost = $Server }
else {
    try {
        $polHost = if ($Domain) {
            $ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $Domain)
            ([System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($ctx)).PdcRoleOwner.Name
        } else {
            ([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()).PdcRoleOwner.Name
        }
    } catch { Write-Verbose "PDC resolution failed; falling back to \\$($gpo.DomainName) DFS path: $_" }
}
$polPath = "\\$polHost\SYSVOL\$($gpo.DomainName)\Policies\{$($gpo.Id)}\$sideFolder\registry.pol"

# ---- Registry.pol reader (state inspection: existing entries, root values, AE) -------------
function Read-PolRecords {
    param([string]$Path)
    $recs = New-Object System.Collections.Generic.List[object]
    if (-not (Test-Path -LiteralPath $Path)) { return $recs }
    # Bound the work before reading: a registry.pol is kilobytes; anything approaching the cap is
    # not a policy file this script should load into an elevated process and walk record by record.
    $fileLen = (Get-Item -LiteralPath $Path).Length
    if ($fileLen -gt 64MB) { throw "registry.pol at $Path is $fileLen bytes - far beyond any plausible policy file; refusing to parse it. Inspect the GPO before using this script." }
    $b = $null
    for ($try = 1; $try -le 3; $try++) {
        try { $b = [System.IO.File]::ReadAllBytes($Path); break }
        catch [System.IO.IOException] {
            if ($try -eq 3) { throw "Cannot read $Path (locked by another process?): $_" }
            Start-Sleep -Milliseconds (400 * $try)
        }
    }
    if ($b.Length -lt 8) { return $recs }
    # [MS-GPREG] header: the 'PReg' signature followed by version 1. Anything else is not a
    # Registry.pol (or a damaged one) and is refused rather than parsed on a best-effort basis.
    if ($b[0] -ne 0x50 -or $b[1] -ne 0x52 -or $b[2] -ne 0x65 -or $b[3] -ne 0x67 -or [BitConverter]::ToUInt32($b, 4) -ne 1) {
        throw "registry.pol at $Path does not carry the PReg/version-1 header - not a Registry.pol file, or corrupt. Repair or re-author the GPO before using this script."
    }
    $pos = 8
    $corrupt = { throw "registry.pol at $Path appears truncated or corrupt (offset $pos). Repair or re-author the GPO before using this script." }
    while ($pos -le $b.Length - 2) {
        if ([BitConverter]::ToUInt16($b, $pos) -ne 0x5B) { & $corrupt }   # every record opens with '['; trailing junk is corruption, not padding
        $pos += 2
        $sb = New-Object System.Text.StringBuilder
        while ($true) {
            if ($pos + 2 -gt $b.Length) { & $corrupt }
            $c = [BitConverter]::ToUInt16($b, $pos)
            if ($c -eq 0) { break }
            [void]$sb.Append([char]$c); $pos += 2
        }
        $pos += 4    # NUL + ';'
        $keyName = $sb.ToString()
        $sb = New-Object System.Text.StringBuilder
        while ($true) {
            if ($pos + 2 -gt $b.Length) { & $corrupt }
            $c = [BitConverter]::ToUInt16($b, $pos)
            if ($c -eq 0) { break }
            [void]$sb.Append([char]$c); $pos += 2
        }
        $pos += 4
        $valName = $sb.ToString()
        if ($pos + 12 -gt $b.Length) { & $corrupt }
        $type = [BitConverter]::ToUInt32($b, $pos); $pos += 6
        $size = [BitConverter]::ToUInt32($b, $pos); $pos += 6
        if ($size -gt ($b.Length - $pos)) { & $corrupt }
        $data = New-Object byte[] $size
        if ($size -gt 0) { [Array]::Copy($b, $pos, $data, 0, $size) }
        $pos += $size + 2   # data + ']'
        $val = switch ($type) {
            1 { [System.Text.Encoding]::Unicode.GetString($data).TrimEnd([char]0) }
            4 { if ($size -ge 4) { [BitConverter]::ToUInt32($data, 0) } else { $null } }
            default { $null }
        }
        $recs.Add([pscustomobject]@{ Key = $keyName; ValueName = $valName; Type = $type; Data = $val; Index = $recs.Count })
    }
    return $recs
}
$relBase = 'Software\Policies\Microsoft\Cryptography\PolicyServers'
$relAe   = 'Software\Policies\Microsoft\Cryptography\AutoEnrollment'
function Get-PolEffectiveValues([object[]]$Recs, [string]$RelKey) {
    # Replays the records of ONE key in file order - the order the client Registry CSE applies
    # them - and returns what a client ends up with: @{ Values = name -> data of the values that
    # survive; Deleted = names that were written and then deleted by a LATER record }.
    #   plain value        a later record for the same name replaces an earlier one
    #   **del.<name>       deletes that one value
    #   **delvals.         deletes every value written so far under the key
    #   **DeleteValues     deletes each name in its ';'-separated data
    #   **DeleteKeys       deletes whole keys: its ';'-separated data lists FULL key paths (without
    #                      the hive) and the record's own key field is IGNORED by the Group Policy
    #                      engine (contrary to Microsoft's documentation; verified by the LGPO
    #                      author's corrections) - so an item naming this key, or an ancestor of
    #                      it, wipes every value written so far, wherever the record sits
    #   **soft.<name>      writes <name> only if no value of that name exists yet
    #   **Comment: ...     a comment; **SecureKey a permissions instruction - neither is a value
    # Any other **-prefixed instruction is unknown to this evaluator and makes it FAIL CLOSED:
    # guessing would let a deleted endpoint look present to the prerequisite and replacement guards.
    # Value names compare case-insensitively (registry semantics). Only the first/older record
    # would be seen by a naive first-match read - and that can be the obsolete one.
    $values = @{}; $deleted = @{}
    $relLower = $RelKey.ToLowerInvariant().Trim('\')
    foreach ($r in $Recs) {
        $vn = "$($r.ValueName)"
        if ($vn -match '^(?i)\*\*deletekeys$') {
            # Key field ignored on purpose (see above). An item deletes this key when it IS this
            # key or an ancestor of it (then this key goes with it).
            foreach ($item in @("$($r.Data)" -split ';' | ForEach-Object { $_.Trim().Trim('\') } | Where-Object { $_ })) {
                $i = $item.ToLowerInvariant()
                if ($relLower -eq $i -or $relLower.StartsWith($i + '\')) { foreach ($k in @($values.Keys)) { $deleted[$k] = $true }; $values.Clear(); break }
            }
            continue
        }
        if ($r.Key -notmatch ('^(?i)' + [regex]::Escape($RelKey) + '$')) { continue }
        if ($vn.StartsWith('**')) {
            # Anchored checks, NOT -like '**del*': in -like the asterisks are wildcards, so that
            # pattern also matches an ordinary value whose name merely contains "del".
            if ($vn -match '^(?i)\*\*(comment:|securekey)') { continue }
            if ($vn -match '^(?i)\*\*soft\.(.+)$') {
                if (-not $values.ContainsKey($Matches[1])) { $values[$Matches[1]] = $r.Data }
                continue
            }
            $targets = @()
            if     ($vn -match '^(?i)\*\*del\.(.+)$')      { $targets = @($Matches[1]) }
            elseif ($vn -match '^(?i)\*\*delvals\.?$')     { $targets = @($values.Keys) }
            elseif ($vn -match '^(?i)\*\*deletevalues$')   { $targets = @("$($r.Data)" -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
            else { throw "registry.pol carries a policy instruction this script does not understand ('$vn' under '$($r.Key)'); refusing to reason about the GPO's effective state. Inspect the GPO in GPME before using this script." }
            foreach ($t in $targets) {
                if ($values.ContainsKey($t)) { $values.Remove($t); $deleted[$t] = $true }
            }
            continue
        }
        $values[$vn] = $r.Data
        if ($deleted.ContainsKey($vn)) { $deleted.Remove($vn) }   # re-written after a deletion: it survives
    }
    @{ Values = $values; Deleted = $deleted }
}
function Get-PolEntries([object[]]$Recs) {
    # One entry per 40-hex subkey, built from the EFFECTIVE values (last record wins, deletions
    # honoured) so an entry a client would not end up with is not reported as present.
    $leaves = @($Recs | ForEach-Object {
            if ($_.Key -match ('^(?i)' + [regex]::Escape($relBase) + '\\([0-9a-f]{40})$')) { $Matches[1].ToLowerInvariant() }
        } | Sort-Object -Unique)
    foreach ($leaf in $leaves) {
        $eff = (Get-PolEffectiveValues $Recs "$relBase\$leaf").Values
        if (-not $eff.Count) { continue }
        [pscustomobject]@{ Key = $leaf; URL = $eff['URL']; PolicyID = $eff['PolicyID'] }
    }
}
function Test-EntryRecordsPresent([object[]]$Recs, [string]$RelKey) {
    # PHYSICAL presence: does registry.pol carry ANY record at this key (values or deletion
    # instructions)? Used by -Remove, which must be able to clear a key whose values a later
    # **delvals. wiped (no EFFECTIVE values, but records to remove all the same).
    @($Recs | Where-Object { $_.Key -match ('^(?i)' + [regex]::Escape($RelKey) + '$') }).Count -gt 0
}
function Get-PolValue([object[]]$Recs, [string]$RelKey, [string]$ValueName) {
    $eff = (Get-PolEffectiveValues $Recs $RelKey).Values
    if ($eff.ContainsKey($ValueName)) { return $eff[$ValueName] }
    return $null
}
function Test-PolDeletionOrder([object[]]$Recs, [string]$RelKey, [string]$Label) {
    # A **del./**delVals. record ORDERED AFTER a value's last record makes the client CSE write
    # the value and then delete it - it is absent on every client. Judged per value, in file
    # order: a deletion BETWEEN an obsolete record and its replacement is legitimate (the
    # replacement survives), so only values that are actually lost are reported.
    $lost = @((Get-PolEffectiveValues $Recs $RelKey).Deleted.Keys | ForEach-Object { if ($_ -eq '') { '(Default)' } else { $_ } } | Sort-Object)
    if ($lost.Count) {
        $notes.Add("DAMAGED registry.pol ordering for ${Label}: a GP deletion record (**del*) follows the value(s) $($lost -join ', '), so clients will DELETE them when applying the GPO. Fix by re-authoring the entry: run -Remove for it, then add it again.")
    }
}
function Test-GpoEntry([string]$K, [hashtable]$Expect) {
    $vals = Get-GPRegistryValue @wr -Key $K
    $map = @{}
    foreach ($v in $vals) { $map[$v.ValueName] = $v.Value }
    $bad = @()
    foreach ($n in $Expect.Keys) {
        if (-not $map.ContainsKey($n) -or $null -eq $map[$n]) { $bad += $n; continue }
        $exp = $Expect[$n]; $got = $map[$n]
        $ok = if ($exp -is [int] -or $exp -is [uint32] -or $exp -is [long]) {
            # DWORDs compare bit-exactly: Get-GPRegistryValue returns 0xFFFFFFFF as Int32 -1
            $expN = if ($exp -is [int]) { [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$exp), 0) } else { [uint32]$exp }
            $gotN = if ($got -is [int]) { [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$got), 0) } else { try { [uint32]$got } catch { $null } }
            $gotN -eq $expN
        } else { "$got" -ceq "$exp" }
        if (-not $ok) { $bad += $n }
    }
    if ($bad) { throw "Post-write verification failed for value(s) $($bad -join ', ') under $K - a concurrent edit of this GPO may have raced this run. Re-run the script." }
}
function Invoke-GPWrite {
    param([scriptblock]$Op, [switch]$TolerateNotFound)
    # Retries transient SYSVOL/registry.pol contention (sharing violations, brief access-denied
    # blips from AV/backup/replication). Set-GPRegistryValue blocks are idempotent rewrites;
    # Remove-GPRegistryValue blocks pass -TolerateNotFound so a retry after a half-committed
    # removal (or an already-absent target) counts as success.
    $attempt = 0
    while ($true) {
        $attempt++
        try { & $Op; return } catch {
            if ($TolerateNotFound -and $_.Exception.Message -match 'was not found') { return }
            if ($attempt -lt 3 -and $_.Exception.Message -match '0x80070020|0x80070005|Access is denied|being used by another process|sharing violation') {
                Start-Sleep -Milliseconds (400 * $attempt); continue
            }
            throw
        }
    }
}
function Get-RootFlagsDisplay([object[]]$Recs) {
    $v = Get-PolValue $Recs $relBase 'Flags'
    if ($null -ne $v) { '0x{0:X}' -f [int64]$v } else { '(absent)' }
}

$preRecs = Read-PolRecords -Path $polPath

# ============================ REMOVE MODE ===================================================
if ($PSCmdlet.ParameterSetName -eq 'Remove') {
    $removedEntry = $false; $defaultCleared = $false
    $entries0 = @(Get-PolEntries $preRecs)
    $mine = $entries0 | Where-Object { $_.Key -eq $hash } | Select-Object -First 1
    $marker0 = Get-PolValue $preRecs $relBase ''
    # Removal is about the PHYSICAL records at the hashed key, not the effective client state: an
    # entry whose values were later wiped by a **delvals. (the damaged ordering this script warns
    # about) is invisible to clients, yet its records still occupy the key and the documented
    # recovery - run -Remove, then add it again - has to be able to clear them.
    if (-not (Test-EntryRecordsPresent $preRecs "$relBase\$hash")) {
        $notes.Add("No entry for this URL (key $hash) in $gpoLabel ($Scope scope) - nothing to remove.")
    } else {
        # Identity for the marker logic and the prompt: the effective values when the entry is
        # live, else the last raw record of each (a damaged entry still names what it was).
        $rawOf = { param($name) @($preRecs | Where-Object { $_.Key -match ('^(?i)' + [regex]::Escape("$relBase\$hash") + '$') -and $_.ValueName -eq $name } | Select-Object -Last 1 -ExpandProperty Data) }
        $entryPid = if ($mine) { $mine.PolicyID } else { & $rawOf 'PolicyID' }
        $entryUrl = if ($mine) { $mine.URL }      else { & $rawOf 'URL' }
        if (-not $mine) { $notes.Add("The entry at key $hash has no EFFECTIVE values (its records are followed by a deletion instruction - clients never see it); removing its physical records.") }
        $markerMatches = $marker0 -and ("$marker0" -eq "$entryPid")
        $survivorServes = @($entries0 | Where-Object { $_.Key -ne $hash -and "$($_.PolicyID)" -eq "$marker0" }).Count -gt 0
        if ($PSCmdlet.ShouldProcess($gpoLabel, "Remove CEP entry $entryKey (URL=$entryUrl, PolicyID=$entryPid)")) {
            Invoke-GPWrite -TolerateNotFound { Remove-GPRegistryValue @wr -Key $entryKey | Out-Null }
            $removedEntry = $true
        }
        if ($markerMatches -and $survivorServes) {
            $notes.Add("(Default) marker kept: another entry still serves PolicyID $marker0 (redundant endpoint).")
        }
        if ($markerMatches -and -not $survivorServes -and ($removedEntry -or $WhatIfPreference)) {
            if ($PSCmdlet.ShouldProcess($gpoLabel, "Clear (Default) marker (would point at removed PolicyID $entryPid)")) {
                Invoke-GPWrite -TolerateNotFound { Remove-GPRegistryValue @wr -Key $baseKey -ValueName '' | Out-Null }
                $defaultCleared = $true
            }
        }
    }
    if ($ClearDefault -and -not $defaultCleared -and (Get-PolValue $preRecs $relBase '')) {
        if ($PSCmdlet.ShouldProcess($gpoLabel, 'Clear (Default) marker')) {
            Invoke-GPWrite -TolerateNotFound { Remove-GPRegistryValue @wr -Key $baseKey -ValueName '' | Out-Null }
            $defaultCleared = $true
        }
    }
    $postRecs = Read-PolRecords -Path $polPath
    $left = @(Get-PolEntries $postRecs)
    if ($removedEntry) { $left = @($left | Where-Object { $_.Key -ne $hash }) }
    if ($left.Count -gt 0) { $notes.Add("Remaining entries in this scope: $(($left | ForEach-Object { $_.URL }) -join ' | ')") }
    if ($null -ne (Get-PolValue $postRecs $relAe 'AEPolicy')) { $notes.Add('The Auto-Enrollment setting (AEPolicy) is still configured in this GPO scope - remove it separately if decommissioning (see NOTES).') }
    if ($left.Count -eq 0 -and (($null -ne (Get-PolValue $postRecs $relBase 'Flags')) -or (Get-PolValue $postRecs $relBase ''))) {
        $notes.Add('No entries remain, but the PolicyServers root values (Flags and/or (Default)) are still present - clients still treat GP CEP configuration as present. See NOTES for full teardown.')
    }
    foreach ($n in $notes) { Write-Warning $n }
    return [pscustomobject]@{
        Mode = 'Remove'; Gpo = $gpo.DisplayName; GpoId = $gpo.Id; Scope = $Scope; Key = $entryKey
        RemovedEntry = $removedEntry; DefaultCleared = $defaultCleared
        RootFlags = Get-RootFlagsDisplay $postRecs; DefaultMarker = "$(Get-PolValue $postRecs $relBase '')"; Notes = @($notes)
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
if (-not $NoClientId)      { $flags = $flags -bor 0x4  }  # PsfUseClientId (GPME default -> 0x14)
if ($AllowUntrustedIssuer) { $flags = $flags -bor 0x20 }  # PsfAllowUnTrustedCA

$entryApplied = $false; $adRow = 'pending'; $defaultChanged = $false; $aeApplied = $false; $dupRemoved = @()

# ---- 0. prerequisite for the AD enrollment policy row - resolved BEFORE any write ----------
# A GPO that delivers a CEP entry but no LDAP: row takes the AD enrollment policy away from every
# client in scope (see -SkipADPolicy). The domain objectGUID that row needs is therefore resolved
# first, and a lookup failure aborts the run while NOTHING has been written - it must not become
# a warning issued after the CEP entry already exists in a published GPO.
$adPid = $null
if (-not $SkipADPolicy) {
    try {
        $ldapBase = if ($Server) { "LDAP://$Server/RootDSE" } elseif ($Domain) { "LDAP://$Domain/RootDSE" } else { 'LDAP://RootDSE' }
        $dn  = ([ADSI]$ldapBase).defaultNamingContext.Value
        if (-not $dn) { throw 'RootDSE returned no defaultNamingContext' }
        $domPath = if ($Server) { "LDAP://$Server/$dn" } else { "LDAP://$dn" }
        $dom = [ADSI]$domPath
        $adPid = '{' + (New-Object Guid (, ([byte[]]$dom.Properties['objectGUID'][0]))).ToString().ToUpper() + '}'
    } catch {
        throw "Could not resolve the domain objectGUID needed for the AD Enrollment Policy row ($_). Nothing was written. Without that row, clients applying this GPO would LOSE the AD enrollment policy (autoenrollment against AD-published templates stops), so the run stops here: fix the lookup (-Server / -Domain, connectivity, permissions), or pass -SkipADPolicy to omit the row deliberately."
    }
}

# ---- 1. AD enrollment policy row FIRST (prevents losing the AD default policy fleet-wide) --
# Written and verified before the CEP entry: the GPO must never be published in the state "CEP
# entry present, LDAP: row absent" - not through a write failure (which now aborts before the CEP
# entry exists) and not through a declined confirmation (the CEP step below refuses to run then).
# "Already present" means a COMPLETE, correct row: URL 'LDAP:' AND the PolicyID this domain
# actually has. A half-written row (an interrupted earlier run) or one carrying another domain's
# GUID would not restore the AD enrollment policy on clients and must not satisfy the prerequisite.
$adRowPresent = $null -ne $adPid -and @(Get-PolEntries $preRecs | Where-Object {
        $_.Key -eq $AD_KEY -and "$($_.URL)" -eq 'LDAP:' -and "$($_.PolicyID)" -eq "$adPid" }).Count -gt 0
if ($SkipADPolicy) { $adRow = 'skipped (-SkipADPolicy)' }
else {
    $adTarget = "$baseKey\$AD_KEY"
    if ($PSCmdlet.ShouldProcess($gpoLabel, "Ensure AD Enrollment Policy row under $adTarget (URL=LDAP:, PolicyID=$adPid, Flags=0x14, Cost=0xFFFFFFFF)")) {
        try {
            Invoke-GPWrite {
                Set-GPRegistryValue @wr -Key $adTarget -ValueName URL          -Type String -Value 'LDAP:' | Out-Null
                Set-GPRegistryValue @wr -Key $adTarget -ValueName PolicyID     -Type String -Value $adPid | Out-Null
                Set-GPRegistryValue @wr -Key $adTarget -ValueName FriendlyName -Type String -Value 'Active Directory Enrollment Policy' | Out-Null
                Set-GPRegistryValue @wr -Key $adTarget -ValueName Flags        -Type DWord -Value 0x14 | Out-Null
                Set-GPRegistryValue @wr -Key $adTarget -ValueName AuthFlags    -Type DWord -Value 2 | Out-Null
                Set-GPRegistryValue @wr -Key $adTarget -ValueName Cost         -Type DWord -Value ([uint32]4294967295) | Out-Null
            }
        } catch { throw "AD policy row write to $gpoLabel failed: $_ (the CEP entry was NOT written)" }
        Test-GpoEntry $adTarget @{ URL = 'LDAP:'; PolicyID = $adPid; Flags = 0x14; AuthFlags = 2; Cost = [uint32]4294967295 }
        $adRow = 'applied'
    } else { $adRow = 'not run' }
}

function New-AddSummary([object[]]$Recs, [string]$RootApplied, [string[]]$Removed) {
    # The Add-mode result object, built from whichever registry.pol snapshot is current: the
    # final one after a full run, or the pre-run one when the run stopped at the prerequisite.
    [pscustomobject]@{
        Mode            = 'Add'
        Gpo             = $gpo.DisplayName
        GpoId           = $gpo.Id
        Scope           = $Scope
        Key             = $entryKey
        Url             = $Url
        PolicyID        = $PolicyId
        FriendlyName    = $PolicyName
        Flags           = '0x{0:X}' -f $flags
        Authentication  = '{0} (0x{1:X})' -f $Authentication, $authFlags
        Cost            = '0x{0:X}' -f $Cost
        EntryApplied    = $entryApplied
        ADPolicyRow     = $adRow
        RootFlags       = '{0} ({1})' -f (Get-RootFlagsDisplay $Recs), $RootApplied
        DefaultMarker   = "$(Get-PolValue $Recs $relBase '')"
        DefaultChanged  = $defaultChanged
        AutoEnrollment  = if ($EnableAutoEnrollmentPolicy) { "AEPolicy=$AEPolicy, $AEExpirationPercent%, '$AEStore' (applied=$aeApplied)" } else { 'not touched' }
        DuplicatesRemoved = @($Removed)
        Notes           = @($notes)
    }
}

# ---- 2. the CEP entry (individual writes - the cmdlet's list form plants **delVals.) -------
# Dependent on step 1: without an LDAP: row in the GPO (just written, already present, or
# deliberately omitted with -SkipADPolicy) the entry is NOT written and the Add workflow STOPS
# here - a declined AD-row prompt therefore declines the CEP entry and everything that builds on
# it (root Flags, (Default) marker, Auto-Enrollment, sibling cleanup). -WhatIf previews all
# steps regardless (nothing is written).
$adRowSatisfied = $SkipADPolicy -or $adRow -eq 'applied' -or $adRowPresent
if (-not $adRowSatisfied -and -not $WhatIfPreference) {
    $notes.Add("CEP entry NOT written and the run stopped here: the AD Enrollment Policy row was declined and $gpoLabel carries none. Publishing the entry without it would make clients in scope LOSE the AD enrollment policy (autoenrollment against AD-published templates stops). Nothing else was changed. Re-run and accept the AD row, or pass -SkipADPolicy to omit it deliberately.")
    foreach ($n in $notes) { Write-Warning $n }
    return New-AddSummary -Recs $preRecs -RootApplied 'unchanged' -Removed @()
}
$action = "Write CEP entry '{0}' under {1} (URL={2}, PolicyID={3}, Flags=0x{4:X}, AuthFlags=0x{5:X} {6}, Cost=0x{7:X})" -f `
          $PolicyName, $entryKey, $Url, $PolicyId, $flags, $authFlags, $Authentication, $Cost
if ($PSCmdlet.ShouldProcess($gpoLabel, $action)) {
    try {
        Invoke-GPWrite {
            Set-GPRegistryValue @wr -Key $entryKey -ValueName URL          -Type String -Value $Url | Out-Null
            Set-GPRegistryValue @wr -Key $entryKey -ValueName PolicyID     -Type String -Value $PolicyId | Out-Null
            Set-GPRegistryValue @wr -Key $entryKey -ValueName FriendlyName -Type String -Value $PolicyName | Out-Null
            Set-GPRegistryValue @wr -Key $entryKey -ValueName Flags        -Type DWord -Value ([int]$flags) | Out-Null
            Set-GPRegistryValue @wr -Key $entryKey -ValueName AuthFlags    -Type DWord -Value ([int]$authFlags) | Out-Null
            Set-GPRegistryValue @wr -Key $entryKey -ValueName Cost         -Type DWord -Value ([uint32]$Cost) | Out-Null
        }
    } catch { throw "CEP entry write to $gpoLabel failed (the entry may be partially written at an already-published GPO version - rerun after fixing the cause): $_" }
    Test-GpoEntry $entryKey @{ URL = $Url; PolicyID = $PolicyId; FriendlyName = $PolicyName; Flags = [int]$flags; AuthFlags = [int]$authFlags; Cost = [uint32]$Cost }
    $entryApplied = $true
}
# The steps that point AT the entry - the (Default) marker and the removal of the entry's stale
# siblings - are only meaningful when an entry for THIS URL that serves THIS PolicyID exists:
# written just now, or already present in the GPO from an earlier run with the same PolicyID. A
# declined CEP prompt must not leave a marker pointing at nothing, and an existing entry that
# still serves a DIFFERENT PolicyID (the declined update would have changed it) must not let the
# marker or the sibling cleanup act for the requested one. (-WhatIf still previews them.)
# A pre-existing entry counts only when COMPLETE for this request - URL and PolicyID both as
# requested (an interrupted earlier write can leave a key with a PolicyID but no URL, which no
# client can use). Re-evaluated from a fresh registry.pol read right before the sibling cleanup.
function Test-EntryComplete([object[]]$Recs) {
    @(Get-PolEntries $Recs | Where-Object { $_.Key -eq $hash -and "$($_.URL)" -eq $Url -and "$($_.PolicyID)" -eq "$PolicyId" }).Count -gt 0
}
function Test-SiblingStillSuperseded([string]$LeafKey) {
    # Live (GPMC API) check that a candidate for -ReplaceExisting removal is still what made it a
    # candidate: a different key than ours, serving the requested PolicyID under a DIFFERENT URL.
    if ($LeafKey -eq $hash -or $LeafKey -eq $AD_KEY) { return $false }
    try {
        $vals = Get-GPRegistryValue @wr -Key "$baseKey\$LeafKey" -ErrorAction Stop
        $map = @{}; foreach ($v in $vals) { $map[$v.ValueName] = $v.Value }
        return ("$($map['PolicyID'])" -eq "$PolicyId") -and ("$($map['URL'])" -ne $Url)
    } catch { return $false }
}
function Test-EntryCompleteLive {
    # Same test, but against the GPO as the GroupPolicy cmdlets see it RIGHT NOW (the API the
    # writes went through) rather than a registry.pol snapshot - used before destructive cleanup.
    try {
        $vals = Get-GPRegistryValue @wr -Key $entryKey -ErrorAction Stop
        $map = @{}; foreach ($v in $vals) { $map[$v.ValueName] = $v.Value }
        return ("$($map['URL'])" -eq $Url) -and ("$($map['PolicyID'])" -eq "$PolicyId")
    } catch { return $false }
}
$entryExists = $entryApplied -or $WhatIfPreference -or (Test-EntryComplete $preRecs)

# ---- 3. root Flags (DISABLE bits; existing bits preserved; bit-safe for high-bit values) ---
$existingRoot = Get-PolValue $preRecs $relBase 'Flags'
$newRoot = if ($null -ne $existingRoot) { [int64]$existingRoot } else { [int64]0 }
if ($newRoot -band 0x2) {
    Write-Warning 'Existing root Flags had bit 0x2 set (clients IGNORE the GP-provided policy list). Clearing it.'
    $newRoot = $newRoot -band (-bnot [int64]0x2)
}
if ($DisableUserConfigured) { $newRoot = $newRoot -bor 0x4 }
if ($EnableUserConfigured)  { $newRoot = $newRoot -band (-bnot [int64]0x4) }
$rootApplied = 'unchanged'
if (($null -eq $existingRoot) -or ([int64]$existingRoot -ne $newRoot)) {
    $from = if ($null -ne $existingRoot) { '0x{0:X}' -f [int64]$existingRoot } else { '(absent)' }
    $rootApplied = 'not run'
    if ($PSCmdlet.ShouldProcess($gpoLabel, ('Set root Flags {0} -> 0x{1:X} on {2} (disable bits: 0x2 ignore GP list, 0x4 ignore user-configured)' -f $from, $newRoot, $baseKey))) {
        Invoke-GPWrite { Set-GPRegistryValue @wr -Key $baseKey -ValueName Flags -Type DWord -Value ([uint32]$newRoot) | Out-Null }
        $rootApplied = 'applied'
    }
}

# ---- 4. (Default) marker -------------------------------------------------------------------
if ($SetAsDefault) {
    if (-not $entryExists) {
        $notes.Add("(Default) marker NOT set: the CEP entry was declined and does not exist in this GPO scope, so the marker would point at nothing.")
    }
    elseif ($PSCmdlet.ShouldProcess($gpoLabel, "Set (Default) marker = $PolicyId on $baseKey (default enrollment policy = '$PolicyName')")) {
        Invoke-GPWrite { Set-GPRegistryValue @wr -Key $baseKey -ValueName '' -Type String -Value $PolicyId | Out-Null }
        $defaultChanged = $true
    }
}
if ($ClearDefault -and (Get-PolValue $preRecs $relBase '')) {
    if ($PSCmdlet.ShouldProcess($gpoLabel, 'Clear (Default) marker')) {
        Invoke-GPWrite -TolerateNotFound { Remove-GPRegistryValue @wr -Key $baseKey -ValueName '' | Out-Null }
        $defaultChanged = $true
    }
}

# ---- 5. Auto-Enrollment (existing values respected: warn before changing) ------------------
if ($EnableAutoEnrollmentPolicy) {
    $oldAe = Get-PolValue $preRecs $relAe 'AEPolicy'
    $oldPct = Get-PolValue $preRecs $relAe 'OfflineExpirationPercent'
    $oldStore = Get-PolValue $preRecs $relAe 'OfflineExpirationStoreNames'
    if (($null -ne $oldAe -and [int64]$oldAe -ne $AEPolicy) -or ($null -ne $oldPct -and [int64]$oldPct -ne $AEExpirationPercent) -or ($null -ne $oldStore -and "$oldStore" -cne $AEStore)) {
        Write-Warning ("This GPO already carries Auto-Enrollment settings (AEPolicy={0}, {1}%, '{2}') which will be changed to (AEPolicy={3}, {4}%, '{5}')." -f $oldAe, $oldPct, $oldStore, $AEPolicy, $AEExpirationPercent, $AEStore)
    }
    if ($PSCmdlet.ShouldProcess($gpoLabel, "Enable Auto-Enrollment under $aeKey (AEPolicy=$AEPolicy, notify at $AEExpirationPercent% on '$AEStore')")) {
        Invoke-GPWrite {
            Set-GPRegistryValue @wr -Key $aeKey -ValueName AEPolicy -Type DWord -Value ([int]$AEPolicy) | Out-Null
            Set-GPRegistryValue @wr -Key $aeKey -ValueName OfflineExpirationPercent -Type DWord -Value ([int]$AEExpirationPercent) | Out-Null
            Set-GPRegistryValue @wr -Key $aeKey -ValueName OfflineExpirationStoreNames -Type String -Value $AEStore | Out-Null
        }
        $aeApplied = $true
    }
}

# ---- 6. consistency: stale duplicates, orphaned (Default), damaged deletion ordering -------
$postRecs = Read-PolRecords -Path $polPath
$entries = @(Get-PolEntries $postRecs)
# Sibling removal is judged on the state as it is NOW, read back through the same GPMC API the
# write went through: the replacing entry must be complete at this moment - "written and verified
# earlier this run" does not count, because another writer may have removed or changed it since -
# or the "superseded" siblings are the only working endpoints for this PolicyID and must stay.
$entryExists = $WhatIfPreference -or (Test-EntryCompleteLive)
$dups = @($entries | Where-Object { "$($_.PolicyID)" -eq "$PolicyId" -and $_.Key -ne $hash -and $_.Key -ne $AD_KEY })
foreach ($d in $dups) {
    if ($ReplaceExisting -and -not $entryExists) {
        $notes.Add("Superseded entry '$($d.URL)' (key $($d.Key)) NOT removed: the replacing CEP entry was declined and does not exist, so removing it would leave no endpoint for PolicyID $PolicyId.")
    }
    elseif ($ReplaceExisting) {
        if ($PSCmdlet.ShouldProcess($gpoLabel, "Remove superseded entry $baseKey\$($d.Key) (same PolicyID, URL=$($d.URL))")) {
            # Re-validate AFTER the approval - a -Confirm prompt can stay open for any length of
            # time, and another writer may have acted meanwhile: the replacement must still be
            # complete, and this sibling must still serve the requested PolicyID under a different
            # URL, both as the GPMC API sees them right now. (Invoke-GPWrite's retries only re-run
            # the removal itself, sub-second, on transient errors.)
            if (-not (Test-EntryCompleteLive)) {
                $notes.Add("Superseded entry '$($d.URL)' (key $($d.Key)) NOT removed: the replacing CEP entry is no longer complete in the GPO (changed or removed while the prompt was open).")
                continue
            }
            if (-not (Test-SiblingStillSuperseded -LeafKey $d.Key)) {
                $notes.Add("Entry at key $($d.Key) NOT removed: it no longer serves PolicyID $PolicyId under a different URL (changed or removed while the prompt was open).")
                continue
            }
            Invoke-GPWrite -TolerateNotFound { Remove-GPRegistryValue @wr -Key "$baseKey\$($d.Key)" | Out-Null }
            $dupRemoved += $d.URL
        }
    } else {
        $notes.Add("Another entry shares PolicyID $PolicyId with a different URL: '$($d.URL)' (key $($d.Key)). Stale/typo entry? Rerun with -ReplaceExisting. Intended redundant endpoint? Ignore this.")
    }
}
$finalRecs = Read-PolRecords -Path $polPath
$polKeys = @($finalRecs | ForEach-Object { $_.Key } | Where-Object { $_ -match ('^(?i)' + [regex]::Escape($relBase)) } | Sort-Object -Unique)
foreach ($pk in $polKeys) {
    $label = if ($pk -match '\\([0-9a-f]{40})$') { "entry key $($Matches[1])" } else { 'the PolicyServers root key' }
    Test-PolDeletionOrder $finalRecs $pk $label
}
$marker = Get-PolValue $finalRecs $relBase ''
if ($marker) {
    $entriesNow = @(Get-PolEntries $finalRecs)
    if (-not ($entriesNow | Where-Object { "$($_.PolicyID)" -eq "$marker" })) {
        $notes.Add("The (Default) marker points at PolicyID '$marker', which matches NO entry in this scope. Fix with -SetAsDefault on the right policy or -ClearDefault.")
    }
}
foreach ($n in $notes) { Write-Warning $n }

New-AddSummary -Recs $finalRecs -RootApplied $rootApplied -Removed $dupRemoved
