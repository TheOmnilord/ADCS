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
    per PolicyID is also the legitimate redundant-endpoint pattern.

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
        if (-not (Test-Path -LiteralPath $parent)) { New-Item -Path $parent -Force | Out-Null }  # -Force only creates missing parents here (guarded)
        New-Item -Path $EntryPath | Out-Null   # leaf WITHOUT -Force: never recreate/wipe an existing key
    }
    foreach ($n in $Strings.Keys) { Set-ItemProperty -LiteralPath $EntryPath -Name $n -Value $Strings[$n] }
    foreach ($n in $Dwords.Keys)  { New-ItemProperty -LiteralPath $EntryPath -Name $n -Value $Dwords[$n] -PropertyType DWord -Force | Out-Null }
    $chk = Get-Item -LiteralPath $EntryPath
    $bad = @()
    foreach ($n in $Strings.Keys) {
        $got = $chk.GetValue($n)
        if ($null -eq $got -or [string]$got -cne [string]$Strings[$n]) { $bad += $n }
    }
    foreach ($n in $Dwords.Keys) {
        $got = $chk.GetValue($n)
        if ($null -eq $got -or [int]$got -ne [int]$Dwords[$n]) { $bad += $n }
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
            try { Remove-Item -LiteralPath $target -Recurse -Force } catch { throw "Failed to remove ${target}: $_" }
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

# ---- 1. the CEP entry ----------------------------------------------------------------------
$verb = if (Test-Path -LiteralPath $target) { 'Update' } else { 'Create' }
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

# ---- 2. AD enrollment policy row (GP locations; prevents losing the AD default policy) -----
if ($isGP) {
    if ($SkipADPolicy) { $adRow = 'skipped (-SkipADPolicy)' }
    else {
        $adPid = $null
        try {
            $dn  = ([ADSI]'LDAP://RootDSE').defaultNamingContext.Value
            $dom = [ADSI]("LDAP://$dn")
            $adPid = '{' + (New-Object Guid (, ([byte[]]$dom.Properties['objectGUID'][0]))).ToString().ToUpper() + '}'
        } catch {
            Write-Warning "Could not resolve the domain objectGUID (workgroup machine?). Skipping the AD policy row - clients using this GP configuration will NOT see the AD enrollment policy. ($_)"
            $adRow = 'skipped (no domain)'
        }
        if ($adPid) {
            $adTarget = "$hive\$AD_KEY"
            if ($PSCmdlet.ShouldProcess($adTarget, "Ensure AD Enrollment Policy row (URL=LDAP:, PolicyID=$adPid, Flags=0x14, Cost=0xFFFFFFFF)")) {
                try {
                    Write-CepEntry -EntryPath $adTarget `
                        -Strings @{ URL = 'LDAP:'; PolicyID = $adPid; FriendlyName = 'Active Directory Enrollment Policy' } `
                        -Dwords  @{ Flags = 0x14; AuthFlags = 2; Cost = (ConvertTo-DwordInt 4294967295) }
                } catch { throw "AD policy row write to $adTarget failed: $_" }
                $adRow = 'applied'
            } else { $adRow = 'not run' }
        }
    }
}

# ---- 3. root Flags (GP locations; DISABLE bits, preserved across runs) ---------------------
if ($isGP) {
    $existing = if (Test-Path -LiteralPath $hive) { (Get-Item -LiteralPath $hive).GetValue('Flags') } else { $null }
    $newFlags = if ($null -ne $existing) { [int]$existing } else { 0 }
    if ($newFlags -band 0x2) {
        Write-Warning 'Existing root Flags had bit 0x2 set (clients IGNORE the GP-provided policy list). Clearing it.'
        $newFlags = $newFlags -band (-bnot 0x2)
    }
    if ($DisableUserConfigured) { $newFlags = $newFlags -bor 0x4 }
    if ($EnableUserConfigured)  { $newFlags = $newFlags -band (-bnot 0x4) }
    if (($null -eq $existing) -or ([int]$existing -ne $newFlags)) {
        $from = if ($null -ne $existing) { '0x{0:X}' -f [int]$existing } else { '(absent)' }
        if ($PSCmdlet.ShouldProcess($hive, ('Set root Flags {0} -> 0x{1:X} (disable bits: 0x2 ignore GP list, 0x4 ignore user-configured)' -f $from, $newFlags))) {
            if (-not (Test-Path -LiteralPath $hive)) { New-Item -Path $hive -Force | Out-Null }
            New-ItemProperty -LiteralPath $hive -Name Flags -Value $newFlags -PropertyType DWord -Force | Out-Null
        }
    }
}

# ---- 4. (Default) marker -------------------------------------------------------------------
if ($SetAsDefault) {
    if ($PSCmdlet.ShouldProcess($hive, "Set (Default) marker = $PolicyId (default enrollment policy = '$PolicyName')")) {
        if (-not (Test-Path -LiteralPath $hive)) { New-Item -Path $hive -Force | Out-Null }
        Set-ItemProperty -LiteralPath $hive -Name '(default)' -Value $PolicyId
        $defaultChanged = $true
    }
}
if ($ClearDefault -and (Get-DefaultMarker)) {
    if ($PSCmdlet.ShouldProcess($hive, 'Clear (Default) marker')) { Remove-DefaultMarker; $defaultChanged = $true }
}

# ---- 5. consistency checks: stale duplicates and orphaned default --------------------------
$entries = @(Get-CepEntries)
$dups = @($entries | Where-Object { "$($_.PolicyID)" -eq "$PolicyId" -and $_.Key -ne $key -and $_.Key -ne $AD_KEY })
foreach ($d in $dups) {
    if ($ReplaceExisting) {
        if ($PSCmdlet.ShouldProcess("$hive\$($d.Key)", "Remove superseded entry with same PolicyID (URL=$($d.URL))")) {
            Remove-Item -LiteralPath "$hive\$($d.Key)" -Recurse -Force
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
    DuplicatesRemoved = @($dupRemoved)
    Notes          = @($notes)
}
