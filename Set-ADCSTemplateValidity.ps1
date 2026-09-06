<#PSScriptInfo
.VERSION 1.0.4
.GUID 48b937ae-18bd-4710-9de9-5ae76f7c9a72
.AUTHOR Sveinung Svea
.PROJECTURI https://github.com/TheOmnilord/ADCS
.LICENSEURI https://github.com/TheOmnilord/ADCS/blob/main/LICENSE
.TAGS ADCS PKI CertificateServices
.RELEASENOTES
1.0.4 - The script is now a flat body instead of begin/process/end: under Windows PowerShell 5.1, `powershell.exe -File` with a non-console stdin (a scheduler, CI, WinRM/psexec, or the `< NUL` idiom) never ran the process{} block, so the run searched nothing, changed nothing, printed nothing and exited 0. A confirmation failure (non-interactive host, ConfirmImpact High, no -Confirm:$false) is now caught, counted and emitted as an Error row instead of escaping the loop uncounted. The run-level failure is raised with a non-terminating error plus `exit 1` rather than `throw`, so the structured report is preserved for a caller that captures or pipes it while automation still sees a non-zero exit code
1.0.3 - Help text only: -OverlapPeriod documents the retained-overlap refusal; no code change
1.0.2 - When the overlap is not being set, a template whose EXISTING renewal overlap is not shorter than the new validity is reported as an error and left unchanged (previously the validity was shortened beneath the retained overlap, an invalid pair; the begin-block check only covered an explicitly supplied overlap)
1.0.1 - A run in which any search or template update failed now ends with a terminating error (non-zero exit) after the summary, instead of exit 0 with errors only in the console; help no longer advertises ? as a wildcard (LDAP substring filters only know *, so ? always matched literally - documentation corrected)
1.0.0 - Initial release
#>

<#
.SYNOPSIS
    Sets the validity period (and optionally the renewal overlap period) on one or more ADCS certificate templates.

.DESCRIPTION
    Queries Active Directory for certificate templates matching the specified name pattern(s) (wildcards supported)
    and updates their pKIExpirationPeriod attribute. Optionally updates pKIOverlapPeriod as well.

    Uses System.DirectoryServices directly - no ActiveDirectory PowerShell module required.

    After modifying templates, you may need to run 'certutil -pulse' on the CA server(s) to pick up changes.

.PARAMETER TemplateName
    One or more template CN names to match. Supports the LDAP wildcard * (any run of characters,
    including none). There is no single-character wildcard in LDAP filters: a ? matches a literal
    question mark. Examples: "WebServer", "User*", "*VPN*"

.PARAMETER ValidityPeriod
    The numeric value for the new validity period (1–9999).

.PARAMETER ValidityPeriodUnit
    The unit for ValidityPeriod: Years, Months, Weeks, Days, or Hours.
    AD uses 365 days/year and 30 days/month.

.PARAMETER OverlapPeriod
    Optional. The numeric value for the renewal overlap period (1–9999).
    Must be specified together with OverlapPeriodUnit, and must be shorter than the validity period.
    When omitted, each template keeps its existing overlap - and a template whose existing overlap
    is not shorter than the new validity (a 30-day validity over a stock 6-week overlap) is reported
    as an error and left unchanged; pass a shorter overlap to change both together.

.PARAMETER OverlapPeriodUnit
    Optional. The unit for OverlapPeriod: Years, Months, Weeks, Days, or Hours.

.PARAMETER Server
    Optional. Target a specific domain controller (not CA server) for the LDAP connection.
    Example: dc01.domain.com. This is the DC to query/write AD objects, not the Certificate Authority.

.EXAMPLE
    .\Set-ADCSTemplateValidity.ps1 -TemplateName "Web*" -ValidityPeriod 2 -ValidityPeriodUnit Years -WhatIf
    Preview which templates would be changed.

.EXAMPLE
    .\Set-ADCSTemplateValidity.ps1 -TemplateName "User*","Computer*" -ValidityPeriod 1 -ValidityPeriodUnit Years -OverlapPeriod 6 -OverlapPeriodUnit Weeks
    Set validity to 1 year and overlap to 6 weeks on all User* and Computer* templates.

.EXAMPLE
    .\Set-ADCSTemplateValidity.ps1 -TemplateName "ExactTemplate" -ValidityPeriod 365 -ValidityPeriodUnit Days -Server dc01.domain.com -Confirm:$false
    Set validity on a specific DC without confirmation prompt.
#>

# NOTE: '#Requires' deliberately sits AFTER the help comment - placed before it, Get-Help
# fails to bind the comment-based help and shows only auto-generated syntax (verified).
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory, Position = 0)]
    [string[]]$TemplateName,

    [Parameter(Mandatory)]
    [ValidateRange(1, 9999)]
    [int]$ValidityPeriod,

    [Parameter(Mandatory)]
    [ValidateSet('Years', 'Months', 'Weeks', 'Days', 'Hours')]
    [string]$ValidityPeriodUnit,

    [Parameter()]
    [ValidateRange(1, 9999)]
    [int]$OverlapPeriod,

    [Parameter()]
    [ValidateSet('Years', 'Months', 'Weeks', 'Days', 'Hours')]
    [string]$OverlapPeriodUnit,

    [Parameter()]
    [string]$Server
)

#region Helpers

function ConvertTo-PKIPeriodDays {
    param(
        [int]$Period,
        [string]$PeriodUnit
    )
    switch ($PeriodUnit) {
        'Years'  { $Period * 365 }
        'Months' { $Period * 30 }
        'Weeks'  { $Period * 7 }
        'Days'   { $Period }
        'Hours'  { $Period / 24.0 }
    }
}

function ConvertTo-PKIPeriodBytes {
    param(
        [int]$Period,
        [string]$PeriodUnit
    )
    $days = ConvertTo-PKIPeriodDays -Period $Period -PeriodUnit $PeriodUnit
    $ticks = [long][Math]::Round($days * 24 * 60 * 60 * 1e7)
    [System.BitConverter]::GetBytes(-$ticks)
}

function ConvertFrom-PKIPeriodBytes {
    param([byte[]]$Bytes)
    if ($null -eq $Bytes -or $Bytes.Length -ne 8) { return 'N/A' }
    $ticks = [System.BitConverter]::ToInt64($Bytes, 0)
    $hours = [long][Math]::Round([Math]::Abs($ticks) / (60.0 * 60 * 1e7))
    if ($hours % 24 -ne 0) {
        return "$hours hour(s)"
    }
    $days = [long]($hours / 24)
    if ($days -ge 365 -and $days % 365 -eq 0) {
        return "$([int]($days / 365)) year(s)"
    }
    if ($days -ge 30 -and $days % 30 -eq 0) {
        return "$([int]($days / 30)) month(s)"
    }
    if ($days -ge 7 -and $days % 7 -eq 0) {
        return "$([int]($days / 7)) week(s)"
    }
    return "$([int]$days) day(s)"
}

function ConvertFrom-PKIPeriodBytesToDays {
    # The exact length in days (fractional for hour-based periods) of a pKI*Period value, or
    # $null when the attribute is absent or malformed. ConvertFrom-PKIPeriodBytes renders
    # display text in the largest clean unit; this one is for comparisons.
    param([byte[]]$Bytes)
    if ($null -eq $Bytes -or $Bytes.Length -ne 8) { return $null }
    [Math]::Abs([System.BitConverter]::ToInt64($Bytes, 0)) / (24.0 * 60 * 60 * 1e7)
}

function ConvertTo-LdapFilterValue {
    # Escapes RFC 4515 filter metacharacters while preserving the * wildcard. ? is not an
    # LDAP metacharacter (RFC 4515 has no single-character wildcard), so it needs no escaping
    # and is matched literally by the server.
    param([string]$Value)
    $Value -replace '\\', '\5c' -replace '\(', '\28' -replace '\)', '\29' -replace "`0", '\00'
}

#endregion

$setOverlap = $PSBoundParameters.ContainsKey('OverlapPeriod')
$setOverlapUnit = $PSBoundParameters.ContainsKey('OverlapPeriodUnit')

if ($setOverlap -xor $setOverlapUnit) {
    throw 'OverlapPeriod and OverlapPeriodUnit must both be specified together.'
}

$validityDays = ConvertTo-PKIPeriodDays -Period $ValidityPeriod -PeriodUnit $ValidityPeriodUnit
if ($setOverlap) {
    $overlapDays = ConvertTo-PKIPeriodDays -Period $OverlapPeriod -PeriodUnit $OverlapPeriodUnit
    if ($overlapDays -ge $validityDays) {
        throw "OverlapPeriod ($OverlapPeriod $OverlapPeriodUnit) must be shorter than ValidityPeriod ($ValidityPeriod $ValidityPeriodUnit)."
    }
}

$newExpirationBytes = ConvertTo-PKIPeriodBytes -Period $ValidityPeriod -PeriodUnit $ValidityPeriodUnit
$newOverlapBytes = $null
if ($setOverlap) {
    $newOverlapBytes = ConvertTo-PKIPeriodBytes -Period $OverlapPeriod -PeriodUnit $OverlapPeriodUnit
}

Write-Verbose "New validity period : $ValidityPeriod $ValidityPeriodUnit"
if ($setOverlap) {
    Write-Verbose "New overlap period  : $OverlapPeriod $OverlapPeriodUnit"
}

# Connect to AD and resolve the Certificate Templates container
try {
    $rootDSE = if ($Server) { [ADSI]"LDAP://$Server/RootDSE" } else { [ADSI]'LDAP://RootDSE' }
    $configNC = $rootDSE.configurationNamingContext.Value
    $templateBaseDN = "CN=Certificate Templates,CN=Public Key Services,CN=Services,$configNC"
    $ldapPath = if ($Server) { "LDAP://$Server/$templateBaseDN" } else { "LDAP://$templateBaseDN" }
    $baseEntry = [ADSI]$ldapPath
    if ($null -eq $baseEntry.distinguishedName) {
        throw "Could not bind to $ldapPath"
    }
    Write-Verbose "Connected to: $ldapPath"
}
catch {
    throw "Failed to connect to Active Directory Certificate Templates container: $_"
}

$processedDNs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$modifiedCount = 0
$alreadySetCount = 0
$skippedCount = 0
$errorCount = 0
$totalMatched = 0

foreach ($pattern in $TemplateName) {
    $filter = "(&(objectClass=pKICertificateTemplate)(cn=$(ConvertTo-LdapFilterValue $pattern)))"
    Write-Verbose "Searching with filter: $filter"

    $searcher = [System.DirectoryServices.DirectorySearcher]::new($baseEntry, $filter)
    $searcher.PropertiesToLoad.AddRange(@(
        'cn', 'displayName', 'distinguishedName',
        'pKIExpirationPeriod', 'pKIOverlapPeriod',
        'msPKI-Template-Minor-Revision'
    ))
    $searcher.PageSize = 1000

    $results = $null
    try {
        try {
            $results = $searcher.FindAll()
        }
        catch {
            Write-Error "LDAP search failed for pattern '$pattern': $_"
            $errorCount++
            continue
        }

        $matchCount = 0
        foreach ($result in $results) {
            $dn = $result.Properties['distinguishedname'][0]
            $cn = $result.Properties['cn'][0]
            $displayName = if ($result.Properties['displayname'].Count -gt 0) { $result.Properties['displayname'][0] } else { $cn }

            # Deduplication. The duplicate still counts toward this pattern's match tally -
            # the pattern DID match; without this, a pattern whose every hit was already
            # processed by an earlier pattern would emit a false "no templates found" warning.
            if (-not $processedDNs.Add($dn)) {
                Write-Verbose "Skipping duplicate: $cn"
                $matchCount++
                continue
            }

            $matchCount++
            $totalMatched++

            # Decode current values
            $currentExpirationBytes = if ($result.Properties['pkiexpirationperiod'].Count -gt 0) {
                [byte[]]$result.Properties['pkiexpirationperiod'][0]
            } else { $null }
            $currentOverlapBytes = if ($result.Properties['pkioverlapperiod'].Count -gt 0) {
                [byte[]]$result.Properties['pkioverlapperiod'][0]
            } else { $null }

            $currentValidity = ConvertFrom-PKIPeriodBytes -Bytes $currentExpirationBytes
            $currentOverlap = ConvertFrom-PKIPeriodBytes -Bytes $currentOverlapBytes

            $newValidityDisplay = "$ValidityPeriod $ValidityPeriodUnit"
            $newOverlapDisplay = if ($setOverlap) { "$OverlapPeriod $OverlapPeriodUnit" } else { '(unchanged)' }

            # Skip if values are already equal
            $validityEqual = $null -ne $currentExpirationBytes -and
                [System.Linq.Enumerable]::SequenceEqual([byte[]]$currentExpirationBytes, [byte[]]$newExpirationBytes)
            $overlapEqual = (-not $setOverlap) -or (
                $null -ne $currentOverlapBytes -and
                [System.Linq.Enumerable]::SequenceEqual([byte[]]$currentOverlapBytes, [byte[]]$newOverlapBytes)
            )
            if ($validityEqual -and $overlapEqual) {
                Write-Verbose "Skipping '$cn' - already set to $newValidityDisplay"
                [PSCustomObject]@{
                    TemplateName     = $cn
                    DisplayName      = $displayName
                    PreviousValidity = $currentValidity
                    NewValidity      = $newValidityDisplay
                    PreviousOverlap  = $currentOverlap
                    NewOverlap       = $newOverlapDisplay
                    Status           = 'Already set'
                }
                $alreadySetCount++
                continue
            }

            # The renewal overlap must stay shorter than the validity. When the overlap is not
            # being set, the template's EXISTING overlap is kept - and shortening the validity
            # beneath it (30 days over a stock 6-week overlap) would leave an invalid pair that
            # the -OverlapPeriod check in begin{} never sees. Such a template is reported as an
            # error (counted in the exit code) and left untouched; pass a shorter -OverlapPeriod.
            if (-not $setOverlap) {
                $currentOverlapDays = ConvertFrom-PKIPeriodBytesToDays -Bytes $currentOverlapBytes
                if ($null -ne $currentOverlapDays -and $currentOverlapDays -ge $validityDays) {
                    Write-Error "Template '$cn': its existing renewal overlap ($currentOverlap) is not shorter than the new validity ($newValidityDisplay); left unchanged. Pass -OverlapPeriod/-OverlapPeriodUnit with a value shorter than the validity to change both together."
                    $errorCount++
                    [PSCustomObject]@{
                        TemplateName     = $cn
                        DisplayName      = $displayName
                        PreviousValidity = $currentValidity
                        NewValidity      = $newValidityDisplay
                        PreviousOverlap  = $currentOverlap
                        NewOverlap       = $newOverlapDisplay
                        Status           = 'Error: existing overlap not shorter than the new validity'
                    }
                    continue
                }
            }

            $target = "'$cn' ($displayName) -Validity: $currentValidity -> $newValidityDisplay"
            if ($setOverlap) {
                $target += ", Overlap: $currentOverlap -> $newOverlapDisplay"
            }
            $action = 'Set certificate template validity period'

            # The try WRAPS ShouldProcess as well as the update: on a non-interactive host
            # (ConfirmImpact High, no -Confirm:$false) ShouldProcess THROWS, and without this
            # that confirmation failure would escape the loop uncounted - the run would print
            # "Errors: 0" and, worse, could reach the summary as if nothing had failed. Now it
            # is caught, counted, and emitted as an Error row like any other update failure.
            try {
                if ($PSCmdlet.ShouldProcess($target, $action)) {
                    $ldapDN = if ($Server) { "LDAP://$Server/$dn" } else { "LDAP://$dn" }
                    $entry = [ADSI]$ldapDN

                    # Assign via the property cache; InvokeSet would unroll the byte
                    # array into 8 separate values through its params object[] binding
                    $entry.Properties['pKIExpirationPeriod'].Value = [byte[]]$newExpirationBytes

                    if ($setOverlap) {
                        $entry.Properties['pKIOverlapPeriod'].Value = [byte[]]$newOverlapBytes
                    }

                    # Bump minor revision so CAs detect the change
                    $currentRevision = 0
                    if ($entry.Properties['msPKI-Template-Minor-Revision'].Count -gt 0) {
                        $currentRevision = [int]$entry.Properties['msPKI-Template-Minor-Revision'][0]
                    }
                    $entry.Properties['msPKI-Template-Minor-Revision'].Value = $currentRevision + 1

                    # $null-assign: through the ADSI COM adapter, SetInfo() emits a $null
                    # onto the pipeline (verified on 5.1), which would corrupt this script's
                    # structured output with a null row per modified template.
                    $null = $entry.SetInfo()
                    $modifiedCount++

                    [PSCustomObject]@{
                        TemplateName     = $cn
                        DisplayName      = $displayName
                        PreviousValidity = $currentValidity
                        NewValidity      = $newValidityDisplay
                        PreviousOverlap  = $currentOverlap
                        NewOverlap       = $newOverlapDisplay
                        Status           = 'Modified'
                    }

                    Write-Verbose "Successfully updated: $cn"
                }
                else {
                    $skippedCount++
                    [PSCustomObject]@{
                        TemplateName     = $cn
                        DisplayName      = $displayName
                        PreviousValidity = $currentValidity
                        NewValidity      = $newValidityDisplay
                        PreviousOverlap  = $currentOverlap
                        NewOverlap       = $newOverlapDisplay
                        Status           = 'Skipped'
                    }
                }
            }
            catch {
                Write-Error "Failed to update template '$cn': $_"
                $errorCount++
                [PSCustomObject]@{
                    TemplateName     = $cn
                    DisplayName      = $displayName
                    PreviousValidity = $currentValidity
                    NewValidity      = $newValidityDisplay
                    PreviousOverlap  = $currentOverlap
                    NewOverlap       = $newOverlapDisplay
                    Status           = "Error: $_"
                }
            }
        }

        if ($matchCount -eq 0) {
            Write-Warning "No certificate templates found matching '$pattern'."
        }
    }
    finally {
        if ($null -ne $results) { $results.Dispose() }
        $searcher.Dispose()
    }
}

if ($totalMatched -gt 0) {
    Write-Host ''
    Write-Host '--- Summary ---' -ForegroundColor White
    Write-Host "  Total matched : $totalMatched" -ForegroundColor White
    Write-Host "  Modified      : $modifiedCount" -ForegroundColor $(if ($modifiedCount -gt 0) { 'Green' } else { 'White' })
    Write-Host "  Already set   : $alreadySetCount" -ForegroundColor $(if ($alreadySetCount -gt 0) { 'Yellow' } else { 'White' })
    Write-Host "  Skipped       : $skippedCount" -ForegroundColor $(if ($skippedCount -gt 0) { 'Yellow' } else { 'White' })
    Write-Host "  Errors        : $errorCount" -ForegroundColor $(if ($errorCount -gt 0) { 'Red' } else { 'White' })
    if ($modifiedCount -gt 0) {
        Write-Host "  Run 'certutil -pulse' on CA server(s) to refresh." -ForegroundColor Cyan
    }
}
# Per-template and per-pattern failures are reported as they happen (Write-Error, non-
# terminating, so the remaining templates are still processed) and counted. Automation
# gates on the exit code, not on a colored summary line - so a run in which anything failed
# must end as a failure, after the structured output and the summary are complete. A
# terminating `throw` here would tear the pipeline down and DISCARD every emitted row (a
# caller's `$report = .\script ...` would be $null), so instead this writes a non-terminating
# error and sets a non-zero process exit code - preserving the structured report AND failing
# the unattended scheduler. (Verified: both preserve the rows and exit 1 on 5.1 and 7.)
if ($errorCount -gt 0) {
    Write-Error "$errorCount search/update error(s) occurred - see the errors above. Templates that did not fail were updated." -ErrorAction Continue
    exit 1
}
