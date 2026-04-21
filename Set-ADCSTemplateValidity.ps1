#Requires -Version 5.1
<#
.SYNOPSIS
    Sets the validity period (and optionally the renewal overlap period) on one or more ADCS certificate templates.

.DESCRIPTION
    Queries Active Directory for certificate templates matching the specified name pattern(s) (wildcards supported)
    and updates their pKIExpirationPeriod attribute. Optionally updates pKIOverlapPeriod as well.

    Uses System.DirectoryServices directly - no ActiveDirectory PowerShell module required.

    After modifying templates, you may need to run 'certutil -pulse' on the CA server(s) to pick up changes.

.PARAMETER TemplateName
    One or more template CN names to match. Supports LDAP wildcards (* and ?).
    Examples: "WebServer", "User*", "*VPN*"

.PARAMETER ValidityPeriod
    The numeric value for the new validity period (1–9999).

.PARAMETER ValidityPeriodUnit
    The unit for ValidityPeriod: Years, Months, Weeks, Days, or Hours.
    AD uses 365 days/year and 30 days/month.

.PARAMETER OverlapPeriod
    Optional. The numeric value for the renewal overlap period (1–9999).
    Must be specified together with OverlapPeriodUnit.

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

begin {
    #region Helpers

    function ConvertTo-PKIPeriodBytes {
        param(
            [int]$Period,
            [string]$PeriodUnit
        )
        $days = switch ($PeriodUnit) {
            'Years'  { $Period * 365 }
            'Months' { $Period * 30 }
            'Weeks'  { $Period * 7 }
            'Days'   { $Period }
            'Hours'  { $Period / 24.0 }
        }
        $ticks = [long]($days * 24 * 60 * 60 * 1e7)
        [System.BitConverter]::GetBytes(-$ticks)
    }

    function ConvertFrom-PKIPeriodBytes {
        param([byte[]]$Bytes)
        if ($null -eq $Bytes -or $Bytes.Length -ne 8) { return 'N/A' }
        $ticks = [System.BitConverter]::ToInt64($Bytes, 0)
        $days = [Math]::Abs($ticks) / (24.0 * 60 * 60 * 1e7)
        if ($days -ge 365 -and $days % 365 -eq 0) {
            $val = [int]($days / 365)
            return "$val year(s)"
        }
        elseif ($days -ge 30 -and $days % 30 -eq 0) {
            $val = [int]($days / 30)
            return "$val month(s)"
        }
        elseif ($days -ge 7 -and $days % 7 -eq 0) {
            $val = [int]($days / 7)
            return "$val week(s)"
        }
        elseif ($days -eq [Math]::Floor($days)) {
            return "$([int]$days) day(s)"
        }
        else {
            $hours = $days * 24
            return "$([int]$hours) hour(s)"
        }
    }

    #endregion

    $setOverlap = $PSBoundParameters.ContainsKey('OverlapPeriod')
    $setOverlapUnit = $PSBoundParameters.ContainsKey('OverlapPeriodUnit')

    if ($setOverlap -xor $setOverlapUnit) {
        throw 'OverlapPeriod and OverlapPeriodUnit must both be specified together.'
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
        $rootDSE = [ADSI]'LDAP://RootDSE'
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
}

process {
    foreach ($pattern in $TemplateName) {
        $filter = "(&(objectClass=pKICertificateTemplate)(cn=$pattern))"
        Write-Verbose "Searching with filter: $filter"

        $searcher = [System.DirectoryServices.DirectorySearcher]::new($baseEntry, $filter)
        $searcher.PropertiesToLoad.AddRange(@(
            'cn', 'displayName', 'distinguishedName',
            'pKIExpirationPeriod', 'pKIOverlapPeriod',
            'msPKI-Template-Minor-Revision'
        ))
        $searcher.PageSize = 1000

        try {
            $results = $searcher.FindAll()
        }
        catch {
            Write-Error "LDAP search failed for pattern '$pattern': $_"
            continue
        }

        $matchCount = 0
        foreach ($result in $results) {
            $dn = $result.Properties['distinguishedname'][0]
            $cn = $result.Properties['cn'][0]
            $displayName = if ($result.Properties['displayname'].Count -gt 0) { $result.Properties['displayname'][0] } else { $cn }

            # Deduplication
            if (-not $processedDNs.Add($dn)) {
                Write-Verbose "Skipping duplicate: $cn"
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

            $target = "'$cn' ($displayName) -Validity: $currentValidity -> $newValidityDisplay"
            if ($setOverlap) {
                $target += ", Overlap: $currentOverlap -> $newOverlapDisplay"
            }
            $action = 'Set certificate template validity period'

            if ($PSCmdlet.ShouldProcess($target, $action)) {
                try {
                    $ldapDN = if ($Server) { "LDAP://$Server/$dn" } else { "LDAP://$dn" }
                    $entry = [ADSI]$ldapDN

                    $entry.InvokeSet('pKIExpirationPeriod', [byte[]]$newExpirationBytes)

                    if ($setOverlap) {
                        $entry.InvokeSet('pKIOverlapPeriod', [byte[]]$newOverlapBytes)
                    }

                    # Bump minor revision so CAs detect the change
                    $currentRevision = 0
                    if ($entry.Properties['msPKI-Template-Minor-Revision'].Count -gt 0) {
                        $currentRevision = [int]$entry.Properties['msPKI-Template-Minor-Revision'][0]
                    }
                    $entry.Properties['msPKI-Template-Minor-Revision'].Value = $currentRevision + 1

                    $entry.SetInfo()
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

        $results.Dispose()
        $searcher.Dispose()

        if ($matchCount -eq 0) {
            Write-Warning "No certificate templates found matching '$pattern'."
        }
    }
}

end {
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
}
