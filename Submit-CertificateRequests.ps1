<#
.SYNOPSIS
    Batch submission and retrieval of certificates via ADCS (certreq.exe).

.DESCRIPTION
    Submits all .req/.csr/.txt files from a folder to an ADCS CA,
    tracks request IDs in a CSV file, and can retrieve issued certificates
    later based on stored request IDs.

.PARAMETER InputPath
    Folder containing .req/.csr/.txt request files for submission.

.PARAMETER CAConfig
    CA configuration string for certreq, e.g. "CA01.domain.com\Contoso Issuing CA 1".

.PARAMETER CertificateTemplate
    Certificate template name used for submission.

.PARAMETER TrackingFile
    Path to the CSV file that tracks request IDs and statuses.

.PARAMETER OutputFolder
    Folder where issued certificates (.cer) are saved.

.PARAMETER Mode
    Submit   = Submit new certificate requests.
    Retrieve = Retrieve issued certificates for pending requests.
    Both     = Run Submit, then Retrieve.

.PARAMETER KeepRspFile
    By default, the .rsp file created next to each retrieved .cer is deleted.
    Specify -KeepRspFile to leave it in place.

.PARAMETER Force
    Resubmit request files that already have a tracked RequestID without prompting.
    Without -Force, the script asks y/n for each already-submitted file (default = No / skip).

.EXAMPLE
    .\Submit-CertificateRequests.ps1 -InputPath "C:\CSRs" `
        -CAConfig "CA01.domain.com\Contoso Issuing CA 1" `
        -CertificateTemplate "WebServer" -Mode Submit

.EXAMPLE
    .\Submit-CertificateRequests.ps1 -InputPath "C:\CSRs" `
        -CAConfig "CA01.domain.com\Contoso Issuing CA 1" `
        -CertificateTemplate "WebServer" -Mode Retrieve

.EXAMPLE
    .\Submit-CertificateRequests.ps1 -InputPath "C:\CSRs" `
        -CAConfig "CA01.domain.com\Contoso Issuing CA 1" `
        -CertificateTemplate "WebServer" -Mode Submit -WhatIf
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory)]
    [string]$InputPath,

    [Parameter(Mandatory)]
    [string]$CAConfig,

    [Parameter(Mandatory)]
    [string]$CertificateTemplate,

    [string]$TrackingFile = ".\CertTracking.csv",

    [string]$OutputFolder = ".\Certificates",

    [ValidateSet("Submit", "Retrieve", "Both")]
    [string]$Mode = "Submit",

    [switch]$KeepRspFile,

    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:LogFile = ".\CertBatch_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date)

#region Functions

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$timestamp] [$Level] $Message"

    switch ($Level) {
        'Warning' { Write-Warning $Message }
        'Error'   { Write-Host $entry -ForegroundColor Red }
        default   { Write-Host $entry }
    }

    $entry | Out-File -FilePath $script:LogFile -Append -Encoding utf8
}

function Test-CAConnectivity {
    param([string]$CAConfig)

    Write-Log "Testing connectivity to CA: $CAConfig"
    try {
        $output = & certutil.exe -ping -config $CAConfig 2>&1
        $exitCode = $LASTEXITCODE
        if ($exitCode -ne 0) {
            Write-Log "certutil -ping failed (exit $exitCode): $($output -join ' ')" -Level Error
            return $false
        }
        Write-Log "CA connectivity OK."
        return $true
    }
    catch {
        Write-Log "Could not reach CA: $_" -Level Error
        return $false
    }
}

function Get-RequestFiles {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        throw "InputPath does not exist: $Path"
    }

    if ((Get-Item $Path).PSIsContainer) {
        $files = @(
            Get-ChildItem -Path $Path -Filter '*.req' -File
            Get-ChildItem -Path $Path -Filter '*.csr' -File
            Get-ChildItem -Path $Path -Filter '*.txt' -File
        )
        if ($files.Count -eq 0) {
            throw "No .req/.csr/.txt files found in: $Path"
        }
        return $files
    }
    else {
        throw "InputPath must be a folder: $Path"
    }
}

function Get-RequestIdFromOutput {
    param([string[]]$Output)

    foreach ($line in $Output) {
        if ($line -match 'RequestId:\s*"?(\d+)"?') {
            return [int]$Matches[1]
        }
    }
    return $null
}

function Get-DispositionFromOutput {
    param([string[]]$Output)

    $joined = $Output -join "`n"
    if ($joined -match 'Certificate retrieved\(Issued\)') { return 'Issued' }
    if ($joined -match 'retrieved\(Issued\)')             { return 'Issued' }
    if ($joined -match 'pending|Taken Under Submission')  { return 'Pending' }
    if ($joined -match 'denied|Denied')                   { return 'Denied' }
    return 'Unknown'
}

function Get-FriendlyErrorHint {
    param(
        [string[]]$Output,
        [string]$CertificateTemplate,
        [string]$CAConfig
    )

    $joined = $Output -join "`n"
    $hints = @()

    if ($joined -match '0x80094800|CERTSRV_E_UNSUPPORTED_CERT_TYPE|template that is not supported|certificate template is not supported') {
        $hints += "Template '$CertificateTemplate' is not supported by CA '$CAConfig'. Likely causes:"
        $hints += "  - Template name is misspelled (use the template's internal name, e.g. 'WebServer', not the display name)."
        $hints += "  - Template is not published on this CA (Certification Authority MMC -> Certificate Templates -> New -> Certificate Template to Issue)."
        $hints += "  - You accidentally included the 'CertificateTemplate:' prefix in -CertificateTemplate. Pass only the template name."
        $hints += "  Verify with:  certutil -config `"$CAConfig`" -CATemplates"
    }
    elseif ($joined -match '0x80094012|CERTSRV_E_TEMPLATE_DENIED|Denied by Policy Module.*[Aa]ccess|not have permission|not allowed') {
        $hints += "Enrollment was denied by policy. Likely causes:"
        $hints += "  - The calling account lacks Enroll permission on template '$CertificateTemplate'."
        $hints += "  - Template requires Autoenroll/manager approval/signature you are not providing."
        $hints += "  - Try running the script as an account that has Enroll rights on the template."
    }
    elseif ($joined -match '0x80094004|CERTSRV_E_BAD_REQUESTSUBJECT|subject') {
        $hints += "Request subject was rejected. The CSR subject/SAN may not match what the template requires."
    }
    elseif ($joined -match '0x80070005|Access is denied|access denied') {
        $hints += "Access denied by the CA. The calling account likely lacks 'Request Certificates' rights on CA '$CAConfig'."
    }

    return $hints
}

function Import-TrackingData {
    param([string]$Path)

    if (Test-Path $Path) {
        return @(Import-Csv -Path $Path -Encoding utf8)
    }
    return @()
}

function Export-TrackingData {
    param(
        [object[]]$Data,
        [string]$Path
    )

    $filtered = @($Data | Where-Object { $_ -ne $null })
    if ($filtered.Count -eq 0) {
        Write-Log "No tracking data to export." -Level Warning
        return
    }
    $tempFile = "$Path.tmp"
    $filtered | Export-Csv -Path $tempFile -NoTypeInformation -Encoding utf8
    Move-Item -Path $tempFile -Destination $Path -Force
}

function Submit-SingleRequest {
    param(
        [System.IO.FileInfo]$RequestFile,
        [string]$CAConfig,
        [string]$CertificateTemplate,
        [string]$OutputFolder
    )

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($RequestFile.Name)
    $cerPath = Join-Path $OutputFolder "$baseName.cer"

    Write-Log "Submitting: $($RequestFile.Name)"

    $stdoutFile = [System.IO.Path]::GetTempFileName()
    $stderrFile = [System.IO.Path]::GetTempFileName()

    try {
        $proc = Start-Process -FilePath 'certreq.exe' -ArgumentList @(
            '-submit', '-f',
            '-config', "`"$CAConfig`"",
            '-attrib', "`"CertificateTemplate:$CertificateTemplate`"",
            "`"$($RequestFile.FullName)`"",
            "`"$cerPath`""
        ) -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile

        $stdout = @(Get-Content -Path $stdoutFile -ErrorAction SilentlyContinue)
        $stderr = @(Get-Content -Path $stderrFile -ErrorAction SilentlyContinue)

        $requestId = Get-RequestIdFromOutput $stdout
        $disposition = Get-DispositionFromOutput $stdout
        $errorMsg = ''

        if ($proc.ExitCode -ne 0 -and $disposition -ne 'Pending') {
            $errorMsg = ($stderr + $stdout) -join ' '
            if (-not $requestId) {
                $disposition = 'Error'
            }
            Write-Log "certreq failed for $($RequestFile.Name): $errorMsg" -Level Warning
        }

        if ($requestId) {
            Write-Log "  RequestID: $requestId - Status: $disposition"
        }
        else {
            Write-Log "  Could not parse RequestID from output" -Level Warning
            $disposition = 'Error'
            $errorMsg = "No RequestID in output: $($stdout -join ' ')"
        }

        if ($disposition -in 'Denied', 'Error') {
            $hints = Get-FriendlyErrorHint -Output ($stdout + $stderr) `
                -CertificateTemplate $CertificateTemplate -CAConfig $CAConfig
            foreach ($h in $hints) {
                Write-Log $h -Level Error
            }
        }

        return [PSCustomObject]@{
            RequestFile    = $RequestFile.FullName
            RequestID      = $requestId
            SubmitTime     = (Get-Date -Format 'o')
            Status         = $disposition
            OutputCertFile = $cerPath
            LastCheckTime  = (Get-Date -Format 'o')
            ErrorMessage   = $errorMsg
        }
    }
    finally {
        Remove-Item -Path $stdoutFile, $stderrFile -ErrorAction SilentlyContinue
    }
}

function Get-IssuedCertificate {
    param(
        [PSCustomObject]$Record,
        [string]$CAConfig,
        [switch]$KeepRspFile
    )

    Write-Log "Retrieving certificate for RequestID: $($Record.RequestID)"

    $stdoutFile = [System.IO.Path]::GetTempFileName()
    $stderrFile = [System.IO.Path]::GetTempFileName()

    try {
        $proc = Start-Process -FilePath 'certreq.exe' -ArgumentList @(
            '-retrieve', '-f',
            '-config', "`"$CAConfig`"",
            "$($Record.RequestID)",
            "`"$($Record.OutputCertFile)`""
        ) -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile

        $stdout = @(Get-Content -Path $stdoutFile -ErrorAction SilentlyContinue)
        $stderr = @(Get-Content -Path $stderrFile -ErrorAction SilentlyContinue)

        $disposition = Get-DispositionFromOutput $stdout
        $Record.Status = $disposition
        $Record.LastCheckTime = (Get-Date -Format 'o')

        switch ($disposition) {
            'Issued' {
                if (Test-Path $Record.OutputCertFile) {
                    Write-Log "  Certificate retrieved: $($Record.OutputCertFile)"
                }
                else {
                    Write-Log "  Status Issued, but .cer file was not created" -Level Warning
                }
            }
            'Pending' {
                Write-Log "  Still pending" -Level Warning
            }
            'Denied' {
                $Record.ErrorMessage = ($stderr + $stdout) -join ' '
                Write-Log "  Request DENIED" -Level Error
                $hints = Get-FriendlyErrorHint -Output ($stdout + $stderr) `
                    -CertificateTemplate '(see original submit)' -CAConfig $CAConfig
                foreach ($h in $hints) {
                    Write-Log $h -Level Error
                }
            }
            default {
                $Record.ErrorMessage = ($stderr + $stdout) -join ' '
                Write-Log "  Unknown status: $disposition" -Level Warning
            }
        }

        return $Record
    }
    finally {
        Remove-Item -Path $stdoutFile, $stderrFile -ErrorAction SilentlyContinue

        if (-not $KeepRspFile) {
            $rspPath = [System.IO.Path]::ChangeExtension($Record.OutputCertFile, '.rsp')
            if (Test-Path $rspPath) {
                Remove-Item -Path $rspPath -Force -ErrorAction SilentlyContinue
                Write-Log "  Deleted .rsp file: $rspPath"
            }
        }
    }
}

function Write-Summary {
    param(
        [PSCustomObject[]]$RunData,
        [PSCustomObject[]]$AllData,
        [string]$RunLabel = 'This run'
    )

    $run = @($RunData | Where-Object { $_ -ne $null })
    $all = @($AllData | Where-Object { $_ -ne $null })

    Write-Log "--- Summary: $RunLabel ---"
    if ($run.Count -eq 0) {
        Write-Log "  (no records processed in this run)"
    }
    else {
        $run | Group-Object Status | Sort-Object Name | ForEach-Object {
            Write-Log "  $($_.Name): $($_.Count)"
        }
        Write-Log "  Total processed this run: $($run.Count)"
    }

    Write-Log "--- Tracking file total (all history) ---"
    if ($all.Count -eq 0) {
        Write-Log "  (tracking file is empty)"
    }
    else {
        $all | Group-Object Status | Sort-Object Name | ForEach-Object {
            Write-Log "  $($_.Name): $($_.Count)"
        }
    }
    Write-Log "Tracking file: $TrackingFile"
}

#endregion

#region Main

# Validation
if (-not (Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
    Write-Log "Created output folder: $OutputFolder"
}

# CA connectivity test
if (-not (Test-CAConnectivity -CAConfig $CAConfig)) {
    throw "Cannot reach CA. Aborting."
}

# Submit mode
if ($Mode -in 'Submit', 'Both') {
    $requestFiles = @(Get-RequestFiles -Path $InputPath)
    $existingData = Import-TrackingData -Path $TrackingFile
    $tracking = [System.Collections.ArrayList]@($existingData)

    $alreadySubmitted = @($tracking | Where-Object { $_.RequestID } | Select-Object -ExpandProperty RequestFile)
    $runResults = [System.Collections.ArrayList]@()

    Write-Log "Found $($requestFiles.Count) request file(s) in $InputPath"

    foreach ($file in $requestFiles) {
        if ($file.FullName -in $alreadySubmitted) {
            $existing = @($tracking | Where-Object { $_.RequestFile -eq $file.FullName -and $_.RequestID }) |
                        Select-Object -Last 1
            $prevId = $existing.RequestID
            $prevStatus = $existing.Status

            if ($Force) {
                Write-Log "Resubmitting (-Force): $($file.Name) [previous RequestID: $prevId, Status: $prevStatus]" -Level Warning
            }
            else {
                $yes = New-Object System.Management.Automation.Host.ChoiceDescription '&Yes', 'Resubmit as a new certificate request'
                $no  = New-Object System.Management.Automation.Host.ChoiceDescription '&No',  'Skip this file'
                $choices = [System.Management.Automation.Host.ChoiceDescription[]]@($yes, $no)
                $message = "File '$($file.Name)' was already submitted (RequestID: $prevId, Status: $prevStatus). Resubmit as a new request?"
                $decision = $Host.UI.PromptForChoice('Already submitted', $message, $choices, 1)
                if ($decision -ne 0) {
                    Write-Log "Skipping (already submitted): $($file.Name)"
                    continue
                }
                Write-Log "Resubmitting on user confirmation: $($file.Name) [previous RequestID: $prevId, Status: $prevStatus]" -Level Warning
            }
        }

        if ($file.Length -eq 0) {
            Write-Log "Skipping (empty file): $($file.Name)" -Level Warning
            continue
        }

        if ($PSCmdlet.ShouldProcess($file.Name, "Submit certificate request to $CAConfig")) {
            try {
                $result = Submit-SingleRequest -RequestFile $file `
                    -CAConfig $CAConfig `
                    -CertificateTemplate $CertificateTemplate `
                    -OutputFolder $OutputFolder

                [void]$tracking.Add($result)
                [void]$runResults.Add($result)
            }
            catch {
                Write-Log "Error submitting $($file.Name): $_" -Level Error
                $errRecord = [PSCustomObject]@{
                    RequestFile    = $file.FullName
                    RequestID      = $null
                    SubmitTime     = (Get-Date -Format 'o')
                    Status         = 'Error'
                    OutputCertFile = ''
                    LastCheckTime  = (Get-Date -Format 'o')
                    ErrorMessage   = $_.ToString()
                }
                [void]$tracking.Add($errRecord)
                [void]$runResults.Add($errRecord)
            }
        }
    }

    Export-TrackingData -Data @($tracking) -Path $TrackingFile
    Write-Summary -RunData @($runResults) -AllData @($tracking) -RunLabel 'Submit'
}

# Retrieve mode
if ($Mode -in 'Retrieve', 'Both') {
    $tracking = @(Import-TrackingData -Path $TrackingFile)

    if ($tracking.Count -eq 0) {
        Write-Log "No data in tracking file. Run Submit first." -Level Warning
        return
    }

    $pending = @($tracking | Where-Object { $_.Status -eq 'Pending' })
    Write-Log "Found $($pending.Count) pending request(s)"

    $runResults = [System.Collections.ArrayList]@()

    foreach ($record in $pending) {
        if (-not $record.RequestID) {
            Write-Log "Skipping row without RequestID: $($record.RequestFile)" -Level Warning
            continue
        }

        if ($PSCmdlet.ShouldProcess("RequestID $($record.RequestID)", "Retrieve certificate from $CAConfig")) {
            try {
                $updated = Get-IssuedCertificate -Record $record -CAConfig $CAConfig -KeepRspFile:$KeepRspFile
            }
            catch {
                Write-Log "Error retrieving RequestID $($record.RequestID): $_" -Level Error
                $record.ErrorMessage = $_.ToString()
                $record.LastCheckTime = (Get-Date -Format 'o')
                $record.Status = 'Error'
            }
            [void]$runResults.Add($record)
        }
    }

    Export-TrackingData -Data $tracking -Path $TrackingFile
    Write-Summary -RunData @($runResults) -AllData $tracking -RunLabel 'Retrieve'
}

Write-Log "Done. Log file: $script:LogFile"

#endregion
