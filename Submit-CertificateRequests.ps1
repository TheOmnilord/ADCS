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
    [string]$Mode = "Submit"
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
        [string]$CAConfig
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
    }
}

function Write-Summary {
    param([PSCustomObject[]]$Data)

    Write-Log "--- Summary ---"
    $Data | Group-Object Status | ForEach-Object {
        Write-Log "  $($_.Name): $($_.Count)"
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

    Write-Log "Found $($requestFiles.Count) request file(s) in $InputPath"

    foreach ($file in $requestFiles) {
        if ($file.FullName -in $alreadySubmitted) {
            Write-Log "Skipping (already submitted): $($file.Name)"
            continue
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
            }
            catch {
                Write-Log "Error submitting $($file.Name): $_" -Level Error
                [void]$tracking.Add([PSCustomObject]@{
                    RequestFile    = $file.FullName
                    RequestID      = $null
                    SubmitTime     = (Get-Date -Format 'o')
                    Status         = 'Error'
                    OutputCertFile = ''
                    LastCheckTime  = (Get-Date -Format 'o')
                    ErrorMessage   = $_.ToString()
                })
            }
        }
    }

    Export-TrackingData -Data @($tracking) -Path $TrackingFile
    Write-Summary -Data @($tracking)
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

    foreach ($record in $pending) {
        if (-not $record.RequestID) {
            Write-Log "Skipping row without RequestID: $($record.RequestFile)" -Level Warning
            continue
        }

        if ($PSCmdlet.ShouldProcess("RequestID $($record.RequestID)", "Retrieve certificate from $CAConfig")) {
            try {
                $updated = Get-IssuedCertificate -Record $record -CAConfig $CAConfig
            }
            catch {
                Write-Log "Error retrieving RequestID $($record.RequestID): $_" -Level Error
                $record.ErrorMessage = $_.ToString()
                $record.LastCheckTime = (Get-Date -Format 'o')
            }
        }
    }

    Export-TrackingData -Data $tracking -Path $TrackingFile
    Write-Summary -Data $tracking
}

Write-Log "Done. Log file: $script:LogFile"

#endregion
