<#PSScriptInfo
.VERSION 1.0.1
.GUID 6f98f16e-0c56-4a72-ba31-443938175c06
.AUTHOR Sveinung Svea
.PROJECTURI https://github.com/TheOmnilord/ADCS
.LICENSEURI https://github.com/TheOmnilord/ADCS/blob/main/LICENSE
.TAGS ADCS PKI CertificateServices
.RELEASENOTES
1.0.1 - Retrieve mode honours an explicit -OutputFolder (previously the path recorded at submit time was always used)
1.0.0 - Initial release
#>

<#
.SYNOPSIS
    Batch submission and retrieval of certificates via ADCS (certreq.exe).

.DESCRIPTION
    Submits all .req/.csr/.txt files from a folder to an ADCS CA,
    tracks request IDs in a CSV file, and can retrieve issued certificates
    later based on stored request IDs.

.PARAMETER InputPath
    Folder containing .req/.csr/.txt request files for submission.
    Required when -Mode is Submit or Both. Not used, and not required, for -Mode Retrieve.

.PARAMETER CAConfig
    CA configuration string for certreq, e.g. "CA01.domain.com\Contoso Issuing CA 1".
    Always required.

.PARAMETER CertificateTemplate
    Certificate template name used for submission.
    Required when -Mode is Submit or Both. Not used, and not required, for -Mode Retrieve.

.PARAMETER TrackingFile
    Path to the CSV file that tracks request IDs and statuses.
    Relative paths are resolved against the current directory and stored as absolute paths.

.PARAMETER OutputFolder
    Folder where issued certificates (.cer) are saved.
    Relative paths are resolved against the current directory and stored as absolute paths.
    In Retrieve mode each row is written to the path recorded at submit time; pass -OutputFolder
    explicitly to redirect retrieved .cer files there instead (the tracking row is updated).

.PARAMETER Mode
    Submit   = Submit new certificate requests.
    Retrieve = Retrieve issued certificates for unresolved requests.
    Both     = Run Submit, then Retrieve.

.PARAMETER KeepRspFile
    By default, the .rsp file created next to each .cer (at both submit and retrieve) is deleted.
    Specify -KeepRspFile to leave it in place.

.PARAMETER Force
    Resubmit request files that already have a tracked RequestID without prompting.
    Without -Force, the script asks y/n for each already-submitted file (default = No / skip).

.EXAMPLE
    .\Submit-CertificateRequests.ps1 -InputPath "C:\CSRs" `
        -CAConfig "CA01.domain.com\Contoso Issuing CA 1" `
        -CertificateTemplate "WebServer" -Mode Submit

.EXAMPLE
    .\Submit-CertificateRequests.ps1 -CAConfig "CA01.domain.com\Contoso Issuing CA 1" -Mode Retrieve

.EXAMPLE
    .\Submit-CertificateRequests.ps1 -InputPath "C:\CSRs" `
        -CAConfig "CA01.domain.com\Contoso Issuing CA 1" `
        -CertificateTemplate "WebServer" -Mode Submit -WhatIf
#>

# NOTE: '#Requires' deliberately sits AFTER the help comment - placed before it, Get-Help
# fails to bind the comment-based help and shows only auto-generated syntax (verified).
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [string]$InputPath,

    [Parameter(Mandatory)]
    [string]$CAConfig,

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

#region Functions

function Resolve-FullPath {
    param([Parameter(Mandatory)][string]$Path)

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }
    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location).ProviderPath $Path))
}

function Write-BatchLog {
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
        'Error'   { Write-Error -Message $entry -ErrorAction Continue }
        default   { Write-Host $entry }
    }

    if (-not $script:SuppressLogFile) {
        $entry | Out-File -FilePath $script:LogFile -Append -Encoding utf8
    }
}

function Test-CAConnectivity {
    param([string]$CAConfig)

    Write-BatchLog "Testing connectivity to CA: $CAConfig"
    try {
        # Localized EAP: with $ErrorActionPreference = 'Stop', 2>&1 turns native
        # stderr lines into terminating errors in Windows PowerShell 5.1
        $output = & {
            $ErrorActionPreference = 'Continue'
            certutil.exe -ping -config $CAConfig 2>&1 | ForEach-Object { "$_" }
        }
        $exitCode = $LASTEXITCODE
        if ($exitCode -ne 0) {
            Write-BatchLog "certutil -ping failed (exit $exitCode): $($output -join ' ')" -Level Error
            return $false
        }
        Write-BatchLog "CA connectivity OK."
        return $true
    }
    catch {
        Write-BatchLog "Could not reach CA: $_" -Level Error
        return $false
    }
}

function Get-RequestFiles {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        Write-BatchLog "InputPath does not exist: $Path" -Level Error
        return @()
    }

    if (-not (Get-Item $Path).PSIsContainer) {
        Write-BatchLog "InputPath must be a folder, not a file: $Path" -Level Error
        return @()
    }

    $files = @(
        Get-ChildItem -Path $Path -Filter '*.req' -File
        Get-ChildItem -Path $Path -Filter '*.csr' -File
        Get-ChildItem -Path $Path -Filter '*.txt' -File
    )

    if ($files.Count -eq 0) {
        Write-BatchLog "No .req/.csr/.txt files found in: $Path" -Level Warning
        $other = @(Get-ChildItem -Path $Path -File -ErrorAction SilentlyContinue)
        if ($other.Count -gt 0) {
            $sample = ($other | Select-Object -First 5 -ExpandProperty Name) -join ', '
            $suffix = if ($other.Count -gt 5) { ", ..." } else { '' }
            Write-BatchLog "  Folder contains $($other.Count) other file(s): $sample$suffix" -Level Warning
            Write-BatchLog "  Rename them to .req/.csr/.txt or move CSRs into this folder." -Level Warning
        }
        else {
            Write-BatchLog "  Folder is empty." -Level Warning
        }
    }

    return $files
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
    if ($joined -match 'retrieved\(Issued\)')            { return 'Issued' }
    if ($joined -match 'pending|Taken Under Submission') { return 'Pending' }
    if ($joined -match 'denied')                         { return 'Denied' }
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
        Write-BatchLog "No tracking data to export." -Level Warning
        return
    }
    $tempFile = "$Path.tmp"
    $filtered | Export-Csv -Path $tempFile -NoTypeInformation -Encoding utf8
    Move-Item -Path $tempFile -Destination $Path -Force
}

function Remove-RspFile {
    param([string]$CerPath)

    $rspPath = [System.IO.Path]::ChangeExtension($CerPath, '.rsp')
    if (Test-Path $rspPath) {
        Remove-Item -Path $rspPath -Force -ErrorAction SilentlyContinue
        Write-BatchLog "  Deleted .rsp file: $rspPath"
    }
}

function Submit-SingleRequest {
    param(
        [System.IO.FileInfo]$RequestFile,
        [string]$CAConfig,
        [string]$CertificateTemplate,
        [string]$CerPath,
        [switch]$KeepRspFile
    )

    Write-BatchLog "Submitting: $($RequestFile.Name)"

    $stdoutFile = [System.IO.Path]::GetTempFileName()
    $stderrFile = [System.IO.Path]::GetTempFileName()

    try {
        # -q is load-bearing: without it certreq may pop a MODAL GUI dialog on some error paths
        # (observed live: an unsupported-template denial), hanging a batch/headless run forever.
        $proc = Start-Process -FilePath 'certreq.exe' -ArgumentList @(
            '-submit', '-q', '-f',
            '-config', "`"$CAConfig`"",
            '-attrib', "`"CertificateTemplate:$CertificateTemplate`"",
            "`"$($RequestFile.FullName)`"",
            "`"$CerPath`""
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
            Write-BatchLog "certreq failed for $($RequestFile.Name): $errorMsg" -Level Warning
        }

        if ($requestId) {
            Write-BatchLog "  RequestID: $requestId - Status: $disposition"
        }
        else {
            Write-BatchLog "  Could not parse RequestID from output" -Level Warning
            $disposition = 'Error'
            $errorMsg = "No RequestID in output: $($stdout -join ' ')"
        }

        if ($disposition -in 'Denied', 'Error') {
            $hints = Get-FriendlyErrorHint -Output ($stdout + $stderr) `
                -CertificateTemplate $CertificateTemplate -CAConfig $CAConfig
            foreach ($h in $hints) {
                Write-BatchLog $h -Level Error
            }
        }

        return [PSCustomObject]@{
            RequestFile    = $RequestFile.FullName
            RequestID      = $requestId
            SubmitTime     = (Get-Date -Format 'o')
            Status         = $disposition
            OutputCertFile = $CerPath
            LastCheckTime  = (Get-Date -Format 'o')
            ErrorMessage   = $errorMsg
        }
    }
    finally {
        Remove-Item -Path $stdoutFile, $stderrFile -ErrorAction SilentlyContinue

        if (-not $KeepRspFile) {
            Remove-RspFile -CerPath $CerPath
        }
    }
}

function Get-IssuedCertificate {
    param(
        [PSCustomObject]$Record,
        [string]$CAConfig,
        [switch]$KeepRspFile
    )

    Write-BatchLog "Retrieving certificate for RequestID: $($Record.RequestID)"

    $stdoutFile = [System.IO.Path]::GetTempFileName()
    $stderrFile = [System.IO.Path]::GetTempFileName()

    try {
        # -q for the same reason as in Submit-SingleRequest: never let certreq raise UI.
        $null = Start-Process -FilePath 'certreq.exe' -ArgumentList @(
            '-retrieve', '-q', '-f',
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
                    Write-BatchLog "  Certificate retrieved: $($Record.OutputCertFile)"
                }
                else {
                    Write-BatchLog "  Status Issued, but .cer file was not created" -Level Warning
                }
            }
            'Pending' {
                Write-BatchLog "  Still pending" -Level Warning
            }
            'Denied' {
                $Record.ErrorMessage = ($stderr + $stdout) -join ' '
                Write-BatchLog "  Request DENIED" -Level Error
                $hints = Get-FriendlyErrorHint -Output ($stdout + $stderr) `
                    -CertificateTemplate '(see original submit)' -CAConfig $CAConfig
                foreach ($h in $hints) {
                    Write-BatchLog $h -Level Error
                }
            }
            default {
                $Record.ErrorMessage = ($stderr + $stdout) -join ' '
                Write-BatchLog "  Unknown status: $disposition" -Level Warning
            }
        }
    }
    finally {
        Remove-Item -Path $stdoutFile, $stderrFile -ErrorAction SilentlyContinue

        if (-not $KeepRspFile) {
            Remove-RspFile -CerPath $Record.OutputCertFile
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

    Write-BatchLog "--- Summary: $RunLabel ---"
    if ($run.Count -eq 0) {
        Write-BatchLog "  (no records processed in this run)"
    }
    else {
        $run | Group-Object Status | Sort-Object Name | ForEach-Object {
            Write-BatchLog "  $($_.Name): $($_.Count)"
        }
        Write-BatchLog "  Total processed this run: $($run.Count)"
    }

    Write-BatchLog "--- Tracking file total (all history) ---"
    if ($all.Count -eq 0) {
        Write-BatchLog "  (tracking file is empty)"
    }
    else {
        $all | Group-Object Status | Sort-Object Name | ForEach-Object {
            Write-BatchLog "  $($_.Name): $($_.Count)"
        }
    }
    Write-BatchLog "Tracking file: $TrackingFile"
}

#endregion

#region Main

# Resolve to absolute paths so tracking records stay valid when later runs
# use a different working directory
$TrackingFile = Resolve-FullPath -Path $TrackingFile
$OutputFolder = Resolve-FullPath -Path $OutputFolder
$script:LogFile = Resolve-FullPath -Path (".\CertBatch_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))
$script:SuppressLogFile = [bool]$WhatIfPreference

# Validation
if ($Mode -in 'Submit', 'Both') {
    if (-not $InputPath) {
        throw "-InputPath is required when -Mode is '$Mode'."
    }
    if (-not $CertificateTemplate) {
        throw "-CertificateTemplate is required when -Mode is '$Mode'."
    }
}

if (-not (Test-Path $OutputFolder)) {
    if ($PSCmdlet.ShouldProcess($OutputFolder, 'Create output folder')) {
        New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
        Write-BatchLog "Created output folder: $OutputFolder"
    }
}

# CA connectivity test
if (-not (Test-CAConnectivity -CAConfig $CAConfig)) {
    throw "Cannot reach CA. Aborting."
}

# Submit mode
if ($Mode -in 'Submit', 'Both') {
    $requestFiles = @(Get-RequestFiles -Path $InputPath)

    if ($requestFiles.Count -eq 0) {
        Write-BatchLog "Nothing to submit. Skipping Submit phase." -Level Warning
    }
    else {
        $existingData = Import-TrackingData -Path $TrackingFile
        $tracking = [System.Collections.ArrayList]@($existingData)

        $alreadySubmitted = @($tracking | Where-Object { $_.RequestID } | Select-Object -ExpandProperty RequestFile)
        $runResults = [System.Collections.ArrayList]@()

        # Files sharing a base name (a.req + a.csr) would collide on a.cer;
        # those keep the full file name in their .cer name instead
        $duplicateBaseNames = @(
            $requestFiles |
                Group-Object { [System.IO.Path]::GetFileNameWithoutExtension($_.Name) } |
                Where-Object { $_.Count -gt 1 } |
                Select-Object -ExpandProperty Name
        )

        Write-BatchLog "Found $($requestFiles.Count) request file(s) in $InputPath"

        foreach ($file in $requestFiles) {
            if ($file.Length -eq 0) {
                Write-BatchLog "Skipping (empty file): $($file.Name)" -Level Warning
                continue
            }

            if (-not $PSCmdlet.ShouldProcess($file.Name, "Submit certificate request to $CAConfig")) {
                continue
            }

            if ($file.FullName -in $alreadySubmitted) {
                $existing = @($tracking | Where-Object { $_.RequestFile -eq $file.FullName -and $_.RequestID }) |
                            Select-Object -Last 1
                $prevId = $existing.RequestID
                $prevStatus = $existing.Status

                if ($Force) {
                    Write-BatchLog "Resubmitting (-Force): $($file.Name) [previous RequestID: $prevId, Status: $prevStatus]" -Level Warning
                }
                else {
                    $yes = New-Object System.Management.Automation.Host.ChoiceDescription '&Yes', 'Resubmit as a new certificate request'
                    $no  = New-Object System.Management.Automation.Host.ChoiceDescription '&No',  'Skip this file'
                    $choices = [System.Management.Automation.Host.ChoiceDescription[]]@($yes, $no)
                    $message = "File '$($file.Name)' was already submitted (RequestID: $prevId, Status: $prevStatus). Resubmit as a new request?"
                    try {
                        $decision = $Host.UI.PromptForChoice('Already submitted', $message, $choices, 1)
                    }
                    catch {
                        Write-BatchLog "Host cannot prompt; skipping already-submitted file: $($file.Name). Use -Force to resubmit." -Level Warning
                        continue
                    }
                    if ($decision -ne 0) {
                        Write-BatchLog "Skipping (already submitted): $($file.Name)"
                        continue
                    }
                    Write-BatchLog "Resubmitting on user confirmation: $($file.Name) [previous RequestID: $prevId, Status: $prevStatus]" -Level Warning
                }
            }

            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
            $cerName = if ($baseName -in $duplicateBaseNames) { "$($file.Name).cer" } else { "$baseName.cer" }
            $cerPath = Join-Path $OutputFolder $cerName

            try {
                $result = Submit-SingleRequest -RequestFile $file `
                    -CAConfig $CAConfig `
                    -CertificateTemplate $CertificateTemplate `
                    -CerPath $cerPath `
                    -KeepRspFile:$KeepRspFile
            }
            catch {
                Write-BatchLog "Error submitting $($file.Name): $_" -Level Error
                $result = [PSCustomObject]@{
                    RequestFile    = $file.FullName
                    RequestID      = $null
                    SubmitTime     = (Get-Date -Format 'o')
                    Status         = 'Error'
                    OutputCertFile = ''
                    LastCheckTime  = (Get-Date -Format 'o')
                    ErrorMessage   = $_.ToString()
                }
            }

            [void]$tracking.Add($result)
            [void]$runResults.Add($result)

            # Persist after every file so request IDs survive a mid-batch crash
            Export-TrackingData -Data @($tracking) -Path $TrackingFile
        }

        Write-Summary -RunData @($runResults) -AllData @($tracking) -RunLabel 'Submit'
    }
}

# Retrieve mode
if ($Mode -in 'Retrieve', 'Both') {
    $tracking = @(Import-TrackingData -Path $TrackingFile)

    if ($tracking.Count -eq 0) {
        Write-BatchLog "No data in tracking file. Run Submit first." -Level Warning
    }
    else {
        # Retry anything with a RequestID that is not finally resolved - including
        # rows stuck in 'Unknown' or 'Error' that the CA may have issued since
        $unresolved = @($tracking | Where-Object { $_.RequestID -and $_.Status -notin 'Issued', 'Denied' })
        Write-BatchLog "Found $($unresolved.Count) unresolved request(s) with a RequestID"

        $runResults = [System.Collections.ArrayList]@()

        # The .cer path is recorded per row at submit time. An explicit -OutputFolder on a
        # Retrieve run overrides it (the file name is kept); the default is not applied here,
        # so a Retrieve from another working directory still lands where Submit put it.
        $redirectOutput = $PSBoundParameters.ContainsKey('OutputFolder')
        if ($redirectOutput) {
            Write-BatchLog "Retrieved certificates will be written to: $OutputFolder"
        }

        foreach ($record in $unresolved) {
            if ($redirectOutput) {
                $cerName = if ($record.OutputCertFile) {
                    [System.IO.Path]::GetFileName($record.OutputCertFile)
                }
                else {
                    [System.IO.Path]::GetFileNameWithoutExtension($record.RequestFile) + '.cer'
                }
                $record.OutputCertFile = Join-Path $OutputFolder $cerName
            }

            if ($PSCmdlet.ShouldProcess("RequestID $($record.RequestID)", "Retrieve certificate from $CAConfig")) {
                try {
                    Get-IssuedCertificate -Record $record -CAConfig $CAConfig -KeepRspFile:$KeepRspFile
                }
                catch {
                    Write-BatchLog "Error retrieving RequestID $($record.RequestID): $_" -Level Error
                    $record.ErrorMessage = $_.ToString()
                    $record.LastCheckTime = (Get-Date -Format 'o')
                    $record.Status = 'Error'
                }
                [void]$runResults.Add($record)

                # Persist after every record so status updates survive a mid-batch crash
                Export-TrackingData -Data $tracking -Path $TrackingFile
            }
        }

        Write-Summary -RunData @($runResults) -AllData $tracking -RunLabel 'Retrieve'
    }
}

if ($script:SuppressLogFile) {
    Write-BatchLog "Done. (-WhatIf run: log file not written)"
}
else {
    Write-BatchLog "Done. Log file: $script:LogFile"
}

#endregion
