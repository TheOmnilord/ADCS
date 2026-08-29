<#
.SYNOPSIS
    Pester suite for Submit-CertificateRequests.ps1. Requires Pester 5+.

.DESCRIPTION
    Three always-on tiers plus one opt-in tier:

      -Tag Unit    Pure helpers extracted from the script by AST (so the REAL code runs, never a
                   copy): path resolution, certreq output parsing (RequestID/disposition), the
                   friendly-error-hint mapping, the tracking-CSV round-trip, and .rsp cleanup.
                   No CA, no network, no changes outside $TestDrive.
      -Tag Static  The script parses and its comment-based help binds. No CA.
      -Tag Guard   Validation throws that fire BEFORE any CA is contacted (missing -InputPath /
                   -CertificateTemplate), plus the connectivity pre-check against a CA that can
                   never exist (localhost) - certutil.exe is standard on Windows, so this runs on
                   CI. All Guard invocations use -WhatIf, so no folder or log file is created.
      -Tag Lab     LIVE submissions to a real Enterprise CA. Skipped unless -RunLab is passed.
                   Requires an enrollable template (default WebServer - subject supplied in the
                   request, Enroll for Domain Admins). Surgical by construction: CSR subjects and
                   key containers carry a per-run PESTER-<hex> prefix, every CA RequestID this run
                   creates is tracked and its CA database row deleted in teardown by exact ID
                   (certutil -deleterow), pending-request store entries are removed by the
                   run-scoped subject filter (which also removes the key container), and all
                   files live in $TestDrive (the working directory is pushed there so the
                   script's per-run log files land in it too).

.EXAMPLE
    Invoke-Pester -Path .\Tests\Submit-CertificateRequests.Tests.ps1 -ExcludeTag Lab

.EXAMPLE
    # Full run against a live CA (auto-discovers the CA config from AD Enrollment Services):
    $cfg = New-PesterContainer -Path .\Tests\Submit-CertificateRequests.Tests.ps1 -Data @{ RunLab = $true }
    Invoke-Pester -Container $cfg
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '',
    Justification = 'container parameters are consumed inside Pester Describe/BeforeAll scriptblocks, which the analyzer cannot see through')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingEmptyCatchBlock', '',
    Justification = 'best-effort teardown paths (AfterAll CA-row/key cleanup) deliberately swallow per-item errors')]
param(
    [bool]   $RunLab      = $false,
    [string] $ScriptPath  = (Join-Path (Split-Path $PSScriptRoot -Parent) 'Submit-CertificateRequests.ps1'),
    [string] $CAConfig    = '',            # Lab: '<host>\<CA name>'; auto-discovered from AD Enrollment Services when empty
    [string] $LabTemplate = 'WebServer'    # Lab: an enrollable subject-in-request template published on the CA
)

BeforeDiscovery {
    $script:LabReady = $RunLab
}

Describe 'Submit-CertificateRequests' {

    BeforeAll {
        $script:Submit = $ScriptPath
        $script:Submit | Should -Exist

        # --- AST-extract the pure helpers so the Unit tier exercises the REAL code ------------
        # (The script has a mandatory -CAConfig and pings the CA on load, so it cannot be
        # dot-sourced wholesale; extracting the function bodies runs them with no side effects.)
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:Submit, [ref]$null, [ref]$null)
        foreach ($name in 'Resolve-FullPath', 'Write-BatchLog', 'Get-RequestIdFromOutput', 'Get-DispositionFromOutput',
                          'Get-FriendlyErrorHint', 'Import-TrackingData', 'Export-TrackingData', 'Remove-RspFile') {
            $def = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
            if ($def) { . ([scriptblock]::Create($def[0].Extent.Text)) }
        }
        # Write-BatchLog (called by Export-TrackingData/Remove-RspFile) writes to $script:LogFile
        # unless suppressed - suppress it so Unit tests never touch the filesystem outside TestDrive.
        $script:SuppressLogFile = $true
    }

    Context 'Unit: pure helpers' -Tag 'Unit' {

        It 'Resolve-FullPath anchors a relative path at the current directory' {
            $r = Resolve-FullPath -Path '.\x\y.csv'
            $r | Should -BeExactly ([System.IO.Path]::GetFullPath((Join-Path (Get-Location).ProviderPath '.\x\y.csv')))
        }

        It 'Resolve-FullPath returns a rooted path unchanged (normalized)' {
            Resolve-FullPath -Path 'C:\a\..\b\t.csv' | Should -BeExactly 'C:\b\t.csv'
        }

        It 'Get-RequestIdFromOutput parses plain and quoted RequestId lines' {
            Get-RequestIdFromOutput @('RequestId: 42') | Should -Be 42
            Get-RequestIdFromOutput @('noise', 'RequestId: "17"') | Should -Be 17
        }

        It 'Get-RequestIdFromOutput rejects a malformed RequestId and plain prose' {
            Get-RequestIdFromOutput @('RequestId: r42') | Should -BeNullOrEmpty
            Get-RequestIdFromOutput @('Certificate retrieved(Issued) Issued') | Should -BeNullOrEmpty
        }

        It 'Get-DispositionFromOutput maps certreq wording to statuses' {
            Get-DispositionFromOutput @('Certificate retrieved(Issued) Issued')          | Should -Be 'Issued'
            Get-DispositionFromOutput @('Certificate request is pending: Taken Under Submission') | Should -Be 'Pending'
            Get-DispositionFromOutput @('Certificate request is denied')                 | Should -Be 'Denied'
            Get-DispositionFromOutput @('something else entirely')                       | Should -Be 'Unknown'
        }

        It 'Get-FriendlyErrorHint recognizes each documented failure and suggests a fix' {
            $unsupported = (Get-FriendlyErrorHint -Output @('Error 0x80094800') -CertificateTemplate T -CAConfig C) -join ' '
            $unsupported | Should -Match 'not supported'
            $unsupported | Should -Match 'CATemplates'   # the actionable "verify with" line must survive edits
            (Get-FriendlyErrorHint -Output @('CERTSRV_E_TEMPLATE_DENIED') -CertificateTemplate T -CAConfig C) -join ' ' |
                Should -Match 'Enroll permission'
            (Get-FriendlyErrorHint -Output @('0x80094004') -CertificateTemplate T -CAConfig C) -join ' ' |
                Should -Match 'subject'
            (Get-FriendlyErrorHint -Output @('0x80070005 Access is denied.') -CertificateTemplate T -CAConfig C) -join ' ' |
                Should -Match 'Request Certificates'
        }

        It 'Get-FriendlyErrorHint stays silent on unrecognized output' {
            @(Get-FriendlyErrorHint -Output @('some benign text') -CertificateTemplate T -CAConfig C).Count | Should -Be 0
        }

        It 'Export/Import-TrackingData round-trips records and leaves no temp file' {
            $path = Join-Path $TestDrive 'rt.csv'
            $rows = @(
                [pscustomobject]@{ RequestFile = 'a.req'; RequestID = 5; SubmitTime = 't'; Status = 'Issued'; OutputCertFile = 'a.cer'; LastCheckTime = 't'; ErrorMessage = '' }
                [pscustomobject]@{ RequestFile = 'b.req'; RequestID = 6; SubmitTime = 't'; Status = 'Pending'; OutputCertFile = 'b.cer'; LastCheckTime = 't'; ErrorMessage = 'x, y' }
            )
            Export-TrackingData -Data $rows -Path $path
            Test-Path "$path.tmp" | Should -BeFalse
            $back = Import-TrackingData -Path $path
            $back.Count | Should -Be 2
            $back[1].ErrorMessage | Should -BeExactly 'x, y'
            $back[0].RequestID | Should -Be '5'
        }

        It 'Import-TrackingData returns an empty array for a missing file' {
            @(Import-TrackingData -Path (Join-Path $TestDrive 'nope.csv')).Count | Should -Be 0
        }

        It 'Remove-RspFile deletes the companion .rsp and leaves the .cer alone' {
            $cer = Join-Path $TestDrive 'x.cer'; $rsp = Join-Path $TestDrive 'x.rsp'
            'c' | Out-File $cer; 'r' | Out-File $rsp
            Remove-RspFile -CerPath $cer
            Test-Path $rsp | Should -BeFalse
            Test-Path $cer | Should -BeTrue
        }
    }

    Context 'Static: parse and help' -Tag 'Static' {

        It 'parses without errors' {
            $errs = $null
            $null = [System.Management.Automation.Language.Parser]::ParseFile($script:Submit, [ref]$null, [ref]$errs)
            $errs | Should -BeNullOrEmpty
        }

        It 'comment-based help binds (Synopsis is real, not auto-generated syntax)' {
            $syn = (Get-Help $script:Submit).Synopsis.Trim()
            $syn | Should -Not -BeNullOrEmpty
            # When help fails to bind, Get-Help synthesizes the syntax line as the synopsis -
            # catch that instead of passing vacuously.
            $syn | Should -Not -Match '\[\[-|\[<CommonParameters>\]'
        }

        It 'documents every non-common parameter' {
            $cmd = Get-Command $script:Submit
            $common = [System.Management.Automation.PSCmdlet]::CommonParameters + [System.Management.Automation.PSCmdlet]::OptionalCommonParameters
            $documented = @((Get-Help $script:Submit).parameters.parameter.name)
            foreach ($p in $cmd.Parameters.Keys | Where-Object { $_ -notin $common }) {
                $documented | Should -Contain $p -Because "parameter -$p should have a .PARAMETER help entry"
            }
        }

        It 'provides examples for Submit and Retrieve modes' {
            $ex = (Get-Help $script:Submit -Examples | Out-String)
            $ex | Should -Match '-Mode Submit'
            $ex | Should -Match '-Mode Retrieve'
        }
    }

    Context 'Guard: validation before any CA contact' -Tag 'Guard' {

        It 'requires -InputPath for -Mode Submit' {
            { & $script:Submit -CAConfig 'x\y' -CertificateTemplate T -Mode Submit -WhatIf `
                  -TrackingFile (Join-Path $TestDrive 'g1.csv') -OutputFolder (Join-Path $TestDrive 'g1') } |
                Should -Throw -ExpectedMessage '*-InputPath is required*'
        }

        It 'requires -CertificateTemplate for -Mode Both' {
            { & $script:Submit -CAConfig 'x\y' -InputPath $TestDrive -Mode Both -WhatIf `
                  -TrackingFile (Join-Path $TestDrive 'g2.csv') -OutputFolder (Join-Path $TestDrive 'g2') } |
                Should -Throw -ExpectedMessage '*-CertificateTemplate is required*'
        }

        It 'aborts when the CA cannot be reached (connectivity pre-check)' {
            # localhost runs no CA service, so certutil -ping fails fast with no external traffic.
            { & $script:Submit -CAConfig 'localhost\PESTER-NoSuchCA' -Mode Retrieve -WhatIf `
                  -TrackingFile (Join-Path $TestDrive 'g3.csv') -OutputFolder (Join-Path $TestDrive 'g3') 3>$null 2>$null 6>$null } |
                Should -Throw -ExpectedMessage '*Cannot reach CA*'
        }
    }

    # -------------------------------------------------------------------------------------------
    # Lab tier: live submissions to a real Enterprise CA. Opt-in (-RunLab). Tests are SEQUENTIAL:
    # submit -> resume-skip -> -Force resubmit -> Retrieve -> edge cases, sharing one tracking CSV
    # the way real batch runs do.
    # -------------------------------------------------------------------------------------------
    Context 'Lab: live submission to an Enterprise CA' -Tag 'Lab' -Skip:(-not $script:LabReady) {

        BeforeAll {
            # Prefix and teardown-consumed state FIRST - before anything that can throw.
            # AfterAll runs even when BeforeAll dies, and its cleanup must never see an unset
            # (= unscoped) prefix or dereference a collection that was never created.
            $script:Prefix = "PESTER-$([guid]::NewGuid().ToString('N').Substring(0,8))"
            # Exact CA RequestIDs this run creates - the ONLY rows teardown deletes.
            $script:CaRequestIds = New-Object System.Collections.Generic.List[int]
            # Key containers created by certreq -new (teardown backstop; the request-store removal
            # normally takes the key with it). Container name = CSR subject CN, so this list also
            # drives the exact-CommonName CA backstop query in teardown.
            $script:KeyContainers = New-Object System.Collections.Generic.List[string]

            # Everything this tier writes locally lives in TestDrive - including the script's
            # per-run CertBatch_*.log files, which land in the CURRENT directory.
            $script:PrevDir = (Get-Location).ProviderPath
            Set-Location $TestDrive

            # Resolve the CA: explicit container arg, or auto-discover from AD Enrollment Services.
            $script:Ca = $CAConfig
            if (-not $script:Ca) {
                Import-Module ActiveDirectory -ErrorAction Stop
                $cfgNc = (Get-ADRootDSE).configurationNamingContext
                $es = Get-ADObject -SearchBase "CN=Enrollment Services,CN=Public Key Services,CN=Services,$cfgNc" `
                    -LDAPFilter '(objectClass=pKIEnrollmentService)' -Properties dNSHostName, cn | Select-Object -First 1
                $es | Should -Not -BeNullOrEmpty -Because 'the Lab tier needs an Enterprise CA registered in AD'
                $script:Ca = "$($es.dNSHostName)\$($es.cn)"
            }
            (certutil -config $script:Ca -CATemplates 2>&1 | Out-String) | Should -Match ([regex]::Escape("${LabTemplate}:")) `
                -Because "template '$LabTemplate' must be published on CA '$script:Ca'"

            $script:CsrDir   = Join-Path $TestDrive 'csrs'
            $script:CertDir  = Join-Path $TestDrive 'certs'
            $script:Tracking = Join-Path $TestDrive 'tracking.csv'
            New-Item -ItemType Directory -Force $script:CsrDir | Out-Null

            function script:New-LabCsr {
                param([string]$BaseName, [string]$OutDir = $script:CsrDir)
                $inf = Join-Path $TestDrive "$BaseName.inf"
                @"
[NewRequest]
Subject = "CN=$BaseName"
KeyLength = 2048
Exportable = FALSE
MachineKeySet = FALSE
KeyContainer = "$BaseName"
RequestType = PKCS10
"@ | Out-File $inf -Encoding ascii
                $out = certreq -new -f -q $inf (Join-Path $OutDir "$BaseName.req") 2>&1 | Out-String
                if ($LASTEXITCODE -ne 0) { throw "certreq -new failed for ${BaseName}: $out" }
                $script:KeyContainers.Add($BaseName)
            }

            # Run the script capturing ALL streams as one string, then re-read the tracking CSV.
            function script:Invoke-Submit {
                param([hashtable]$Params)
                $base = @{ CAConfig = $script:Ca; TrackingFile = $script:Tracking; OutputFolder = $script:CertDir; Confirm = $false }
                $out = & $script:Submit @base @Params *>&1 | Out-String -Width 400
                $rows = if (Test-Path $script:Tracking) { @(Import-Csv $script:Tracking) } else { @() }
                foreach ($r in $rows) {
                    if ($r.RequestID -match '^\d+$' -and [int]$r.RequestID -notin $script:CaRequestIds) {
                        $script:CaRequestIds.Add([int]$r.RequestID)
                    }
                }
                [pscustomobject]@{ Output = $out; Rows = $rows }
            }

            script:New-LabCsr -BaseName "$script:Prefix-1"
            script:New-LabCsr -BaseName "$script:Prefix-2"
        }

        AfterAll {
            # STRUCTURAL GUARD: cleanup runs only with a fully-formed run prefix - an unset one
            # would unscope the request-store subject filter. Never widen this.
            if ($script:Prefix -match '^PESTER-[0-9a-f]{8}$') {
                # 0a) Re-harvest RequestIDs from the run's tracking CSV: the script persists each
                #     ID crash-safely after every file, so the CSV is authoritative even when an
                #     invocation died before Invoke-Submit's in-memory harvest ran. The CSV lives
                #     in TestDrive and is this run's alone, so no foreign ID can enter the list.
                if ($script:Tracking -and (Test-Path $script:Tracking)) {
                    foreach ($r in @(Import-Csv $script:Tracking)) {
                        if ($r.RequestID -match '^\d+$' -and [int]$r.RequestID -notin $script:CaRequestIds) {
                            $script:CaRequestIds.Add([int]$r.RequestID)
                        }
                    }
                }
                # 0b) Backstop for rows the CA created without a parsable RequestId reaching the
                #     CSV (e.g. a connection drop mid-submit): query the CA DB by the EXACT
                #     run-unique CommonNames this run requested and harvest their row IDs.
                foreach ($cn in $script:KeyContainers) {
                    try {
                        $view = certutil -config $script:Ca -view -restrict "CommonName=$cn" -out RequestId 2>&1 | Out-String
                        foreach ($m in [regex]::Matches($view, 'Request ID:\s*0x([0-9a-f]+)')) {
                            $rid = [Convert]::ToInt32($m.Groups[1].Value, 16)
                            if ($rid -notin $script:CaRequestIds) { $script:CaRequestIds.Add($rid) }
                        }
                    } catch { }
                }
                # 1) CA database rows: delete ONLY the exact RequestIDs harvested above.
                foreach ($id in $script:CaRequestIds) {
                    try { $null = certutil -config $script:Ca -deleterow $id 2>&1 } catch { }
                }
                # 2) Pending-request store entries from certreq -new, by run-scoped subject
                #    (removing the entry also removes its key container).
                $filter = "CN=$script:Prefix-*"
                foreach ($e in @(Get-ChildItem Cert:\CurrentUser\Request -ErrorAction SilentlyContinue |
                                 Where-Object { $_.Subject -like $filter })) {
                    try { Remove-Item -LiteralPath ('Cert:\CurrentUser\Request\' + $e.Thumbprint) } catch { }
                }
                # 3) Key-container backstop (exact tracked names; usually already gone).
                foreach ($k in $script:KeyContainers) {
                    try { $null = certutil -user -csp 'Microsoft Software Key Storage Provider' -delkey $k 2>&1 } catch { }
                }
            } else {
                Write-Warning "Lab cleanup sweep skipped: run prefix is unset or malformed ('$script:Prefix')."
            }

            # Leave TestDrive LAST, and guarded - a failed restore must never abort the CA
            # cleanup above (which is why it is not first), and TestDrive teardown only needs
            # the CWD moved out by the end of this block.
            try { if ($script:PrevDir) { Set-Location $script:PrevDir } }
            catch { Set-Location "$env:SystemDrive\" }
        }

        It 'submits a batch and the CA issues: rows, files, and certificates all check out' {
            $r = script:Invoke-Submit @{ InputPath = $script:CsrDir; CertificateTemplate = $LabTemplate; Mode = 'Submit' }
            $r.Rows.Count | Should -Be 2
            foreach ($row in $r.Rows) {
                $row.Status | Should -BeExactly 'Issued'
                $row.RequestID | Should -Match '^\d+$'
                Test-Path $row.OutputCertFile | Should -BeTrue
                $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($row.OutputCertFile)
                $cert.Subject | Should -BeExactly ('CN=' + [System.IO.Path]::GetFileNameWithoutExtension($row.RequestFile))
                # default behavior: the .rsp certreq writes next to the .cer is deleted
                Test-Path ([System.IO.Path]::ChangeExtension($row.OutputCertFile, '.rsp')) | Should -BeFalse
            }
            @($r.Rows.RequestID | Sort-Object -Unique).Count | Should -Be 2
        }

        It 'rerunning without -Force skips already-submitted files (non-interactive prompt falls back to skip)' {
            $r = script:Invoke-Submit @{ InputPath = $script:CsrDir; CertificateTemplate = $LabTemplate; Mode = 'Submit' }
            $r.Rows.Count | Should -Be 2
            $r.Output | Should -Match 'cannot prompt|Skipping \(already submitted\)'
        }

        It '-Force resubmits tracked files as new requests with new RequestIDs' {
            $before = @(Import-Csv $script:Tracking | ForEach-Object { [int]$_.RequestID })
            $r = script:Invoke-Submit @{ InputPath = $script:CsrDir; CertificateTemplate = $LabTemplate; Mode = 'Submit'; Force = $true }
            $r.Rows.Count | Should -Be 4
            $newIds = @($r.Rows | ForEach-Object { [int]$_.RequestID } | Where-Object { $_ -notin $before })
            $newIds.Count | Should -Be 2
            $oldMax = ($before | Measure-Object -Maximum).Maximum
            foreach ($id in $newIds) { $id | Should -BeGreaterThan $oldMax }
        }

        It 'Retrieve mode re-resolves a row parked as Unknown (only -CAConfig and the tracking file needed)' {
            $rows = @(Import-Csv $script:Tracking)
            $rows[0].Status = 'Unknown'
            Remove-Item -LiteralPath $rows[0].OutputCertFile -Force
            $rows | Export-Csv $script:Tracking -NoTypeInformation -Encoding utf8

            $r = script:Invoke-Submit @{ Mode = 'Retrieve' }
            $updated = $r.Rows | Where-Object { $_.RequestID -eq $rows[0].RequestID } | Select-Object -First 1
            $updated.Status | Should -BeExactly 'Issued'
            Test-Path $updated.OutputCertFile | Should -BeTrue
        }

        It 'skips an empty request file without creating a row' {
            $emptyDir = Join-Path $TestDrive 'empty-input'
            New-Item -ItemType Directory -Force $emptyDir | Out-Null
            New-Item -ItemType File (Join-Path $emptyDir "$script:Prefix-empty.req") | Out-Null
            $before = @(Import-Csv $script:Tracking).Count
            $r = script:Invoke-Submit @{ InputPath = $emptyDir; CertificateTemplate = $LabTemplate; Mode = 'Submit' }
            $r.Rows.Count | Should -Be $before
            $r.Output | Should -Match 'empty file'
        }

        It 'an unsupported template yields an Error row, the raw code, and the friendly hint' {
            $badDir = Join-Path $TestDrive 'bad-template'
            New-Item -ItemType Directory -Force $badDir | Out-Null
            script:New-LabCsr -BaseName "$script:Prefix-3" -OutDir $badDir
            $r = script:Invoke-Submit @{ InputPath = $badDir; CertificateTemplate = "$script:Prefix-NoSuchTemplate"; Mode = 'Submit' }
            $row = $r.Rows | Where-Object { $_.RequestFile -like "*$script:Prefix-3.req" } | Select-Object -First 1
            $row | Should -Not -BeNullOrEmpty
            $row.Status | Should -BeIn @('Error', 'Denied')
            ($row.ErrorMessage + $r.Output) | Should -Match '0x80094800|UNSUPPORTED_CERT_TYPE|not supported'
            $r.Output | Should -Match 'CATemplates'   # the actionable part of the friendly hint
        }
    }
}
