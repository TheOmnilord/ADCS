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
      -Tag Lab     LIVE submissions to a real Enterprise CA. Skipped unless -RunLab is passed
                   TOGETHER WITH an explicit -CAConfig naming the lab CA: the tier deletes the CA
                   database rows it creates, so it never auto-discovers a CA (the first CA
                   registered in AD could be production). Requires an enrollable template
                   (default WebServer - subject supplied in the request, Enroll for Domain
                   Admins). Surgical by construction: CSR subjects and
                   key containers carry a per-run PESTER-<hex> prefix, every CA RequestID this run
                   creates is tracked and its CA database row deleted in teardown by exact ID
                   (certutil -deleterow), pending-request store entries are removed by the
                   run-scoped subject filter (which also removes the key container), and all
                   files live in $TestDrive (the working directory is pushed there so the
                   script's per-run log files land in it too).

.EXAMPLE
    Invoke-Pester -Path .\Tests\Submit-CertificateRequests.Tests.ps1 -ExcludeTag Lab

.EXAMPLE
    # Full run against a live LAB CA - the CA must be named explicitly (no auto-discovery):
    $cfg = New-PesterContainer -Path .\Tests\Submit-CertificateRequests.Tests.ps1 -Data @{
        RunLab = $true; CAConfig = 'LAB-CA01.lab.example\Lab Issuing CA'
    }
    Invoke-Pester -Container $cfg
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '',
    Justification = 'container parameters are consumed inside Pester Describe/BeforeAll scriptblocks, which the analyzer cannot see through')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingEmptyCatchBlock', '',
    Justification = 'best-effort teardown paths (AfterAll CA-row/key cleanup) deliberately swallow per-item errors')]
param(
    [bool]   $RunLab      = $false,
    [string] $ScriptPath  = (Join-Path (Split-Path $PSScriptRoot -Parent) 'Submit-CertificateRequests.ps1'),
    [string] $CAConfig    = '',            # Lab: '<host>\<CA name>' of the LAB CA. REQUIRED for -RunLab (never auto-discovered).
    [string] $LabTemplate = 'WebServer'    # Lab: an enrollable subject-in-request template published on the CA
)

BeforeDiscovery {
    # The Lab tier mutates a CA database (submits, then deletes its own rows), so it runs only
    # against a CA the operator NAMED - never one picked up from AD, which could be production.
    $script:LabReady = $RunLab -and -not [string]::IsNullOrWhiteSpace($CAConfig)
    if ($RunLab -and -not $script:LabReady) {
        Write-Warning 'Submit-CertificateRequests Lab tier skipped: -RunLab needs an explicit -CAConfig naming the lab CA (auto-discovery is deliberately not offered for a tier that deletes CA database rows).'
    }
}

Describe 'Submit-CertificateRequests' {

    BeforeAll {
        $script:Submit = $ScriptPath
        $script:Submit | Should -Exist

        # --- AST-extract the pure helpers so the Unit tier exercises the REAL code ------------
        # (The script has a mandatory -CAConfig and pings the CA on load, so it cannot be
        # dot-sourced wholesale; extracting the function bodies runs them with no side effects.)
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:Submit, [ref]$null, [ref]$null)
        foreach ($name in 'Resolve-FullPath', 'Resolve-TrackingFilePath', 'Assert-SafeNativeArgument', 'Assert-CertificateOutputPath', 'Assert-ProtectedDirectoryChain', 'Move-StaleCertificateAside', 'Remove-AsideIfIdentical', 'Move-RetrievedCertificate', 'New-TempCertificatePath',
                          'Get-TrustedPrincipalSet', 'ConvertTo-PrincipalLabel', 'Get-UntrustedGrant', 'Get-UntrustedOwner',
                          'Write-BatchLog', 'Get-RequestIdFromOutput', 'Get-DispositionFromOutput',
                          'Get-FriendlyErrorHint', 'Import-TrackingData', 'Export-TrackingData', 'Remove-RspFile', 'Resolve-CertificateOutputNames', 'Get-DestinationOwnerConflict', 'Get-RequestFiles') {
            $def = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
            if ($def) { . ([scriptblock]::Create($def[0].Extent.Text)) }
        }
        # Write-BatchLog (called by Export-TrackingData/Remove-RspFile) writes to $script:LogFile
        # unless suppressed - suppress it so Unit tests never touch the filesystem outside TestDrive.
        $script:SuppressLogFile = $true
        # Assert-CertificateOutputPath reads the -AllowUnprotectedOutputFolder decision and the
        # trusted-principal set from here (as the script's main block sets them).
        $script:AllowUnprotectedOutput = $false
        $script:TrustedSids = Get-TrustedPrincipalSet
        # The machine's TEMP chain (where TestDrive lives) may legitimately grant delete/rename
        # rights to principals the script does not know (e.g. a sandbox account on a dev box). The
        # tests exercise the trust LOGIC, not this machine's ACLs, so whatever the pre-existing
        # environment grants on TestDrive's ancestors are treated as trusted for the run - exactly
        # what an operator would do with -TrustedOutputPrincipal. Folders the tests create carry
        # only the ACEs the tests put there (inheritance is cut), so the assertions stay exact.
        $script:EnvTrusted = @()
        $probe = $TestDrive
        while ($probe) {
            try {
                foreach ($label in @(Get-UntrustedGrant -Path $probe -Kind Swap)) { if ($label -match 'S-1-[\d-]+') { $script:EnvTrusted += $Matches[0] } }
                $owner = Get-UntrustedOwner -Path $probe
                if ($owner -and $owner -match 'S-1-[\d-]+') { $script:EnvTrusted += $Matches[0] }
            } catch { }
            $probe = [System.IO.Path]::GetDirectoryName($probe)
        }
        $script:EnvTrusted = @($script:EnvTrusted | Sort-Object -Unique)
        $script:TrustedSids = Get-TrustedPrincipalSet -Extra $script:EnvTrusted
    }

    Context 'Unit: pure helpers' -Tag 'Unit' {

        It 'Resolve-FullPath anchors a relative path at the current directory' {
            $r = Resolve-FullPath -Path '.\x\y.csv'
            $r | Should -BeExactly ([System.IO.Path]::GetFullPath((Join-Path (Get-Location).ProviderPath '.\x\y.csv')))
        }

        It 'Resolve-FullPath returns a rooted path unchanged (normalized)' {
            Resolve-FullPath -Path 'C:\a\..\b\t.csv' | Should -BeExactly 'C:\b\t.csv'
        }

        It 'Resolve-TrackingFilePath canonicalizes an 8.3 alias to the long name (same lock for both spellings)' {
            $long = Join-Path $TestDrive 'CertificateTracking.csv'; 'x' | Out-File $long
            $short = (New-Object -ComObject Scripting.FileSystemObject).GetFile($long).ShortPath
            if ($short -eq $long) { Set-ItResult -Skipped -Because '8.3 names are disabled on this volume'; return }
            Resolve-TrackingFilePath -Path $short | Should -BeExactly (Resolve-TrackingFilePath -Path $long)
            (Resolve-TrackingFilePath -Path $short) | Should -Not -Match '~'
            # a not-yet-existing file simply resolves to its absolute path
            Resolve-TrackingFilePath -Path (Join-Path $TestDrive 'new.csv') | Should -BeExactly (Join-Path $TestDrive 'new.csv')
        }

        It 'Resolve-TrackingFilePath refuses a hard-linked or symlinked tracking file' {
            $orig = Join-Path $TestDrive 'linked.csv'; 'x' | Out-File $orig
            $link = Join-Path $TestDrive 'alias.csv'
            try { New-Item -ItemType HardLink -Path $link -Target $orig -ErrorAction Stop | Out-Null }
            catch { Set-ItResult -Skipped -Because "cannot create a hard link here: $_"; return }
            { Resolve-TrackingFilePath -Path $link } | Should -Throw -ExpectedMessage '*HardLink*'
            { Resolve-TrackingFilePath -Path $orig } | Should -Throw -ExpectedMessage '*HardLink*'   # both names are the same file
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

        It 'Export-TrackingData and Write-BatchLog suppress nested confirmation (the checkpoint is never a separate prompt)' {
            # Under the script's -Confirm, $ConfirmPreference is Low and Export-Csv / Out-File would
            # prompt on their own; declining THAT after certreq has submitted would lose the
            # RequestID. Both must pass -Confirm:$false explicitly.
            Mock Export-Csv { }
            Mock Out-File { }
            $rows = @([pscustomobject]@{ RequestFile = 'a.req'; RequestID = 1; SubmitTime = 't'; Status = 'Issued'; OutputCertFile = 'a.cer'; LastCheckTime = 't'; ErrorMessage = '' })
            try { Export-TrackingData -Data $rows -Path (Join-Path $TestDrive 'nc.csv') } catch { }   # the (mocked-away) temp file makes the rename fail; irrelevant here
            Should -Invoke Export-Csv -Times 1 -ParameterFilter { $null -ne $Confirm -and -not [bool]$Confirm }
            $script:SuppressLogFile = $false
            $script:LogFile = Join-Path $TestDrive 'nc.log'
            try { Write-BatchLog 'checkpoint' } finally { $script:SuppressLogFile = $true }
            Should -Invoke Out-File -Times 1 -ParameterFilter { $null -ne $Confirm -and -not [bool]$Confirm }
        }

        It 'Import-TrackingData returns an empty array for a missing file' {
            @(Import-TrackingData -Path (Join-Path $TestDrive 'nope.csv')).Count | Should -Be 0
        }

        It 'Assert-SafeNativeArgument rejects quotes and control characters, accepts ordinary CA/template values' {
            { Assert-SafeNativeArgument -Name T -Value 'CA01.domain.com\Contoso Issuing CA 1' } | Should -Not -Throw
            { Assert-SafeNativeArgument -Name T -Value 'x" -foo "y' } | Should -Throw -ExpectedMessage '*double quotes*'
            { Assert-SafeNativeArgument -Name T -Value "a`tb" }      | Should -Throw -ExpectedMessage '*control characters*'
            { Assert-SafeNativeArgument -Name T -Value '' }          | Should -Not -Throw
        }

        It 'Import-TrackingData reads a tracking file whose name carries wildcard characters literally' {
            $p = Join-Path $TestDrive 'tracking[1].csv'
            [pscustomobject]@{ RequestFile = 'C:\r\a.req'; RequestID = '7'; SubmitTime = ''; Status = 'Pending'; OutputCertFile = 'C:\r\a.cer'; LastCheckTime = ''; ErrorMessage = '' } |
                Export-Csv -LiteralPath $p -NoTypeInformation -Encoding utf8
            $rows = @(Import-TrackingData -Path $p)
            $rows.Count | Should -Be 1 -Because "a wildcard-aware Test-Path looks for 'tracking1.csv', reports the file absent and returns an empty history - every request would be resubmitted"
            $rows[0].RequestID | Should -Be '7'
        }

        It 'Export-TrackingData writes a tracking path containing wildcard characters literally (the checkpoint is not lost)' {
            $p = Join-Path $TestDrive 'Cert[1]Tracking.csv'
            $rows = @([pscustomobject]@{ RequestFile = 'C:\r\a.req'; RequestID = '9'; SubmitTime = ''; Status = 'Issued'; OutputCertFile = 'C:\r\a.cer'; LastCheckTime = ''; ErrorMessage = ''; CAConfig = 'ca\CA' })
            $script:SuppressLogFile = $true
            try { Export-TrackingData -Data $rows -Path $p 3>$null } finally { $script:SuppressLogFile = $false }
            Test-Path -LiteralPath $p | Should -BeTrue -Because 'Export-Csv -Path would glob the [1] and silently fail to write the checkpoint, losing the RequestID of an already-submitted request'
            (Import-Csv -LiteralPath $p).RequestID | Should -Be '9'
        }

        It 'Get-RequestFiles reads a bracketed drop folder literally and matches only EXACT .req/.csr/.txt extensions' {
            $dir = Join-Path $TestDrive 'CSR[prod]'   # a real folder whose name contains a wildcard metacharacter
            [void][System.IO.Directory]::CreateDirectory($dir)   # New-Item has no -LiteralPath; -Path would glob the brackets
            foreach ($n in 'a.req', 'b.csr', 'note.txt', 'a.reqbak', 'c.request', 'readme.md') { Set-Content -LiteralPath (Join-Path $dir $n) -Value 'x' }
            $script:SuppressLogFile = $true
            try { $files = @(Get-RequestFiles -Path $dir 3>$null) } finally { $script:SuppressLogFile = $false }
            @($files | ForEach-Object { $_.Name } | Sort-Object) | Should -Be (@('a.req', 'b.csr', 'note.txt') | Sort-Object) -Because 'the folder is read with -LiteralPath (not globbed as a character class) and only exact extensions match'
            $files.Name | Should -Not -Contain 'a.reqbak' -Because "-Filter '*.req' would match this backup, submitting it to the CA"
            $files.Name | Should -Not -Contain 'c.request'
        }

        It 'a row from an older tracking file (no CAConfig property) is readable under strict mode' {
            Set-StrictMode -Version Latest
            $row = [pscustomobject]@{ RequestFile = 'C:\r\a.req'; RequestID = '7'; Status = 'Pending'; OutputCertFile = 'C:\r\a.cer' }
            { if ($row.PSObject.Properties['CAConfig']) { "$($row.CAConfig)" } else { '' } } | Should -Not -Throw
            { "$($row.CAConfig)" } | Should -Throw -Because 'this is what the Retrieve loop must NOT do on legacy rows'
        }

        It 'Export-TrackingData writes the CAConfig column even when the first row comes from an older file without it, and keeps extra columns' {
            $p = Join-Path $TestDrive 'mixed.csv'
            $old = [pscustomobject]@{ RequestFile = 'C:\r\old.req'; RequestID = '1'; SubmitTime = ''; Status = 'Issued'; OutputCertFile = 'C:\r\old.cer'; LastCheckTime = ''; ErrorMessage = ''; Note = 'kept' }
            $new = [pscustomobject]@{ RequestFile = 'C:\r\new.req'; RequestID = '2'; SubmitTime = ''; Status = 'Pending'; OutputCertFile = 'C:\r\new.cer'; LastCheckTime = ''; ErrorMessage = ''; CAConfig = 'ca1\CA' }
            $script:SuppressLogFile = $true
            try { Export-TrackingData -Data @($old, $new) -Path $p 3>$null } finally { $script:SuppressLogFile = $false }
            $rows = @(Import-Csv -LiteralPath $p)
            ($rows | Where-Object RequestID -eq '2').CAConfig | Should -Be 'ca1\CA' -Because 'Export-Csv takes its columns from the FIRST object; an older first row would otherwise drop the CA binding of every newer row'
            ($rows | Where-Object RequestID -eq '1').Note | Should -Be 'kept'
        }

        It 'Write-BatchLog folds CR/LF and control characters so a tracking field cannot forge extra log lines' {
            $log = Join-Path $TestDrive 'forge.log'
            $saved = $script:LogFile; $script:LogFile = $log; $script:SuppressLogFile = $false
            try { Write-BatchLog ("row for 'x.req'`r`n[2026-01-01 00:00:00] [Info] forged line" + [char]27 + "[0m") 6>$null }
            finally { $script:LogFile = $saved }
            $lines = @(Get-Content -LiteralPath $log)
            $lines.Count | Should -Be 1 -Because 'one message must always be exactly one log record'
            $lines[0] | Should -Match 'forged line'
            $lines[0].Contains([string][char]27) | Should -BeFalse -Because 'control characters are replaced, not passed through to the log'
            ($lines[0].Contains("`r") -or $lines[0].Contains("`n")) | Should -BeFalse
        }

        It 'Resolve-CertificateOutputNames gives every request a unique .cer name, within the batch and against other files'' recorded destinations' {
            # prod.req + prod.csr share the base name -> full names; prod.req.txt has the BASE name
            # 'prod.req', exactly what prod.req falls back to - the old rule mapped both to prod.req.cer
            $names = Resolve-CertificateOutputNames -RequestPaths 'C:\in\prod.req', 'C:\in\prod.csr', 'C:\in\prod.req.txt' -ExistingRows @()
            $names['C:\in\prod.req']     | Should -Be 'prod.req.cer'
            $names['C:\in\prod.csr']     | Should -Be 'prod.csr.cer'
            $names['C:\in\prod.req.txt'] | Should -Be 'prod.req.txt.cer'
            @($names.Values | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique).Count | Should -Be 3
            # a plain file keeps the short name
            (Resolve-CertificateOutputNames -RequestPaths 'C:\in\web01.req' -ExistingRows @())['C:\in\web01.req'] | Should -Be 'web01.cer'
            # a base name already recorded for a DIFFERENT request file (an earlier run's web01.csr -> web01.cer) is not reused
            $rows = @([pscustomobject]@{ RequestFile = 'C:\in\web01.csr'; OutputCertFile = 'C:\out\web01.cer' })
            (Resolve-CertificateOutputNames -RequestPaths 'C:\in\web01.req' -ExistingRows $rows)['C:\in\web01.req'] | Should -Be 'web01.req.cer'
            # ...while the same file's own row does not block it (a -Force resubmit)
            $own = @([pscustomobject]@{ RequestFile = 'C:\in\web01.req'; OutputCertFile = 'C:\out\web01.cer' })
            (Resolve-CertificateOutputNames -RequestPaths 'C:\in\web01.req' -ExistingRows $own)['C:\in\web01.req'] | Should -Be 'web01.cer'
            # a name that an OLDER run recorded for TWO other files (the pre-1.0.4 collision) blocks
            # both of them, not just the first one listed
            $twoOwners = @([pscustomobject]@{ RequestFile = 'C:\in\prod.req'; OutputCertFile = 'C:\out\prod.req.cer' },
                           [pscustomobject]@{ RequestFile = 'C:\in\prod.req.txt'; OutputCertFile = 'C:\out\prod.req.cer' })
            # prod.req + prod.csr collide on prod.cer, so prod.req must fall back to prod.req.cer - which
            # prod.req.txt's row still points at. No safe name exists: the batch is refused, not guessed.
            { Resolve-CertificateOutputNames -RequestPaths 'C:\in\prod.req', 'C:\in\prod.csr' -ExistingRows $twoOwners } |
                Should -Throw -ExpectedMessage '*already records for*prod.req.txt*'
            # alone (no prod.csr in the batch) prod.req keeps its natural prod.cer, which nobody recorded
            (Resolve-CertificateOutputNames -RequestPaths 'C:\in\prod.req' -ExistingRows $twoOwners)['C:\in\prod.req'] | Should -Be 'prod.cer'
            # an unresolvable clash aborts before anything is submitted
            $clash = @([pscustomobject]@{ RequestFile = 'C:\in\other.csr'; OutputCertFile = 'C:\out\web01.req.cer' })
            { Resolve-CertificateOutputNames -RequestPaths 'C:\in\web01.req', 'C:\in\web01.csr' -ExistingRows $clash } | Should -Throw -ExpectedMessage '*never share a destination*'
            # a tracking row whose recorded destination carries a Win32-invalid path char (| < >) must
            # NOT throw (GetFileName throws on 5.1): the row is simply not registered as taken.
            $badRow = @([pscustomobject]@{ RequestFile = 'C:\in\other.req'; OutputCertFile = 'C:\out\a|b.cer' })
            { Resolve-CertificateOutputNames -RequestPaths 'C:\in\web01.req' -ExistingRows $badRow } | Should -Not -Throw
            (Resolve-CertificateOutputNames -RequestPaths 'C:\in\web01.req' -ExistingRows $badRow)['C:\in\web01.req'] | Should -Be 'web01.cer'
        }

        It 'Get-DestinationOwnerConflict finds another request (different RequestID, Issued/Undelivered) already owning the destination' {
            $rows = @(
                [pscustomobject]@{ RequestFile = 'C:\in\srv.req'; RequestID = '41'; Status = 'Pending'; OutputCertFile = 'C:\out\srv.cer'; SubmitTime = '2026-09-01T10:00:00' }
                [pscustomobject]@{ RequestFile = 'C:\in\srv.req'; RequestID = '42'; Status = 'Issued';  OutputCertFile = 'C:\out\SRV.cer'; SubmitTime = '2026-09-02T10:00:00' }
                [pscustomobject]@{ RequestFile = 'C:\in\web.req'; RequestID = '43'; Status = 'Issued';  OutputCertFile = 'C:\out\web.cer'; SubmitTime = '2026-09-02T11:00:00' }
            )
            (Get-DestinationOwnerConflict -Record $rows[0] -Tracking $rows).RequestID | Should -Be '42' -Because 'the newer request already delivered srv.cer (paths compare case-insensitively); retrieving 41 into it must be refused'
            Get-DestinationOwnerConflict -Record $rows[2] -Tracking $rows | Should -BeNullOrEmpty
            # a row never conflicts with itself, and a Pending/Error row owns nothing yet
            Get-DestinationOwnerConflict -Record $rows[1] -Tracking $rows | Should -BeNullOrEmpty
            $undelivered = [pscustomobject]@{ RequestFile = 'C:\in\web.req'; RequestID = '44'; Status = 'Undelivered'; OutputCertFile = 'C:\out\web.cer'; SubmitTime = '' }
            (Get-DestinationOwnerConflict -Record $rows[2] -Tracking ($rows + $undelivered)).RequestID | Should -Be '44' -Because 'an Undelivered row has a certificate staged for that destination'
        }

        It 'Get-DestinationOwnerConflict is DIRECTIONAL: an older delivered request never blocks a newer one, and it fails closed on unparseable timestamps' {
            # The -Force renewal case: 41 Issued (older) at srv.cer, 42 (newer) resubmitted to the
            # SAME destination. The newer request must be deliverable; only a strictly newer row owns.
            $renew = @(
                [pscustomobject]@{ RequestFile = 'C:\in\srv.req'; RequestID = '41'; Status = 'Issued';  OutputCertFile = 'C:\out\srv.cer'; SubmitTime = '2026-09-01T10:00:00' }
                [pscustomobject]@{ RequestFile = 'C:\in\srv.req'; RequestID = '42'; Status = 'Pending'; OutputCertFile = 'C:\out\srv.cer'; SubmitTime = '2026-09-02T10:00:00' }
            )
            Get-DestinationOwnerConflict -Record $renew[1] -Tracking $renew | Should -BeNullOrEmpty -Because 'the newer request 42 may deliver; the older Issued 41 does NOT own the destination'
            # ...but once 42 has delivered (Issued), retrieving the OLDER still-pending... (make 41 the one retrieved)
            $renew[1].Status = 'Issued'
            (Get-DestinationOwnerConflict -Record $renew[0] -Tracking $renew).RequestID | Should -Be '42' -Because 'retrieving the OLDER 41 over the newer, already delivered 42 is refused'
            # -Destination override checks the effective (redirected) path without mutating the row
            Get-DestinationOwnerConflict -Record $renew[0] -Tracking $renew -Destination 'D:\Redirect\elsewhere.cer' |
                Should -BeNullOrEmpty -Because 'redirected to a different physical path, there is no shared destination'
            # unparseable/missing SubmitTime -> fail closed (refuse rather than guess the order)
            $noTime = @(
                [pscustomobject]@{ RequestFile = 'C:\in\x.req'; RequestID = '50'; Status = 'Issued';  OutputCertFile = 'C:\out\x.cer'; SubmitTime = '' }
                [pscustomobject]@{ RequestFile = 'C:\in\x.req'; RequestID = '51'; Status = 'Pending'; OutputCertFile = 'C:\out\x.cer'; SubmitTime = 'not-a-date' }
            )
            (Get-DestinationOwnerConflict -Record $noTime[1] -Tracking $noTime).RequestID | Should -Be '50' -Because 'when the order cannot be established, the conflict is refused (fail closed)'

            # CROSS-CA: RequestIDs are per CA, so two DIFFERENT requests can share the number. CA-A/41
            # (older, Pending) and CA-B/41 (newer, Issued) share a destination; retrieving CA-A/41 must
            # see CA-B/41 as the owner (they are NOT the same request despite the same ID).
            $crossCa = @(
                [pscustomobject]@{ RequestFile = 'C:\in\srv.req'; RequestID = '41'; CAConfig = 'CA-A\Issuing'; Status = 'Pending'; OutputCertFile = 'C:\out\srv.cer'; SubmitTime = '2026-09-01T10:00:00' }
                [pscustomobject]@{ RequestFile = 'C:\in\srv.req'; RequestID = '41'; CAConfig = 'CA-B\Issuing'; Status = 'Issued';  OutputCertFile = 'C:\out\srv.cer'; SubmitTime = '2026-09-02T10:00:00' }
            )
            (Get-DestinationOwnerConflict -Record $crossCa[0] -Tracking $crossCa).CAConfig | Should -Be 'CA-B\Issuing' -Because 'a different CA''s request with the same number is a different request and still owns the shared destination'
            # a row still does not conflict with ITSELF (reference identity, not RequestID equality)
            Get-DestinationOwnerConflict -Record $crossCa[1] -Tracking @($crossCa[1]) | Should -BeNullOrEmpty

            # DST: instants must be compared, not local wall-clock. On the Oslo fall-back night an
            # older submission with the summer offset is an EARLIER instant than a later one with the
            # winter offset; a [datetime] comparison reverses them. The older must not own the dest.
            $dst = @(
                [pscustomobject]@{ RequestFile = 'C:\in\d.req'; RequestID = '60'; Status = 'Issued'; OutputCertFile = 'C:\out\d.cer'; SubmitTime = '2026-10-25T02:45:00.0000000+02:00' }  # 00:45 UTC (older instant)
                [pscustomobject]@{ RequestFile = 'C:\in\d.req'; RequestID = '61'; Status = 'Issued'; OutputCertFile = 'C:\out\d.cer'; SubmitTime = '2026-10-25T02:15:00.0000000+01:00' }  # 01:15 UTC (newer instant)
            )
            Get-DestinationOwnerConflict -Record $dst[1] -Tracking $dst | Should -BeNullOrEmpty -Because 'row 61 is the later INSTANT (01:15 UTC) despite the earlier wall-clock; the older 60 must not own the destination'
            (Get-DestinationOwnerConflict -Record $dst[0] -Tracking $dst).RequestID | Should -Be '61' -Because 'retrieving the older 60 over the newer 61 is refused (instants compared, not local wall-clock)'
        }

        It 'Assert-CertificateOutputPath confines a tracking-row path to a rooted .cer beneath an allowed root' {
            $root = Join-Path $TestDrive 'batch'; $sub = Join-Path $root 'Certificates'
            New-Item -ItemType Directory -Force $sub | Out-Null
            { Assert-CertificateOutputPath -Name T -Path (Join-Path $sub 'srv1.cer') -AllowedRoots @($root) } | Should -Not -Throw
            { Assert-CertificateOutputPath -Name T -Path (Join-Path $root 'srv1.cer') -AllowedRoots @($root) } | Should -Not -Throw   # directly in the root
            # a script/config/other type can never be the delivery or move-aside target
            { Assert-CertificateOutputPath -Name T -Path (Join-Path $sub 'startup.ps1') -AllowedRoots @($root) } | Should -Throw -ExpectedMessage '*must be a .cer file*'
            # OUTSIDE the boundary - a sibling folder, and a '..' escape that textually starts inside
            $outside = Join-Path $TestDrive 'elsewhere'; New-Item -ItemType Directory -Force $outside | Out-Null
            { Assert-CertificateOutputPath -Name T -Path (Join-Path $outside 'x.cer') -AllowedRoots @($root) }        | Should -Throw -ExpectedMessage '*outside the allowed output locations*'
            { Assert-CertificateOutputPath -Name T -Path (Join-Path $root '..\elsewhere\x.cer') -AllowedRoots @($root) } | Should -Throw -ExpectedMessage '*outside the allowed output locations*'
            # a root that merely PREFIXES the name is not containment (C:\batch2 vs C:\batch)
            $root2 = "${root}2"; New-Item -ItemType Directory -Force $root2 | Out-Null
            { Assert-CertificateOutputPath -Name T -Path (Join-Path $root2 'x.cer') -AllowedRoots @($root) } | Should -Throw -ExpectedMessage '*outside the allowed output locations*'
            # any of several roots may contain it
            { Assert-CertificateOutputPath -Name T -Path (Join-Path $outside 'x.cer') -AllowedRoots @($root, $outside) } | Should -Not -Throw
            # relative paths are not accepted from the CSV (rows are stored absolute)
            { Assert-CertificateOutputPath -Name T -Path '.\x.cer' -AllowedRoots @($root) } | Should -Throw -ExpectedMessage '*rooted*'
            # the folder must already exist
            { Assert-CertificateOutputPath -Name T -Path (Join-Path $sub 'nope\x.cer') -AllowedRoots @($root) } | Should -Throw -ExpectedMessage '*does not exist*'
            # and the command-line rule still applies
            { Assert-CertificateOutputPath -Name T -Path (Join-Path $sub 'a"b.cer') -AllowedRoots @($root) } | Should -Throw -ExpectedMessage '*double quotes*'
        }

        It 'Assert-CertificateOutputPath rejects a destination that is itself a folder or a link named *.cer' {
            $root = Join-Path $TestDrive 'droot'; New-Item -ItemType Directory -Force $root | Out-Null
            $dirNamedCer = Join-Path $root 'trap.cer'; New-Item -ItemType Directory -Force $dirNamedCer | Out-Null
            { Assert-CertificateOutputPath -Name T -Path $dirNamedCer -AllowedRoots @($root) } | Should -Throw -ExpectedMessage '*existing folder*'
            $away = Join-Path $TestDrive 'daway'; New-Item -ItemType Directory -Force $away | Out-Null
            $junctionCer = Join-Path $root 'jtrap.cer'
            try { New-Item -ItemType Junction -Path $junctionCer -Target $away -ErrorAction Stop | Out-Null }
            catch { Set-ItResult -Skipped -Because "cannot create a junction here: $_"; return }
            # a junction is a folder AND a reparse point - either rule must stop it
            { Assert-CertificateOutputPath -Name T -Path $junctionCer -AllowedRoots @($root) } | Should -Throw
            # RACE: a folder/junction planted under the destination name AFTER validation (simulated by
            # skipping the check) must receive NO file - the no-overwrite rename fails instead.
            $tmp = Join-Path $TestDrive 'certreq-d.cer'; 'cert' | Out-File $tmp -Encoding ascii
            { Move-RetrievedCertificate -TempCer $tmp -Destination $dirNamedCer } | Should -Throw
            @(Get-ChildItem $dirNamedCer -Force).Count | Should -Be 0 -Because 'nothing may be moved INTO a folder planted under the destination name'
            Test-Path $tmp | Should -BeTrue -Because 'the undelivered certificate stays in its staging file for recovery'
            { Move-RetrievedCertificate -TempCer $tmp -Destination $junctionCer } | Should -Throw
            @(Get-ChildItem $away -Force).Count | Should -Be 0 -Because 'a raced junction receives no file'
            # a plain file planted under the name is not overwritten either
            $planted = Join-Path $root 'planted.cer'; 'theirs' | Out-File $planted -Encoding ascii
            $tmp2 = Join-Path $TestDrive 'certreq-d2.cer'; 'cert' | Out-File $tmp2 -Encoding ascii
            Move-RetrievedCertificate -TempCer $tmp2 -Destination $planted 6>$null
            (Get-Content $planted -Raw).Trim() | Should -BeExactly 'cert' -Because 'a pre-existing plain file is moved aside (kept), then replaced'
        }

        It 'Assert-CertificateOutputPath rejects a junction below the root (a redirecting folder is not containment)' {
            $root = Join-Path $TestDrive 'jroot'; $away = Join-Path $TestDrive 'jaway'
            New-Item -ItemType Directory -Force $root, $away | Out-Null
            $junction = Join-Path $root 'Certificates'
            try { New-Item -ItemType Junction -Path $junction -Target $away -ErrorAction Stop | Out-Null }
            catch { Set-ItResult -Skipped -Because "cannot create a junction here: $_"; return }
            { Assert-CertificateOutputPath -Name T -Path (Join-Path $junction 'x.cer') -AllowedRoots @($root) } | Should -Throw -ExpectedMessage '*reparse point*'
        }

        It 'Move-RetrievedCertificate re-validates the destination right before delivery and leaves the temp file when it fails' {
            $root = Join-Path $TestDrive 'rv-root'; $outside = Join-Path $TestDrive 'rv-outside'
            New-Item -ItemType Directory -Force $root, $outside | Out-Null
            $tmp = Join-Path $TestDrive 'certreq-rv.cer'; 'cert' | Out-File $tmp -Encoding ascii
            { Move-RetrievedCertificate -TempCer $tmp -Destination (Join-Path $outside 'x.cer') -AllowedRoots @($root) } |
                Should -Throw -ExpectedMessage '*outside the allowed output locations*'
            Test-Path $tmp | Should -BeTrue -Because 'an undelivered certificate must survive for recovery'
            Test-Path (Join-Path $outside 'x.cer') | Should -BeFalse
            Move-RetrievedCertificate -TempCer $tmp -Destination (Join-Path $root 'x.cer') -AllowedRoots @($root)
            Test-Path (Join-Path $root 'x.cer') | Should -BeTrue
        }

        It 'Get-UntrustedGrant separates swap-capable (delete/rename/write-data) from create-subfolder-only rights and ignores trusted principals' {
            $tight = Join-Path $TestDrive 'acl-tight'; $create = Join-Path $TestDrive 'acl-create'; $swap = Join-Path $TestDrive 'acl-swap'
            New-Item -ItemType Directory -Force $tight, $create, $swap | Out-Null
            $me = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
            foreach ($d in $tight, $create, $swap) {
                $acl = Get-Acl -LiteralPath $d
                $acl.SetAccessRuleProtection($true, $false)                       # drop inherited ACEs
                foreach ($r in @($acl.Access)) { [void]$acl.RemoveAccessRule($r) }
                $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($me, 'FullControl', 'ContainerInherit, ObjectInherit', 'None', 'Allow')))
                # SYSTEM and Administrators with Full Control are TRUSTED - never reported
                $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule((New-Object System.Security.Principal.SecurityIdentifier('S-1-5-18')), 'FullControl', 'ContainerInherit, ObjectInherit', 'None', 'Allow')))
                $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule((New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')), 'FullControl', 'ContainerInherit, ObjectInherit', 'None', 'Allow')))
                Set-Acl -LiteralPath $d -AclObject $acl
            }
            @(Get-UntrustedGrant -Path $tight -Kind Swap).Count   | Should -Be 0
            @(Get-UntrustedGrant -Path $tight -Kind Create).Count | Should -Be 0
            # the C:\ default: Authenticated Users may create folders - create-only, NOT swap-capable
            $acl = Get-Acl -LiteralPath $create
            $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
                (New-Object System.Security.Principal.SecurityIdentifier('S-1-5-11')), 'CreateDirectories', 'ContainerInherit', 'None', 'Allow')))
            Set-Acl -LiteralPath $create -AclObject $acl
            @(Get-UntrustedGrant -Path $create -Kind Create) -join ' ' | Should -Match 'Authenticated Users'
            @(Get-UntrustedGrant -Path $create -Kind Swap).Count | Should -Be 0
            # ...but "create FILES" (FILE_WRITE_DATA) or write-attributes IS swap-capable: a directory
            # handle with either right can set a reparse point, turning an empty folder into a
            # junction in place with nothing deleted or renamed
            foreach ($right in 'CreateFiles', 'WriteAttributes') {
                $wd = Join-Path $TestDrive "acl-$right"; New-Item -ItemType Directory -Force $wd | Out-Null
                $acl = Get-Acl -LiteralPath $wd
                $acl.SetAccessRuleProtection($true, $false); foreach ($r in @($acl.Access)) { [void]$acl.RemoveAccessRule($r) }
                $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($me, 'FullControl', 'ContainerInherit, ObjectInherit', 'None', 'Allow')))
                $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
                    (New-Object System.Security.Principal.SecurityIdentifier('S-1-5-11')), $right, 'None', 'None', 'Allow')))
                Set-Acl -LiteralPath $wd -AclObject $acl
                @(Get-UntrustedGrant -Path $wd -Kind Swap) -join ' ' | Should -Match 'Authenticated Users' -Because "$right lets an untrusted user set a reparse point on the folder"
            }
            # Modify (includes Delete) for a NON-builtin, unresolvable principal (a "Domain Users"
            # of some domain) IS swap-capable and untrusted - a fixed list of broad groups would miss it
            $acl = Get-Acl -LiteralPath $swap
            $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
                (New-Object System.Security.Principal.SecurityIdentifier('S-1-5-21-111-222-333-513')), 'Modify', 'ContainerInherit, ObjectInherit', 'None', 'Allow')))
            Set-Acl -LiteralPath $swap -AclObject $acl
            @(Get-UntrustedGrant -Path $swap -Kind Swap) -join ' ' | Should -Match 'S-1-5-21-111-222-333-513'
            # ...unless that principal is named as trusted
            $saved = $script:TrustedSids
            try {
                $script:TrustedSids = Get-TrustedPrincipalSet -Extra 'S-1-5-21-111-222-333-513'
                @(Get-UntrustedGrant -Path $swap -Kind Swap).Count | Should -Be 0
            } finally { $script:TrustedSids = $saved }
            # read-only for a broad group is neither
            $acl = Get-Acl -LiteralPath $tight
            $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
                (New-Object System.Security.Principal.SecurityIdentifier('S-1-1-0')), 'ReadAndExecute', 'ContainerInherit, ObjectInherit', 'None', 'Allow')))
            Set-Acl -LiteralPath $tight -AclObject $acl
            @(Get-UntrustedGrant -Path $tight -Kind Swap).Count   | Should -Be 0
            @(Get-UntrustedGrant -Path $tight -Kind Create).Count | Should -Be 0
        }

        It 'Get-UntrustedOwner flags a folder owned by an untrusted principal (an attacker-created subfolder)' {
            $d = Join-Path $TestDrive 'owned-by-users'; New-Item -ItemType Directory -Force $d | Out-Null
            Get-UntrustedOwner -Path $d | Should -BeNullOrEmpty -Because 'a folder we created is owned by the running account or Administrators'
            $acl = Get-Acl -LiteralPath $d
            try { $acl.SetOwner((New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-545'))); Set-Acl -LiteralPath $d -AclObject $acl -ErrorAction Stop }
            catch { Set-ItResult -Skipped -Because "cannot reassign ownership here: $_"; return }
            (Get-UntrustedOwner -Path $d) | Should -Match 'Users'
            { Assert-ProtectedDirectoryChain -Name T -Directory $d -Path (Join-Path $d 'x.cer') } | Should -Throw -ExpectedMessage '*OWNED by untrusted principal*'
            # restore ownership so TestDrive cleanup is unaffected
            $acl = Get-Acl -LiteralPath $d
            $acl.SetOwner([System.Security.Principal.WindowsIdentity]::GetCurrent().User); Set-Acl -LiteralPath $d -AclObject $acl
        }

        It 'Assert-CertificateOutputPath REFUSES a swappable folder on the path unless the override is set' {
            $root = Join-Path $TestDrive 'sw-root'; $sub = Join-Path $root 'drop'
            New-Item -ItemType Directory -Force $sub | Out-Null
            $me = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
            foreach ($d in $root, $sub) {
                $acl = Get-Acl -LiteralPath $d
                $acl.SetAccessRuleProtection($true, $false)
                foreach ($r in @($acl.Access)) { [void]$acl.RemoveAccessRule($r) }
                $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($me, 'FullControl', 'ContainerInherit, ObjectInherit', 'None', 'Allow')))
                Set-Acl -LiteralPath $d -AclObject $acl
            }
            { Assert-CertificateOutputPath -Name T -Path (Join-Path $sub 'x.cer') -AllowedRoots @($root) } | Should -Not -Throw
            # Users may delete/rename the intermediate folder -> refused (fail closed)
            $acl = Get-Acl -LiteralPath $sub
            $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
                (New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-545')), 'Modify', 'ContainerInherit, ObjectInherit', 'None', 'Allow')))
            Set-Acl -LiteralPath $sub -AclObject $acl
            { Assert-CertificateOutputPath -Name T -Path (Join-Path $sub 'x.cer') -AllowedRoots @($root) } | Should -Throw -ExpectedMessage '*can delete, rename or write to*'
            # ...also when it is the ROOT itself that is swappable
            { Assert-CertificateOutputPath -Name T -Path (Join-Path $sub 'x.cer') -AllowedRoots @($sub) } | Should -Throw -ExpectedMessage '*can delete, rename or write to*'
            # the explicit override turns the refusal into acceptance
            $script:AllowUnprotectedOutput = $true
            try { { Assert-CertificateOutputPath -Name T -Path (Join-Path $sub 'x.cer') -AllowedRoots @($root) } | Should -Not -Throw }
            finally { $script:AllowUnprotectedOutput = $false }
        }

        It 'Assert-CertificateOutputPath checks the chain ABOVE the root too: a junction ancestor is refused' {
            $away = Join-Path $TestDrive 'anc-away'; New-Item -ItemType Directory -Force (Join-Path $away 'Certificates') | Out-Null
            $jparent = Join-Path $TestDrive 'anc-jparent'
            try { New-Item -ItemType Junction -Path $jparent -Target $away -ErrorAction Stop | Out-Null }
            catch { Set-ItResult -Skipped -Because "cannot create a junction here: $_"; return }
            # the allowed root itself is a real folder... reached THROUGH a junction one level up
            $root = Join-Path $jparent 'Certificates'
            { Assert-CertificateOutputPath -Name T -Path (Join-Path $root 'x.cer') -AllowedRoots @($root) } | Should -Throw -ExpectedMessage '*reparse point*'
        }

        It 'Assert-ProtectedDirectoryChain fails CLOSED when an ACL cannot be read (override turns it into a warning)' {
            $d = Join-Path $TestDrive 'acl-unreadable'; New-Item -ItemType Directory -Force $d | Out-Null
            Mock Get-Acl { throw 'simulated: access to the security descriptor is denied' }
            { Assert-ProtectedDirectoryChain -Name T -Directory $d -Path (Join-Path $d 'x.cer') } | Should -Throw -ExpectedMessage '*could not be read*'
            $script:AllowUnprotectedOutput = $true
            try { { Assert-ProtectedDirectoryChain -Name T -Directory $d -Path (Join-Path $d 'x.cer') 3>$null } | Should -Not -Throw }
            finally { $script:AllowUnprotectedOutput = $false }
        }

        It 'Get-UntrustedGrant ignores inherit-only ACEs when judging the folder itself' {
            $d = Join-Path $TestDrive 'acl-inheritonly'; New-Item -ItemType Directory -Force $d | Out-Null
            $me = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
            $acl = Get-Acl -LiteralPath $d
            $acl.SetAccessRuleProtection($true, $false)
            foreach ($r in @($acl.Access)) { [void]$acl.RemoveAccessRule($r) }
            $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($me, 'FullControl', 'ContainerInherit, ObjectInherit', 'None', 'Allow')))
            # the C:\-style "Modify, subfolders and files only" entry: applies to CHILDREN, not to this folder
            $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
                (New-Object System.Security.Principal.SecurityIdentifier('S-1-5-11')), 'Modify', 'ContainerInherit, ObjectInherit', 'InheritOnly', 'Allow')))
            Set-Acl -LiteralPath $d -AclObject $acl
            @(Get-UntrustedGrant -Path $d -Kind Swap).Count | Should -Be 0 -Because 'an inherit-only ACE grants nothing on the folder itself'
            # ...but a child created beneath it inherits the effective Modify and IS swappable
            $child = Join-Path $d 'child'; New-Item -ItemType Directory -Force $child | Out-Null
            @(Get-UntrustedGrant -Path $child -Kind Swap) -join ' ' | Should -Match 'Authenticated Users'
            # ...and a FILE created in the folder inherits it too - which is what the File kind and
            # the delivery-folder check exist for: the folder passes every swap check, yet the
            # staging file and the delivered certificate would be writable by Authenticated Users
            @(Get-UntrustedGrant -Path $d -Kind File) -join ' ' | Should -Match 'Authenticated Users'
            { Assert-ProtectedDirectoryChain -Name T -Directory $d -Path (Join-Path $d 'x.cer') } | Should -Throw -ExpectedMessage '*FILES created inside it*'
            $script:AllowUnprotectedOutput = $true
            try { { Assert-ProtectedDirectoryChain -Name T -Directory $d -Path (Join-Path $d 'x.cer') 3>$null } | Should -Not -Throw }
            finally { $script:AllowUnprotectedOutput = $false }
        }

        It 'Get-UntrustedGrant -Kind File judges what the delivery folder hands to FILES: files-only Modify and inheritable append-data are refused, a folders-only entry is not' {
            $me = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
            $authUsers = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-11')
            $mk = {
                param($name, $rights, $inherit, $prop)
                $d = Join-Path $TestDrive "acl-file-$name"; New-Item -ItemType Directory -Force $d | Out-Null
                $acl = Get-Acl -LiteralPath $d
                $acl.SetAccessRuleProtection($true, $false); foreach ($r in @($acl.Access)) { [void]$acl.RemoveAccessRule($r) }
                $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($me, 'FullControl', 'ContainerInherit, ObjectInherit', 'None', 'Allow')))
                $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($authUsers, $rights, $inherit, $prop, 'Allow')))
                Set-Acl -LiteralPath $d -AclObject $acl
                $d
            }
            # 1. "Modify - files only" (ObjectInherit + InheritOnly): nothing on the folder, Modify on every file in it
            $filesOnly = & $mk 'filesonly' 'Modify' 'ObjectInherit' 'InheritOnly'
            @(Get-UntrustedGrant -Path $filesOnly -Kind Swap).Count   | Should -Be 0 -Because 'the folder itself grants nothing'
            @(Get-UntrustedGrant -Path $filesOnly -Kind Create).Count | Should -Be 0
            @(Get-UntrustedGrant -Path $filesOnly -Kind File) -join ' ' | Should -Match 'Authenticated Users'
            $f = Join-Path $filesOnly 'proof.staging.cer'; Set-Content -LiteralPath $f -Value 'x'
            $inherited = @((Get-Acl -LiteralPath $f).GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]) |
                Where-Object { $_.IdentityReference.Value -eq 'S-1-5-11' -and $_.IsInherited })
            $inherited.Count | Should -BeGreaterThan 0 -Because 'a file created in the folder inherits the files-only entry'
            ([int]$inherited[0].FileSystemRights -band [int][System.Security.AccessControl.FileSystemRights]::WriteData) | Should -Not -Be 0
            { Assert-ProtectedDirectoryChain -Name T -Directory $filesOnly -Path (Join-Path $filesOnly 'x.cer') } | Should -Throw -ExpectedMessage '*FILES created inside it*'
            # 2. create-subfolder (FILE_APPEND_DATA) with ObjectInherit: tolerated on the folder (Create only) but APPEND on a file
            $append = & $mk 'append' 'CreateDirectories' 'ContainerInherit, ObjectInherit' 'None'
            @(Get-UntrustedGrant -Path $append -Kind Swap).Count | Should -Be 0
            @(Get-UntrustedGrant -Path $append -Kind Create) -join ' ' | Should -Match 'Authenticated Users'
            @(Get-UntrustedGrant -Path $append -Kind File) -join ' ' | Should -Match 'Authenticated Users' -Because 'FILE_APPEND_DATA on a file lets the holder append to the certificate'
            # 3. Modify for SUBFOLDERS only (ContainerInherit + InheritOnly): files never inherit it - not a file-tamper grant
            $foldersOnly = & $mk 'foldersonly' 'Modify' 'ContainerInherit' 'InheritOnly'
            @(Get-UntrustedGrant -Path $foldersOnly -Kind File).Count | Should -Be 0 -Because 'a ContainerInherit-only entry does not propagate to files'
            $f2 = Join-Path $foldersOnly 'proof.cer'; Set-Content -LiteralPath $f2 -Value 'x'
            @((Get-Acl -LiteralPath $f2).GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]) | Where-Object { $_.IdentityReference.Value -eq 'S-1-5-11' }).Count | Should -Be 0
            { Assert-ProtectedDirectoryChain -Name T -Directory $foldersOnly -Path (Join-Path $foldersOnly 'x.cer') } | Should -Not -Throw
            # 4. read-only for Everyone on files is not a write-class grant
            $ro = & $mk 'readonly' 'ReadAndExecute' 'ContainerInherit, ObjectInherit' 'None'
            @(Get-UntrustedGrant -Path $ro -Kind File).Count | Should -Be 0
            # 4b. write-ATTRIBUTES on files only: not content, but FSCTL_SET_REPARSE_POINT accepts a
            #     file handle opened with FILE_WRITE_ATTRIBUTES alone, so the delivered certificate
            #     could be turned into a reparse point after the plain-file post-condition
            $wa = & $mk 'writeattr' 'WriteAttributes' 'ObjectInherit' 'InheritOnly'
            @(Get-UntrustedGrant -Path $wa -Kind Swap).Count | Should -Be 0
            @(Get-UntrustedGrant -Path $wa -Kind File) -join ' ' | Should -Match 'Authenticated Users'
            # 4c. the C:\ default "CREATOR OWNER: Full Control, subfolders and files only" - a
            #     placeholder the file system replaces with the CREATING account, i.e. the running
            #     account certreq runs as. Every unprotected folder under C:\ carries it; it must NOT
            #     be refused, or no real output folder would pass.
            $co = Join-Path $TestDrive 'acl-file-creatorowner'; New-Item -ItemType Directory -Force $co | Out-Null
            $acl = Get-Acl -LiteralPath $co
            $acl.SetAccessRuleProtection($true, $false); foreach ($r in @($acl.Access)) { [void]$acl.RemoveAccessRule($r) }
            $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($me, 'FullControl', 'ContainerInherit, ObjectInherit', 'None', 'Allow')))
            $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
                (New-Object System.Security.Principal.SecurityIdentifier('S-1-3-0')), 'FullControl', 'ContainerInherit, ObjectInherit', 'InheritOnly', 'Allow')))
            Set-Acl -LiteralPath $co -AclObject $acl
            @(Get-UntrustedGrant -Path $co -Kind File).Count | Should -Be 0 -Because 'CREATOR OWNER resolves to the creating (running, trusted) account'
            { Assert-ProtectedDirectoryChain -Name T -Directory $co -Path (Join-Path $co 'x.cer') } | Should -Not -Throw
            $f3 = Join-Path $co 'proof.cer'; Set-Content -LiteralPath $f3 -Value 'x'
            $onFile = @((Get-Acl -LiteralPath $f3).GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]) | ForEach-Object { $_.IdentityReference.Value })
            $onFile | Should -Not -Contain 'S-1-3-0' -Because 'the placeholder never survives onto the file'
            $onFile | Should -Contain $me.Value -Because 'it resolves to the account that created the file'
            # ...whereas CREATOR GROUP resolves to the running account's PRIMARY GROUP (possibly Domain Users) and stays untrusted
            $cg = Join-Path $TestDrive 'acl-file-creatorgroup'; New-Item -ItemType Directory -Force $cg | Out-Null
            $acl = Get-Acl -LiteralPath $cg
            $acl.SetAccessRuleProtection($true, $false); foreach ($r in @($acl.Access)) { [void]$acl.RemoveAccessRule($r) }
            $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($me, 'FullControl', 'ContainerInherit, ObjectInherit', 'None', 'Allow')))
            $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
                (New-Object System.Security.Principal.SecurityIdentifier('S-1-3-1')), 'Modify', 'ObjectInherit', 'InheritOnly', 'Allow')))
            Set-Acl -LiteralPath $cg -AclObject $acl
            @(Get-UntrustedGrant -Path $cg -Kind File) -join ' ' | Should -Match 'CREATOR GROUP'
            # 5. ...and the trusted set applies to the File kind as well
            $saved = $script:TrustedSids
            try {
                $script:TrustedSids = Get-TrustedPrincipalSet -Extra 'S-1-5-11'
                @(Get-UntrustedGrant -Path $filesOnly -Kind File).Count | Should -Be 0
            } finally { $script:TrustedSids = $saved }
        }

        It 'Assert-CertificateOutputPath picks the DEEPEST matching root and refuses a root that is a junction' {
            $outer = Join-Path $TestDrive 'deep-outer'; $inner = Join-Path $outer 'Certificates'
            New-Item -ItemType Directory -Force $inner | Out-Null
            # both roots contain the file; the walk must stop at the inner one (no intermediate components)
            { Assert-CertificateOutputPath -Name T -Path (Join-Path $inner 'x.cer') -AllowedRoots @($outer, $inner) } | Should -Not -Throw
            $away = Join-Path $TestDrive 'deep-away'; New-Item -ItemType Directory -Force $away | Out-Null
            $jroot = Join-Path $TestDrive 'deep-jroot'
            try { New-Item -ItemType Junction -Path $jroot -Target $away -ErrorAction Stop | Out-Null }
            catch { Set-ItResult -Skipped -Because "cannot create a junction here: $_"; return }
            { Assert-CertificateOutputPath -Name T -Path (Join-Path $jroot 'x.cer') -AllowedRoots @($jroot) } | Should -Throw -ExpectedMessage '*reparse point*'
        }

        It 'Move-RetrievedCertificate refuses to place the .rsp through a folder or link named x.rsp' {
            $dest = Join-Path $TestDrive 'rsp-dest.cer'
            $trap = Join-Path $TestDrive 'rsp-dest.rsp'; New-Item -ItemType Directory -Force $trap | Out-Null
            $tmp = Join-Path $TestDrive 'certreq-rsp.cer'; 'cert' | Out-File $tmp -Encoding ascii
            'rsp' | Out-File ([System.IO.Path]::ChangeExtension($tmp, '.rsp')) -Encoding ascii
            Move-RetrievedCertificate -TempCer $tmp -Destination $dest -KeepRspFile 3>$null 6>$null
            Test-Path $dest -PathType Leaf | Should -BeTrue
            @(Get-ChildItem $trap).Count | Should -Be 0 -Because 'nothing may be moved INTO a folder named x.rsp'
        }

        It 'Export-TrackingData never moves the CSV into a folder occupying the tracking-file name' {
            $trap = Join-Path $TestDrive 'trap.csv'; New-Item -ItemType Directory -Force $trap | Out-Null
            $rows = @([pscustomobject]@{ RequestFile = 'a.req'; RequestID = 1; SubmitTime = 't'; Status = 'Issued'; OutputCertFile = 'a.cer'; LastCheckTime = 't'; ErrorMessage = '' })
            { Export-TrackingData -Data $rows -Path $trap } | Should -Throw
            @(Get-ChildItem $trap -Force).Count | Should -Be 0
            @(Get-ChildItem $TestDrive -Filter 'trap.csv.*.tmp').Count | Should -Be 0 -Because 'the temp file is cleaned up on failure'
        }

        It 'Move-RetrievedCertificate delivers the temp file, moves a predecessor aside, and places/drops the .rsp' {
            $dest = Join-Path $TestDrive 'deliver.cer'
            'old' | Out-File $dest -Encoding ascii
            $tmp = Join-Path $TestDrive 'certreq-tmp.cer'; 'new' | Out-File $tmp -Encoding ascii
            'rsp' | Out-File ([System.IO.Path]::ChangeExtension($tmp, '.rsp')) -Encoding ascii
            Move-RetrievedCertificate -TempCer $tmp -Destination $dest -KeepRspFile
            (Get-Content $dest -Raw).Trim() | Should -BeExactly 'new'
            Test-Path $tmp | Should -BeFalse
            @(Get-ChildItem $TestDrive -Filter 'deliver.superseded-*.cer').Count | Should -Be 1 -Because 'the predecessor is kept, not deleted'
            Test-Path (Join-Path $TestDrive 'deliver.rsp') | Should -BeTrue -Because '-KeepRspFile places the .rsp beside the destination'
            # without -KeepRspFile the .rsp stays with the staging file for the caller''s cleanup - the destination gets none
            $dest2 = Join-Path $TestDrive 'deliver2.cer'
            $tmp2 = Join-Path $TestDrive 'certreq-tmp2.cer'; 'new2' | Out-File $tmp2 -Encoding ascii
            'rsp' | Out-File ([System.IO.Path]::ChangeExtension($tmp2, '.rsp')) -Encoding ascii
            Move-RetrievedCertificate -TempCer $tmp2 -Destination $dest2
            Test-Path $dest2 | Should -BeTrue
            Test-Path (Join-Path $TestDrive 'deliver2.rsp') | Should -BeFalse
        }

        It 'New-TempCertificatePath stages INSIDE the destination folder (never %TEMP%), under a random .cer name' {
            $dest = Join-Path $TestDrive 'out\srv1.cer'
            $a = New-TempCertificatePath -Destination $dest
            $b = New-TempCertificatePath -Destination $dest
            [System.IO.Path]::GetDirectoryName($a) | Should -BeExactly (Join-Path $TestDrive 'out')
            # TestDrive itself lives under %TEMP%, so test the FOLDER identity, not a prefix: the
            # staging file sits in the destination's own folder, not directly in the TEMP folder
            [System.IO.Path]::GetDirectoryName($a) | Should -Not -Be ([System.IO.Path]::GetTempPath().TrimEnd('\'))
            $a | Should -BeLike (Join-Path $TestDrive 'out\srv1.*.staging.cer')
            $a | Should -Not -Be $b -Because 'a random name cannot be pre-planted'
        }

        It 'Move-RetrievedCertificate leaves the delivered file inheriting the folder ACL (no explicit ACEs carried over)' {
            $dir = Join-Path $TestDrive 'acl-inherit'; New-Item -ItemType Directory -Force $dir | Out-Null
            $tmp = Join-Path $dir 'x.abc.staging.cer'; 'cert' | Out-File $tmp -Encoding ascii
            # give the staging file an explicit, loose ACE - what a shared TEMP would have handed it
            $acl = Get-Acl -LiteralPath $tmp
            $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
                (New-Object System.Security.Principal.SecurityIdentifier('S-1-5-11')), 'Modify', 'Allow')))
            Set-Acl -LiteralPath $tmp -AclObject $acl
            $dest = Join-Path $dir 'x.cer'
            Move-RetrievedCertificate -TempCer $tmp -Destination $dest
            $explicit = @((Get-Acl -LiteralPath $dest).Access | Where-Object { -not $_.IsInherited })
            $explicit.Count | Should -Be 0 -Because 'the delivered certificate must be governed by its folder alone'
            (Get-Acl -LiteralPath $dest).AreAccessRulesProtected | Should -BeFalse
        }

        It 'Move-StaleCertificateAside clears a pre-existing .cer without deleting it; no-op when absent' {
            $cer = Join-Path $TestDrive 'stale.cer'
            Move-StaleCertificateAside -Path $cer | Should -BeNullOrEmpty          # nothing there -> nothing moved
            'old cert' | Out-File $cer -Encoding ascii
            $aside = Move-StaleCertificateAside -Path $cer
            Test-Path $cer   | Should -BeFalse -Because 'the destination must be empty before certreq runs'
            Test-Path $aside | Should -BeTrue  -Because 'the earlier certificate is preserved, never deleted'
            $aside | Should -BeLike (Join-Path $TestDrive 'stale.superseded-*.cer')
            (Get-Content $aside -Raw).Trim() | Should -BeExactly 'old cert'
        }

        It 'Remove-AsideIfIdentical drops the aside copy only when the fresh file is byte-identical' {
            $cer = Join-Path $TestDrive 'ident.cer'
            'same' | Out-File $cer -Encoding ascii
            $aside = Move-StaleCertificateAside -Path $cer
            'different' | Out-File $cer -Encoding ascii                             # a resubmission: new certificate
            Remove-AsideIfIdentical -Path $cer -Aside $aside
            Test-Path $aside | Should -BeTrue -Because 'a different predecessor is kept'
            'same' | Out-File $cer -Encoding ascii                                  # the same certificate retrieved again
            Remove-AsideIfIdentical -Path $cer -Aside $aside
            Test-Path $aside | Should -BeFalse -Because 'an identical copy is redundant'
            Test-Path $cer   | Should -BeTrue
            { Remove-AsideIfIdentical -Path $cer -Aside $null } | Should -Not -Throw   # nothing was moved aside
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

        It 'carries a PSScriptInfo header (Test-ScriptFileInfo parses it; Version is semver)' {
            $info = Test-ScriptFileInfo -Path $script:Submit -ErrorAction Stop
            $info.Version | Should -Match '^\d+\.\d+\.\d+$'
            $info.Guid    | Should -Not -BeNullOrEmpty
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
            # -TrustedOutputPrincipal: the output-folder trust check runs BEFORE the CA check and
            # must pass on this machine's TEMP chain (see BeforeAll) for the CA failure to be reached.
            { & $script:Submit -CAConfig 'localhost\PESTER-NoSuchCA' -Mode Retrieve -WhatIf -TrustedOutputPrincipal $script:EnvTrusted `
                  -TrackingFile (Join-Path $TestDrive 'g3.csv') -OutputFolder (Join-Path $TestDrive 'g3') 3>$null 2>$null 6>$null } |
                Should -Throw -ExpectedMessage '*Cannot reach CA*'
        }

        It 'refuses to start when an output location is swappable by an untrusted principal, unless overridden' {
            $loose = Join-Path $TestDrive 'guard-loose'; New-Item -ItemType Directory -Force $loose | Out-Null
            $acl = Get-Acl -LiteralPath $loose
            $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
                (New-Object System.Security.Principal.SecurityIdentifier('S-1-5-21-111-222-333-513')), 'Modify', 'ContainerInherit, ObjectInherit', 'None', 'Allow')))
            Set-Acl -LiteralPath $loose -AclObject $acl
            { & $script:Submit -CAConfig 'localhost\PESTER-NoSuchCA' -Mode Retrieve -WhatIf -TrustedOutputPrincipal $script:EnvTrusted `
                  -TrackingFile (Join-Path $TestDrive 'g4.csv') -OutputFolder $loose 3>$null 6>$null } |
                Should -Throw -ExpectedMessage '*can delete, rename or write to*'
            # naming the principal as trusted, or accepting the risk, lets the run proceed to the CA check
            { & $script:Submit -CAConfig 'localhost\PESTER-NoSuchCA' -Mode Retrieve -WhatIf -TrustedOutputPrincipal (@($script:EnvTrusted) + 'S-1-5-21-111-222-333-513') `
                  -TrackingFile (Join-Path $TestDrive 'g4.csv') -OutputFolder $loose 3>$null 2>$null 6>$null } |
                Should -Throw -ExpectedMessage '*Cannot reach CA*'
            { & $script:Submit -CAConfig 'localhost\PESTER-NoSuchCA' -Mode Retrieve -WhatIf -AllowUnprotectedOutputFolder `
                  -TrackingFile (Join-Path $TestDrive 'g4.csv') -OutputFolder $loose 3>$null 2>$null 6>$null } |
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

            # The CA is always the one the operator NAMED (see BeforeDiscovery) - no auto-discovery
            # for a tier that deletes CA database rows.
            $script:Ca = $CAConfig
            $script:Ca | Should -Not -BeNullOrEmpty -Because 'the Lab tier requires an explicit -CAConfig'
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
                $call = @{ CAConfig = $script:Ca; TrackingFile = $script:Tracking; OutputFolder = $script:CertDir; Confirm = $false
                           TrustedOutputPrincipal = $script:EnvTrusted }   # this machine's TEMP-chain grants (see the outer BeforeAll)
                foreach ($k in $Params.Keys) { $call[$k] = $Params[$k] }   # a test may override a default (e.g. OutputFolder)
                # A run with failed/attention rows ends with a terminating error (non-zero exit for
                # automation). Collect the streamed output first, then the error, so the tests can
                # still judge both the rows and what was logged.
                $lines = [System.Collections.Generic.List[string]]::new()
                try { & $script:Submit @call *>&1 | ForEach-Object { $lines.Add("$_") } }
                catch { $lines.Add("TERMINATING: $($_.Exception.Message)") }
                $out = $lines -join "`n"
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

        It 'Retrieve mode with an explicit -OutputFolder writes the .cer there and updates the row' {
            $rows = @(Import-Csv $script:Tracking)
            $rows[0].Status = 'Unknown'
            Remove-Item -LiteralPath $rows[0].OutputCertFile -Force
            $rows | Export-Csv $script:Tracking -NoTypeInformation -Encoding utf8
            $altDir = Join-Path $TestDrive 'certs-alt'

            $r = script:Invoke-Submit @{ Mode = 'Retrieve'; OutputFolder = $altDir }
            $updated = $r.Rows | Where-Object { $_.RequestID -eq $rows[0].RequestID } | Select-Object -First 1
            $updated.Status | Should -BeExactly 'Issued'
            $updated.OutputCertFile | Should -Be (Join-Path $altDir ([IO.Path]::GetFileName($rows[0].OutputCertFile)))
            Test-Path $updated.OutputCertFile | Should -BeTrue
            Test-Path $rows[0].OutputCertFile | Should -BeFalse
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
