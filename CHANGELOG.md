# Changelog

All notable changes to this repository are recorded here. The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/); release tags follow [Semantic Versioning](https://semver.org/).

Each script also carries its own version in the `PSScriptInfo` header at the top of the file (`.VERSION`), bumped only when that script changes. The table under each release lists the script versions it ships, so a deployed copy can be checked without opening it:

```powershell
Test-ScriptFileInfo .\Submit-CertificateRequests.ps1 | Select-Object Name, Version
```

## [1.0.2] — 2026-09-05

### Fixed
- **Submit-CertificateRequests.ps1 → 1.0.2**
  - A retrieval whose certreq output said *Issued* but produced no `.cer` file was recorded as a final `Issued` and never retried. It is now recorded as `Error` (with the certreq output) and picked up again by the next `-Mode Retrieve`. The same rule applies at submit time.
  - `Issued` is now decided by certreq's exit code plus the presence of the certificate it wrote, independent of certreq's localized console wording. certreq writes into a private, randomly named staging file inside the (already verified) destination folder — never %TEMP%, whose inherited file ACL could let another account tamper with the certificate and, since a same-volume rename keeps a file's DACL, would travel onto the delivered file — and only a successful write is delivered to the destination (whose ACL the delivered file is then reset to inherit), where a file already present (a `-Force` resubmit; a retry of an unresolved row) is moved aside as `<name>.superseded-<UTC stamp>.cer` — never deleted — and that copy is removed again only if the fresh certificate is byte-identical. A pending, denied or failed request never touches the destination. A certificate written without a parsable `RequestId` line keeps `Issued` (and the row is treated as submitted on later runs, so it is not resubmitted) with the gap noted in `ErrorMessage`.
  - Values that reach the certreq command line (`-CAConfig`, `-CertificateTemplate`, and the `RequestID` / `OutputCertFile` read from the tracking CSV) are validated: a double quote or control character is rejected before certreq runs (`Start-Process -ArgumentList` joins its arguments into one command line, so an embedded quote could have been read as extra switches).
  - Concurrent runs on one tracking file are refused: an exclusive lock file (`<TrackingFile>.lock`, removed when the run ends) is held for the whole run. The lock is enforced by the file system, so it also covers runs from other machines and every alias of the folder path (drive letter vs. UNC, junctions); the tracking file's own name is canonicalized first (an 8.3 short name resolves to the long name, so both spellings share one lock) and a hard-linked or symlinked tracking file is refused. Two such runs previously overwrote each other's rows, and a later run would have resubmitted the lost requests. The tracking file is now replaced via a uniquely named temp file instead of a fixed `.tmp`.
  - A tracking row's `OutputCertFile` is treated as an identifier inside an operator-chosen boundary, not as an authority: it must be a rooted `.cer` path in an existing folder that resolves canonically beneath the tracking file's folder or the run's `-OutputFolder`, with no junction/symlink between root and file — checked when the row is read and again immediately before delivery. Rows failing this are skipped with an error, so a tampered CSV cannot steer the privileged delivery or move-aside at a file elsewhere. Every folder from the destination up to the volume/share root is checked, and the script **refuses** to deliver when any of them is a reparse point (a pre-planted or swapped-in junction), is owned by an untrusted principal (an attacker-created subfolder), can be deleted, renamed or written to by one (delete/rename lets such a user swap it for a junction inside the remaining check-then-write window; write-data or write-attributes access is all `FSCTL_SET_REPARSE_POINT` needs, so "create files" rights let them turn an *empty* folder, such as a freshly created output folder, into a junction in place), or has a security descriptor that cannot be read. Trusted principals are SYSTEM, Administrators, TrustedInstaller, the running account, its Domain/Enterprise Admins and the new `-TrustedOutputPrincipal` list — anything else, not just a fixed list of broad groups, is untrusted; the new `-AllowUnprotectedOutputFolder` switch downgrades these refusals to warnings. Only the create-*subfolder* right by itself (what the `C:\` root grants Users on itself) and inherit-only ACEs are tolerated, and the residual race they leave — a folder or junction planted under the exact destination name between check and delivery — is closed by delivering with a no-overwrite rename (`File.Move`, which fails when anything occupies the name) instead of `Move-Item -Force` (which would move the certificate *into* it). The `.rsp` companion, the move-aside and the tracking-file replacement use the same primitive. The `.rsp` written with `-KeepRspFile` gets the same destination-shape checks as the `.cer`.
  - An issued certificate that cannot be delivered (locked destination, denied rename, failed re-validation) now yields a new `Undelivered` status that keeps the RequestID and leaves the certificate in its staging file (path in `ErrorMessage`). `Undelivered` counts as submitted on later runs — even without a RequestID — so it is never resubmitted automatically; `Retrieve` re-fetches it when a RequestID is present and reports the rest for manual reconciliation. Previously the delivery exception produced an `Error` row with no RequestID, which a later Submit would have resubmitted as a duplicate request.
  - Under `-Confirm`, the cmdlets inside an already-approved action (`Export-Csv` writing the tracking checkpoint, `Out-File` for the log, `Remove-Item`, `New-Item`) inherited the low confirm preference and raised their own prompts; declining the checkpoint prompt after certreq had already submitted would have lost the RequestID and caused a duplicate request on the next run. They now pass `-Confirm:$false`: the checkpoint is a mandatory part of the approved submission.
  - The destination path itself must be a plain file or absent: an existing folder or a link named `x.cer` is refused (Move-Item onto a folder would have moved the certificate *into* it, and a junction there would carry it anywhere), and delivery verifies that a plain file resulted before the row is finalized.
- **Sync-ADCSTemplate.ps1 → 1.0.1**
  - The template create runs with `-ErrorAction Stop` and checks the returned object. A non-terminating `New-ADObject` failure previously skipped the catch block and the companion-OID rollback, printed a green *Created template* line with an empty DN, and returned `$null` — read by the caller as "declined at the prompt" — so the run exited 0 with an orphaned companion OID object.
  - After the create, the OID is re-queried on the same server; if another template claimed it concurrently (the pre-flight check is check-then-create and AD does not enforce OID uniqueness), the run fails instead of leaving two templates sharing one OID. Any failure after the create — the collision, or the re-query itself failing — rolls the new template back by its captured GUID; the companion OID object is removed only once the template is gone, and if the template cannot be removed both survivors are reported for manual cleanup instead of half of the pair being deleted.
- **Add-CertificateEnrollmentPolicyServerToGpo.ps1 → 1.0.1**
  - The domain objectGUID for the AD Enrollment Policy row is resolved **before** any GPO write, and a lookup failure aborts the run with nothing written. The row itself is now written and verified **before** the CEP entry; when the row's confirmation was declined and the GPO carries no complete one (URL `LDAP:` *and* this domain's PolicyID — a half-written row from an interrupted run does not count), the CEP entry is not written and the run stops there (no root Flags, marker, Auto-Enrollment or sibling changes follow). The `(Default)` marker and `-ReplaceExisting` removals run only when a **complete** CEP entry for this URL exists (URL and PolicyID as requested). For the sibling cleanup that is checked *live* through the GPMC API immediately beforehand and again after each confirmation is accepted (a `-Confirm` prompt can stay open indefinitely), together with a live re-check that the candidate sibling still serves this PolicyID under a different URL, regardless of whether this run wrote the entry — a concurrent writer may have removed or changed either since — so a declined entry prompt, a half-written entry from an interrupted run, or an entry that vanished mid-run cannot leave a marker pointing at nothing, delete the only working endpoint, or act for a PolicyID the retained entry does not serve. Previously the CEP entry was written first and a failed lookup became a warning — leaving a published GPO that removes the AD enrollment policy from every client in scope (the very outage the row exists to prevent); a declined AD-row prompt produced the same half-state.
  - A GUID-shaped `-GpoName` falls back to the GPO ID only after an independent `Get-GPO -All` listing proves no GPO carries that display name; previously *any* name-lookup failure (authorization, transient RPC/SYSVOL error) triggered the fallback and could retarget the run at whichever GPO had that ID.
  - `registry.pol` state reads (`Get-PolValue`, `Get-PolEntries`) returned the *first* matching record; they now replay the records in file order (last write wins; `**del.<name>`, `**delvals.`, `**DeleteValues` and `**DeleteKeys` honoured — the latter with the Group Policy engine's real semantics, verified against the LGPO author's corrections to MS-GPREG: its data is a list of *full* hive-relative key paths and the record's own key field is ignored, so an item naming the key or any ancestor deletes it — and `**soft.<name>` writes only an absent value, while `**Comment:` and `**SecureKey` are not values at all). A deleted LDAP row or endpoint is therefore not reported as present to the prerequisite and replacement guards — while `-Remove` judges presence by the *physical* records at the hashed key, so an entry whose values a later `**delvals.` wiped (the damaged ordering the script warns about) can still be removed and re-added as the recovery text says — and any other `**`-prefixed instruction makes the read fail closed instead of being skipped. The deletion-order damage check is judged per value, so a deletion between an obsolete record and its replacement no longer raises a false *DAMAGED* warning. Value names now compare case-insensitively, as the registry does.
  - The `registry.pol` parser validates the `PReg`/version-1 header, refuses files over 64 MB before reading them, and treats trailing bytes after the last record as corruption instead of silently ignoring them.
- **Add-CertificateEnrollmentPolicyServerOffline.ps1 → 1.0.1** — for the GP locations the domain objectGUID for the AD Enrollment Policy row is resolved **before** anything is written and the row is written before the CEP entry; on a domain-joined machine a lookup failure now aborts the run with nothing written, and a declined AD-row confirmation stops the run before the CEP entry (same half-state as above, limited to one machine); the `(Default)` marker and `-ReplaceExisting` removals run only when a complete CEP entry for this URL exists (URL and PolicyID as requested; for the cleanup, checked in the live registry right beforehand and again after each confirmation is accepted, together with a re-check of the candidate sibling, regardless of whether this run wrote it). A workgroup machine still skips the row with a warning.
- **Set-ADCSTemplateValidity.ps1 → 1.0.1**
  - A run in which any search or template update failed now ends with a terminating error (non-zero exit) after the structured output and summary; previously the errors were only written to the console and the process exited 0, so automation would treat a partial privileged update as success.
  - Help and README advertised `?` as an LDAP wildcard; LDAP substring filters only know `*`, so `?` always matched literally. Documentation corrected (behaviour unchanged).
- README: the *How It Works* section for Set-ADCSTemplateValidity described the byte array as written via `DirectoryEntry.InvokeSet()`; the script deliberately assigns through the property cache (InvokeSet unrolls the array). Corrected.
- CI: both Pester legs now also fail when the run result is not `Passed` (a discovery error or a container/BeforeAll failure produces no failed *test*, so the previous `FailedCount` check alone could let the step go green). `actions/checkout` is pinned to a commit SHA and Pester / PSScriptAnalyzer to exact versions (6.1.0 / 1.25.0), so a moved tag or a new gallery release cannot change what runs on the runner.
- Tests: the Submit-CertificateRequests Lab tier no longer auto-discovers a CA from AD — `-RunLab` now requires an explicit `-CAConfig` naming the lab CA (the tier deletes the CA database rows it creates; the first CA registered in AD could be production).

### Added
- **Submit-CertificateRequests.ps1**: `-AllowUnprotectedOutputFolder` switch and `-TrustedOutputPrincipal` list (see above).
- Tests: effective-value / deletion-record fixtures and header/trailing-junk cases for the GPO `registry.pol` reader; a `Get-AttrCanonical` regression test for the backslash-plus-separator collision (verified not to collide); the `?`-is-literal assertion for the validity script's LDAP escaper; unit tests for the Submit script's argument, output-path and move-aside helpers.

| Script | Version |
|---|---|
| Set-ADCSTemplateValidity.ps1 | **1.0.1** |
| Submit-CertificateRequests.ps1 | **1.0.2** |
| Sync-ADCSTemplate.ps1 | **1.0.1** |
| Add-CertificateEnrollmentPolicyServerOffline.ps1 | **1.0.1** |
| Add-CertificateEnrollmentPolicyServerToGpo.ps1 | **1.0.1** |

## [1.0.1] — 2026-09-02

### Fixed
- **Submit-CertificateRequests.ps1 → 1.0.1** — `-Mode Retrieve` ignored `-OutputFolder`: retrieved `.cer` files always went to the path recorded in the tracking CSV at submit time. An explicit `-OutputFolder` on a Retrieve run now redirects each file there (file name kept) and updates the tracking row; without it, the recorded path is still used, so a Retrieve from another working directory keeps landing where Submit put the files.

### Added
- `PSScriptInfo` version header on all five scripts (`.VERSION`, `.GUID`, `.AUTHOR`, `.PROJECTURI`, `.LICENSEURI`, `.RELEASENOTES`).
- This changelog.
- A Static-tier test in every suite asserting the header parses via `Test-ScriptFileInfo` and the version is semver.
- Lab-tier test: Retrieve with an explicit `-OutputFolder` writes the `.cer` there and updates the row.

| Script | Version |
|---|---|
| Set-ADCSTemplateValidity.ps1 | 1.0.0 |
| Submit-CertificateRequests.ps1 | **1.0.1** |
| Sync-ADCSTemplate.ps1 | 1.0.0 |
| Add-CertificateEnrollmentPolicyServerOffline.ps1 | 1.0.0 |
| Add-CertificateEnrollmentPolicyServerToGpo.ps1 | 1.0.0 |

## [1.0.0] — 2026-08-30

Initial release.

- **Set-ADCSTemplateValidity.ps1** — bulk validity / renewal-overlap updates on certificate templates (wildcard matching, `-WhatIf`).
- **Submit-CertificateRequests.ps1** — batch CSR submission via `certreq.exe` with resume-safe CSV tracking and later retrieval.
- **Sync-ADCSTemplate.ps1** — cross-forest template copy (JSON export/import or direct forest-to-forest), OID handling modes, composable enrollment ACL; target forest needs no AD CS.
- **Add-CertificateEnrollmentPolicyServerOffline.ps1** — register an EJBCA/MS-XCEP enrollment-policy server on one machine, computed offline.
- **Add-CertificateEnrollmentPolicyServerToGpo.ps1** — author the same setting into a domain GPO for fleet-wide rollout.
- **Template library** (`Templates/`) — all 33 certutil default templates as importable JSON, EJBCA-ready variants, and latest-compatibility copies.
- Four-tier Pester suites (Unit / Static / Guard / Lab) per script, PSScriptAnalyzer settings, CI on both PowerShell 5.1 and 7.

| Script | Version |
|---|---|
| Set-ADCSTemplateValidity.ps1 | 1.0.0 |
| Submit-CertificateRequests.ps1 | 1.0.0 |
| Sync-ADCSTemplate.ps1 | 1.0.0 |
| Add-CertificateEnrollmentPolicyServerOffline.ps1 | 1.0.0 |
| Add-CertificateEnrollmentPolicyServerToGpo.ps1 | 1.0.0 |

[1.0.2]: https://github.com/TheOmnilord/ADCS/compare/v1.0.1...v1.0.2
[1.0.1]: https://github.com/TheOmnilord/ADCS/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/TheOmnilord/ADCS/releases/tag/v1.0.0