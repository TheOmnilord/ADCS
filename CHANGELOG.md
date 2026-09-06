# Changelog

All notable changes to this repository are recorded here. The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/); release tags follow [Semantic Versioning](https://semver.org/).

Each script also carries its own version in the `PSScriptInfo` header at the top of the file (`.VERSION`), bumped only when that script changes. The table under each release lists the script versions it ships, so a deployed copy can be checked without opening it:

```powershell
Test-ScriptFileInfo .\Submit-CertificateRequests.ps1 | Select-Object Name, Version
```

## [1.0.6] — 2026-09-06

Two adversarial review rounds after v1.0.5 — a Codex/gpt-6-astra full-repository review, then a thorough Claude review (13 finders with 3-refuter verification). Every finding was verified against the code before fixing.

### Fixed — Codex (gpt-6-astra) full-repository review
- **Sync-ADCSTemplate.ps1**
  - Every known attribute of an import is validated for type, shape and range and converted before anything is created; a malformed value **refuses the import** with the attribute named. The casts ran bare, so under the default error preference a failed `[int]` cast was only statement-terminating and a tampered export silently *dropped* the attribute — `msPKI-RA-Signature`, a CA-enforced control, among them. JSON shapes a bare cast coerced (`0.4` to `0`, a three-element period array, `"5"` to `5`) are refused; OID-list attributes must hold dotted OIDs; a giant or non-finite number (`1e40`, Infinity, NaN) is refused with the attribute named rather than a raw .NET error.
  - An issuance policy OID (`msPKI-Certificate-Policy` / `msPKI-RA-Policies`) that the **target** forest already links to a group through Authentication Mechanism Assurance refuses the import unless the new `-AllowLinkedIssuancePolicy` switch is given: certificates from the copy would otherwise grant that group's membership at logon to everyone the copy's enrollment ACL admits, with no link ever copied. Validate's throwaway copies are exempt.
- **Add-CertificateEnrollmentPolicyServerOffline.ps1**
  - Before any write, every existing key from the hive root down to the target is checked for a registry **symbolic link**, an untrusted owner, or write-class rights for an untrusted principal. A link planted where `PolicyServers` did not exist yet would have carried an elevated first-time write to whatever key it pointed at; `-LiteralPath` stops wildcards, not link traversal.
  - String values and the `(Default)` marker are written as `REG_SZ` explicitly and every value's **kind** is verified (`Set-ItemProperty` without `-Type` kept an existing wrong kind); the cmdlets inside an approved action pass `-Confirm:$false` and a removal is verified before it is reported.
- **Add-CertificateEnrollmentPolicyServerToGpo.ps1**
  - The gates that set the `(Default)` marker and let `-ReplaceExisting` delete siblings require the replacing entry to be complete in **both** the live GPMC view and the effective `registry.pol` replay; a freshly written AD row or CEP entry must also show in the replay, or the run fails. The mutation cmdlets pass `-Confirm:$false` and a removal or marker clear is verified from `registry.pol` before it is reported.
- **Submit-CertificateRequests.ps1** — `-Mode Retrieve` refuses to deliver a request whose destination already holds the certificate of a *different* request for the same request file, so an older request no longer overwrites a newer certificate.

### Fixed — thorough Claude review (20 confirmed findings)
- **Set-ADCSTemplateValidity.ps1 → 1.0.4**
  - **(high)** The script was `begin`/`process`/`end`; under Windows PowerShell 5.1, `powershell.exe -File` with a non-console stdin (a scheduler, CI, WinRM/psexec, or the `< NUL` idiom) never ran the `process{}` block, so a privileged unattended run searched nothing, changed nothing, printed nothing and exited 0. It is now a flat body.
  - A confirmation failure (non-interactive host, `ConfirmImpact` High, no `-Confirm:$false`) is caught, counted and emitted as an `Error` row instead of escaping the loop uncounted.
  - The run-level failure is raised with a non-terminating error plus `exit 1` rather than `throw`, so a caller that captures or pipes the structured report keeps it while automation still sees a non-zero exit.
- **Submit-CertificateRequests.ps1 → 1.0.7**
  - `Get-RequestFiles` reads the drop folder with `-LiteralPath` and an exact-extension filter: a folder named e.g. `CSR[prod]` was globbed as a character class and matched nothing (exit 0, work undone), and `-Filter '*.req'` also matched longer extensions like `.reqbak` / `.request`, submitting stray backups to the CA.
  - `Export-TrackingData` writes the checkpoint with `-LiteralPath` and `Remove-RspFile` removes with `-LiteralPath`: a tracking path or request name containing `[ ]` lost the RequestID (resubmitted as a duplicate) or left the `.rsp` behind.
  - certreq's redirected stdout/stderr are read with `-Encoding Oem`: the default was ANSI on 5.1 but UTF-8 on 7, so a localized CA's `ErrorMessage` differed by engine and was lossy on 7.
  - `Get-DestinationOwnerConflict` is now **directional** (only a strictly newer request owns the shared destination), so a `-Force` renewal that goes `Pending` can be retrieved instead of blocking every later run forever; and it guards its `[System.IO.Path]` calls (an invalid path character in a hand-edited row threw on 5.1 and aborted the whole batch).
  - `-Mode Retrieve` computes the redirected destination and stamps a legacy row's `CAConfig` only inside the approved `ShouldProcess` branch, so a declined row is left byte-identical in the tracking file.
- **Sync-ADCSTemplate.ps1 → 1.0.5**
  - The documented `DOMAIN\user@domain` principal form now takes the UPN-only resolution and sAMAccountName shadow check (matching on the raw key let a prefixed key skip to the sAMAccountName lookup, so a planted account could still capture the grant).
  - The dotted-OID validation regexes are anchored with the true end-of-string anchor instead of the end-of-line anchor, so a trailing newline in a tampered `msPKI-Cert-Template-OID` can no longer pass validation and bypass the template-OID uniqueness search.
  - `-UpgradeCompatibility` refuses a schema-2 source carrying `msPKI-RA-Application-Policies` — its encoding differs at v3/v4, so upgrading in place would silently drop the RA-signature application-policy requirement.
- **Add-CertificateEnrollmentPolicyServerToGpo.ps1 → 1.0.5** / **Add-CertificateEnrollmentPolicyServerOffline.ps1 → 1.0.5**
  - Root Flags are no longer written when no usable policy-server entry exists in scope (the CEP entry declined **and** no AD Enrollment Policy row): writing PolicyServers root values there activated GP CEP configuration with zero servers, so clients lost the AD enrollment policy fleet-wide.
  - (GPO) The Auto-Enrollment write is now verified and its key is included in the deletion-order damage scan — a deletion record ordered after the AE values defeated autoenrollment silently while the run reported `applied=True`.
- **Templates/**: `MaxCompat/Default/CrossCA.json` is restored to its stock schema v2 (it cannot be upgraded in place — see Sync 1.0.5), and `Templates/README.md`'s provenance count is corrected to the true **8 upgraded defaults + 4 EJBCA variants**.
- Tests: coverage added for every fix — wildcard-path and exact-extension cases for `Get-RequestFiles` / `Export-TrackingData`, directional and fail-closed cases for `Get-DestinationOwnerConflict`, the anchored-OID assertion, a guard that the validity script carries no begin/process/end blocks, and a Lab test for the retained-overlap refusal.

| Script | Version |
|---|---|
| Set-ADCSTemplateValidity.ps1 | **1.0.4** |
| Submit-CertificateRequests.ps1 | **1.0.7** |
| Sync-ADCSTemplate.ps1 | **1.0.5** |
| Add-CertificateEnrollmentPolicyServerOffline.ps1 | **1.0.5** |
| Add-CertificateEnrollmentPolicyServerToGpo.ps1 | **1.0.5** |

## [1.0.5] — 2026-09-06

Help text only — no code changes. Each script's comment-based help now describes the behaviour its 1.0.2–1.0.4 fixes introduced, matching the README:

- **Submit-CertificateRequests.ps1 → 1.0.5** — the notes cover the `CAConfig` column and the CA check on Retrieve, up-front unique certificate names, the `Unknown`-without-RequestID rule, and the non-zero exit on failed or attention-needing rows.
- **Sync-ADCSTemplate.ps1 → 1.0.3** — `-EnrollPrincipals` documents the UPN-only resolution of `user@domain` keys, `-UpgradeCompatibility` the legacy-provider rule, `-Mode Validate` the read-back failure.
- **Set-ADCSTemplateValidity.ps1 → 1.0.3** — `-OverlapPeriod` documents the retained-overlap refusal.
- **Add-CertificateEnrollmentPolicyServerToGpo.ps1 → 1.0.3** and **…Offline.ps1 → 1.0.3** — `-ReplaceExisting` documents the complete-row requirement.

| Script | Version |
|---|---|
| Set-ADCSTemplateValidity.ps1 | **1.0.3** |
| Submit-CertificateRequests.ps1 | **1.0.5** |
| Sync-ADCSTemplate.ps1 | **1.0.3** |
| Add-CertificateEnrollmentPolicyServerOffline.ps1 | **1.0.3** |
| Add-CertificateEnrollmentPolicyServerToGpo.ps1 | **1.0.3** |

## [1.0.4] — 2026-09-06

Findings of a full-repository adversarial review (Codex, gpt-6-astra, high effort) after v1.0.3, each verified against the code before fixing.

### Fixed
- **Submit-CertificateRequests.ps1 → 1.0.4**
  - Certificate file names are allocated **before** anything is submitted and must be unique within the batch and against the destinations the tracking file already records for *other* request files. The old rule (files sharing a base name keep their full name) still mapped two requests onto one `.cer`: `prod.req.txt` has the base name `prod.req`, exactly what `prod.req` fell back to, so the later request's certificate replaced the earlier one's while both rows said `Issued`. A base name already recorded for a different request file is not reused either; an unresolvable clash aborts the run.
  - The tracking file is read with `-LiteralPath`. Its name is canonicalized and locked literally, but the read used a wildcard-aware `Test-Path`, so `tracking[1].csv` was reported absent, its history came back empty, every request was resubmitted and the original file was replaced.
  - Every row now records the CA it was submitted to (`CAConfig` column), and `-Mode Retrieve` refuses rows submitted to a different CA: RequestIDs are per CA, so during a migration CA-B's request 42 would have been delivered under CA-A's row as `Issued`. Rows from older files record no CA; they are assumed to belong to the run's `-CAConfig` and stamped. `Export-Csv` derives its columns from the first object, so a mixed old/new list led by an old row would have dropped the column silently; every row is now projected onto the full schema.
  - A submission that certreq reports as *successful* but whose reply yields neither a RequestID nor a certificate (localized output) is recorded as `Unknown` and counts as submitted: it is never resubmitted automatically (a later Submit asks, or needs `-Force`) and Retrieve lists it for reconciliation. Previously it became an `Error` row with no RequestID, which the next run resubmitted as a duplicate request.
  - A run in which any request failed or needs attention (`Error`, `Denied`, `Undelivered`, `Unknown`, or a Retrieve row skipped as invalid) now ends with a terminating error after the summary and the final checkpoint, in Submit and Retrieve alike. A nonexistent or non-folder `-InputPath` is a terminating error instead of an empty batch (an existing empty folder still is one). Previously a scheduled run in which the CA rejected every CSR, or Retrieve skipped every row, reached "Done" with exit code 0. `Pending` is not a failure.
  - Log messages fold CR/LF and other control characters, so a tracking-file field containing a line break can no longer forge additional, attacker-chosen log records.
- **Sync-ADCSTemplate.ps1 → 1.0.2**
  - A `user@domain` value in `-EnrollPrincipals` resolves **only** as a UPN. `sAMAccountName` may legally contain `@`, and the lookup consulted it first, so an attacker with delegated account-creation rights could plant a principal whose sAMAccountName equals the victim's UPN and capture the grant. A different object carrying the string as its sAMAccountName is now refused as a planted or colliding account.
  - `-UpgradeCompatibility` set `CT_FLAG_USE_LEGACY_PROVIDER` whenever `pKIDefaultCSPs` was populated, switching a schema-3 template that lists a KSP ("Microsoft Software Key Storage Provider") to legacy CryptoAPI key handling. The bit is now set only for a schema-2 source with a provider list (v2 knows only CSPs); a schema-3 source keeps its own bit.
  - `-Mode Validate` ends with a terminating error when the throwaway template cannot be read back after a confirmed create (replication lag on a domain-name `-Server`). Previously a warning, and the run returned normally with exit 0 and nothing validated. The throwaway object is still removed by the cleanup.
- **Add-CertificateEnrollmentPolicyServerToGpo.ps1 → 1.0.2**
  - A `**`-prefixed registry.pol instruction (`**DeleteKeys`, `**DeleteValues`, `**del.`, `**delvals.`) decodes its data as a string whatever the record's type field says, as the Group Policy engine does; `**soft.<name>` keeps its declared type, since it writes a value. A `**DeleteKeys` typed e.g. REG_BINARY previously decoded to no data, deleted nothing in the model and let a deleted AD row satisfy the prerequisite while clients delete the key.
  - A pre-existing AD-policy row or CEP entry counts as complete only with URL, PolicyID, FriendlyName and numeric Flags/AuthFlags/Cost present. URL + PolicyID alone is what an interrupted write leaves behind (the values are written in that order), and it satisfied the prerequisite and the gate that lets the `(Default)` marker be set and `-ReplaceExisting` delete the working siblings.
- **Add-CertificateEnrollmentPolicyServerOffline.ps1 → 1.0.2** — the same completeness rule against the live registry (DWORD kinds checked).
- **Set-ADCSTemplateValidity.ps1 → 1.0.2** — when the overlap is not being set, a template whose *existing* renewal overlap is not shorter than the new validity is reported as an error (counted in the exit code) and left unchanged. Previously a 30-day validity was written beneath a stock 6-week overlap; the begin-block check only covered an explicitly supplied overlap.
- CI: the RSAT/GPMC feature result is checked, both Guard-tier modules are imported in each engine before Pester runs, and any skipped test fails the leg. A Guard context whose module is missing self-skips, and the run result stayed `Passed`.
- Tests: the Lab-tier ACL tests assert the **complete** DACL as a normalized ACE set — exactly one Allow ACE per expected (principal, right, object GUID), no other principal, no inherited or deny entry, and no bit beyond the expected right on any ACE (a generic right expanded by the DS into its specific bits is tolerated; WriteProperty, WriteDacl, an unexpected extended right or Full Control is not) — instead of merely that the expected entries exist. Unit tests added for every fix above: output-name allocation, literal tracking path, mixed-schema export, log folding, KSP/legacy-provider upgrade, REG_BINARY `**DeleteKeys`, row completeness, and the exact-days period decoder.

| Script | Version |
|---|---|
| Set-ADCSTemplateValidity.ps1 | **1.0.2** |
| Submit-CertificateRequests.ps1 | **1.0.4** |
| Sync-ADCSTemplate.ps1 | **1.0.2** |
| Add-CertificateEnrollmentPolicyServerOffline.ps1 | **1.0.2** |
| Add-CertificateEnrollmentPolicyServerToGpo.ps1 | **1.0.2** |

## [1.0.3] — 2026-09-05

### Fixed
- **Submit-CertificateRequests.ps1 → 1.0.3** — the delivery folder is now also judged on what its ACL hands to the *files* created inside it. The 1.0.2 checks judged each folder on the rights it grants on the folder itself and, correctly for folder-swap purposes, ignored inherit-only entries — but an inheritable "files only" entry granting an untrusted principal write, append, delete or re-permission rights is exactly what certreq's staging file and the delivered certificate inherit, so such a user could have altered the certificate's bytes after certreq closed the file, before or after delivery, with every folder check passing (and the post-delivery ACL reset, which re-enables inheritance from the folder, would have kept the grant). Every entry that propagates to files (ObjectInherit, inherit-only or not) is now checked against the file write-class on the delivery folder — write-data, append-data, delete, change permissions, take ownership, and write-attributes (enough on its own to set a reparse point on a file) — and a grant to an untrusted principal is refused — a warning under `-AllowUnprotectedOutputFolder`; `-TrustedOutputPrincipal` applies. The inheritance placeholders are resolved as the file system resolves them: CREATOR OWNER and OWNER RIGHTS become the creating (running, trusted) account, so the inherit-only CREATOR OWNER entry that the `C:\` root propagates to every unprotected folder does not cause a refusal; CREATOR GROUP becomes the running account's primary group and stays untrusted unless `S-1-3-1` is named in `-TrustedOutputPrincipal`. Found by adversarial review after v1.0.2.
- Tests: file-inheritance regression cases for the delivery folder — an inherit-only *files* Modify entry, an object-inheritable append-data entry (tolerated as create-subfolder on the folder, but append on a file), a files-only write-attributes entry, a folders-only entry that files do not inherit, and the `C:\`-style inherit-only CREATOR OWNER Full Control entry that must *not* be refused (with the file it produces proving it resolves to the running account) — with a real file created inside proving what is inherited.

| Script | Version |
|---|---|
| Submit-CertificateRequests.ps1 | **1.0.3** |

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

[1.0.6]: https://github.com/TheOmnilord/ADCS/compare/v1.0.5...v1.0.6
[1.0.5]: https://github.com/TheOmnilord/ADCS/compare/v1.0.4...v1.0.5
[1.0.4]: https://github.com/TheOmnilord/ADCS/compare/v1.0.3...v1.0.4
[1.0.3]: https://github.com/TheOmnilord/ADCS/compare/v1.0.2...v1.0.3
[1.0.2]: https://github.com/TheOmnilord/ADCS/compare/v1.0.1...v1.0.2
[1.0.1]: https://github.com/TheOmnilord/ADCS/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/TheOmnilord/ADCS/releases/tag/v1.0.0