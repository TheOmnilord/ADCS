<a id="top"></a>

# ADCS — Active Directory Certificate Services PowerShell Tools

[![CI](https://github.com/TheOmnilord/ADCS/actions/workflows/ci.yml/badge.svg)](https://github.com/TheOmnilord/ADCS/actions/workflows/ci.yml)
[![Latest release](https://img.shields.io/github/v/release/TheOmnilord/ADCS?sort=semver&label=release)](https://github.com/TheOmnilord/ADCS/releases/latest)

A set of PowerShell scripts for administering **Active Directory Certificate Services** (AD CS / ADCS) from the command line. Built for Windows PKI administrators who need to manage certificate templates, process certificate requests, and configure enrollment clients at scale — without clicking through the Certificate Templates MMC snap-in, the Certification Authority console, or the per-machine "Certificate Enrollment Policy" dialog.

Currently includes:

- **Bulk certificate template validity updates** — useful for rolling out the CA/Browser Forum **SC-081** validity reductions (200 days from March 2026, 100 days from March 2027, 47 days from March 2029) across many templates at once.
- **Batch CSR submission to an Enterprise CA** via `certreq.exe`, with resume-safe CSV tracking of request IDs and automated retrieval of issued certificates.
- **Cross-forest certificate template sync** — copy a template between forests at the directory level (all access over ADWS), either through a JSON export/import or **directly forest-to-forest in one run** (`-Mode Sync`, with optional explicit credentials per side, so no trust is required). Optional rename, controlled OID handling, and a composable enrollment ACL. Works even when the target forest has **no AD CS installed** — e.g. to publish a template that an external CA reads for enrollment authorization.
- **Client enrollment-policy (CEP) configuration** — point Windows enrollment clients at an **EJBCA** (or other MS-XCEP) policy server so they enroll against the synced template, computing every registry value **offline** (no "Validate Server" round-trip). Apply it to a single machine or a local GP hive, or fleet-wide by authoring the setting straight into a domain **GPO** (with optional Auto-Enrollment). This is the client half of the template-sync story above.
- **A ready-to-import template library** ([`Templates/`](./Templates/README.md)) — every certutil default template as JSON (template OIDs stripped, so nothing fingerprints a forest and importers mint their own), plus **EJBCA-ready variants** of the templates whose Subject is empty at creation, and a **latest-compatibility** (Windows Server 2016 / Windows 10) copy of both.

Works on Windows PowerShell 5.1 and PowerShell 7+. `Set-ADCSTemplateValidity` and `Submit-CertificateRequests` have no AD PowerShell module dependency; `Sync-ADCSTemplate` requires the RSAT ActiveDirectory module and `Add-CertificateEnrollmentPolicyServerToGpo` the GroupPolicy module (see each script's requirements); `Add-CertificateEnrollmentPolicyServerOffline` needs no module.

**Versioning.** Each script carries its own version in a `PSScriptInfo` header at the top of the file, bumped only when that script changes; releases are tagged on the repository (badge above) and [CHANGELOG.md](./CHANGELOG.md) lists which script versions each release ships. To check whether a deployed copy is current without opening it:

```powershell
Test-ScriptFileInfo .\Submit-CertificateRequests.ps1 | Select-Object Name, Version
```

## Scripts

Jump to a script — or straight to a section within it. Each script title links to its full documentation; the source file is linked alongside.

- **[Set-ADCSTemplateValidity.ps1](#set-adcstemplatevalidityps1)** &nbsp;·&nbsp; [source](./Set-ADCSTemplateValidity.ps1)
  Bulk-update the validity period (and optionally the renewal overlap period) on one or more certificate templates, with wildcard name matching.
  <br>↳ [Why you need this](#why-you-need-this) · [Features](#features) · [Requirements](#requirements) · [Parameters](#parameters) · [Usage](#usage) · [Output](#output) · [Notes](#notes) · [How It Works](#how-it-works)
- **[Submit-CertificateRequests.ps1](#submit-certificaterequestsps1)** &nbsp;·&nbsp; [source](./Submit-CertificateRequests.ps1)
  Batch-submit `.req`/`.csr`/`.txt` files to an ADCS CA via `certreq.exe`, track request IDs in a CSV, and later retrieve the issued certificates.
  <br>↳ [Features](#features-1) · [Requirements](#requirements-1) · [Parameters](#parameters-1) · [Usage](#usage-1) · [Friendly Error Hints](#friendly-error-hints) · [Run Summary](#run-summary) · [Tracking CSV Schema](#tracking-csv-schema) · [Notes](#notes-1)
- **[Sync-ADCSTemplate.ps1](#sync-adcstemplateps1)** &nbsp;·&nbsp; [source](./Sync-ADCSTemplate.ps1)
  Copy the Kerberos Authentication (or any other) certificate template between forests — through a JSON file or directly forest-to-forest in one run — with optional rename, four OID-handling modes, per-side credentials, a composable enrollment ACL, and a round-trip validation mode. Target forest does not need AD CS.
  <br>↳ [Why you need this](#why-you-need-this-1) · [Features](#features-2) · [Requirements](#requirements-2) · [Parameters](#parameters-2) · [Usage](#usage-2) · [Using with EJBCA](#using-the-template-with-ejbca) · [Notes](#notes-2) · [Tests](#tests)
- **[Add-CertificateEnrollmentPolicyServerOffline.ps1](#add-certificateenrollmentpolicyserverofflineps1)** &nbsp;·&nbsp; [source](./Add-CertificateEnrollmentPolicyServerOffline.ps1)
  Register (or remove) an EJBCA/MSAE enrollment-policy server on one machine — in the user-configured store or a Group Policy hive — writing the exact registry values the CEP dialog produces, computed entirely offline.
  <br>↳ [Why you need this](#why-you-need-this-2) · [Features](#features-3) · [Requirements](#requirements-3) · [Parameters](#parameters-3) · [Usage](#usage-3) · [Notes](#notes-3) · [Tests](#tests-1)
- **[Add-CertificateEnrollmentPolicyServerToGpo.ps1](#add-certificateenrollmentpolicyservertogpops1)** &nbsp;·&nbsp; [source](./Add-CertificateEnrollmentPolicyServerToGpo.ps1)
  Author (or remove) the same enrollment-policy setting directly in a domain **GPO** via `Set-GPRegistryValue`, for fleet-wide rollout — with optional Auto-Enrollment, AD-policy-row preservation, and Registry.pol safety checks.
  <br>↳ [Why you need this](#why-you-need-this-3) · [Features](#features-4) · [Requirements](#requirements-4) · [Parameters](#parameters-4) · [Usage](#usage-4) · [Notes](#notes-4) · [Tests](#tests-2)

Also in the repo: the **[Template library](#template-library)** ([`Templates/`](./Templates/README.md)) — importable JSON exports of all default templates, plus EJBCA-ready variants.

---

## Set-ADCSTemplateValidity.ps1

Modifies the `pKIExpirationPeriod` (and optionally `pKIOverlapPeriod`) attribute on ADCS certificate templates in Active Directory. Supports wildcard template name matching so you can update many templates in one go.

### Why you need this

The CA/Browser Forum (ballot **SC-081**, passed April 2025) mandates a phased reduction of the maximum validity period for publicly-trusted TLS server certificates. While ADCS is typically used for internal PKI, many organizations mirror these limits on their internal CAs to keep templates aligned with industry best practice (and to be ready if any templates ever feed into publicly-trusted chains).

| Effective date | Maximum validity (TLS server certs) |
| --- | --- |
| **15 March 2026** (current) | **200 days** |
| **15 March 2027** | **100 days** |
| **15 March 2029** | **47 days** |

As the cadence tightens, manually adjusting every template through the Certificate Templates MMC snap-in becomes painful. This script lets you update dozens of templates in seconds:

```powershell
# March 2026 rollover: drop TLS templates to 200 days
.\Set-ADCSTemplateValidity.ps1 -TemplateName "*Web*","*TLS*" -ValidityPeriod 200 -ValidityPeriodUnit Days -WhatIf
```

Client authentication, code signing, S/MIME, and other non-TLS templates are **not** covered by SC-081 and can keep longer validity periods. Use targeted wildcards to avoid changing those.

#### The Apple/Safari ceiling: 825 days

Even if you decide to ignore SC-081 on your internal CA, there is a hard cap you cannot ignore: **Safari on macOS and iOS rejects any TLS server certificate with a validity longer than 825 days** (~2 years 3 months), regardless of whether the issuing CA is publicly trusted or a user-/admin-added internal root. Apple's `trustd` daemon has enforced this since iOS 13 / macOS 10.15 (July 2019).

The failure mode is hostile to troubleshoot: Safari shows a generic *"cannot establish a secure connection"* error with no override option, while the same certificate works fine in Chrome and Firefox on the same machine. So even if your internal policy allows 5- or 10-year templates, anything over 825 days will silently break for Apple users.

Sources: [michalspacek.com](https://www.michalspacek.com/validity-period-of-https-certificates-issued-from-a-user-added-ca-is-essentially-2-years), [certkit.io](https://www.certkit.io/blog/apple-doesnt-care-who-signed-your-certificate).

### Features

- **Wildcard matching** on template CN (e.g. `User*`, `*Web*`, `*VPN*`)
- **Human-readable durations** (`Years`, `Months`, `Weeks`, `Days`, `Hours`)
- **`-WhatIf` / `-Confirm`** support with `ConfirmImpact = 'High'`
- **No AD PowerShell module required** - uses `System.DirectoryServices` directly
- **Skips templates already set** to the requested value (byte-array compare)
- **Auto-increments** `msPKI-Template-Minor-Revision` so CAs detect the change
- **Deduplication** when multiple patterns match the same template
- **Summary output** with counts of Modified / Already set / Skipped / Errors
- **Pipeline-friendly output** as `PSCustomObject` per template

### Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- A domain-joined machine (or use `-Server` to target a specific DC)
- Permissions to modify certificate templates (typically Enterprise Admin or delegated rights on the `CN=Certificate Templates` container in the Configuration naming context)

### Parameters

| Parameter | Type | Required | Description |
| --- | --- | --- | --- |
| `-TemplateName` | `string[]` | Yes | One or more template CN names. Supports the LDAP wildcard `*` (LDAP has no single-character wildcard; a `?` matches literally). |
| `-ValidityPeriod` | `int` (1-9999) | Yes | Numeric value for the new validity period. |
| `-ValidityPeriodUnit` | `Years` / `Months` / `Weeks` / `Days` / `Hours` | Yes | Unit for `-ValidityPeriod`. AD uses 365 days/year and 30 days/month. |
| `-OverlapPeriod` | `int` (1-9999) | No | Numeric value for the renewal overlap period. |
| `-OverlapPeriodUnit` | `Years` / `Months` / `Weeks` / `Days` / `Hours` | No | Unit for `-OverlapPeriod`. Required if `-OverlapPeriod` is set. |
| `-Server` | `string` | No | Target a specific **domain controller** (not a CA server) for the LDAP connection, e.g. `dc01.domain.com`. |
| `-WhatIf` | switch | No | Preview changes without making them. |
| `-Confirm` | switch | No | Prompt before each change. |

### Usage

**Preview which templates would be changed (WhatIf):**
```powershell
.\Set-ADCSTemplateValidity.ps1 -TemplateName "Web*" -ValidityPeriod 2 -ValidityPeriodUnit Years -WhatIf
```

**Set validity and overlap on multiple wildcard patterns:**
```powershell
.\Set-ADCSTemplateValidity.ps1 `
    -TemplateName "User*","Computer*" `
    -ValidityPeriod 1 -ValidityPeriodUnit Years `
    -OverlapPeriod 6 -OverlapPeriodUnit Weeks
```

**Set validity on a single exact template, target a specific DC, skip confirmation:**
```powershell
.\Set-ADCSTemplateValidity.ps1 `
    -TemplateName "WebServer" `
    -ValidityPeriod 365 -ValidityPeriodUnit Days `
    -Server dc01.domain.com `
    -Confirm:$false
```

**Preview all templates matching a pattern and capture the output:**
```powershell
$report = .\Set-ADCSTemplateValidity.ps1 -TemplateName "*" -ValidityPeriod 1 -ValidityPeriodUnit Years -WhatIf
$report | Format-Table -AutoSize
```

### Output

Each matched template produces a `PSCustomObject` with:

| Property | Description |
| --- | --- |
| `TemplateName` | Template CN |
| `DisplayName` | Template display name |
| `PreviousValidity` | Current validity period (human-readable) |
| `NewValidity` | Requested new validity period |
| `PreviousOverlap` | Current overlap period (human-readable) |
| `NewOverlap` | Requested new overlap period, or `(unchanged)` |
| `Status` | `Modified`, `Already set`, `Skipped`, or `Error: <message>` |

A color-coded summary is printed at the end:

```
--- Summary ---
  Total matched : 12
  Modified      : 7
  Already set   : 3
  Skipped       : 2
  Errors        : 0
  Run 'certutil -pulse' on CA server(s) to refresh.
```

### Notes

- After modifying templates, run `certutil -pulse` on each CA server for changes to be picked up immediately. Otherwise AD replication + CA cache refresh will eventually apply them.
- The `-Server` parameter is for a **domain controller**, not the CA server. Certificate templates are AD objects stored in the forest's Configuration naming context.
- Changes replicate forest-wide from the Configuration NC. Allow normal AD replication time.
- The script increments `msPKI-Template-Minor-Revision` on each change so issuing CAs detect the update.
- `-WhatIf` templates show as `Skipped` in the output (they would have been modified but weren't due to the WhatIf flag).
- Failures are reported per template as they happen and the remaining templates are still processed — but a run in which anything failed ends with a terminating error (non-zero exit) after the summary, so automation cannot mistake a partial update for success.
- When `-OverlapPeriod` is not given, each template keeps its existing renewal overlap — and if that overlap is not shorter than the new validity (a 30-day validity over a stock 6-week overlap), the template is reported as an error and left unchanged; pass a shorter `-OverlapPeriod` / `-OverlapPeriodUnit` to change both together.

### How It Works

1. Connects to `RootDSE` to resolve the Configuration naming context.
2. Searches `CN=Certificate Templates,CN=Public Key Services,CN=Services,<ConfigNC>` for templates matching the pattern(s) using an LDAP filter.
3. For each match, decodes the current `pKIExpirationPeriod` / `pKIOverlapPeriod` (8-byte little-endian negative FILETIME ticks).
4. Compares against the requested value; skips if already equal.
5. Assigns the new byte array through the property cache (`DirectoryEntry.Properties['pKIExpirationPeriod'].Value`) and calls `SetInfo()` to commit. (`InvokeSet()` is deliberately *not* used: its `params object[]` binding unrolls the 8-byte array into eight separate values.)

<sub>[↑ Back to top](#top)</sub>

---

## Submit-CertificateRequests.ps1

Batch-submits certificate signing requests (`.req` / `.csr` / `.txt`) from a folder to an ADCS CA using `certreq.exe`, tracks each submission's request ID in a CSV file, and can later retrieve the issued certificates.

### Features

- **Batch submit** all request files in a folder in one run
- **CSV tracking file** records request ID, submit time, status, error messages per file
- **Resume-safe** - files already present in the tracking CSV are skipped on re-run; with `-Force` (or an interactive y/n confirmation) you can resubmit a tracked file as a new request
- **Retrieve mode** picks up any tracked request that isn't yet finally resolved (`Pending`, `Unknown`, or `Error`) and pulls issued `.cer` files; only needs `-CAConfig` and a tracking file, not `-InputPath` / `-CertificateTemplate` A retrieval whose destination already holds the certificate of a *different* request for the same CSR (a `-Force` resubmission that was issued first) is refused and reported, so an older request never replaces a newer certificate.
- **Both mode** submits then retrieves in a single invocation
- **Connectivity pre-check** via `certutil -ping` before any submissions
- **Per-run timestamped log file** (`CertBatch_yyyyMMdd_HHmmss.log`)
- **Friendly error hints** for common ADCS failures (unsupported template, denied by policy, bad subject, access denied) — instead of just dumping the raw certreq output
- **Helpful input-folder diagnostics** — if no `.req`/`.csr`/`.txt` files are found, the script lists what *is* in the folder (or notes that it is empty) and skips Submit cleanly instead of throwing a PowerShell stack trace
- **Dual-section summary** — separates *this run's* results from the *cumulative tracking-file totals*, so historical errors don't look like new ones
- **Automatic `.rsp` cleanup** after retrieval (override with `-KeepRspFile`)
- **`-WhatIf` / `-Confirm`** support
- Handles empty files, missing request IDs, denied requests gracefully

### Requirements

- Windows with `certreq.exe` and `certutil.exe` available (standard on Windows)
- Windows PowerShell 5.1 or PowerShell 7+
- Permissions to submit to the target CA and template
- Network connectivity to the CA

### Parameters

| Parameter | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| `-InputPath` | `string` | Submit/Both only | | Folder containing `.req` / `.csr` / `.txt` request files. Not used, and not required, for `-Mode Retrieve`. |
| `-CAConfig` | `string` | Yes | | CA configuration string, e.g. `CA01.domain.com\Contoso Issuing CA 1`. Always required. |
| `-CertificateTemplate` | `string` | Submit/Both only | | Certificate template name (the CN, not the display name). Not used, and not required, for `-Mode Retrieve`. |
| `-TrackingFile` | `string` | No | `.\CertTracking.csv` | CSV file used to track request IDs and statuses across runs. |
| `-OutputFolder` | `string` | No | `.\Certificates` | Folder where issued `.cer` files are saved (one per request, named after the request file). In `Retrieve` mode each row is written to the path recorded at submit time unless `-OutputFolder` is passed explicitly, which redirects the retrieved files (and updates the tracking row). |
| `-Mode` | `Submit` / `Retrieve` / `Both` | No | `Submit` | `Submit` = submit new requests only; `Retrieve` = pull certs for previously-pending requests; `Both` = do both. |
| `-KeepRspFile` | switch | No | | By default the `.rsp` file `certreq` writes next to each retrieved `.cer` is deleted. Specify this switch to leave it in place. |
| `-Force` | switch | No | | Resubmit request files that already have a tracked RequestID without prompting. Without `-Force`, the script asks y/n for each already-submitted file (default = No / skip). |
| `-AllowUnprotectedOutputFolder` | switch | No | | By default the script refuses to deliver into or through a folder that an untrusted principal owns, or can delete, rename or write to (such a user could swap it for a junction during a delivery, or turn an empty folder into one with mere write rights), and refuses a delivery folder whose ACL would hand an untrusted principal write/append/delete/write-attributes rights on the *files* created inside it (the staging file and the delivered certificate inherit such an entry; CREATOR OWNER entries resolve to the running account and are fine). Pass this to accept the risk; the conditions are then only warned about. |
| `-TrustedOutputPrincipal` | `string[]` | No | | Additional principals (SIDs or `DOMAIN\Group` names) that may own, or hold delete/rename/write rights on, the delivery folders (including file-inheritable write rights in the delivery folder itself), on top of SYSTEM, Administrators, TrustedInstaller, the running account and its Domain/Enterprise Admins. |
| `-WhatIf` | switch | No | | Preview without submitting/retrieving. |
| `-Confirm` | switch | No | | Prompt before each action. |

### Usage

**Submit all CSRs in a folder:**
```powershell
.\Submit-CertificateRequests.ps1 `
    -InputPath "C:\CSRs" `
    -CAConfig "CA01.domain.com\Contoso Issuing CA 1" `
    -CertificateTemplate "WebServer" `
    -Mode Submit
```

**Retrieve any issued certificates for previously-unresolved requests** (only `-CAConfig` and the tracking file are needed — `-InputPath` / `-CertificateTemplate` aren't used in this mode):
```powershell
.\Submit-CertificateRequests.ps1 `
    -CAConfig "CA01.domain.com\Contoso Issuing CA 1" `
    -Mode Retrieve
```

Retrieved `.cer` files land at the path recorded for each row at submit time. To put them somewhere else, pass `-OutputFolder` explicitly on the `Retrieve` run:
```powershell
.\Submit-CertificateRequests.ps1 `
    -CAConfig "CA01.domain.com\Contoso Issuing CA 1" `
    -Mode Retrieve `
    -OutputFolder "C:\Certificates\Issued"
```

**Submit and retrieve in one go:**
```powershell
.\Submit-CertificateRequests.ps1 `
    -InputPath "C:\CSRs" `
    -CAConfig "CA01.domain.com\Contoso Issuing CA 1" `
    -CertificateTemplate "WebServer" `
    -Mode Both
```

**Preview what would be submitted:**
```powershell
.\Submit-CertificateRequests.ps1 `
    -InputPath "C:\CSRs" `
    -CAConfig "CA01.domain.com\Contoso Issuing CA 1" `
    -CertificateTemplate "WebServer" `
    -Mode Submit -WhatIf
```

**Resubmit already-tracked files (non-interactive) and keep `.rsp` files after retrieval:**
```powershell
.\Submit-CertificateRequests.ps1 `
    -InputPath "C:\CSRs" `
    -CAConfig "CA01.domain.com\Contoso Issuing CA 1" `
    -CertificateTemplate "WebServer" `
    -Mode Both -Force -KeepRspFile
```

### Friendly Error Hints

When `certreq` rejects a request, the raw output is dumped as a warning *and* the script appends an actionable hint based on the underlying error code. Patterns currently recognized:

| Detected pattern | Hint |
| --- | --- |
| `0x80094800` / `CERTSRV_E_UNSUPPORTED_CERT_TYPE` | Template name misspelled, not published on this CA, or you accidentally included the `CertificateTemplate:` prefix. Suggests `certutil -config "<CA>" -CATemplates` to list valid names. |
| `0x80094012` / `CERTSRV_E_TEMPLATE_DENIED` | Calling account lacks Enroll permission on the template, or the template requires approval/signature. |
| `0x80094004` / `CERTSRV_E_BAD_REQUESTSUBJECT` | CSR subject/SAN does not match what the template requires. |
| `0x80070005` / access denied | Calling account lacks "Request Certificates" rights on the CA. |

### Run Summary

The end-of-run summary is split into two sections so you can tell *this run's* results apart from the *cumulative state* of the tracking file:

```
--- Summary: Retrieve ---
  Issued: 7
  Total processed this run: 7
--- Tracking file total (all history) ---
  Error: 1
  Issued: 27
Tracking file: .\CertTracking.csv
```

In this example the lone `Error: 1` is a historical row from an earlier session, not something that happened in the current run.

### Tracking CSV Schema

| Column | Description |
| --- | --- |
| `RequestFile` | Full path of the source `.req`/`.csr`/`.txt` file |
| `RequestID` | Numeric request ID assigned by the CA |
| `SubmitTime` | ISO-8601 submission timestamp |
| `Status` | `Issued`, `Pending`, `Denied`, `Error`, `Unknown`, or `Undelivered` (the CA issued, but the `.cer` could not be delivered to its destination; counts as submitted, re-fetched by `Retrieve` when a RequestID is present). An `Unknown` row **without** a RequestID (certreq reported success but neither a RequestID nor a certificate could be read) also counts as submitted and is listed by `Retrieve` for manual reconciliation. |
| `OutputCertFile` | Full path where the issued `.cer` is saved |
| `LastCheckTime` | ISO-8601 timestamp of last status check |
| `ErrorMessage` | Error output from `certreq` if the submission or retrieval failed |
| `CAConfig` | The CA the request was submitted to. `-Mode Retrieve` refuses rows submitted to a different CA (RequestIDs are per CA); rows from older files are assumed to belong to the run's `-CAConfig` and stamped. |

### Notes

- The tracking CSV is the source of truth for resume behavior. Deleting it will cause the script to re-submit all files (and the CA may issue duplicates).
- One run per tracking file: the script holds an exclusive lock file (`<TrackingFile>.lock`, removed when the run ends) for its whole run and refuses to start while another run holds it — on this or any other machine, via any alias of the path — because two concurrent writers would silently drop each other's rows. The tracking file's name is canonicalized first (an 8.3 short name resolves to the long name, so both spellings share one lock), and a hard-linked or symlinked tracking file is refused. A `-WhatIf` run takes no lock.
- certreq always writes into a private, randomly named staging file **inside the destination folder** (never %TEMP%, whose inherited file ACL could let another account tamper with the certificate and would travel onto the delivered file); the destination is touched only after a successful write, so a pending, denied or failed request never disturbs it, and the delivered file is reset to inherit the folder's ACL. A row's `OutputCertFile` is an identifier inside an operator-chosen boundary, not an authority: it must be a rooted `.cer` path in an existing folder beneath the tracking file's folder or the run's `-OutputFolder` (canonically, with no junction/symlink in between, re-checked right before delivery), and `RequestID` must be numeric; rows failing this are skipped with an error and never reach certreq. The destination itself must be a plain file or absent (a folder or link named `x.cer` is refused, and delivery verifies that a plain file resulted). If an issued certificate cannot be delivered (locked destination, denied rename), the row becomes `Undelivered` with its RequestID kept and the certificate left in its staging file beside the destination (path in `ErrorMessage`): it counts as submitted on later runs (never resubmitted automatically) and `Retrieve` re-fetches it; an `Undelivered` row with no RequestID is reported for manual reconciliation.
- Output folders must not be swappable by untrusted users. Every folder from the destination up to the volume/share root is checked, and the script **refuses** to deliver when any of them is a reparse point (junction/symlink/mount point), is owned by an untrusted principal, can be deleted, renamed or written to by one (delete/rename lets such a user replace it with a junction between the path check and the privileged delivery; write-data or write-attributes access is all `FSCTL_SET_REPARSE_POINT` needs, so "create files" rights let them turn an *empty* folder, such as a freshly created output folder, into a junction in place), or has a security descriptor that cannot be read. Trusted principals are SYSTEM, `BUILTIN\Administrators`, TrustedInstaller, the running account, its Domain Admins / Enterprise Admins, and anything named in `-TrustedOutputPrincipal`; `-AllowUnprotectedOutputFolder` downgrades the refusals to warnings. Only the create-*subfolder* right by itself (what the `C:\` root grants Users on itself) and inherit-only ACEs are tolerated on the folder chain: neither can set a reparse point on the folder, a pre-planted junction is refused by the reparse-point checks, an attacker-created subfolder is refused by the owner check, and a folder or junction planted under the exact destination name between check and delivery receives no file, because every delivery is a no-overwrite rename (`File.Move`) that fails when anything occupies the name instead of moving the file into it. The `.rsp` written with `-KeepRspFile` and the tracking-file replacement follow the same rule. The delivery folder itself is additionally judged on what its ACL hands to the *files* created inside it: every entry that propagates to files (ObjectInherit, inherit-only or not) is checked against the file write-class (write-data, append-data, delete, change permissions, take ownership, and write-attributes, which alone suffices to set a reparse point on a file), and a grant to an untrusted principal is refused, because certreq's staging file and the delivered certificate inherit it and such a user could otherwise alter the certificate's bytes, or turn the delivered file into a reparse point, before or after delivery with every folder check passing. The inheritance placeholders are resolved the way the file system resolves them: CREATOR OWNER and OWNER RIGHTS become the creating account (the trusted running account, since certreq creates the file), so the inherit-only CREATOR OWNER entry the `C:\` root hands to every unprotected folder is fine; CREATOR GROUP becomes the running account's primary group, which can be as broad as Domain Users, and stays untrusted unless `S-1-3-1` is named in `-TrustedOutputPrincipal`.
- `Issued` is decided by certreq's exit code plus the presence of the certificate it wrote (into a fresh temp file, so it is unambiguously this run's output; language-independent). On success the certificate is delivered to the destination; a file already there (a `-Force` resubmit, or a retry of an unresolved row) is moved aside as `<name>.superseded-<UTC stamp>.cer` — never deleted — and that copy is removed again only if the fresh certificate is byte-identical. A retrieval that reports Issued without producing the file is recorded as `Error` and retried on the next `Retrieve`. The RequestID and the `Pending`/`Denied` dispositions are parsed from certreq's console text, which Windows localizes; on a non-English system they may parse as `Unknown`, and a missing RequestID is flagged in `ErrorMessage` so it can be filled in from the CA database.
- Certificate file names are allocated before anything is submitted and must be unique within the batch and against the destinations already recorded for other request files (`prod.req`, `prod.csr` and `prod.req.txt` no longer map two requests onto one `.cer`); a clash aborts the run. A submission certreq reports as successful but whose reply yields neither a RequestID nor a certificate is recorded as `Unknown`, counts as submitted, and is never resubmitted automatically.
- A run in which any request failed or needs attention (`Error`, `Denied`, `Undelivered`, `Unknown`, or a Retrieve row skipped as invalid) ends with a terminating error after the summary, so automation gating on the exit code does not treat a partial batch as success. `Pending` is not a failure.
- Values that reach the certreq command line (`-CAConfig`, `-CertificateTemplate`, the `.cer` paths and `RequestID`) must not contain double quotes or control characters; such values are rejected before certreq runs.
- `Pending` typically means the CA requires manager approval. Re-run in `Retrieve` mode after approval to pull the issued cert.
- Issued `.cer` files are named after the source request file (e.g. `server1.req` -> `server1.cer`).
- Empty request files are skipped with a warning.
- A timestamped log file is created in the working directory for each run.

<sub>[↑ Back to top](#top)</sub>

---

## Sync-ADCSTemplate.ps1

*(Formerly `Sync-KerberosAuthTemplate.ps1` — renamed because it copies **any** certificate template; the Kerberos Authentication template merely remains its default.)*

Copies a certificate template (by default the built-in **Kerberos Authentication** template) between AD forests. It reads the template's functional attributes from the source forest and recreates the template in the target forest with `New-ADObject` (all directory access over ADWS) — deriving the container DN, `objectCategory`, and (optionally) a fresh template OID from the **target** forest, then applying a composable enrollment ACL. Two interchangeable flows share the same pipeline:

- **File-based** (`-Mode Export` / `-Mode Import`) — serialize to a JSON file in the source forest, import it in the target forest. Right for air-gapped or change-controlled environments.
- **Direct** (`-Mode Sync`) — read from a DC in the source forest (`-SourceServer`) and write to the target side in a single run, no intermediate file. With a trust, the current identity can usually do both sides; without one, pass `-SourceCredential` and/or `-Credential` — explicit credentials against explicitly named DCs need no trust at all.

The flows mix freely: a JSON exported earlier imports into a remote forest with `-Mode Import -Server <target DC> -Credential (...)`, and `-Mode Export` equally accepts `-Server`/`-Credential` to read from a remote source forest.

### Why you need this

There is no supported UI path for moving a template definition between forests, and the classic `certutil -dsTemplate` / `-dsAddTemplate` round-trip cannot rename the template, always carries the source OID, and leaves the ACL to you. This script does the whole job, including two scenarios the certutil approach cannot handle:

- **Target forest without AD CS.** The template objects live in the (CA-independent) `Certificate Templates` container that every forest has, so you can publish a template into a forest that never had AD CS — e.g. so an external CA can read the template and its ACL as the enrollment-authorization source. (Using this with **EJBCA** has a few specifics of its own — see [Using with EJBCA](#using-the-template-with-ejbca).)
- **Renamed copies with controlled identity.** Import under a new cn/display name, and choose whether the OID is carried over or freshly minted.

The ACL is the part an external CA actually consumes, so it gets first-class treatment: bases (`-AclBase`) that either replace or extend the schema-default DACL, plus per-principal additions (`-EnrollPrincipals`) for enrollees and template admins. Principal resolution is fail-closed: a name matching BOTH a well-known token and a directory object with a *different* SID, or matching more than one object, is refused rather than guessed (disambiguate with a SID or `DOMAIN\` prefix); built-in groups' own names like `Domain Admins` resolve normally, since both readings give the same SID. Well-known SID/RID tokens are language-invariant, so the script also works on **non-English forests** where group names are localized.

### Features

- **Direct attribute copy** — no `certutil`, no text-dump parsing; JSON file or direct forest-to-forest, everything over ADWS (TCP 9389) only
- **`-Mode Sync`**: one-run direct sync between forests, no intermediate file — with a same-forest guard so a forgotten `-Server` cannot silently write the copy back into the source forest
- **Per-side credentials** (`-Credential` for the target, `-SourceCredential` for the source) — works across a trust with the current identity, or with **no trust at all** using explicit credentials
- **Works against a target forest with no AD CS** (default OID mode needs no PKI OID root)
- **Rename on import** (`-NewTemplateName` / `-NewDisplayName`)
- **Four OID modes** (`-OidHandling`): `Preserve` (default; carry source OID), `Generate` (mint under the target forest's real OID root), `GenerateFromRoot` (mint under a base you supply), `GenerateRandom` (mint under a synthesized base) — every mode registers the companion OID "display" object so Windows resolves the OID to the template name
- **Composable ACL**: `-AclBase Standard | Schema | SchemaPlusStandard | PrincipalsOnly` plus additive `-EnrollPrincipals` (principal → Read/Write/Enroll/Autoenroll/FullControl) — a `user@domain` key resolves **only** as a UPN, and a different object carrying that string as its sAMAccountName is refused (sAMAccountName may contain `@`, so a planted account could otherwise capture the grant)
- **`-UpgradeCompatibility`**: raise the imported copy to the newest compatibility (CA *Windows Server 2016*, recipient *Windows 10 / Windows Server 2016*) as it is created — schema v2/v3 → v4 with the matching private-key-flag bits (v1 built-ins and already-v4 templates are left as-is). The source is never modified; only the target copy is upgraded. The legacy-provider bit (`CT_FLAG_USE_LEGACY_PROVIDER`) is set only for a schema-2 source with a provider list; a schema-3 source keeps its own bit, so a v3 template that lists a KSP stays CNG.
- **`-Mode Validate`**: proves round-trip fidelity of **both** pipelines — the JSON file flow *and* the direct in-memory flow Sync uses — by importing a throwaway copy per pipeline and diffing every PKI attribute of the source (byte-array attributes included), then cleaning up. A mismatch fails the run with a non-zero exit code, so Validate can gate automation A throwaway that cannot be read back after its confirmed creation fails the run as well (terminating error), so an exit-code gate never accepts an unvalidated copy.
- **Fail-fast pre-flights**: refuses duplicate cn, duplicate template OID, missing containers, unresolvable principals — all before anything is created; every known attribute of an import is validated for type, shape and range (a malformed or tampered export is **refused**, never silently dropped or coerced), and an issuance policy OID that the *target* forest links to a group via AMA refuses the import unless `-AllowLinkedIssuancePolicy` is given
- **`-WhatIf` / `-Confirm`** support end to end; UTF-8 BOM file format safe across PowerShell 5.1 ↔ 7

### Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- **RSAT ActiveDirectory PowerShell module** (this script is the exception to the repo's otherwise module-free approach — the module is what makes typed attribute writes, `-Server` pinning, and clean rollback practical)
- No CA role, no RSAT AD CS Tools, and no reachable CA in either forest
- **Export** (and the read side of Sync): read access to the template (Authenticated Users has this by default)
- **Import/Validate** (and the write side of Sync): Enterprise Admin, or delegated write access to the `CN=Certificate Templates` **and** `CN=OID` containers in the Configuration naming context (every OID mode - including the default `Preserve` - registers a companion OID display object when the carried OID has none yet; only a forest without the `CN=OID` container skips it)
- **Sync**: ADWS (TCP 9389) reachability to a DC in *each* forest from the machine it runs on; a trust between the forests **or** explicit `-SourceCredential`/`-Credential`. All modes use ADWS only — no LDAP (389) access is needed. `-Server`/`-SourceServer` values are always used **verbatim** — the script never substitutes an endpoint you did not type. Name **one DC**: a domain name (DNS or NetBIOS) locates a different DC per connection and draws a loud warning, since the create/read-back/ACL steps could hit different replicas (a lagging read-back fails the run rather than leaving a template mis-secured)

### Parameters

| Parameter | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| `-Mode` | `Export` / `Import` / `Sync` / `Validate` | Yes | | `Export` writes the JSON (source forest); `Import` recreates the template + ACL (target forest); `Sync` does both directly forest-to-forest with no file; `Validate` round-trips a template into a throwaway copy and diffs it. |
| `-Path` | `string` | Export/Import | | JSON file to write (Export) or read (Import). Optional for Validate (temp file used); not used by Sync. |
| `-TemplateName` | `string` | No | `KerberosAuthentication` | Export/Validate/Sync: the cn (internal name) of the source template. |
| `-NewTemplateName` | `string` | No | source's `name` | Import/Sync: new cn in the target forest (letters incl. non-ASCII, digits, non-edge spaces and `._-()` allowed; DN metacharacters `, + = " \ ; < >`, the LDAP wildcard `*`, and `/ #` rejected). |
| `-NewDisplayName` | `string` | No | source's `displayName` | Import/Sync: new display name in the target forest. |
| `-StripIdentity` | switch | No | | Export: omit `name`/`displayName` from the file, forcing explicit naming on import. |
| `-StripOid` | switch | No | | Export: omit the source OID (then import needs a Generate mode). |
| `-OidHandling` | `Preserve` / `Generate` / `GenerateFromRoot` / `GenerateRandom` | No | `Preserve` | Import/Sync: how the template OID is chosen (see Features). |
| `-OidRoot` | `string` | With `GenerateFromRoot` | | Base OID to mint under, e.g. `1.3.6.1.4.1.311.21.8.<arcs>`. Rejected with any other `-OidHandling` (never silently ignored). |
| `-AclBase` | `Standard` / `Schema` / `SchemaPlusStandard` / `PrincipalsOnly` | No | `Standard` | ACL foundation. `Standard` writes the stock Kerberos Authentication ACL, replacing the schema default; `Schema` keeps the schema default; `SchemaPlusStandard` keeps it and adds the standard set; `PrincipalsOnly` writes exactly `-EnrollPrincipals`. The default (`Standard`) is Kerberos-Authentication-specific and DC-oriented — the script warns when it applies only by default to a template not named like a Kerberos Authentication copy. |
| `-EnrollPrincipals` | `hashtable` | With `PrincipalsOnly` | | Principal → rights map added on top of the base, e.g. `@{ 'DomainControllers'='Enroll','Autoenroll'; 'PKI-Admins'='FullControl' }`. Keys: SID, sAMAccountName, UPN (user@domain), or well-known token; a bare string matching both a token and an object with a different SID, or more than one object, is refused (disambiguate with a SID or DOMAIN\ prefix). A `user@domain` key resolves only as a UPN; a different object carrying that string as its sAMAccountName is refused. |
| `-SkipAcl` | switch | No | | Import/Sync: skip the permission step entirely. Mutually exclusive with `-EnrollPrincipals` and an explicit `-AclBase`. |
| `-UpgradeCompatibility` | switch | No | | Import/Sync: raise the created copy to the latest compatibility (CA Windows Server 2016 / recipient Windows 10 / Windows Server 2016; schema v2/v3 → v4). v1 templates and already-v4 templates are imported unchanged (with a note). The legacy-provider bit is set only for a schema-2 source with a provider list; a schema-3 source keeps its own. |
| `-AllowLinkedIssuancePolicy` | switch | No | | Import/Sync: accept an issuance policy OID (in `msPKI-Certificate-Policy`, stamped into the issued certificate) that the *target* forest already links to a group via Authentication Mechanism Assurance. Refused by default, because certificates from the copy would grant that group's membership at logon to everyone the copy's enrollment ACL admits. |
| `-KeepArtifacts` | switch | No | | Validate: keep the two throwaway templates and the export file for inspection (no companion OID object is ever created for them). |
| `-Server` | `string` | No | auto-discover | Pin all (target-side, for Sync) operations to a specific writable DC — point it at another forest's DC to operate there. Required together with `-Credential`. |
| `-Credential` | `pscredential` | No | current identity | Credentials used against `-Server` — the target side for Import/Sync/Validate, the *source* side for Export. With `-Server`, enables operating on a forest you are not logged on to — no trust needed. |
| `-SourceServer` | `string` | Sync | | Sync: DC (or domain name) in the **source** forest to read the template from. |
| `-SourceCredential` | `pscredential` | No | current identity | Sync: credentials for the source-side read. |
| `-WhatIf` / `-Confirm` | switch | No | | Preview / prompt. `-WhatIf` shows the planned template, OID object, and exact ACL grants. |

### Usage

**Source forest — export:**
```powershell
.\Sync-ADCSTemplate.ps1 -Mode Export -Path .\KerberosAuth.json
```

**Target forest (no AD CS needed) — import with the standard ACL:**
```powershell
.\Sync-ADCSTemplate.ps1 -Mode Import -Path .\KerberosAuth.json
```

**Import as a renamed copy with a freshly minted (forest-independent) OID:**
```powershell
.\Sync-ADCSTemplate.ps1 -Mode Import -Path .\KerberosAuth.json -OidHandling GenerateRandom `
    -NewTemplateName "YY-KerberosAuthentication" -NewDisplayName "YY-Kerberos Authentication"
```

**Standard ACL plus a template-admin group (what an external CA will read):**
```powershell
.\Sync-ADCSTemplate.ps1 -Mode Import -Path .\KerberosAuth.json -EnrollPrincipals @{
    'CONTOSO\PKI-Admins' = 'FullControl'
}
```

**Direct sync — run in the target forest, pull from the source forest over the trust (no file):**
```powershell
.\Sync-ADCSTemplate.ps1 -Mode Sync -SourceServer dc01.source.example `
    -NewTemplateName "YY-KerberosAuthentication" -NewDisplayName "YY-Kerberos Authentication"
```

**Direct sync from a third machine with explicit credentials on both sides (no trust needed):**
```powershell
.\Sync-ADCSTemplate.ps1 -Mode Sync `
    -SourceServer dc01.a.example -SourceCredential (Get-Credential A\template.reader) `
    -Server dc01.b.example -Credential (Get-Credential B\ent.admin)
```

**Mixed flow — import a previously exported JSON straight into another forest:**
```powershell
.\Sync-ADCSTemplate.ps1 -Mode Import -Path .\KerberosAuth.json `
    -Server dc01.b.example -Credential (Get-Credential B\ent.admin)
```

**Prove round-trip fidelity in the source forest first (creates and removes throwaway copies — checks both the file pipeline and the direct Sync pipeline):**
```powershell
.\Sync-ADCSTemplate.ps1 -Mode Validate -TemplateName "KerberosAuthentication"
```

### Using the template with EJBCA

**EJBCA** can use an AD certificate template as its enrollment-authorization source through Microsoft auto-enrollment (**MSAE**): it reads the template object and its ACL straight from AD, so a forest that never had AD CS can still host the template. The general cross-forest mechanics are covered above; the points below are what's specific to EJBCA.

**Start here — you may not need to build a template at all.** This repo ships **ready-to-import, EJBCA-ready** templates (the full-DN Subject already applied) under [`Templates/EJBCA/`](./Templates/README.md) — all four defaults whose Subject would otherwise be empty, including **Kerberos Authentication** (the modern DC template used as the field-validated example below). Import one directly and you can skip the preparation steps in this section entirely — nothing has to be installed or edited locally with `certutil`. For domain controllers:

```powershell
.\Sync-ADCSTemplate.ps1 -Mode Import -Path .\Templates\EJBCA\KerberosAuthentication-EJBCA.json -OidHandling GenerateRandom
```

The rest of this section explains how to prepare your *own* template when you need one the shipped set doesn't cover, and the mechanics behind why the EJBCA variants are built the way they are.

> **Field-validated end to end:** a Kerberos Authentication template with the full-DN Subject enabled was exported, imported into a separate forest with **no CA**, and delivered to clients with [`Add-CertificateEnrollmentPolicyServerToGpo`](#add-certificateenrollmentpolicyservertogpops1) — the DCs in scope each received a Kerberos Authentication certificate issued by EJBCA. The same template with an **empty Subject** was confirmed to fail, which is exactly the first point below.

- **The Subject cannot be empty.** The built-in **Kerberos Authentication** template issues certificates with an **empty Subject** — the identity lives entirely in the SAN as DNS entries — and EJBCA cannot use a template that produces no Subject. Fix it on the **source** template, **before** you export/sync: Certificate Templates MMC → *Properties* → **Subject Name** → *Build from this Active Directory information* → set **Subject name format** to **Fully distinguished name**. That sets the `CT_FLAG_SUBJECT_REQUIRE_DIRECTORY_PATH` bit (`0x80000000`) in `msPKI-Certificate-Name-Flag`; the script copies that attribute verbatim, so the setting rides into the synced copy — but only if it is on the source template first.
  - **This is not unique to Kerberos Authentication.** Any template that would otherwise issue an **empty** Subject needs the same treatment before EJBCA can use it.
  - **Full DN vs. common name:** the hard rule is only that the Subject must not be empty. Whether a bare **Common Name** is enough — or the **full distinguished name** is specifically required — depends on your EJBCA end-entity profile; "Fully distinguished name" is the safe choice that always populates the Subject.
- **The ACL is what EJBCA consumes** for enrollment authorization — build it with `-AclBase` / `-EnrollPrincipals` (see [Features](#features-2) and [Parameters](#parameters-2)).
- **Publishing is separate.** Import/Sync does **not** publish the template to any CA; you configure the template mapping in EJBCA itself.
- **No default templates to export from?** First, you may not need to produce any — the repo's [`Templates/`](./Templates/README.md) ships every default (and the EJBCA-ready variants) as importable JSON, so importing straight from there is usually the fastest path. If you do want to generate them yourself: the default templates (Kerberos Authentication included) normally arrive in AD when the first Enterprise CA is installed, so a forest that never had a CA won't have them — but they are plain AD objects, so no CA is needed to hold them. `certutil -InstallDefaultTemplates` writes the standard set into AD and **works with no AD CS role installed at all** (verified) — run it as an **Enterprise Admin** (add `-dc <DCName>` to target a specific DC). That gives you a stock template to edit (turn on the full-DN Subject above) and then export or sync.
- **Then point the clients at EJBCA.** Publishing the template is only half the job — Windows clients still need to be told to enroll against the EJBCA policy server. That is what [`Add-CertificateEnrollmentPolicyServerOffline`](#add-certificateenrollmentpolicyserverofflineps1) (per machine) and [`Add-CertificateEnrollmentPolicyServerToGpo`](#add-certificateenrollmentpolicyservertogpops1) (fleet-wide via GPO) do, computing the CEP registry values offline from the same EJBCA alias.

### Notes

- Import/Sync **refuse** to overwrite an existing template (same cn) or to duplicate an existing template OID — delete or rename instead of clobbering.
- The companion `msPKI-Enterprise-Oid` "display" object is the only OID-container object involved, and Import/Sync register it automatically; the source forest's own OID object is deliberately **not** copied verbatim (its other attributes are forest-specific).
- **Authentication Mechanism Assurance (AMA) links are not carried over** — `msDS-OIDToGroupLink` points at a group DN in the *source* forest. If you use AMA, recreate the link in the target forest against a local universal group. The reverse case **is** checked: when the copy carries an issuance policy OID that the *target* forest already links to a group, the import is refused unless `-AllowLinkedIssuancePolicy` is given, because certificates from the copy would grant that group's membership at logon.
- v1 templates copy too (with an advisory warning): the object round-trips faithfully — live-verified — but Windows fixes v1 semantics in code (name-matched, not editable, no autoenrollment), so import a v1 copy under its **original** name; non-Windows consumers that read the object/ACL directly are unaffected. The ACL is intentionally not exported — it is forest-specific and is rebuilt from `-AclBase`/`-EnrollPrincipals` on import.
- Import does **not** publish the template to any CA; with a Microsoft CA that remains a separate "Certificate Templates to Issue" step. (For EJBCA, template mapping is configured there — see [Using with EJBCA](#using-the-template-with-ejbca).)
- The JSON is written with a UTF-8 BOM so localized/accented display names survive a PowerShell 7 → 5.1 round-trip.
- Run `-Mode Validate` in a lab or the source forest before the first production import or sync — it exercises both the export→import file pipeline and the direct in-memory pipeline Sync uses, and diffs every PKI attribute the source carries for each.
- Parameters a mode does not consume are rejected up front (e.g. `-StripOid` with `-Mode Import`, `-OidRoot` without `GenerateFromRoot`, `-AclBase` with `-SkipAcl`) instead of being silently ignored.
- PKI attributes the built-in type lists don't know (a genuine schema extension linked to the template class) are typed automatically from the *target* forest's schema; an attribute is dropped with a warning when the schema cannot type it, does not permit it on the template class, or types it single-valued while the source value is empty or multi-valued. (v3/v4 CNG algorithm settings are not separate attributes — they travel packed inside `msPKI-RA-Application-Policies`.)
- v3/v4 templates can carry a **private-key SDDL** (`msPKI-Key-Security-Descriptor`) packed inside `msPKI-RA-Application-Policies`; it copies verbatim, and any domain SIDs in it are source-forest SIDs — the script warns so you can review the key ACL in the target forest.

### Tests

A Pester suite ([`Tests/Sync-ADCSTemplate.Tests.ps1`](./Tests/Sync-ADCSTemplate.Tests.ps1)) covers the script in four tiers (Pester **5+**; `Install-Module Pester -MinimumVersion 5.0 -Scope CurrentUser`):

| Tier (`-Tag`) | Needs | Changes anything? |
| --- | --- | --- |
| `Unit` | nothing (pure helpers, extracted from the script by AST so the real code runs) | no |
| `Static` | nothing (parse + comment-based help) | no |
| `Guard` | RSAT ActiveDirectory module (no reachable DC needed) | no |
| `Lab` | a lab forest — opt-in via `-RunLab` | **creates and removes AD objects** |

```powershell
# Safe tiers only — no changes (Unit + Static + Guard):
Invoke-Pester -Path .\Tests\Sync-ADCSTemplate.Tests.ps1 -ExcludeTag Lab

# Full run against a lab (targets are configurable; child/cross-forest are optional):
$cfg = New-PesterContainer -Path .\Tests\Sync-ADCSTemplate.Tests.ps1 -Data @{
    RunLab          = $true
    AronsServer     = 'dc1.lab.example'          # target DC (a single DC)
    ChildServer     = 'childdc.child.lab.example' # optional: child-domain root-SID path
    NorefjellServer = '10.0.0.9'                  # optional: a SEPARATE forest without AD CS
}
Invoke-Pester -Container $cfg
```

The `Lab` tier is surgical: every object it creates carries a unique per-run `PESTER-<hex>` prefix, is tracked by exact DN, and is removed in teardown (with a prefix-scoped safety-net sweep as backstop); pre-existing objects are never touched. Read-backs poll with retry so a target forest that lags briefly over ADWS after a write does not cause false failures.

<sub>[↑ Back to top](#top)</sub>

---

## Add-CertificateEnrollmentPolicyServerOffline.ps1

Registers (or removes) a **Certificate Enrollment Policy (CEP)** server in the registry entirely **offline** — no "Validate Server" round-trip, no contact with the policy server at all. It writes the same values the *"Certificate Services Client – Certificate Enrollment Policy"* dialog produces, but computes everything locally: the SHA-1 subkey name from the URL, and (for EJBCA/MSAE) the PolicyID from the alias's Policy Name.

### Why you need this

The built-in dialog (and the `X509Enrollment` COM path behind it) insists on reaching the policy server's MS-XCEP `GetPolicies` endpoint before it will save anything. That is exactly what you cannot do when you are staging a machine before the PKI is reachable, building a golden image, working in an air-gapped or change-controlled environment, or scripting an identical config across many machines. This script derives every value locally and writes it directly, so enrollment-policy configuration becomes a repeatable, unattended step.

It also handles the parts the dialog hides: on the Group Policy hives it preserves the built-in **AD enrollment policy row** (without it, enabling GP-based CEP silently removes the Active Directory enrollment policy and autoenrollment against AD-published templates stops), and it manages the `PolicyServers` root **Flags** DISABLE bits.

### Features

- **Fully offline derivation** — SHA-1 over the UTF-16LE bytes of the invariant-lowercased URL for the subkey; EJBCA MSAE PolicyID = Java `String.hashCode()` of the Policy Name (or pass `-PolicyId` for a GUID-returning CEP such as Microsoft's)
- **Four target locations** — `LocalMachine` / `LocalUser` (the user-configured store `certlm.msc` / `certmgr.msc` manages) and `GPMachine` / `GPUser` (the Group Policy hives)
- **AD enrollment policy row preserved** on the GP locations (opt out with `-SkipADPolicy`), so GP-based CEP does not knock out the AD default policy
- **Complete-row gates** — a pre-existing AD row or CEP entry satisfies the AD-row prerequisite, the `(Default)` marker and `-ReplaceExisting` only when it is *usable*: URL and PolicyID as requested **and** `FriendlyName` plus DWORD-typed `Flags` / `AuthFlags` / `Cost` present
- **Root Flags handled as DISABLE bits** — clears the "ignore GP list" bit (`0x2`) if present; `-DisableUserConfigured` / `-EnableUserConfigured` toggle the "ignore user-configured servers" bit (`0x4`); existing bits are preserved across runs
- **Read-back verification** of every value written (missing values detected); the summary reports the *actual* registry state and per-gate outcomes
- **Protected registry path** — before any write, every existing key from the hive root down to the target is checked: a registry **symbolic link**, an untrusted owner, or write-class rights for an untrusted principal refuse the run (a link planted where `PolicyServers` does not exist yet would carry an elevated first-time write to whatever key it points at). String values and the `(Default)` marker are written as `REG_SZ` explicitly and every value's *kind* is verified, not just its content; a removal is verified before it is reported
- **`-SetAsDefault` / `-ClearDefault`** for the unnamed `(Default)` interactive-enrollment marker, and **`-ReplaceExisting`** to clean up stale same-PolicyID siblings from a superseded URL
- **`-Remove`** mode, **`-WhatIf` / `-Confirm`** support, and structured `PSCustomObject` output

### Requirements

- Windows PowerShell 5.1 or PowerShell 7+ — **no PowerShell module required**
- An **elevated** session for `-Location LocalMachine`, `GPMachine`, or `GPUser` (`LocalUser` needs no elevation)
- Domain connectivity only for the GP-location AD policy row (on a workgroup machine that lookup is skipped with a warning)
- **Tattooing caveat:** writing directly into the GP hives (`GPMachine` / `GPUser`) on a **domain member** produces pseudo-policy backed by no GPO — invisible to RSoP/`gpresult`, not reverted by `gpupdate`, read-only in the certificate MMC. On domain members use [`Add-CertificateEnrollmentPolicyServerToGpo`](#add-certificateenrollmentpolicyservertogpops1) instead; the GP locations here are for standalone/workgroup machines and lab work.

### Parameters

| Parameter | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| `-Url` | `string` | Yes | | Full CEP URI, e.g. `https://pki.example.net/ejbca/msae/CEPService?alias`. Used verbatim for the SHA-1 subkey and by clients' `GetPolicies` calls; must be an absolute http/https URI. |
| `-PolicyName` | `string` | Add only | | EJBCA MSAE alias "Policy Name". Becomes `FriendlyName` and (unless `-PolicyId` is set) is hashed **verbatim** to the PolicyID — keep it identical to EJBCA. |
| `-PolicyId` | `string` | No | hash of `-PolicyName` | Explicit PolicyID for non-EJBCA servers (must match the server's `GetPolicies` response). |
| `-Location` | `LocalMachine` / `LocalUser` / `GPMachine` / `GPUser` | No | `LocalMachine` | Which store to write. See the tattooing caveat for the GP hives. |
| `-Authentication` | `Anonymous` / `Kerberos` / `UsernamePassword` / `Certificate` | No | `Kerberos` | Client authentication type for the endpoint (`Kerberos` = "Windows integrated"). |
| `-Cost` | `long` (1–4294967295) | No | `0x7FFFFFFD` | Priority; lower = preferred among endpoints sharing a PolicyID. Pass large values in decimal. |
| `-NoAutoEnroll` | switch | No | | Leave "Enable for automatic enrollment and renewal" off (clears Flags bit `0x10`). |
| `-AllowUntrustedIssuer` | switch | No | | Uncheck "Require strong validation during enrollment" (sets Flags bit `0x20`). |
| `-NoClientId` | switch | No | | Do not send the ClientId attribute (clears Flags bit `0x4`; default `0x14` matches the GPO editor). |
| `-SetAsDefault` / `-ClearDefault` | switch | No | | Set / clear the unnamed `(Default)` marker (interactive-enrollment preselection only). |
| `-SkipADPolicy` | switch | No | | GP locations: do **not** write the AD enrollment policy row (only when that removal is intended). |
| `-ReplaceExisting` | switch | No | | Remove sibling entries with the same PolicyID but a different URL (stale/typo). The AD row is never removed. |
| `-DisableUserConfigured` / `-EnableUserConfigured` | switch | No | | GP locations: set / clear root Flags bit `0x4` (ignore user-configured servers). |
| `-Remove` | switch | Remove mode | | Delete the entry for `-Url` from the chosen location and clear a now-orphaned `(Default)` marker. |
| `-WhatIf` / `-Confirm` | switch | No | | Preview / prompt. `-WhatIf` shows the computed subkey and PolicyID and writes nothing. |

### Usage

**Preview everything (computed subkey + PolicyID), no writes:**
```powershell
.\Add-CertificateEnrollmentPolicyServerOffline.ps1 `
    -Url 'https://pki.example.net/ejbca/msae/CEPService?alias' -PolicyName 'Example PKI Service' -WhatIf
```

**Configure the per-user store and mark it the default enrollment policy:**
```powershell
.\Add-CertificateEnrollmentPolicyServerOffline.ps1 `
    -Url 'https://pki.example.net/ejbca/msae/CEPService?alias' -PolicyName 'Example PKI Service' `
    -Location LocalUser -SetAsDefault
```

**Remove that entry again:**
```powershell
.\Add-CertificateEnrollmentPolicyServerOffline.ps1 `
    -Url 'https://pki.example.net/ejbca/msae/CEPService?alias' -Location LocalUser -Remove
```

### Notes

- `-PolicyName` is hashed **verbatim** — renaming the alias in EJBCA changes the PolicyID and orphans already-deployed entries.
- Rerunning with a **different** URL does not remove the old entry (multiple URLs per PolicyID is also the legitimate redundant-endpoint pattern); use `-ReplaceExisting` to clean up a superseded one.
- GP locations: the domain objectGUID for the AD Enrollment Policy row is resolved **before** anything is written. On a domain-joined machine a failed lookup aborts the run with nothing written (a GP configuration without that row removes the AD enrollment policy); on a workgroup machine the row is skipped with a warning. The row is written before the CEP entry; if the row's confirmation was declined and the hive carries none, the CEP entry is not written and the run stops there. The `(Default)` marker and `-ReplaceExisting` removals run only when a complete CEP entry for this URL exists (URL and PolicyID as requested, re-read right before the cleanup). Pass `-SkipADPolicy` to omit it deliberately.
- `-Remove` handles a single entry and its `(Default)` marker; shared/root configuration (root Flags, the AD row, autoenrollment) is left in place — full manual teardown is documented in the script's `.NOTES`.

### Tests

The Pester suite ([`Tests/Add-CertificateEnrollmentPolicyServerOffline.Tests.ps1`](./Tests/Add-CertificateEnrollmentPolicyServerOffline.Tests.ps1)) runs three always-safe tiers — **Unit** (the real subkey/PolicyID/flag derivations, exercised through `-WhatIf` so nothing is written), **Static** (parse + comment-based help), and **Guard** (parameter-conflict validation) — plus an opt-in **Lab** tier of live registry round-trips (add → idempotent update → default marker → replace-sibling → remove, all verified against independent oracles). The safe tiers need no module, no AD, no elevation, and change nothing:

```powershell
Invoke-Pester -Path .\Tests\Add-CertificateEnrollmentPolicyServerOffline.Tests.ps1 -ExcludeTag Lab
```

The Lab tier **writes to this machine's registry** — only to the user-configured stores (`LocalUser` always, `LocalMachine` when elevated), never to the GP hives (on a domain member those tattoo pseudo-policy; the GPO suite covers Group Policy delivery instead). It is surgical: entries carry a per-run `PESTER-<hex>` name and a URL under the RFC-reserved `.invalid` TLD, every created key is tracked by exact path and removed in teardown, a pre-existing `(Default)` marker is snapshotted and restored, and a `PolicyServers` base key is removed only if the run created it and it ends the run empty:

```powershell
$cfg = New-PesterContainer -Path .\Tests\Add-CertificateEnrollmentPolicyServerOffline.Tests.ps1 -Data @{ RunLab = $true }
Invoke-Pester -Container $cfg
```

<sub>[↑ Back to top](#top)</sub>

---

## Add-CertificateEnrollmentPolicyServerToGpo.ps1

Writes (or removes) the same **Certificate Enrollment Policy** setting directly in a domain **GPO** — still offline with respect to the policy server (no "Validate Server" round-trip, ever) — for fleet-wide rollout. It authors the values with `Set-GPRegistryValue`, which also does what hand-editing SYSVOL gets wrong: the AD + `GPT.INI` version bumps and Registry CSE registration. The result appears in the GPME *Public Key Policies* dialog exactly as if clicked in.

### Why you need this

Configuring CEP through Group Policy normally means the same server round-trip in the GPME dialog, one GPO at a time, by hand. This script authors the policy directly and correctly for a fleet, and it defends against three sharp edges: it preserves the built-in **AD enrollment policy row** (a GPO carrying only your CEP entry would otherwise silently *remove* the AD enrollment policy from every client in scope), it writes each value individually to avoid the `Set-GPRegistryValue` list-form **`**delVals.`** deletion record that makes clients delete the whole entry, and it reads state from the **PDC emulator** (or `-Server`) so reads and writes see the same replica.

### Features

- **Authored straight into the GPO** with `Set-GPRegistryValue` — AD/`GPT.INI` version bumps and Registry CSE registration handled; shows up normally in GPME
- **Same offline derivations** as the per-machine script (SHA-1 subkey, EJBCA `String.hashCode()` PolicyID, or `-PolicyId`)
- **AD enrollment policy row preserved** by default so the GPO does not take away the AD default policy fleet-wide (opt out with `-SkipADPolicy`)
- **Complete-row gates** — a pre-existing AD row or CEP entry satisfies the AD-row prerequisite, the `(Default)` marker and `-ReplaceExisting` only when it is *usable*: URL and PolicyID as requested **and** `FriendlyName` plus DWORD `Flags` / `AuthFlags` / `Cost` present (URL + PolicyID alone is what an interrupted write leaves behind)
- **Deletion-record-safe** — individual value writes avoid the `**delVals.` trap, and the script **detects and warns** about mis-ordered deletion records already in the GPO; `registry.pol` state is read as a client applies it — records replayed in order, `**del.` / `**delvals.` / `**DeleteValues` / `**DeleteKeys` honoured with their data read as a string whatever the record's type says (`**soft.` keeps its type)
- **Optional Auto-Enrollment** (`-EnableAutoEnrollmentPolicy`) in the same GPO scope (AEPolicy / expiration percent / store), warning before it changes existing values
- **Same-replica state reads** from the PDC emulator (or `-Server`), retry on transient SYSVOL/`registry.pol` contention, and **read-back verification** of every entry
- **Effective-state gates** — the `(Default)` marker and `-ReplaceExisting` act only when the replacing entry is complete in *both* the live GPMC view and the `registry.pol` replay (a row whose authored values a later deletion record wipes does not count), a freshly written row must show in the replay or the run fails, and a removal is verified from `registry.pol` before it is reported
- **Machine or User scope**, root Flags DISABLE-bit handling, `-SetAsDefault` / `-ClearDefault`, `-ReplaceExisting`, a `-Remove` mode, `-WhatIf` / `-Confirm`, and structured output

### Requirements

- Windows PowerShell 5.1 or PowerShell 7+ (on 7 the GroupPolicy module loads through the WinPSCompat shim — expect its one-time compatibility warning; the script deliberately avoids `#Requires -Modules`, which would refuse to run there)
- The **GroupPolicy module** (GPMC / RSAT) and permission to edit the target GPO
- An existing, linked GPO to write into (create one first, e.g. `New-GPO -Name 'PKI - Enrollment Policy' | New-GPLink -Target 'OU=...,DC=...'`)
- Domain connectivity (the script targets the PDC emulator by default, or the DC you pass to `-Server`)
- Do **not** edit the same GPO concurrently from another session or GPME — `Set-GPRegistryValue` is an unlocked read-modify-write; the read-back detects loss of this script's own entry

### Parameters

| Parameter | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| `-GpoName` | `string` | Yes | | Display name of an existing GPO, or its GUID (display name is tried first). |
| `-Url` | `string` | Yes | | Full CEP URI (verbatim; absolute http/https). |
| `-PolicyName` | `string` | Add only | | EJBCA MSAE alias "Policy Name" — `FriendlyName` and (unless `-PolicyId`) the PolicyID hash input, hashed **verbatim**. |
| `-PolicyId` | `string` | No | hash of `-PolicyName` | Explicit PolicyID for a GUID-returning CEP server. |
| `-Scope` | `Machine` / `User` | No | `Machine` | Computer Configuration (HKLM) or User Configuration (HKCU). |
| `-Authentication` | `Anonymous` / `Kerberos` / `UsernamePassword` / `Certificate` | No | `Kerberos` | Endpoint client authentication type. |
| `-Cost` | `long` (1–4294967295) | No | `0x7FFFFFFD` | Priority; lower preferred among endpoints sharing a PolicyID. |
| `-NoAutoEnroll` / `-AllowUntrustedIssuer` / `-NoClientId` | switch | No | | Entry Flags tweaks (clear `0x10` / set `0x20` / clear `0x4`); default `0x14` matches GPME. |
| `-SetAsDefault` / `-ClearDefault` | switch | No | | Set / clear the `(Default)` marker (interactive preselection only). |
| `-SkipADPolicy` | switch | No | | Do **not** write the AD enrollment policy row (only when that removal is intended). |
| `-ReplaceExisting` | switch | No | | Remove same-PolicyID sibling entries under a different URL (the AD row is never removed). |
| `-DisableUserConfigured` / `-EnableUserConfigured` | switch | No | | Set / clear root Flags bit `0x4` (ignore user-configured servers). |
| `-EnableAutoEnrollmentPolicy` | switch | No | | Also write the "Auto-Enrollment" setting into the same scope. |
| `-AEPolicy` / `-AEExpirationPercent` / `-AEStore` | `int` / `int` / `string` | No | `7` / `10` / `MY` | Auto-Enrollment values (only with `-EnableAutoEnrollmentPolicy`). |
| `-Domain` / `-Server` | `string` | No | current / PDC emulator | Domain and DC for the GroupPolicy cmdlets and the `registry.pol` state reads (kept on one replica). |
| `-Remove` | switch | Remove mode | | Delete the entry for `-Url` from the GPO scope and clear a now-orphaned `(Default)` marker. |
| `-WhatIf` / `-Confirm` | switch | No | | Preview / prompt. |

### Usage

**Preview a fleet rollout with Auto-Enrollment enabled:**
```powershell
.\Add-CertificateEnrollmentPolicyServerToGpo.ps1 -GpoName 'PKI - Enrollment Policy' `
    -Url 'https://pki.example.net/ejbca/msae/CEPService?alias' -PolicyName 'Example PKI Service' `
    -EnableAutoEnrollmentPolicy -WhatIf
```

**Author it for real into the Computer configuration:**
```powershell
.\Add-CertificateEnrollmentPolicyServerToGpo.ps1 -GpoName 'PKI - Enrollment Policy' `
    -Url 'https://pki.example.net/ejbca/msae/CEPService?alias' -PolicyName 'Example PKI Service' `
    -EnableAutoEnrollmentPolicy
```

**Remove it from the User scope again:**
```powershell
.\Add-CertificateEnrollmentPolicyServerToGpo.ps1 -GpoName 'PKI - Enrollment Policy' `
    -Url 'https://pki.example.net/ejbca/msae/CEPService?alias' -Scope User -Remove
```

### Notes

- After a change, clients pick it up at the next GP refresh — force with `gpupdate`, then trigger enrollment with `certutil -pulse` (machine) or `certutil -user -pulse` (user).
- `-PolicyName` is hashed **verbatim**; keep it identical to the EJBCA alias or deployed entries orphan.
- `-Remove` clears one entry and its `(Default)` marker; shared configuration (root Flags, the AD row, Auto-Enrollment) is left in place — full teardown via `Remove-GPRegistryValue` is listed in the script's `.NOTES`.
- The domain objectGUID for the AD Enrollment Policy row is resolved **before** any GPO write; if it cannot be resolved the run aborts with nothing written (a GPO carrying a CEP entry but no `LDAP:` row removes the AD enrollment policy from every client in scope). The row is then written and verified **before** the CEP entry; if the row's confirmation was declined and the GPO carries none, the CEP entry is not written and the run stops there (no root Flags, marker, Auto-Enrollment or sibling changes). The `(Default)` marker and `-ReplaceExisting` removals run only when a **complete** CEP entry for this URL exists (URL and PolicyID as requested, re-read right before the cleanup), so a declined entry prompt or a half-written entry cannot leave a marker pointing at nothing, delete the only working endpoint, or act for a PolicyID the retained entry does not serve. Pass `-SkipADPolicy` to omit the row deliberately.
- A GUID-shaped `-GpoName` is read as a GPO ID only after an independent listing proves no GPO carries that display name; any other name-lookup failure aborts instead of silently retargeting.

### Tests

The Pester suite ([`Tests/Add-CertificateEnrollmentPolicyServerToGpo.Tests.ps1`](./Tests/Add-CertificateEnrollmentPolicyServerToGpo.Tests.ps1)) runs three always-safe tiers — **Unit** (the real `Registry.pol` binary parser and entry/value extractors, exercised against a hand-built `.pol` stream; no module needed), **Static** (parse + help), and **Guard** (parameter-conflict validation that throws before any GPO is touched; needs the GroupPolicy module, self-skips without it) — plus an opt-in **Lab** tier that runs the full lifecycle (author → verify via GPMC *and* the real `registry.pol` → default marker + Auto-Enrollment → replace-sibling → remove, Machine and User scope) inside **one throwaway, never-linked GPO** that applies to zero clients and is deleted afterwards by its exact tracked GUID. The safe tiers write no GPO:

```powershell
Invoke-Pester -Path .\Tests\Add-CertificateEnrollmentPolicyServerToGpo.Tests.ps1 -ExcludeTag Lab
```

The Lab tier needs a domain-joined machine, the GroupPolicy module, and permission to create GPOs (a lab DC is ideal); pre-existing GPOs are never touched:

```powershell
$cfg = New-PesterContainer -Path .\Tests\Add-CertificateEnrollmentPolicyServerToGpo.Tests.ps1 -Data @{ RunLab = $true }
Invoke-Pester -Container $cfg
```

<sub>[↑ Back to top](#top)</sub>

---

## Template library

[`Templates/`](./Templates/README.md) ships importable JSON exports of **all 33 certutil default certificate templates** (`Templates/Default/`) and **EJBCA-ready variants** (`Templates/EJBCA/`) of all four defaults whose Subject is empty at creation (`DirectoryEmailReplication`, `DomainControllerAuthentication`, `KerberosAuthentication`, `Workstation`) — the same one-bit full-DN-Subject change the [EJBCA section](#using-the-template-with-ejbca) describes, pre-applied. A parallel **`Templates/MaxCompat/`** carries both sets moved to the newest compatibility (**CA: Windows Server 2016**, **recipient: Windows 10 / Windows Server 2016**) for the 9 templates where that is possible (the 24 schema-v1 built-ins are read-only and carry through unchanged). Template OIDs are stripped from every file (they would fingerprint the source forest; importers mint their own with `-OidHandling Generate` / `GenerateRandom`), names are kept, no file carries a domain SID; the whole library was produced **clean-room** (seeded into a fresh forest with `certutil -InstallDefaultTemplates`) and every file was validated against a live DC. Details, criteria, the compatibility encoding, and the full per-template table live in the folder's [README](./Templates/README.md).

<sub>[↑ Back to top](#top)</sub>

---

## License

[MIT](./LICENSE)

## Contributing

Issues and pull requests welcome — see [CONTRIBUTING.md](./CONTRIBUTING.md) for the house style, the dual-engine (PowerShell 5.1 + 7) and `-WhatIf` expectations, and how to run the analyzer and the four-tier test suites.

## Security

To report a vulnerability, use GitHub's private reporting (**Security → Report a vulnerability**) rather than a public issue — details and the safety model are in [SECURITY.md](./SECURITY.md).
