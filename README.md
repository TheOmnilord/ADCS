# ADCS — Active Directory Certificate Services PowerShell Tools

A set of PowerShell scripts for administering **Active Directory Certificate Services** (AD CS / ADCS) from the command line. Built for Windows PKI administrators who need to manage certificate templates and process certificate requests at scale, without clicking through the Certificate Templates MMC snap-in or the Certification Authority console.

Currently includes:

- **Bulk certificate template validity updates** — useful for rolling out the CA/Browser Forum **SC-081** validity reductions (200 days from March 2026, 100 days from March 2027, 47 days from March 2029) across many templates at once.
- **Batch CSR submission to an Enterprise CA** via `certreq.exe`, with resume-safe CSV tracking of request IDs and automated retrieval of issued certificates.

No AD PowerShell module dependency. Works on Windows PowerShell 5.1 and PowerShell 7+.

## Scripts

| Script | Description |
| --- | --- |
| [`Set-ADCSTemplateValidity.ps1`](./Set-ADCSTemplateValidity.ps1) | Bulk-update the validity period (and optionally the renewal overlap period) on one or more certificate templates, with wildcard name matching. |
| [`Submit-CertificateRequests.ps1`](./Submit-CertificateRequests.ps1) | Batch-submit `.req`/`.csr`/`.txt` files to an ADCS CA via `certreq.exe`, track request IDs in a CSV, and later retrieve the issued certificates. |

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
| `-TemplateName` | `string[]` | Yes | One or more template CN names. Supports LDAP wildcards (`*`, `?`). |
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

### How It Works

1. Connects to `RootDSE` to resolve the Configuration naming context.
2. Searches `CN=Certificate Templates,CN=Public Key Services,CN=Services,<ConfigNC>` for templates matching the pattern(s) using an LDAP filter.
3. For each match, decodes the current `pKIExpirationPeriod` / `pKIOverlapPeriod` (8-byte little-endian negative FILETIME ticks).
4. Compares against the requested value; skips if already equal.
5. Writes the new byte array via `DirectoryEntry.InvokeSet()` and calls `SetInfo()` to commit.

---

## Submit-CertificateRequests.ps1

Batch-submits certificate signing requests (`.req` / `.csr` / `.txt`) from a folder to an ADCS CA using `certreq.exe`, tracks each submission's request ID in a CSV file, and can later retrieve the issued certificates.

### Features

- **Batch submit** all request files in a folder in one run
- **CSV tracking file** records request ID, submit time, status, error messages per file
- **Resume-safe** - files already present in the tracking CSV are skipped on re-run; with `-Force` (or an interactive y/n confirmation) you can resubmit a tracked file as a new request
- **Retrieve mode** picks up previously-submitted `Pending` requests and pulls issued `.cer` files
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
| `-InputPath` | `string` | Yes | | Folder containing `.req` / `.csr` / `.txt` request files. |
| `-CAConfig` | `string` | Yes | | CA configuration string, e.g. `CA01.domain.com\Contoso Issuing CA 1`. |
| `-CertificateTemplate` | `string` | Yes | | Certificate template name (the CN, not the display name). |
| `-TrackingFile` | `string` | No | `.\CertTracking.csv` | CSV file used to track request IDs and statuses across runs. |
| `-OutputFolder` | `string` | No | `.\Certificates` | Folder where issued `.cer` files are saved (one per request, named after the request file). |
| `-Mode` | `Submit` / `Retrieve` / `Both` | No | `Submit` | `Submit` = submit new requests only; `Retrieve` = pull certs for previously-pending requests; `Both` = do both. |
| `-KeepRspFile` | switch | No | | By default the `.rsp` file `certreq` writes next to each retrieved `.cer` is deleted. Specify this switch to leave it in place. |
| `-Force` | switch | No | | Resubmit request files that already have a tracked RequestID without prompting. Without `-Force`, the script asks y/n for each already-submitted file (default = No / skip). |
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

**Retrieve any issued certificates for previously-pending requests:**
```powershell
.\Submit-CertificateRequests.ps1 `
    -InputPath "C:\CSRs" `
    -CAConfig "CA01.domain.com\Contoso Issuing CA 1" `
    -CertificateTemplate "WebServer" `
    -Mode Retrieve
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
| `Status` | `Issued`, `Pending`, `Denied`, `Error`, or `Unknown` |
| `OutputCertFile` | Full path where the issued `.cer` is saved |
| `LastCheckTime` | ISO-8601 timestamp of last status check |
| `ErrorMessage` | Error output from `certreq` if the submission or retrieval failed |

### Notes

- The tracking CSV is the source of truth for resume behavior. Deleting it will cause the script to re-submit all files (and the CA may issue duplicates).
- `Pending` typically means the CA requires manager approval. Re-run in `Retrieve` mode after approval to pull the issued cert.
- Issued `.cer` files are named after the source request file (e.g. `server1.req` -> `server1.cer`).
- Empty request files are skipped with a warning.
- A timestamped log file is created in the working directory for each run.

---

## License

[MIT](./LICENSE)

## Contributing

Issues and pull requests welcome.
