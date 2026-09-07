<a id="top"></a>

# ADCS — Active Directory Certificate Services PowerShell Tools

[![CI](https://github.com/TheOmnilord/ADCS/actions/workflows/ci.yml/badge.svg)](https://github.com/TheOmnilord/ADCS/actions/workflows/ci.yml)
[![Latest release](https://img.shields.io/github/v/release/TheOmnilord/ADCS?sort=semver&label=release)](https://github.com/TheOmnilord/ADCS/releases/latest)

This repository holds a set of PowerShell scripts that administer **Active Directory Certificate Services** (AD CS / ADCS) from the command line. The scripts are for Windows Public Key Infrastructure (PKI) administrators who manage certificate templates, process certificate requests, and configure enrollment clients at scale. With the scripts you do not click through the Certificate Templates MMC snap-in, the Certification Authority console, or the per-machine "Certificate Enrollment Policy" dialog.

The repository currently includes:

- **Bulk certificate template validity updates.** Use this to apply the CA/Browser Forum **SC-081** validity reductions to many templates at once. The limits are 200 days from March 2026, 100 days from March 2027, and 47 days from March 2029.
- **Batch CSR submission to an Enterprise CA** through `certreq.exe`. The script tracks the request IDs in a CSV file, so a run can resume, and it retrieves the issued certificates.
- **Cross-forest certificate template sync** between forests at the directory level, with all access over Active Directory Web Services (ADWS). The script copies through a JSON export and import, or **directly forest-to-forest in one run** with `-Mode Sync`. Sync accepts explicit credentials for each side, so it requires no trust. The script offers an optional rename, controlled OID handling, and a composable enrollment ACL. It also works when the target forest has **no AD CS installed**. An example is a template that an external CA reads for enrollment authorization.
- **Client Certificate Enrollment Policy (CEP) configuration.** The scripts point Windows enrollment clients at an **EJBCA** policy server, or at another MS-XCEP policy server, so the clients enroll against the synced template. They compute every registry value **offline**, with no "Validate Server" round-trip. You can apply the setting to a single machine or to a local Group Policy hive. You can also apply it fleet-wide, when the script writes the setting directly into a domain Group Policy Object (**GPO**), with optional Auto-Enrollment. This is the client half of the template sync above.
- **A ready-to-import template library** ([`Templates/`](./Templates/README.md)). It holds every certutil default template as JSON. The template OIDs are removed, so no file identifies a forest, and an importer generates its own OID. The library also holds **EJBCA-ready variants** of the templates whose Subject is empty at creation. It also holds a **latest-compatibility** copy (Windows Server 2016 / Windows 10) of both sets.

The scripts work on Windows PowerShell 5.1 and PowerShell 7+. `Set-ADCSTemplateValidity` and `Submit-CertificateRequests` need no AD PowerShell module. `Sync-ADCSTemplate` requires the RSAT ActiveDirectory module. `Add-CertificateEnrollmentPolicyServerToGpo` requires the GroupPolicy module. `Add-CertificateEnrollmentPolicyServerOffline` needs no module. See the requirements of each script.

**Versioning.** Each script has its own version in a `PSScriptInfo` header at the top of the file. The version changes only when that script changes. Releases are tagged on the repository (see the badge above). [CHANGELOG.md](./CHANGELOG.md) lists the script versions that each release ships. To check whether a deployed copy is current without opening it, run:

```powershell
Test-ScriptFileInfo .\Submit-CertificateRequests.ps1 | Select-Object Name, Version
```

## Scripts

Jump to a script, or to a section within it. Each script title links to its full documentation. The source file is linked next to it.

- **[Set-ADCSTemplateValidity.ps1](#set-adcstemplatevalidityps1)** &nbsp;·&nbsp; [source](./Set-ADCSTemplateValidity.ps1)
  Bulk-update the validity period, and optionally the renewal overlap period, on one or more certificate templates. The script matches template names with wildcards.
  <br>↳ [Why you need this](#why-you-need-this) · [Features](#features) · [Requirements](#requirements) · [Parameters](#parameters) · [Usage](#usage) · [Output](#output) · [Notes](#notes) · [How It Works](#how-it-works)
- **[Submit-CertificateRequests.ps1](#submit-certificaterequestsps1)** &nbsp;·&nbsp; [source](./Submit-CertificateRequests.ps1)
  Batch-submit `.req`, `.csr` and `.txt` files to an ADCS CA through `certreq.exe`. The script tracks the request IDs in a CSV file and later retrieves the issued certificates.
  <br>↳ [Features](#features-1) · [Requirements](#requirements-1) · [Parameters](#parameters-1) · [Usage](#usage-1) · [Friendly Error Hints](#friendly-error-hints) · [Run Summary](#run-summary) · [Tracking CSV Schema](#tracking-csv-schema) · [Notes](#notes-1)
- **[Sync-ADCSTemplate.ps1](#sync-adcstemplateps1)** &nbsp;·&nbsp; [source](./Sync-ADCSTemplate.ps1)
  Copy the Kerberos Authentication certificate template, or any other template, between forests. The copy goes through a JSON file or directly forest-to-forest in one run. The script offers an optional rename, four OID-handling modes, per-side credentials, a composable enrollment ACL, and a round-trip validation mode. The target forest does not need AD CS.
  <br>↳ [Why you need this](#why-you-need-this-1) · [Features](#features-2) · [Requirements](#requirements-2) · [Parameters](#parameters-2) · [Usage](#usage-2) · [Using with EJBCA](#using-the-template-with-ejbca) · [Notes](#notes-2) · [Tests](#tests)
- **[Add-CertificateEnrollmentPolicyServerOffline.ps1](#add-certificateenrollmentpolicyserverofflineps1)** &nbsp;·&nbsp; [source](./Add-CertificateEnrollmentPolicyServerOffline.ps1)
  Register or remove an EJBCA/MSAE enrollment-policy server on one machine, in the user-configured store or in a Group Policy hive. The script writes the exact registry values that the CEP dialog produces, and it computes them entirely offline.
  <br>↳ [Why you need this](#why-you-need-this-2) · [Features](#features-3) · [Requirements](#requirements-3) · [Parameters](#parameters-3) · [Usage](#usage-3) · [Notes](#notes-3) · [Tests](#tests-1)
- **[Add-CertificateEnrollmentPolicyServerToGpo.ps1](#add-certificateenrollmentpolicyservertogpops1)** &nbsp;·&nbsp; [source](./Add-CertificateEnrollmentPolicyServerToGpo.ps1)
  Write or remove the same enrollment-policy setting directly in a domain **GPO** through `Set-GPRegistryValue`, for a fleet-wide rollout. The script offers optional Auto-Enrollment, keeps the AD policy row, and checks `registry.pol` for safety.
  <br>↳ [Why you need this](#why-you-need-this-3) · [Features](#features-4) · [Requirements](#requirements-4) · [Parameters](#parameters-4) · [Usage](#usage-4) · [Notes](#notes-4) · [Tests](#tests-2)

The repository also has the **[Template library](#template-library)** ([`Templates/`](./Templates/README.md)). It holds importable JSON exports of all default templates, plus EJBCA-ready variants.

---

## Set-ADCSTemplateValidity.ps1

The script modifies the `pKIExpirationPeriod` attribute, and optionally the `pKIOverlapPeriod` attribute, on ADCS certificate templates in Active Directory. It matches template names with wildcards, so you can update many templates in one run.

### Why you need this

The CA/Browser Forum ballot **SC-081**, passed in April 2025, requires a phased reduction of the maximum validity period for publicly-trusted TLS server certificates. Organizations use ADCS mostly for an internal Public Key Infrastructure (PKI). Many organizations still apply these limits on their internal CAs, to keep the templates aligned with industry practice. The limits also prepare the templates for the case where a template later feeds a publicly-trusted chain.

| Effective date | Maximum validity (TLS server certs) |
| --- | --- |
| **15 March 2026** (current) | **200 days** |
| **15 March 2027** | **100 days** |
| **15 March 2029** | **47 days** |

As the limits decrease, a manual change of every template through the Certificate Templates MMC snap-in takes much work. This script updates dozens of templates in seconds:

```powershell
# March 2026 rollover: drop TLS templates to 200 days
.\Set-ADCSTemplateValidity.ps1 -TemplateName "*Web*","*TLS*" -ValidityPeriod 200 -ValidityPeriodUnit Days -WhatIf
```

SC-081 does **not** cover client authentication, code signing, S/MIME, and other non-TLS templates. Those templates can keep longer validity periods. Use targeted wildcards, so that you do not change them.

#### The Apple/Safari ceiling: 825 days

You can decide to ignore SC-081 on your internal CA. There is one limit that you cannot ignore. **Safari on macOS and iOS rejects any TLS server certificate with a validity longer than 825 days** (about 2 years and 3 months). The limit applies whether the issuing CA is publicly trusted or a user-added or admin-added internal root. Apple's `trustd` daemon has enforced this limit since iOS 13 / macOS 10.15 (July 2019).

The failure is hard to troubleshoot. Safari shows a generic *"cannot establish a secure connection"* error with no override option. The same certificate works in Chrome and Firefox on the same machine. Your internal policy can allow 5-year or 10-year templates. A certificate with a validity over 825 days still fails for Apple users, with no clear error.

Sources: [michalspacek.com](https://www.michalspacek.com/validity-period-of-https-certificates-issued-from-a-user-added-ca-is-essentially-2-years), [certkit.io](https://www.certkit.io/blog/apple-doesnt-care-who-signed-your-certificate).

### Features

- **Wildcard matching** on the template CN, for example `User*`, `*Web*`, `*VPN*`
- **Human-readable durations** (`Years`, `Months`, `Weeks`, `Days`, `Hours`)
- **`-WhatIf` / `-Confirm`** support with `ConfirmImpact = 'High'`
- **No AD PowerShell module required**. The script uses `System.DirectoryServices` directly
- **Skips templates already set** to the requested value, after a byte-array compare
- **Auto-increments** `msPKI-Template-Minor-Revision`, so the CAs detect the change
- **Deduplication** when multiple patterns match the same template
- **Summary output** with counts of Modified / Already set / Skipped / Errors
- **Pipeline-friendly output** as `PSCustomObject` per template

### Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- A domain-joined machine, or `-Server` to target a specific domain controller (DC)
- Permission to modify certificate templates. This is usually Enterprise Admin, or delegated rights on the `CN=Certificate Templates` container in the Configuration naming context

### Parameters

| Parameter | Type | Required | Description |
| --- | --- | --- | --- |
| `-TemplateName` | `string[]` | Yes | One or more template CN names. The names support the LDAP wildcard `*`. LDAP has no single-character wildcard, so a `?` matches literally. |
| `-ValidityPeriod` | `int` (1-9999) | Yes | The numeric value of the new validity period. |
| `-ValidityPeriodUnit` | `Years` / `Months` / `Weeks` / `Days` / `Hours` | Yes | The unit of `-ValidityPeriod`. AD uses 365 days per year and 30 days per month. |
| `-OverlapPeriod` | `int` (1-9999) | No | The numeric value of the renewal overlap period. |
| `-OverlapPeriodUnit` | `Years` / `Months` / `Weeks` / `Days` / `Hours` | No | The unit of `-OverlapPeriod`. Required when you set `-OverlapPeriod`. |
| `-Server` | `string` | No | The **domain controller** for the LDAP connection, for example `dc01.domain.com`. Do not name a CA server. |
| `-WhatIf` | switch | No | Preview the changes and make none. |
| `-Confirm` | switch | No | Prompt before each change. |

### Usage

**Preview the templates that the script would change (WhatIf):**
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

**Preview all templates that match a pattern, and capture the output:**
```powershell
$report = .\Set-ADCSTemplateValidity.ps1 -TemplateName "*" -ValidityPeriod 1 -ValidityPeriodUnit Years -WhatIf
$report | Format-Table -AutoSize
```

### Output

The script writes a `PSCustomObject` for each matched template, with these properties:

| Property | Description |
| --- | --- |
| `TemplateName` | The template CN |
| `DisplayName` | The template display name |
| `PreviousValidity` | The current validity period, human-readable |
| `NewValidity` | The requested new validity period |
| `PreviousOverlap` | The current overlap period, human-readable |
| `NewOverlap` | The requested new overlap period, or `(unchanged)` |
| `Status` | `Modified`, `Already set`, `Skipped`, or `Error: <message>` |

The script prints a color-coded summary at the end:

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

- After you modify templates, run `certutil -pulse` on each CA server, so that the CA loads the changes immediately. Otherwise AD replication and the CA cache refresh apply them later.
- The `-Server` parameter names a **domain controller**, not the CA server. Certificate templates are AD objects in the Configuration naming context of the forest.
- AD replicates the changes forest-wide from the Configuration naming context. Allow the normal AD replication time.
- The script increments `msPKI-Template-Minor-Revision` on each change, so the issuing CAs detect the update.
- In a `-WhatIf` run, the output shows the templates as `Skipped`. The script would have modified them, but the WhatIf flag prevented the change.
- The script reports each failure per template as it happens, and it continues with the remaining templates. A run in which any template failed ends with a terminating error (a non-zero exit code) after the summary. So automation cannot mistake a partial update for success.
- When you do not pass `-OverlapPeriod`, each template keeps its existing renewal overlap. When that overlap is not shorter than the new validity, the script reports the template as an error and leaves it unchanged. An example is a 30-day validity with the stock 6-week overlap. Pass a shorter `-OverlapPeriod` / `-OverlapPeriodUnit` to change both together.

### How It Works

1. The script connects to `RootDSE` to resolve the Configuration naming context.
2. It searches `CN=Certificate Templates,CN=Public Key Services,CN=Services,<ConfigNC>` with an LDAP filter for the templates that match the patterns.
3. For each match, it decodes the current `pKIExpirationPeriod` / `pKIOverlapPeriod` value. The value is 8-byte little-endian negative FILETIME ticks.
4. It compares the value with the requested value. It skips the template when the values are equal.
5. It assigns the new byte array through the property cache (`DirectoryEntry.Properties['pKIExpirationPeriod'].Value`) and calls `SetInfo()` to commit. The script does *not* use `InvokeSet()`. Its `params object[]` binding unrolls the 8-byte array into eight separate values.

<sub>[↑ Back to top](#top)</sub>

---

## Submit-CertificateRequests.ps1

The script batch-submits certificate signing requests (CSRs) as `.req` / `.csr` / `.txt` files from a folder to an ADCS CA with `certreq.exe`. It tracks the request ID of each submission in a CSV file. Later, it can retrieve the issued certificates.

### Features

- **Batch submit** all request files in a folder in one run
- **CSV tracking file** that records the request ID, submit time, status, and error message of each file
- **Resume-safe**. On a re-run, the script skips the files that the tracking CSV already lists. With `-Force`, or with an interactive y/n confirmation, you can resubmit a tracked file as a new request
- **Retrieve mode** collects every tracked request that is not yet finally resolved (`Pending`, `Unknown`, or `Error`) and retrieves the issued `.cer` files. It needs only `-CAConfig` and a tracking file, not `-InputPath` / `-CertificateTemplate`. The destination can already hold the certificate of a *different* request for the same CSR, for example a `-Force` resubmission that the CA issued first. Then the script refuses the retrieval and reports it, so an older request never replaces a newer certificate
- **Both mode** submits, then retrieves, in a single run
- **Connectivity pre-check** with `certutil -ping` before any submission
- **Per-run timestamped log file** (`CertBatch_yyyyMMdd_HHmmss_<id>.log`). The name is unique per run, and the script writes the file beside the tracking file
- **Friendly error hints** for common ADCS failures: unsupported template, denied by policy, bad subject, access denied. The script adds the hint to the raw certreq output
- **Helpful input-folder diagnostics**. When the script finds no `.req`/`.csr`/`.txt` files, it lists what *is* in the folder, or notes that the folder is empty. Then it skips Submit without a PowerShell stack trace
- **Dual-section summary**. It separates the results of *this run* from the *cumulative totals of the tracking file*, so historical errors do not look like new ones
- **Automatic `.rsp` cleanup** after retrieval. Pass `-KeepRspFile` to keep the file
- **`-WhatIf` / `-Confirm`** support
- Handles empty files, missing request IDs, and denied requests, and continues with the remaining files

### Requirements

- Windows with `certreq.exe` and `certutil.exe`. Both are standard on Windows
- Windows PowerShell 5.1 or PowerShell 7+
- Permissions to submit to the target CA and template
- Network connectivity to the CA

### Parameters

| Parameter | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| `-InputPath` | `string` | Submit/Both only | | The folder that holds the `.req` / `.csr` / `.txt` request files. `-Mode Retrieve` does not use it and does not require it. |
| `-CAConfig` | `string` | Yes | | The CA configuration string, for example `CA01.domain.com\Contoso Issuing CA 1`. Always required. |
| `-CertificateTemplate` | `string` | Submit/Both only | | The certificate template name: the CN, not the display name. `-Mode Retrieve` does not use it and does not require it. |
| `-TrackingFile` | `string` | No | `.\CertTracking.csv` | The CSV file that tracks the request IDs and statuses across runs. |
| `-OutputFolder` | `string` | No | `.\Certificates` | The folder where the script saves the issued `.cer` files, one per request, named after the request file. In `Retrieve` mode the script writes each row to the path that it recorded at submit time. When you pass `-OutputFolder` explicitly, the script redirects the retrieved files to it and updates the tracking row. |
| `-Mode` | `Submit` / `Retrieve` / `Both` | No | `Submit` | `Submit` submits new requests only. `Retrieve` retrieves the certificates of previously pending requests. `Both` does both. |
| `-KeepRspFile` | switch | No | | By default the script deletes the `.rsp` file that `certreq` writes next to each retrieved `.cer`. Pass this switch to keep the file. |
| `-Force` | switch | No | | Resubmit the request files that already have a tracked RequestID, without a prompt. Without `-Force`, the script asks y/n for each already-submitted file. The default answer is No, which skips the file. |
| `-AllowUnprotectedOutputFolder` | switch | No | | By default the script refuses to deliver into or through a folder that an untrusted principal owns, or can delete, rename or write to. Such a user could replace the folder with a junction during a delivery, or turn an empty folder into a junction with write rights alone. The script also refuses a delivery folder whose ACL would give an untrusted principal write, append, delete or write-attributes rights on the *files* created inside it. The staging file and the delivered certificate inherit such an entry. CREATOR OWNER entries resolve to the running account and are safe. Pass this switch to accept the risk. The script then only warns about these conditions. |
| `-TrustedOutputPrincipal` | `string[]` | No | | Additional principals, as SIDs or `DOMAIN\Group` names, that may own the delivery folders, or hold delete, rename or write rights on them. This includes file-inheritable write rights in the delivery folder itself. The script always trusts SYSTEM, Administrators, TrustedInstaller, the running account, and its Domain Admins and Enterprise Admins. |
| `-WhatIf` | switch | No | | Preview. The script submits and retrieves nothing. |
| `-Confirm` | switch | No | | Prompt before each action. |

> **CAUTION** - `-AllowUnprotectedOutputFolder` turns the folder refusals into warnings. A user who can replace the folder, or write its files, can then redirect the certificate write or alter the certificate before or after delivery. Use it only for a shared drop folder whose ACL you cannot tighten, and accept that risk.
> `-Force` resubmits a request file that the tracking file already records, without a prompt. The CA can then issue a second certificate for the same CSR. Use it only when you want a new request for that file.

### Usage

**Submit all CSRs in a folder:**
```powershell
.\Submit-CertificateRequests.ps1 `
    -InputPath "C:\CSRs" `
    -CAConfig "CA01.domain.com\Contoso Issuing CA 1" `
    -CertificateTemplate "WebServer" `
    -Mode Submit
```

**Retrieve the issued certificates of previously unresolved requests.** This mode needs only `-CAConfig` and the tracking file. It does not use `-InputPath` / `-CertificateTemplate`:
```powershell
.\Submit-CertificateRequests.ps1 `
    -CAConfig "CA01.domain.com\Contoso Issuing CA 1" `
    -Mode Retrieve
```

The script writes each retrieved `.cer` file to the path that it recorded for the row at submit time. To write the files to another folder, pass `-OutputFolder` explicitly on the `Retrieve` run:
```powershell
.\Submit-CertificateRequests.ps1 `
    -CAConfig "CA01.domain.com\Contoso Issuing CA 1" `
    -Mode Retrieve `
    -OutputFolder "C:\Certificates\Issued"
```

**Submit and retrieve in one run:**
```powershell
.\Submit-CertificateRequests.ps1 `
    -InputPath "C:\CSRs" `
    -CAConfig "CA01.domain.com\Contoso Issuing CA 1" `
    -CertificateTemplate "WebServer" `
    -Mode Both
```

**Preview what the script would submit:**
```powershell
.\Submit-CertificateRequests.ps1 `
    -InputPath "C:\CSRs" `
    -CAConfig "CA01.domain.com\Contoso Issuing CA 1" `
    -CertificateTemplate "WebServer" `
    -Mode Submit -WhatIf
```

**Resubmit already-tracked files without a prompt, and keep the `.rsp` files after retrieval:**
```powershell
.\Submit-CertificateRequests.ps1 `
    -InputPath "C:\CSRs" `
    -CAConfig "CA01.domain.com\Contoso Issuing CA 1" `
    -CertificateTemplate "WebServer" `
    -Mode Both -Force -KeepRspFile
```

### Friendly Error Hints

When `certreq` rejects a request, the script writes the raw output as a warning *and* appends a hint based on the error code. The script recognizes these patterns:

| Detected pattern | Hint |
| --- | --- |
| `0x80094800` / `CERTSRV_E_UNSUPPORTED_CERT_TYPE` | The template name is misspelled, the template is not published on this CA, or you included the `CertificateTemplate:` prefix by mistake. The hint suggests `certutil -config "<CA>" -CATemplates` to list the valid names. |
| `0x80094012` / `CERTSRV_E_TEMPLATE_DENIED` | The calling account lacks the Enroll permission on the template, or the template requires approval or a signature. |
| `0x80094004` / `CERTSRV_E_BAD_REQUESTSUBJECT` | The CSR subject or SAN does not match what the template requires. |
| `0x80070005` / access denied | The calling account lacks the "Request Certificates" right on the CA. |

### Run Summary

The script splits the end-of-run summary into two sections. So you can tell the results of *this run* apart from the *cumulative state* of the tracking file:

```
--- Summary: Retrieve ---
  Issued: 7
  Total processed this run: 7
--- Tracking file total (all history) ---
  Error: 1
  Issued: 27
Tracking file: .\CertTracking.csv
```

In this example, the single `Error: 1` is a historical row from an earlier session. It did not happen in the current run.

### Tracking CSV Schema

| Column | Description |
| --- | --- |
| `RequestFile` | The full path of the source `.req`/`.csr`/`.txt` file |
| `RequestID` | The numeric request ID that the CA assigned |
| `SubmitTime` | The ISO-8601 submission timestamp |
| `Status` | `Issued`, `Pending`, `Denied`, `Error`, `Unknown`, or `Undelivered`. `Undelivered` means that the CA issued, but the script could not deliver the `.cer` to its destination. The row counts as submitted, and `Retrieve` retrieves it again when a RequestID is present. An `Unknown` row **without** a RequestID means that certreq reported success, but the script could read neither a RequestID nor a certificate. That row also counts as submitted, and `Retrieve` lists it for manual reconciliation. |
| `OutputCertFile` | The full path where the script saves the issued `.cer` |
| `LastCheckTime` | The ISO-8601 timestamp of the last status check |
| `ErrorMessage` | The error output of `certreq` when the submission or retrieval failed |
| `CAConfig` | The CA that received the request. `-Mode Retrieve` refuses rows that went to a different CA, because RequestIDs are per CA. The script assumes that rows from older files belong to the `-CAConfig` of the run, and it stamps them. |

### Notes

- The tracking CSV is the source of truth for the resume behavior. When you delete it, the script resubmits all files, and the CA can issue duplicates.
- **One run per tracking file.** The script holds an exclusive lock file (`<TrackingFile>.lock`) for its whole run, and removes it when the run ends. It refuses to start while another run holds the lock, on this or any other machine, through any alias of the path. Two concurrent writers would drop each other's rows without an error.
  - The script canonicalizes the name of the tracking file first. An 8.3 short name resolves to the long name, so both spellings share one lock. The script refuses a hard-linked or symlinked tracking file.
  - The script creates the lock and the per-run log only after the tracking folder has passed the same chain check that every delivery gets. It refuses a reparse point that already exists at either name, even a dangling link. It opens both files with `CreateNew`, so it never writes through, or deletes through, a link that another user placed there.
  - A `-WhatIf` run takes no lock and writes no log.
- **Staging and delivery.** `certreq` always writes into a private, randomly named staging file **inside the destination folder**. It never writes into `%TEMP%`. The inherited file ACL of `%TEMP%` could let another account tamper with the certificate, and that ACL would stay on the delivered file. The script touches the destination only after a successful write, so a pending, denied or failed request never disturbs it. The script resets the delivered file to inherit the ACL of the folder.
  - The `OutputCertFile` of a row is an identifier inside a boundary that the operator chose. It is not an authority. It must be a rooted `.cer` path in an existing folder beneath the folder of the tracking file, or beneath the `-OutputFolder` of the run. The script checks the path canonically, with no junction or symbolic link in between, and checks it again right before delivery. `RequestID` must be numeric. The script skips a row that fails these checks and reports an error; the row never reaches certreq.
  - The destination itself must be a plain file or absent. The script refuses a folder or a link named `x.cer`. After delivery, the script verifies that a plain file resulted.
  - When the script cannot deliver an issued certificate, for example a locked destination or a denied rename, the row becomes `Undelivered`. The row keeps its RequestID, and the certificate stays in its staging file beside the destination, with the path in `ErrorMessage`. The row counts as submitted on later runs, and the script never resubmits it automatically. `Retrieve` retrieves it again. The script reports an `Undelivered` row with no RequestID for manual reconciliation.
- **Output folders must not be replaceable by untrusted users.** The script checks every folder from the destination up to the root of the volume or share. It **refuses** to deliver when any folder in the chain meets one of these conditions:
  - The folder is a reparse point: a junction, a symbolic link, or a mount point.
  - An untrusted principal owns the folder.
  - An untrusted principal can delete, rename or write to the folder. With delete or rename rights, such a user can replace the folder with a junction between the path check and the privileged delivery. Write-data or write-attributes access is all that `FSCTL_SET_REPARSE_POINT` needs. So "create files" rights let the user turn an *empty* folder, such as a freshly created output folder, into a junction in place.
  - The script cannot read the security descriptor of the folder.
- The trusted principals are SYSTEM, `BUILTIN\Administrators`, TrustedInstaller, the running account, its Domain Admins and Enterprise Admins, and every principal named in `-TrustedOutputPrincipal`. `-AllowUnprotectedOutputFolder` turns the refusals into warnings.
- On the folder chain, the script permits an untrusted principal only two things: the create-*subfolder* right by itself, and inherit-only ACEs. The `C:\` root grants Users that right on itself. Neither can set a reparse point on the folder. These rights are safe for these reasons:
  - The reparse-point checks refuse a junction that already exists.
  - The owner check refuses a subfolder that an attacker created.
  - A folder or junction that an attacker places under the exact destination name between the check and the delivery receives no file. Every delivery is a no-overwrite rename (`File.Move`). The rename fails when anything occupies the name, instead of moving the file into it.
- The `.rsp` file written with `-KeepRspFile` and the replacement of the tracking file follow the same rule.
- The script also judges the delivery folder itself on the rights that its ACL gives to the *files* created inside it. It checks every entry that propagates to files, ObjectInherit, inherit-only or not, against the file write class. That class is write-data, append-data, delete, change permissions, take ownership, and write-attributes. Write-attributes alone is enough to set a reparse point on a file.
  - The script refuses a grant to an untrusted principal, because the staging file of certreq and the delivered certificate inherit it. Such a user could otherwise alter the bytes of the certificate, or turn the delivered file into a reparse point. That can happen before or after delivery, while every folder check passes.
- The script resolves the inheritance placeholders the way the file system resolves them. CREATOR OWNER and OWNER RIGHTS become the creating account, which is the trusted running account, because certreq creates the file. So the inherit-only CREATOR OWNER entry that the `C:\` root gives to every unprotected folder is safe. CREATOR GROUP becomes the primary group of the running account, which can be as broad as Domain Users. It stays untrusted unless you name `S-1-3-1` in `-TrustedOutputPrincipal`.
- **`Issued`.** The script decides `Issued` from the exit code of certreq plus the presence of the certificate that certreq wrote. `certreq` writes the certificate into a fresh temp file, so it is unambiguously the output of this run, and the check is language-independent. On success the script delivers the certificate to the destination.
  - A file can already be at the destination, from a `-Force` resubmit or from a retry of an unresolved row. The script moves that file aside as `<name>.superseded-<UTC stamp>.cer`. It never deletes that file. It removes that copy again only when the fresh certificate is byte-identical.
  - When a retrieval reports Issued but produces no file, the script records `Error` and retries on the next `Retrieve`.
  - The script parses the RequestID and the `Pending` / `Denied` dispositions from the console text of certreq, which Windows localizes. On a non-English system they can parse as `Unknown`. The script flags a missing RequestID in `ErrorMessage`, so you can fill it in from the CA database.
- The script allocates the certificate file names before it submits anything. The names must be unique within the batch and against the destinations already recorded for other request files. So `prod.req`, `prod.csr` and `prod.req.txt` no longer map two requests onto one `.cer`. A clash stops the run. When certreq reports a submission as successful, but its reply yields neither a RequestID nor a certificate, the script records `Unknown`. That row counts as submitted, and the script never resubmits it automatically.
- A run in which any request failed or needs attention ends with a terminating error after the summary. This covers `Error`, `Denied`, `Undelivered`, `Unknown`, and a Retrieve row skipped as invalid. So automation that uses the exit code to decide does not treat a partial batch as success. `Pending` is not a failure.
- The values that reach the certreq command line must not contain double quotes or control characters. These values are `-CAConfig`, `-CertificateTemplate`, the `.cer` paths and `RequestID`. The script refuses such a value before certreq runs.
- `Pending` usually means that the CA requires manager approval. Run the script again in `Retrieve` mode after the approval to retrieve the issued certificate.
- The script names the issued `.cer` files after the source request file, for example `server1.req` -> `server1.cer`.
- The script skips an empty request file with a warning.
- **Log file.** The script creates a timestamped log file (`CertBatch_yyyyMMdd_HHmmss_<id>.log`) beside the tracking file for each run. The suffix is random, so runs that share a folder never collide. The script never writes the log in the working directory, which it never validates.
- **A missing `-OutputFolder`.** The script creates a missing `-OutputFolder` one component at a time, with an atomic fail-on-existing create. So the script refuses a junction that an attacker places in the missing part of the path after the folder check. It does not create through the junction. The script judges every parent *before* it creates a folder in it. That includes what the new folder would inherit, because an inheritable grant on the parent becomes effective on the new folder the instant it exists.
  - Anywhere the script validates a folder, it refuses a component that ends in a space or a period. Windows resolves such a name differently as a leaf and as a parent.
  - The script refuses, rather than creates, a folder that does not exist yet when its parent would give an untrusted principal rights on it. The `C:\` root does that, with an inheritable `CreateFiles` grant to Users. So pre-create a folder directly under `C:\` with a protected ACL, or pass `-AllowUnprotectedOutputFolder`.
- The script captures the console output of certreq in files beside the log, never under `%TEMP%`. A SYSTEM or service run shares `%TEMP%` with every local user.

<sub>[↑ Back to top](#top)</sub>

---

## Sync-ADCSTemplate.ps1

*The former name of the script was `Sync-KerberosAuthTemplate.ps1`. The new name reflects that it copies **any** certificate template. The Kerberos Authentication template remains only its default.*

The script copies a certificate template between AD forests, by default the built-in **Kerberos Authentication** template. It reads the functional attributes of the template from the source forest. It recreates the template in the target forest with `New-ADObject`, with all directory access over Active Directory Web Services (ADWS). The script derives the container DN, the `objectCategory`, and optionally a fresh template OID from the **target** forest. Then it applies a composable enrollment ACL. Two interchangeable flows share the same pipeline:

- **File-based** (`-Mode Export` / `-Mode Import`). The script serializes the template to a JSON file in the source forest, and imports the file in the target forest. This flow suits air-gapped or change-controlled environments.
- **Direct** (`-Mode Sync`). The script reads from a domain controller (DC) in the source forest (`-SourceServer`). It writes to the target side in the same run, with no intermediate file. With a trust, the current identity can usually do both sides. Without a trust, pass `-SourceCredential`, `-Credential`, or both. Explicit credentials against explicitly named DCs need no trust at all.

You can mix the flows. A JSON file exported earlier imports into a remote forest with `-Mode Import -Server <target DC> -Credential (...)`. `-Mode Export` also accepts `-Server` / `-Credential` to read from a remote source forest.

### Why you need this

There is no supported UI path to move a template definition between forests. The classic `certutil -dsTemplate` / `-dsAddTemplate` round-trip cannot rename the template, always keeps the source OID, and leaves the ACL to you. This script does the whole job, including two scenarios that the certutil approach cannot handle:

- **Target forest without AD CS.** The template objects live in the `Certificate Templates` container, which every forest has and which does not depend on a CA. So you can publish a template into a forest that never had AD CS. For example, an external CA can then read the template and its ACL as the source of enrollment authorization. Use with **EJBCA** has some specifics of its own. See [Using with EJBCA](#using-the-template-with-ejbca).
- **Renamed copies with controlled identity.** Import the template under a new cn or display name. Choose whether the script keeps the source OID or generates a new one.

The ACL is the part that an external CA consumes, so the script treats it with care. The ACL bases (`-AclBase`) either replace or extend the schema-default DACL. Per-principal additions (`-EnrollPrincipals`) add enrollees and template admins. Principal resolution fails closed: the script refuses a name that matches BOTH a well-known token and a directory object with a *different* SID. It also refuses a name that matches more than one object. The script does not guess; disambiguate with a SID or a `DOMAIN\` prefix.

The own names of built-in groups, such as `Domain Admins`, resolve normally, because both readings give the same SID. Well-known SID and RID tokens are language-invariant. So the script also works on **non-English forests** where group names are localized.

### Features

- **Direct attribute copy**. No `certutil`, no text-dump parsing. JSON file or direct forest-to-forest, with everything over ADWS (TCP 9389) only
- **`-Mode Sync`**: a one-run direct sync between forests, with no intermediate file. A same-forest guard stops a forgotten `-Server` from writing the copy back into the source forest without an error
- **Per-side credentials**: `-Credential` for the target, `-SourceCredential` for the source. The script works across a trust with the current identity, or with **no trust at all** with explicit credentials
- **Works against a target forest with no AD CS**. The default OID mode needs no PKI OID root
- **Rename on import** (`-NewTemplateName` / `-NewDisplayName`)
- **Four OID modes** (`-OidHandling`). `Preserve` is the default and keeps the source OID. `Generate` creates the OID under the real OID root of the target forest. `GenerateFromRoot` creates it under a base that you supply. `GenerateRandom` creates it under a synthesized base. Every mode registers the companion OID "display" object, so Windows resolves the OID to the template name
- **Composable ACL**: `-AclBase Standard | Schema | SchemaPlusStandard | PrincipalsOnly` plus additive `-EnrollPrincipals`, which maps a principal to Read, Write, Enroll, Autoenroll, or FullControl. A `user@domain` key resolves **only** as a UPN. The script refuses a different object that has that string as its sAMAccountName. A sAMAccountName can contain `@`, so an account that an attacker created could otherwise capture the grant
- **`-UpgradeCompatibility`**: raise the imported copy to the newest compatibility as the script creates it. That is CA *Windows Server 2016* and recipient *Windows 10 / Windows Server 2016*: schema v2/v3 becomes v4, with the matching private-key-flag bits. The script leaves v1 built-ins and already-v4 templates as they are. The script never modifies the source; it upgrades only the target copy. The script sets the legacy-provider bit (`CT_FLAG_USE_LEGACY_PROVIDER`) only for a schema-2 source with a provider list. A schema-3 source keeps its own bit, so a v3 template that lists a KSP stays CNG
- **`-Mode Validate`**: proves the round-trip fidelity of **both** pipelines, the JSON file flow *and* the direct in-memory flow that Sync uses. The script imports a throwaway copy per pipeline, diffs every PKI attribute of the source, byte-array attributes included, and then cleans up. A mismatch fails the run with a non-zero exit code, so automation can use Validate to decide. A throwaway copy that the script cannot read back after its confirmed creation also fails the run, with a terminating error. So a check on the exit code never accepts an unvalidated copy
- **Fail-fast pre-flights**: the script refuses a duplicate cn, a duplicate template OID, a missing container, and an unresolvable principal, all before it creates anything. It validates every known attribute of an import for type, shape and range. It **refuses** a malformed or tampered export; it never drops or coerces a value without an error. The *target* forest can link an issuance policy OID to a group through Authentication Mechanism Assurance (AMA). Then the script refuses the import, unless you pass `-AllowLinkedIssuancePolicy`
- **`-WhatIf` / `-Confirm`** support end to end. The UTF-8 BOM file format is safe across PowerShell 5.1 and 7

### Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- The **RSAT ActiveDirectory PowerShell module**. This script is the exception to the module-free approach of the repository. The module makes typed attribute writes, `-Server` pinning, and a clean rollback practical
- No CA role, no RSAT AD CS Tools, and no reachable CA in either forest
- **Export**, and the read side of Sync: read access to the template. Authenticated Users has this access by default
- **Import/Validate**, and the write side of Sync: Enterprise Admin, or delegated write access to the `CN=Certificate Templates` **and** `CN=OID` containers in the Configuration naming context. Every OID mode, including the default `Preserve`, registers a companion OID display object when the OID has none yet. Only a forest without the `CN=OID` container skips it
- **Sync**: ADWS (TCP 9389) reachability to a DC in *each* forest from the machine that runs the script. Also a trust between the forests, **or** explicit `-SourceCredential` / `-Credential`. All modes use ADWS only; no LDAP (389) access is needed. The script always uses the `-Server` / `-SourceServer` values **verbatim**. It never substitutes an endpoint that you did not type.
  - Name **one DC**. A domain name (DNS or NetBIOS) locates a different DC per connection, and the script gives a loud warning. The create, read-back and ACL steps could hit different replicas. A lagging read-back fails the run rather than leave a template incorrectly secured.

### Parameters

| Parameter | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| `-Mode` | `Export` / `Import` / `Sync` / `Validate` | Yes | | `Export` writes the JSON in the source forest. `Import` recreates the template and the ACL in the target forest. `Sync` does both directly forest-to-forest, with no file. `Validate` round-trips a template into a throwaway copy and diffs it. |
| `-Path` | `string` | Export/Import | | The JSON file to write (Export) or read (Import). Optional for Validate, which uses a temp file otherwise. Sync does not use it. |
| `-TemplateName` | `string` | No | `KerberosAuthentication` | Export/Validate/Sync: the cn, that is the internal name, of the source template. |
| `-NewTemplateName` | `string` | No | source's `name` | Import/Sync: the new cn in the target forest. The script allows letters, including non-ASCII letters, digits, spaces that are not at an edge, and `._-()`. It refuses the DN metacharacters `, + = " \ ; < >`, the LDAP wildcard `*`, and `/ #`. |
| `-NewDisplayName` | `string` | No | source's `displayName` | Import/Sync: the new display name in the target forest. |
| `-StripIdentity` | switch | No | | Export: omit `name` / `displayName` from the file. The import then requires explicit names. |
| `-StripOid` | switch | No | | Export: omit the source OID. The import then needs a Generate mode. |
| `-OidHandling` | `Preserve` / `Generate` / `GenerateFromRoot` / `GenerateRandom` | No | `Preserve` | Import/Sync: how the script chooses the template OID. See Features. |
| `-OidRoot` | `string` | With `GenerateFromRoot` | | The base OID to create the new OID under, for example `1.3.6.1.4.1.311.21.8.<arcs>`. The script refuses it with any other `-OidHandling`; it never ignores it without an error. |
| `-AclBase` | `Standard` / `Schema` / `SchemaPlusStandard` / `PrincipalsOnly` | No | `Standard` | The ACL foundation. `Standard` writes the stock Kerberos Authentication ACL and replaces the schema default. `Schema` keeps the schema default. `SchemaPlusStandard` keeps it and adds the standard set. `PrincipalsOnly` writes exactly `-EnrollPrincipals`. The default, `Standard`, is specific to Kerberos Authentication and oriented to DCs. The script warns when `Standard` applies only by default, that is without an explicit `-AclBase`, to a template that is not named like a Kerberos Authentication copy. |
| `-EnrollPrincipals` | `hashtable` | With `PrincipalsOnly` | | A map of principal to rights, added on top of the base, for example `@{ 'DomainControllers'='Enroll','Autoenroll'; 'PKI-Admins'='FullControl' }`. A key is a SID, a sAMAccountName, a UPN (user@domain), or a well-known token. The script refuses a bare string that matches both a token and an object with a different SID, or that matches more than one object. Disambiguate with a SID or a DOMAIN\ prefix. A `user@domain` key resolves only as a UPN. The script refuses a different object that has that string as its sAMAccountName. |
| `-SkipAcl` | switch | No | | Import/Sync: skip the permission step entirely. Not allowed together with `-EnrollPrincipals` or an explicit `-AclBase`. |
| `-UpgradeCompatibility` | switch | No | | Import/Sync: raise the created copy to the latest compatibility: CA Windows Server 2016, recipient Windows 10 / Windows Server 2016, schema v2/v3 to v4. The script imports v1 templates and already-v4 templates unchanged, with a note. It sets the legacy-provider bit only for a schema-2 source with a provider list. A schema-3 source keeps its own bit. |
| `-AllowLinkedIssuancePolicy` | switch | No | | Import/Sync: accept an issuance policy OID that the *target* forest already links to a group through Authentication Mechanism Assurance. The OID is in `msPKI-Certificate-Policy`, and the CA stamps it into the issued certificate. The script refuses such an OID by default. Certificates from the copy would grant the membership of that group at logon to everyone that the enrollment ACL of the copy admits. |
| `-KeepArtifacts` | switch | No | | Validate: keep the two throwaway templates and the export file for inspection. The script never creates a companion OID object for them. |
| `-Server` | `string` | No | auto-discover | Pin all operations, the target-side operations for Sync, to a specific writable DC. Point it at a DC of another forest to operate there. Required together with `-Credential`. |
| `-Credential` | `pscredential` | No | current identity | The credentials that the script uses against `-Server`: the target side for Import/Sync/Validate, the *source* side for Export. With `-Server`, you can operate on a forest that you are not logged on to. No trust is needed. |
| `-SourceServer` | `string` | Sync | | Sync: the DC, or domain name, in the **source** forest that the script reads the template from. |
| `-SourceCredential` | `pscredential` | No | current identity | Sync: the credentials for the source-side read. |
| `-WhatIf` / `-Confirm` | switch | No | | Preview / prompt. `-WhatIf` shows the planned template, OID object, and exact ACL grants. |

> **CAUTION** - `-SkipAcl` skips the permission step entirely. The copy then keeps the schema-default ACL, in which nobody has Enroll, so an external CA that reads the ACL authorizes no one. Use it only when you set the ACL by another means.
> `-AllowLinkedIssuancePolicy` accepts an issuance policy OID that the target forest links to a group through Authentication Mechanism Assurance (AMA). Certificates from the copy then grant the membership of that group at logon to everyone that the enrollment ACL admits. Use it only when you intend that grant.

### Usage

**In the source forest, export:**
```powershell
.\Sync-ADCSTemplate.ps1 -Mode Export -Path .\KerberosAuth.json
```

**In the target forest, which needs no AD CS, import with the standard ACL:**
```powershell
.\Sync-ADCSTemplate.ps1 -Mode Import -Path .\KerberosAuth.json
```

**Import as a renamed copy with a new, forest-independent OID:**
```powershell
.\Sync-ADCSTemplate.ps1 -Mode Import -Path .\KerberosAuth.json -OidHandling GenerateRandom `
    -NewTemplateName "YY-KerberosAuthentication" -NewDisplayName "YY-Kerberos Authentication"
```

**Standard ACL plus a template-admin group, which is what an external CA reads:**
```powershell
.\Sync-ADCSTemplate.ps1 -Mode Import -Path .\KerberosAuth.json -EnrollPrincipals @{
    'CONTOSO\PKI-Admins' = 'FullControl'
}
```

**Direct sync. Run it in the target forest, and read from the source forest over the trust, with no file:**
```powershell
.\Sync-ADCSTemplate.ps1 -Mode Sync -SourceServer dc01.source.example `
    -NewTemplateName "YY-KerberosAuthentication" -NewDisplayName "YY-Kerberos Authentication"
```

**Direct sync from a third machine with explicit credentials on both sides. No trust is needed:**
```powershell
.\Sync-ADCSTemplate.ps1 -Mode Sync `
    -SourceServer dc01.a.example -SourceCredential (Get-Credential A\template.reader) `
    -Server dc01.b.example -Credential (Get-Credential B\ent.admin)
```

**Mixed flow. Import a previously exported JSON file directly into another forest:**
```powershell
.\Sync-ADCSTemplate.ps1 -Mode Import -Path .\KerberosAuth.json `
    -Server dc01.b.example -Credential (Get-Credential B\ent.admin)
```

**Prove the round-trip fidelity in the source forest first. This creates and removes throwaway copies, and checks both the file pipeline and the direct Sync pipeline:**
```powershell
.\Sync-ADCSTemplate.ps1 -Mode Validate -TemplateName "KerberosAuthentication"
```

### Using the template with EJBCA

**EJBCA** can use an AD certificate template as its source of enrollment authorization through Microsoft auto-enrollment (**MSAE**). It reads the template object and its ACL directly from AD, so a forest that never had AD CS can still host the template. The sections above cover the general cross-forest mechanics. The points below are specific to EJBCA.

**Start here, because you do not always need to build a template.** This repository ships **ready-to-import, EJBCA-ready** templates, with the full-DN Subject already applied, under [`Templates/EJBCA/`](./Templates/README.md). They are the four defaults whose Subject would otherwise be empty, including **Kerberos Authentication**, the modern DC template that the field-validated example below uses. When you import one directly, you can skip the preparation steps in this section entirely. You install or edit nothing locally with `certutil`. For domain controllers:

```powershell
.\Sync-ADCSTemplate.ps1 -Mode Import -Path .\Templates\EJBCA\KerberosAuthentication-EJBCA.json -OidHandling GenerateRandom
```

The rest of this section explains how to prepare your *own* template when the shipped set does not cover your need. It also explains why the EJBCA variants are built the way they are.

> **Field-validated end to end:** we exported a Kerberos Authentication template with the full-DN Subject enabled, and imported it into a separate forest with **no CA**. We delivered it to clients with [`Add-CertificateEnrollmentPolicyServerToGpo`](#add-certificateenrollmentpolicyservertogpops1). Each DC in scope received a Kerberos Authentication certificate that EJBCA issued. We confirmed that the same template with an **empty Subject** fails, which is exactly the first point below.

- **The Subject must not be empty.** The built-in **Kerberos Authentication** template issues certificates with an **empty Subject**; the identity lives entirely in the SAN as DNS entries. EJBCA cannot use a template that produces no Subject. Fix it on the **source** template, **before** you export or sync. In the Certificate Templates MMC, open *Properties*, then **Subject Name**. Select *Build from this Active Directory information*, and set **Subject name format** to **Fully distinguished name**.
  - That sets the `CT_FLAG_SUBJECT_REQUIRE_DIRECTORY_PATH` bit (`0x80000000`) in `msPKI-Certificate-Name-Flag`. The script copies that attribute verbatim, so the setting goes into the synced copy. But the setting must be on the source template first.
  - **This is not unique to Kerberos Authentication.** Every template that would otherwise issue an **empty** Subject needs the same change before EJBCA can use it.
  - **Full DN or common name:** the hard rule is only that the Subject must not be empty. Your EJBCA end-entity profile decides whether a bare **Common Name** is enough, or whether it requires the **full distinguished name**. "Fully distinguished name" is the safe choice that always populates the Subject.
- **The ACL is what EJBCA consumes** for enrollment authorization. Build it with `-AclBase` / `-EnrollPrincipals`. See [Features](#features-2) and [Parameters](#parameters-2).
- **Publishing is separate.** Import/Sync does **not** publish the template to any CA. You configure the template mapping in EJBCA itself.
- **No default templates to export from?** First, you do not always need to produce any. The [`Templates/`](./Templates/README.md) folder of the repository ships every default, and the EJBCA-ready variants, as importable JSON. An import directly from there is usually the fastest path. You can also generate the templates yourself. The default templates, Kerberos Authentication included, normally arrive in AD when you install the first Enterprise CA.
  - So a forest that never had a CA does not have them. But the templates are plain AD objects, so no CA is needed to hold them. `certutil -InstallDefaultTemplates` writes the standard set into AD, and it **works with no AD CS role installed at all** (verified). Run it as an **Enterprise Admin**. Add `-dc <DCName>` to target a specific DC. That gives you a stock template to edit: turn on the full-DN Subject above, and then export or sync.
- **Then point the clients at EJBCA.** The published template is only half the job. Windows clients still need a setting that tells them to enroll against the EJBCA policy server. [`Add-CertificateEnrollmentPolicyServerOffline`](#add-certificateenrollmentpolicyserverofflineps1) writes that setting per machine. [`Add-CertificateEnrollmentPolicyServerToGpo`](#add-certificateenrollmentpolicyservertogpops1) writes it fleet-wide through a GPO. Both compute the CEP registry values offline from the same EJBCA alias.

### Notes

- Import/Sync **refuse** to overwrite an existing template with the same cn, and refuse to duplicate an existing template OID. Delete or rename the existing template instead.
- The companion `msPKI-Enterprise-Oid` "display" object is the only object in the OID container that is involved. Import/Sync register it automatically. The script deliberately does **not** copy the own OID object of the source forest verbatim, because its other attributes are forest-specific.
- **The script does not copy Authentication Mechanism Assurance (AMA) links.** `msDS-OIDToGroupLink` points at a group DN in the *source* forest. If you use AMA, recreate the link in the target forest against a local universal group. The script **does** check the reverse case: the copy can have an issuance policy OID that the *target* forest already links to a group. Then the script refuses the import unless you pass `-AllowLinkedIssuancePolicy`. Certificates from the copy would grant the membership of that group at logon.
- The script copies v1 templates too, with an advisory warning. The object round-trips faithfully; this is live-verified. But Windows fixes the v1 semantics in code: name-matched, not editable, no autoenrollment. So import a v1 copy under its **original** name. Non-Windows consumers that read the object or the ACL directly are unaffected. The script intentionally does not export the ACL; the ACL is forest-specific, and the script rebuilds it from `-AclBase` / `-EnrollPrincipals` on import.
- Import does **not** publish the template to any CA. With a Microsoft CA, that remains a separate "Certificate Templates to Issue" step. For EJBCA, you configure the template mapping there. See [Using with EJBCA](#using-the-template-with-ejbca).
- The script writes the JSON with a UTF-8 BOM, so localized or accented display names survive a round-trip from PowerShell 7 to 5.1.
- Run `-Mode Validate` in a lab or in the source forest before the first production import or sync. It exercises both the export-to-import file pipeline and the direct in-memory pipeline that Sync uses. For each pipeline it diffs every PKI attribute that the source has.
- The script refuses a parameter that the mode does not consume, before it does anything, instead of ignoring it without an error. Examples are `-StripOid` with `-Mode Import`, `-OidRoot` without `GenerateFromRoot`, and `-AclBase` with `-SkipAcl`.
- The built-in type lists do not know every PKI attribute. A genuine schema extension linked to the template class is such an attribute. The script types it automatically from the schema of the *target* forest. The script drops an attribute with a warning in three cases:
  - The schema cannot type it.
  - The schema does not permit it on the template class.
  - The schema types it single-valued, while the source value is empty or multi-valued.
- The v3/v4 CNG algorithm settings are not separate attributes, so that typing does not apply to them. They are packed inside `msPKI-RA-Application-Policies`.
- A v3/v4 template can have a **private-key SDDL** (`msPKI-Key-Security-Descriptor`) packed inside `msPKI-RA-Application-Policies`. The script copies it verbatim. Every domain SID in it is a source-forest SID, so the script warns, and you can review the key ACL in the target forest.

### Tests

A Pester suite ([`Tests/Sync-ADCSTemplate.Tests.ps1`](./Tests/Sync-ADCSTemplate.Tests.ps1)) covers the script in four tiers. It needs Pester **5+** (`Install-Module Pester -MinimumVersion 5.0 -Scope CurrentUser`):

| Tier (`-Tag`) | Needs | Changes anything? |
| --- | --- | --- |
| `Unit` | nothing. The tier tests pure helpers, extracted from the script by AST, so the real code runs | no |
| `Static` | nothing. The tier parses the script and checks the comment-based help | no |
| `Guard` | the RSAT ActiveDirectory module. No reachable DC is needed | no |
| `Lab` | a lab forest. Opt in with `-RunLab` | **creates and removes AD objects** |

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

The `Lab` tier is precise. Every object that it creates has a unique per-run `PESTER-<hex>` prefix. The tier tracks each object by its exact DN and removes it in teardown, with a prefix-scoped safety-net sweep as a backstop. It never touches pre-existing objects. Read-backs poll with retry, so a target forest that lags briefly over ADWS after a write does not cause false failures.

<sub>[↑ Back to top](#top)</sub>

---

## Add-CertificateEnrollmentPolicyServerOffline.ps1

The script registers or removes a **Certificate Enrollment Policy (CEP)** server in the registry entirely **offline**. It makes no "Validate Server" round-trip and no contact with the policy server at all. It writes the same values that the *"Certificate Services Client – Certificate Enrollment Policy"* dialog produces, but computes everything locally. It computes the SHA-1 subkey name from the URL, and, for EJBCA/MSAE, the PolicyID from the Policy Name of the alias.

### Why you need this

The built-in dialog, and the `X509Enrollment` COM path behind it, must reach the MS-XCEP `GetPolicies` endpoint of the policy server before it saves anything. That is exactly what you cannot do in four cases. You stage a machine before the PKI is reachable. You build a golden image. You work in an air-gapped or change-controlled environment. You script an identical configuration across many machines.

This script derives every value locally and writes it directly. So the enrollment-policy configuration becomes a repeatable, unattended step.

The script also handles the parts that the dialog hides. On the Group Policy hives it keeps the built-in **AD enrollment policy row**. Without that row, GP-based CEP removes the Active Directory enrollment policy without a message, and autoenrollment against AD-published templates stops. The script also manages the DISABLE bits of the root **Flags** value of `PolicyServers`.

### Features

- **Fully offline derivation**. The subkey is the SHA-1 over the UTF-16LE bytes of the invariant-lowercased URL. The EJBCA MSAE PolicyID is the Java `String.hashCode()` of the Policy Name. Pass `-PolicyId` for a CEP server that returns a GUID, such as Microsoft's
- **Four target locations**. `LocalMachine` / `LocalUser` are the user-configured stores that `certlm.msc` / `certmgr.msc` manage. `GPMachine` / `GPUser` are the Group Policy hives
- **AD enrollment policy row kept** on the GP locations, so GP-based CEP does not remove the AD default policy. Opt out with `-SkipADPolicy`
- **Complete-row checks**. A pre-existing AD row or CEP entry satisfies the AD-row prerequisite, the `(Default)` marker and `-ReplaceExisting` only when it is *usable*. Usable means: the URL and PolicyID are as requested, **and** `FriendlyName` plus DWORD-typed `Flags` / `AuthFlags` / `Cost` are present
- **Root Flags handled as DISABLE bits**. The script clears the "ignore GP list" bit (`0x2`) when present. `-DisableUserConfigured` / `-EnableUserConfigured` set or clear the "ignore user-configured servers" bit (`0x4`). The script keeps the existing bits across runs
- **Read-back verification** of every written value, which detects a missing value. The summary reports the *actual* registry state and the outcome of each check
- **Protected registry path**. Before any write, the script checks every existing key from the hive root down to the target. A registry **symbolic link**, an untrusted owner, or write-class rights for an untrusted principal make the script refuse the run. A link placed where `PolicyServers` does not exist yet would send an elevated first-time write to the key that it points at. The script writes string values and the `(Default)` marker as `REG_SZ` explicitly, and it verifies the *kind* of every value, not only its content. It verifies a removal before it reports it
- **`-SetAsDefault` / `-ClearDefault`** for the unnamed `(Default)` interactive-enrollment marker. **`-ReplaceExisting`** removes stale siblings with the same PolicyID from a superseded URL
- **`-Remove`** mode, **`-WhatIf` / `-Confirm`** support, and structured `PSCustomObject` output

### Requirements

- Windows PowerShell 5.1 or PowerShell 7+. **No PowerShell module is required**
- An **elevated** session for `-Location LocalMachine`, `GPMachine`, or `GPUser`. `LocalUser` needs no elevation
- Domain connectivity, only for the AD policy row of the GP locations. On a workgroup machine the script skips that lookup with a warning
- **Tattooing caveat:** a direct write into the GP hives (`GPMachine` / `GPUser`) on a **domain member** produces a pseudo-policy that no GPO backs. RSoP and `gpresult` do not show it, `gpupdate` does not revert it, and the certificate MMC shows it read-only. On domain members use [`Add-CertificateEnrollmentPolicyServerToGpo`](#add-certificateenrollmentpolicyservertogpops1) instead. The GP locations here are for standalone or workgroup machines and for lab work.

### Parameters

| Parameter | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| `-Url` | `string` | Yes | | The full CEP URI, for example `https://pki.example.net/ejbca/msae/CEPService?alias`. The script uses it verbatim for the SHA-1 subkey, and clients use it for their `GetPolicies` calls. It must be an absolute http or https URI. |
| `-PolicyName` | `string` | Add only | | The "Policy Name" of the EJBCA MSAE alias. It becomes `FriendlyName`. Unless you set `-PolicyId`, the script hashes it **verbatim** to the PolicyID, so keep it identical to EJBCA. |
| `-PolicyId` | `string` | No | hash of `-PolicyName` | An explicit PolicyID for non-EJBCA servers. It must match the `GetPolicies` response of the server. |
| `-Location` | `LocalMachine` / `LocalUser` / `GPMachine` / `GPUser` | No | `LocalMachine` | The store to write. See the tattooing caveat for the GP hives. |
| `-Authentication` | `Anonymous` / `Kerberos` / `UsernamePassword` / `Certificate` | No | `Kerberos` | The client authentication type for the endpoint. `Kerberos` is "Windows integrated". |
| `-Cost` | `long` (1–4294967295) | No | `0x7FFFFFFD` | The priority. A lower value is preferred among endpoints that share a PolicyID. Pass large values in decimal. |
| `-NoAutoEnroll` | switch | No | | Leave "Enable for automatic enrollment and renewal" off. This clears Flags bit `0x10`. |
| `-AllowUntrustedIssuer` | switch | No | | Clear "Require strong validation during enrollment". This sets Flags bit `0x20`. |
| `-NoClientId` | switch | No | | Do not send the ClientId attribute. This clears Flags bit `0x4`. The default `0x14` matches the GPO editor. |
| `-SetAsDefault` / `-ClearDefault` | switch | No | | Set or clear the unnamed `(Default)` marker. The marker only preselects the policy for interactive enrollment. |
| `-SkipADPolicy` | switch | No | | GP locations: do **not** write the AD enrollment policy row. Use it only when you intend that removal. |
| `-ReplaceExisting` | switch | No | | Remove the sibling entries with the same PolicyID but a different URL, for example a stale URL or a typo. The script never removes the AD row. |
| `-DisableUserConfigured` / `-EnableUserConfigured` | switch | No | | GP locations: set or clear root Flags bit `0x4`, which means "ignore user-configured servers". |
| `-Remove` | switch | Remove mode | | Delete the entry for `-Url` from the chosen location, and clear a `(Default)` marker that the removal orphans. |
| `-WhatIf` / `-Confirm` | switch | No | | Preview / prompt. `-WhatIf` shows the computed subkey and PolicyID and writes nothing. |

> **CAUTION** - `-ReplaceExisting` removes every sibling entry with the same PolicyID but a different URL. A removed entry can be a working redundant endpoint. Use it only to replace a superseded URL.
> `-Remove` deletes the entry for `-Url` and clears an orphaned `(Default)` marker. Clients that use this location then lose that enrollment policy. Use it only when the entry must go.

### Usage

**Preview everything, the computed subkey and PolicyID, with no writes:**
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

- The script hashes `-PolicyName` **verbatim**. When you rename the alias in EJBCA, the PolicyID changes, and already-deployed entries become orphans.
- A re-run with a **different** URL does not remove the old entry. Multiple URLs per PolicyID is also the legitimate pattern for redundant endpoints. Use `-ReplaceExisting` to remove a superseded entry.
- GP locations: the script resolves the domain objectGUID for the AD Enrollment Policy row **before** it writes anything. On a domain-joined machine, a failed lookup stops the run with nothing written, because a GP configuration without that row removes the AD enrollment policy. On a workgroup machine, the script skips the row with a warning. Pass `-SkipADPolicy` to omit the row deliberately.
  - The script writes the row before the CEP entry. You can decline the confirmation of the row while the hive has no row. Then the script does not write the CEP entry, and the run stops there.
  - The `(Default)` marker and the `-ReplaceExisting` removals run only when a complete CEP entry for this URL exists. Complete means the URL and PolicyID as requested, read again right before the cleanup.
- `-Remove` handles a single entry and its `(Default)` marker. It leaves the shared root configuration in place: the root Flags, the AD row, and autoenrollment. The `.NOTES` of the script document the full manual teardown.

### Tests

The Pester suite ([`Tests/Add-CertificateEnrollmentPolicyServerOffline.Tests.ps1`](./Tests/Add-CertificateEnrollmentPolicyServerOffline.Tests.ps1)) runs three always-safe tiers and one opt-in tier. **Unit** exercises the real subkey, PolicyID and flag derivations through `-WhatIf`, so it writes nothing. **Static** parses the script and checks the comment-based help. **Guard** validates parameter conflicts. The opt-in **Lab** tier runs live registry round-trips: add, idempotent update, default marker, replace-sibling, remove, all verified against independent oracles. The safe tiers need no module, no AD, no elevation, and change nothing:

```powershell
Invoke-Pester -Path .\Tests\Add-CertificateEnrollmentPolicyServerOffline.Tests.ps1 -ExcludeTag Lab
```

The Lab tier **writes to the registry of this machine**, only to the user-configured stores: `LocalUser` always, `LocalMachine` when elevated. It never writes to the GP hives, because on a domain member those tattoo a pseudo-policy; the GPO suite covers Group Policy delivery instead. The tier is precise. Its entries have a per-run `PESTER-<hex>` name and a URL under the RFC-reserved `.invalid` TLD. It tracks every created key by its exact path and removes it in teardown, and it snapshots and restores a pre-existing `(Default)` marker. It removes a `PolicyServers` base key only when the run created it and the key ends the run empty:

```powershell
$cfg = New-PesterContainer -Path .\Tests\Add-CertificateEnrollmentPolicyServerOffline.Tests.ps1 -Data @{ RunLab = $true }
Invoke-Pester -Container $cfg
```

<sub>[↑ Back to top](#top)</sub>

---

## Add-CertificateEnrollmentPolicyServerToGpo.ps1

The script writes or removes the same **Certificate Enrollment Policy** setting directly in a domain Group Policy Object (**GPO**), for a fleet-wide rollout. It stays offline with respect to the policy server; it never makes a "Validate Server" round-trip. It writes the values with `Set-GPRegistryValue`. That cmdlet also does what a manual edit of SYSVOL gets wrong: the AD and `GPT.INI` version increments and the Registry CSE registration. The result appears in the GPME *Public Key Policies* dialog exactly as if you had clicked it in.

### Why you need this

A CEP configuration through Group Policy normally means the same server round-trip in the GPME dialog, one GPO at a time, by hand. This script writes the policy directly and correctly for a fleet, and it defends against three problems. It keeps the built-in **AD enrollment policy row**. A GPO with only your CEP entry would otherwise *remove* the AD enrollment policy from every client in scope, without a message. It writes each value individually, to avoid the list-form **`**delVals.`** deletion record of `Set-GPRegistryValue`, which makes clients delete the whole entry. It reads the state from the **PDC emulator**, or from `-Server`, so reads and writes see the same replica.

### Features

- **Written directly into the GPO** with `Set-GPRegistryValue`. The script handles the AD / `GPT.INI` version increments and the Registry CSE registration. The setting shows normally in GPME
- **Same offline derivations** as the per-machine script: the SHA-1 subkey, the EJBCA `String.hashCode()` PolicyID, or `-PolicyId`
- **AD enrollment policy row kept** by default, so the GPO does not remove the AD default policy fleet-wide. Opt out with `-SkipADPolicy`
- **Complete-row checks**. A pre-existing AD row or CEP entry satisfies the AD-row prerequisite, the `(Default)` marker and `-ReplaceExisting` only when it is *usable*. Usable means: the URL and PolicyID are as requested, **and** `FriendlyName` plus DWORD `Flags` / `AuthFlags` / `Cost` are present. The URL and PolicyID alone are what an interrupted write leaves behind
- **Deletion-record-safe**. Individual value writes avoid the `**delVals.` problem, and the script **detects and warns** about mis-ordered deletion records that are already in the GPO. The script reads the `registry.pol` state as a client applies it. It replays the records in order. It honours `**del.` / `**delvals.` / `**DeleteValues` / `**DeleteKeys`, with their data read as a string, whatever the type of the record says. `**soft.` keeps its type
- **Optional Auto-Enrollment** (`-EnableAutoEnrollmentPolicy`) in the same GPO scope: AEPolicy, expiration percent and store. The script warns before it changes existing values
- **Same-replica state reads** from the PDC emulator, or from `-Server`. The script retries on transient SYSVOL / `registry.pol` contention, and it **verifies every entry with a read-back**
- **Effective-state checks**. The `(Default)` marker and `-ReplaceExisting` act only when the replacing entry is complete in *both* the live GPMC view and the `registry.pol` replay. A row whose written values a later deletion record wipes does not count. A freshly written row must show in the replay, or the run fails. The script verifies a removal from `registry.pol` before it reports it
- **Machine or User scope**, root Flags DISABLE-bit handling, `-SetAsDefault` / `-ClearDefault`, `-ReplaceExisting`, a `-Remove` mode, `-WhatIf` / `-Confirm`, and structured output

### Requirements

- Windows PowerShell 5.1 or PowerShell 7+. On 7, the GroupPolicy module loads through the WinPSCompat shim; expect its one-time compatibility warning. The script deliberately avoids `#Requires -Modules`, which would refuse to run there
- The **GroupPolicy module** (GPMC / RSAT) and permission to edit the target GPO
- An existing, linked GPO to write into. Create one first, for example `New-GPO -Name 'PKI - Enrollment Policy' | New-GPLink -Target 'OU=...,DC=...'`
- Domain connectivity. The script targets the PDC emulator by default, or the DC that you pass to `-Server`
- Do **not** edit the same GPO concurrently from another session or from GPME. `Set-GPRegistryValue` is an unlocked read-modify-write. The read-back detects the loss of the own entry of this script

### Parameters

| Parameter | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| `-GpoName` | `string` | Yes | | The display name of an existing GPO, or its GUID. The script tries the display name first. |
| `-Url` | `string` | Yes | | The full CEP URI. The script uses it verbatim. It must be an absolute http or https URI. |
| `-PolicyName` | `string` | Add only | | The "Policy Name" of the EJBCA MSAE alias. It becomes `FriendlyName`. Unless you set `-PolicyId`, the script hashes it **verbatim** to the PolicyID. |
| `-PolicyId` | `string` | No | hash of `-PolicyName` | An explicit PolicyID for a CEP server that returns a GUID. |
| `-Scope` | `Machine` / `User` | No | `Machine` | Computer Configuration (HKLM) or User Configuration (HKCU). |
| `-Authentication` | `Anonymous` / `Kerberos` / `UsernamePassword` / `Certificate` | No | `Kerberos` | The client authentication type for the endpoint. |
| `-Cost` | `long` (1–4294967295) | No | `0x7FFFFFFD` | The priority. A lower value is preferred among endpoints that share a PolicyID. |
| `-NoAutoEnroll` / `-AllowUntrustedIssuer` / `-NoClientId` | switch | No | | Changes to the entry Flags: clear `0x10`, set `0x20`, or clear `0x4`. The default `0x14` matches GPME. |
| `-SetAsDefault` / `-ClearDefault` | switch | No | | Set or clear the `(Default)` marker. The marker only preselects the policy for interactive enrollment. |
| `-SkipADPolicy` | switch | No | | Do **not** write the AD enrollment policy row. Use it only when you intend that removal. |
| `-ReplaceExisting` | switch | No | | Remove the sibling entries with the same PolicyID under a different URL. The script never removes the AD row. |
| `-DisableUserConfigured` / `-EnableUserConfigured` | switch | No | | Set or clear root Flags bit `0x4`, which means "ignore user-configured servers". |
| `-EnableAutoEnrollmentPolicy` | switch | No | | Also write the "Auto-Enrollment" setting into the same scope. |
| `-AEPolicy` / `-AEExpirationPercent` / `-AEStore` | `int` / `int` / `string` | No | `7` / `10` / `MY` | The Auto-Enrollment values. Only with `-EnableAutoEnrollmentPolicy`. |
| `-Domain` / `-Server` | `string` | No | current / PDC emulator | The domain and the DC for the GroupPolicy cmdlets and the `registry.pol` state reads. The script keeps them on one replica. |
| `-Remove` | switch | Remove mode | | Delete the entry for `-Url` from the GPO scope, and clear a `(Default)` marker that the removal orphans. |
| `-WhatIf` / `-Confirm` | switch | No | | Preview / prompt. |

> **CAUTION** - `-ReplaceExisting` removes every sibling entry with the same PolicyID but a different URL from the GPO. A removed entry can be a working redundant endpoint for every client in scope. Use it only to replace a superseded URL.
> `-Remove` deletes the entry for `-Url` from the GPO scope and clears an orphaned `(Default)` marker. Every client in scope then loses that enrollment policy at the next Group Policy refresh. Use it only when the entry must go.

### Usage

**Preview a fleet rollout with Auto-Enrollment enabled:**
```powershell
.\Add-CertificateEnrollmentPolicyServerToGpo.ps1 -GpoName 'PKI - Enrollment Policy' `
    -Url 'https://pki.example.net/ejbca/msae/CEPService?alias' -PolicyName 'Example PKI Service' `
    -EnableAutoEnrollmentPolicy -WhatIf
```

**Write it into the Computer configuration:**
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

- After a change, clients load it at the next Group Policy refresh. Force the refresh with `gpupdate`. Then trigger enrollment with `certutil -pulse` for the machine, or `certutil -user -pulse` for the user.
- The script hashes `-PolicyName` **verbatim**. Keep it identical to the EJBCA alias, or the deployed entries become orphans.
- `-Remove` clears one entry and its `(Default)` marker. It leaves the shared configuration in place: the root Flags, the AD row, and Auto-Enrollment. The `.NOTES` of the script list the full teardown with `Remove-GPRegistryValue`.
- The script resolves the domain objectGUID for the AD Enrollment Policy row **before** any GPO write. When it cannot resolve the objectGUID, the run stops with nothing written. A GPO with a CEP entry but no `LDAP:` row removes the AD enrollment policy from every client in scope. Pass `-SkipADPolicy` to omit the row deliberately.
  - The script writes and verifies the row **before** the CEP entry. You can decline the confirmation of the row while the GPO has no row. Then the script does not write the CEP entry, and the run stops there. It changes no root Flags, marker, Auto-Enrollment or sibling.
  - The `(Default)` marker and the `-ReplaceExisting` removals run only when a **complete** CEP entry for this URL exists. Complete means the URL and PolicyID as requested, read again right before the cleanup. So a declined entry prompt or a half-written entry cannot leave a marker that points at nothing. It also cannot delete the only working endpoint, or act for a PolicyID that the retained entry does not serve.
- The script reads a GUID-shaped `-GpoName` as a GPO ID only after an independent listing proves that no GPO has that display name. Every other name-lookup failure stops the run; the script never retargets without an error.

### Tests

The Pester suite ([`Tests/Add-CertificateEnrollmentPolicyServerToGpo.Tests.ps1`](./Tests/Add-CertificateEnrollmentPolicyServerToGpo.Tests.ps1)) runs three always-safe tiers and one opt-in tier. **Unit** exercises the real `Registry.pol` binary parser and the entry and value extractors against a hand-built `.pol` stream; it needs no module. **Static** parses the script and checks the help. **Guard** validates parameter conflicts and throws before it touches any GPO; it needs the GroupPolicy module and skips itself without it.

The opt-in **Lab** tier runs the full lifecycle in Machine and User scope. The lifecycle is: write, verify through GPMC *and* the real `registry.pol`, default marker plus Auto-Enrollment, replace-sibling, and remove. It runs inside **one throwaway, never-linked GPO**, which applies to zero clients, and deletes it afterwards by its exact tracked GUID. The safe tiers write no GPO:

```powershell
Invoke-Pester -Path .\Tests\Add-CertificateEnrollmentPolicyServerToGpo.Tests.ps1 -ExcludeTag Lab
```

The Lab tier needs a domain-joined machine, the GroupPolicy module, and permission to create GPOs. A lab DC is ideal. The tier never touches pre-existing GPOs:

```powershell
$cfg = New-PesterContainer -Path .\Tests\Add-CertificateEnrollmentPolicyServerToGpo.Tests.ps1 -Data @{ RunLab = $true }
Invoke-Pester -Container $cfg
```

<sub>[↑ Back to top](#top)</sub>

---

## Template library

[`Templates/`](./Templates/README.md) ships importable JSON exports of **all 33 certutil default certificate templates** (`Templates/Default/`). It also ships **EJBCA-ready variants** (`Templates/EJBCA/`) of all four defaults whose Subject is empty at creation: `DirectoryEmailReplication`, `DomainControllerAuthentication`, `KerberosAuthentication`, `Workstation`. The variants have the same one-bit full-DN-Subject change that the [EJBCA section](#using-the-template-with-ejbca) describes, already applied. A parallel **`Templates/MaxCompat/`** holds both sets moved to the newest compatibility: **CA: Windows Server 2016**, **recipient: Windows 10 / Windows Server 2016**. That is possible for 9 templates. The 24 schema-v1 built-ins are read-only and pass through unchanged.

We removed the template OID from every file, because it would identify the source forest. Importers generate their own OID with `-OidHandling Generate` / `GenerateRandom`. The names are kept, and no file has a domain SID. We produced the whole library **clean-room**: we seeded a fresh forest with `certutil -InstallDefaultTemplates`, and we validated every file against a live DC. The details, the criteria, the compatibility encoding, and the full per-template table are in the [README](./Templates/README.md) of the folder.

<sub>[↑ Back to top](#top)</sub>

---

## License

[MIT](./LICENSE)

## Contributing

Issues and pull requests are welcome. See [CONTRIBUTING.md](./CONTRIBUTING.md) for the house style, the dual-engine (PowerShell 5.1 + 7) and `-WhatIf` expectations, and how to run the analyzer and the four-tier test suites.

**Writing style.** The operator-facing text follows [STYLE.md](./STYLE.md). Check it with the lint: `Invoke-Pester -Path .\Tests\Style.Tests.ps1`.

## Security

To report a vulnerability, use the private reporting of GitHub (**Security → Report a vulnerability**), not a public issue. [SECURITY.md](./SECURITY.md) has the details and the safety model.
