# Changelog

All notable changes to this repository are recorded here. The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/); release tags follow [Semantic Versioning](https://semver.org/).

Each script also carries its own version in the `PSScriptInfo` header at the top of the file (`.VERSION`), bumped only when that script changes. The table under each release lists the script versions it ships, so a deployed copy can be checked without opening it:

```powershell
Test-ScriptFileInfo .\Submit-CertificateRequests.ps1 | Select-Object Name, Version
```

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

[1.0.1]: https://github.com/TheOmnilord/ADCS/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/TheOmnilord/ADCS/releases/tag/v1.0.0