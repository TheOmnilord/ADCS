# Security policy

These are PowerShell tools that administer **Active Directory Certificate Services** and
**enrollment-policy** configuration — they modify certificate templates, submit certificate
requests, and change client/GPO registry settings. Run them only against environments you are
authorized to administer, and prefer `-WhatIf` first.

## Reporting a vulnerability

If you find a security issue in this project, please **do not open a public issue**. Instead use
GitHub's private reporting: on the repository, go to **Security → Report a vulnerability** (GitHub
Private Vulnerability Reporting). Include the affected script, a description, and reproduction
steps or a proof of concept.

Please allow a reasonable period for a fix before any public disclosure. There is no bug-bounty
program; this is a best-effort, community-maintained project.

## Scope and safety model

- **Every state-changing script supports `-WhatIf` / `-Confirm`.** Preview first; the scripts
  are written to make previews side-effect-free.
- **No secrets are handled in plaintext.** Cross-forest operations take a `[pscredential]`
  (`-Credential` / `-SourceCredential`); nothing is written to disk in clear.
- **Exported template JSON never contains ACLs or security descriptors** — permissions are
  rebuilt on import from `-AclBase` / `-EnrollPrincipals`. Template OIDs are stripped from the
  files shipped under [`Templates/`](./Templates/README.md) so they cannot fingerprint a forest.
- **Destructive operations are exact-scoped.** The test suites, when run against a lab
  (`-RunLab`), create and remove only objects carrying a unique per-run `PESTER-<hex>` prefix and
  never touch pre-existing objects; see the per-suite notes under [`Tests/`](./Tests).

## Supported versions

The latest commit on `main` is the supported version. Fixes are applied there; there is no
long-term-support branch.
