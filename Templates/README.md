# Template library

Ready-to-import JSON exports of the **certutil default certificate templates**, for use with
[`Sync-ADCSTemplate.ps1`](../Sync-ADCSTemplate.ps1) — plus **EJBCA-ready variants** of the
templates that need one.

- **`Default/`** — all 33 default templates, exported **as-is** from a
  `certutil -InstallDefaultTemplates` set (Windows Server 2025 schema, exported 2026-08-29).
  > **Provenance caveat:** the source was a working lab forest, presumed stock but not
  > guaranteed pristine — a re-export from a freshly seeded, never-touched forest is planned
  > (`certutil -InstallDefaultTemplates` needs no AD CS role, so any clean forest will do).
  > Until then, diff against your own environment before relying on exact values.
- **`EJBCA/`** — variants of the templates whose Subject is **empty at creation** (Subject
  name format *None*), with the **Fully distinguished name** Subject enabled — the one change
  EJBCA/MSAE requires (see [Using the template with EJBCA](../README.md#using-the-template-with-ejbca)).

## What was deliberately changed or left out

- **Template OIDs are stripped** (`-StripOid`). A template OID's root arcs are minted randomly
  per forest, so publishing them would fingerprint the source forest — and importing a foreign
  OID into your forest is bad hygiene anyway. Import therefore **requires a Generate mode**:
  `-OidHandling Generate` (mints under your forest's OID root) or `GenerateRandom` (works even
  in a forest with **no AD CS**).
- **ACLs are never part of an export** — they are forest-specific and are rebuilt at import
  from `-AclBase` / `-EnrollPrincipals`.
- **Names are kept.** Default template names are public and meaningful; rename at import with
  `-NewTemplateName` / `-NewDisplayName` if needed.
- **Audited:** no `S-1-5-21-*` (domain) SIDs appear in any file. The only embedded security
  descriptor in the set (`OCSPResponseSigning`'s private-key SDDL, packed inside
  `msPKI-RA-Application-Policies`) grants only machine-independent principals: Builtin
  Administrators, SYSTEM, and the OCSP responder's service SID (`S-1-5-80-…`, derived from the
  service name — identical on every Windows machine). The import script still prints its
  generic key-SDDL advisory for that file; it is safe to proceed.

## Importing

```powershell
# Into a forest that never had a CA (e.g. an EJBCA/MSAE enrollment forest):
.\Sync-ADCSTemplate.ps1 -Mode Import -Path .\Templates\EJBCA\DomainControllerAuthentication-EJBCA.json -OidHandling GenerateRandom

# Into a forest with AD CS, minting under its own OID root, renamed:
.\Sync-ADCSTemplate.ps1 -Mode Import -Path .\Templates\Default\Workstation.json -OidHandling Generate `
    -NewTemplateName "YY-Workstation" -NewDisplayName "YY-Workstation"
```

Import refuses a duplicate cn, so importing a default under its original name into a forest
that already carries the default set fails safely — rename, or import into the empty target
forest these files are meant for. **v1 templates** (schema version 1 below) should be imported
under their **original names** when Windows clients must recognize them — Windows fixes v1
semantics in code by name; non-Windows consumers like EJBCA read the object directly and are
unaffected.

## The EJBCA variants

EJBCA cannot use a template that produces an **empty Subject**. A default template gets a
variant here exactly when its Subject at creation is *None* — i.e. `msPKI-Certificate-Name-Flag`
has neither `CT_FLAG_ENROLLEE_SUPPLIES_SUBJECT` (`0x1`) nor any `CT_FLAG_SUBJECT_REQUIRE_*`
bit. The variant is the default file with **one change**:
`CT_FLAG_SUBJECT_REQUIRE_DIRECTORY_PATH` (`0x80000000`, "Fully distinguished name") OR'd into
`msPKI-Certificate-Name-Flag`. Templates that already put something in the Subject (supplied
in the request, or built as CN/DN/DNS/e-mail) need no variant.

| Variant | Name flag change |
| --- | --- |
| `DirectoryEmailReplication-EJBCA.json` | `0x09000000` → `0x89000000` |
| `DomainControllerAuthentication-EJBCA.json` | `0x08000000` → `0x88000000` |
| `Workstation-EJBCA.json` | `0x08000000` → `0x88000000` |

**`KerberosAuthentication` also qualifies but is deliberately not shipped** — it is the worked
example of the main README's [EJBCA section](../README.md#using-the-template-with-ejbca), and
the full-DN variant of it has already been field-validated end to end against EJBCA.

## The full set

Subject-at-creation and supersedence below are read from the objects themselves
(`msPKI-Certificate-Name-Flag`, `msPKI-Supersede-Templates`), not from documentation.

| Template | Ver | Subject at creation | EJBCA variant | Notes |
| --- | --- | --- | --- | --- |
| Administrator | 1 | Built: full DN + e-mail | — | |
| CA | 1 | Supplied in request | — | |
| CAExchange | 2 | Supplied in request | — | |
| CEPEncryption | 1 | Supplied in request | — | |
| ClientAuth | 1 | Built: full DN | — | |
| CodeSigning | 1 | Built: full DN | — | |
| CrossCA | 2 | Supplied in request | — | |
| CTLSigning | 1 | Built: full DN | — | |
| DirectoryEmailReplication | 2 | **None (empty)** | ✔ | supersedes DomainController |
| DomainController | 1 | Built: DNS as CN | — | superseded (see below) |
| DomainControllerAuthentication | 2 | **None (empty)** | ✔ | supersedes DomainController |
| EFS | 1 | Built: full DN | — | |
| EFSRecovery | 1 | Built: full DN | — | |
| EnrollmentAgent | 1 | Built: full DN | — | |
| EnrollmentAgentOffline | 1 | Supplied in request | — | |
| ExchangeUser | 1 | Supplied in request | — | |
| ExchangeUserSignature | 1 | Supplied in request | — | |
| IPSECIntermediateOffline | 1 | Supplied in request | — | |
| IPSECIntermediateOnline | 1 | Built: DNS as CN | — | |
| KerberosAuthentication | 2 | **None (empty)** | deliberate omit | modern DC template |
| KeyRecoveryAgent | 2 | Built: full DN | — | |
| Machine | 1 | Built: DNS as CN | — | |
| MachineEnrollmentAgent | 1 | Built: DNS as CN | — | |
| OCSPResponseSigning | 3 | Built: DNS as CN | — | embedded key SDDL (well-known SIDs only) |
| OfflineRouter | 1 | Supplied in request | — | |
| RASAndIASServer | 2 | Built: CN | — | |
| SmartcardLogon | 1 | Built: full DN | — | |
| SmartcardUser | 1 | Built: full DN + e-mail | — | |
| SubCA | 1 | Supplied in request | — | |
| User | 1 | Built: full DN + e-mail | — | |
| UserSignature | 1 | Built: full DN + e-mail | — | |
| WebServer | 1 | Supplied in request | — | |
| Workstation | 2 | **None (empty)** | ✔ | |

On DC certificates generally: **KerberosAuthentication** is Microsoft's modern template for
domain controllers — for new deployments prefer it (or its full-DN variant for EJBCA) over the
legacy `DomainController` (v1) and `DomainControllerAuthentication` templates.

## Verification performed

Every file in this folder was validated on a live lab forest before publishing: all 36 passed
the import pre-flight (`-Mode Import -WhatIf` with a throwaway name), and a representative
EJBCA variant was imported for real — the full-DN Subject flag persisted verbatim in AD, a
fresh OID was minted, and the throwaway was removed.

<sub>[↑ Back to repo README](../README.md)</sub>
