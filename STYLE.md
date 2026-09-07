# Writing style for operator-facing text

This repository writes its operator-facing text in a style derived from **ASD-STE100 Simplified Technical English** (Issue 9, 2025). The goal is text that a non-native reader, a translator, or a tool can read without ambiguity. We take the STE writing rules that give the most clarity. We do not use the STE dictionary of 900 approved words: the vocabulary of PKI and Windows security is too specific for it.

`Tests\Style.Tests.ps1` checks the rules that a tool can check. CI runs it on Windows PowerShell 5.1 and PowerShell 7. Run it locally with:

```powershell
Invoke-Pester -Path .\Tests\Style.Tests.ps1 -Output Detailed
```

## Scope

The rules apply to:

- The comment-based help of every script (`.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER`, `.EXAMPLE` descriptions, `.NOTES`).
- `README.md`.
- The newest section of `CHANGELOG.md`: `[Unreleased]` while you prepare a release, and the latest release after it ships. The lint reads only that section.

The rules do not apply to:

- Older sections of `CHANGELOG.md` and the `.RELEASENOTES` lines in each script. They are a historical record. Do not rewrite them.
- Code comments. When you write or change a code comment, keep sentences short and use no figurative language (rules 1 and 4). The lint does not check code comments.
- Test names, `-Because` strings, and log or error messages in the scripts.

## Rules

### 1. Keep sentences short

Write sentences of **25 words or fewer**. Write paragraphs of **6 sentences or fewer**. Give one instruction per sentence. When a sentence carries two facts, write two sentences.

- Not this: "The lock is opened only after the tracking folder has passed the chain check, because a link planted there would otherwise carry the privileged write elsewhere."
- This: "The script checks the tracking folder first. Then it opens the lock. A link in an unchecked folder could send the privileged write to another file."

A code span, a path, a URL, or a parameter name counts as one word.

### 2. Use the active voice and name the actor

Say who does what. The actor is usually "the script", "certreq", "the CA", "Windows", or "you".

- Not this: "The folder is refused."
- This: "The script refuses the folder."

Use the present tense for what the script does. Use the imperative for what the reader must do.

### 3. Use must, can, and may with one meaning each

- **must**: a requirement. "The tracking folder must exist."
- **must not**: a prohibition. "The path must not contain a junction."
- **can**: ability or possibility. "A user with this right can replace the folder."
- **may**: permission. "You may pass `-Force` to resubmit."

Do not write "should". Write "must", or write "we recommend that you ...".

### 4. Use no figurative language

Write what happens, not an image of it. The lint refuses the words below. Use the replacement.

| Do not write | Write |
|---|---|
| plant, planted | create, place |
| swap, swappable | replace, "can replace" |
| window (a period of time) | the time between X and Y |
| carry (a write, a change) | apply to, send to |
| land, lands | is written to, goes to |
| walk (a chain, a path) | check each folder from X to Y |
| descend | go into, create inside |
| silently | without an error, without a message |
| tolerate | permit, allow |
| wedge, hang | stop responding |
| pop (a dialog) | open |
| hand (rights to) | give |
| pick up (changes) | load, read, collect |
| gate on (an exit code) | use ... to decide |
| up front | first, before |
| load-bearing | required |
| e.g., i.e., etc., via, and/or | for example, that is, and so on, through, "X, Y, or both" |
| in order to, prior to, utilize, leverage, ensure | to, before, use, use, make sure |
| please, very | (omit) |
| contractions (don't, it's, ...) | the full form |

Defined technical terms are not figurative (see the glossary). "Fail closed" is a defined term.

### 5. Use the glossary terms, and define acronyms

Use one term for one thing, and use it every time. Do not alternate between "folder" and "directory", or between "certificate template" and "template" in one section without reason. Define an acronym the first time you use it in a section: "Authentication Mechanism Assurance (AMA)".

### 6. Mark accepted risk with CAUTION

Some switches make the script accept a risk that it refuses by default. Their help must state the risk in a line that starts with `CAUTION:`, in 25 words or fewer. The README must have a `> **CAUTION**` block that names each such switch. The lint checks both.

Switches that need a CAUTION:

| Script | Switches |
|---|---|
| Submit-CertificateRequests.ps1 | `-AllowUnprotectedOutputFolder`, `-Force` |
| Sync-ADCSTemplate.ps1 | `-SkipAcl`, `-AllowLinkedIssuancePolicy` |
| Add-CertificateEnrollmentPolicyServerOffline.ps1 | `-ReplaceExisting`, `-Remove` |
| Add-CertificateEnrollmentPolicyServerToGpo.ps1 | `-ReplaceExisting`, `-Remove` |

### 7. Keep structure simple

- Use a bulleted list for parallel items. Give each item its own sentence.
- Put essential information in the sentence, not in parentheses.
- Use articles ("the", "a"). Write "the script", not "script".
- Keep noun strings short. Write "the registry key of the policy server", not "the policy server registry key".
- Use a table for values, defaults, and mappings.

### 8. Keep the facts

A rewrite for style must not change a fact. Do not drop a condition, a default, an exception, or a limit. When a rule and a fact conflict, keep the fact and write the sentence another way.

## Glossary

Technical terms that the text uses as they are. Use the exact spelling.

**Windows and file system:** junction, symbolic link, reparse point, mount point, ACL, ACE, DACL, owner, SID, inheritance, inherit-only entry, container-inheritable entry, extended path (`\\?\`), UNC path, `%TEMP%`.

**Principals and trust:** principal, trusted principal (SYSTEM, Administrators, TrustedInstaller, the running account, its Domain Admins and Enterprise Admins, and the principals named in `-TrustedOutputPrincipal`), untrusted principal (every other principal), CREATOR OWNER, CREATOR GROUP, OWNER RIGHTS.

**This repository:** chain check (the check of every folder from a folder up to the root of its volume or share), run lock, per-run log, tracking file, tracking folder, staging file, delivery, destination, drop folder, request file, capture file, output folder, anchor (the deepest existing folder above a folder the script creates), fail closed (the script stops with an error rather than continue with a risk).

**PKI and AD CS:** CA, CSR, certreq, certutil, RequestID, disposition (Issued, Pending, Denied, Error, Unknown, Undelivered), certificate template, schema version, OID, issuance policy, application policy, enrollment agent, Authentication Mechanism Assurance (AMA), Certificate Enrollment Policy (CEP), policy server, Auto-Enrollment, Group Policy Object (GPO), `registry.pol`, forest, domain controller (DC), UPN, sAMAccountName.

**Verbs with one meaning:** refuse (stop with an error and change nothing), validate (check and refuse on failure), canonicalize (reduce to one standard form), resolve (find the object a name refers to), deliver (move a certificate to its destination), retrieve (get an issued certificate from the CA), submit, resubmit, import, export, sync, upgrade, replicate, enroll.

## When you write

1. Write the sentence. Read it aloud. If you take a breath, split it.
2. Name the actor. Replace "is done" with "the script does".
3. Remove every image. Replace it with the event.
4. Check every fact against the code.
5. Run `Invoke-Pester -Path .\Tests\Style.Tests.ps1`.
