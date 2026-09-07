# Project rules

## Writing style

Operator-facing text follows `STYLE.md`, a style derived from ASD-STE100 Simplified Technical English. It applies to the comment-based help of every script, to `README.md`, and to the newest section of `CHANGELOG.md` (`[Unreleased]` while a release is prepared, the latest release after it ships).

- Write sentences of 25 words or fewer and paragraphs of 6 sentences or fewer.
- Use the active voice and name the actor ("the script refuses ...", not "... is refused").
- Use "must" for a requirement, "can" for an ability, "may" for permission. Do not write "should".
- Use no figurative language. `STYLE.md` lists the banned words and their replacements.
- Use the glossary terms in `STYLE.md`. Define an acronym the first time you use it in a section.
- Give every accept-risk switch a `CAUTION:` line in its help and a `> **CAUTION**` block in the README.
- Keep every fact. A rewrite for style must not drop a condition, a default, an exception, or a limit.

`Tests\Style.Tests.ps1` checks these rules; CI runs it on both PowerShell engines. Run it after every change to help or Markdown:

```powershell
Invoke-Pester -Path .\Tests\Style.Tests.ps1 -Output Detailed
```

Code comments are not linted. When you write or change one, keep the sentences short and use no figurative language. Do not rewrite older `CHANGELOG.md` sections or the `.RELEASENOTES` lines: they are a historical record.

## Gate for every change

- PSScriptAnalyzer must be clean (the module is installed for Windows PowerShell 5.1).
- The non-Lab Pester tiers must pass on Windows PowerShell 5.1 and on PowerShell 7 (`Invoke-Pester -Path .\Tests -ExcludeTag Lab`).
- When a script changes, bump its `.VERSION`, add a `.RELEASENOTES` line, and add a bullet to the `[Unreleased]` section of `CHANGELOG.md`.
