# Contributing

Issues and pull requests are welcome. This is a small, focused set of PowerShell tools for
AD CS / EJBCA PKI administration; the notes below keep contributions consistent with the
existing code.

## Ground rules

- **Target both engines.** Everything must work on **Windows PowerShell 5.1** and **PowerShell
  7+**. CI runs the analyzer and the safe test tiers on both.
- **State-changing code supports `-WhatIf` / `-Confirm`** and makes previews side-effect-free.
- **No plaintext secrets.** Take a `[pscredential]` for anything cross-forest or authenticated.
- **Match the house style.** Comment-based help with `.PARAMETER`/`.EXAMPLE` for every
  parameter; structured `PSCustomObject` output where a caller might consume it; clear,
  actionable error messages.
  - Put `#Requires` **after** the comment-based help block, not before it — before it, `Get-Help`
    silently fails to bind (a real bug this repo has already hit).

## Before you open a PR

1. **PSScriptAnalyzer must be clean** with the repo settings:
   ```powershell
   Invoke-ScriptAnalyzer -Path . -Recurse -Settings .\PSScriptAnalyzerSettings.psd1
   ```
   Rule exclusions are centralized in `PSScriptAnalyzerSettings.psd1`; add a *narrow*, justified
   `[Diagnostics.CodeAnalysis.SuppressMessageAttribute]` only for genuine false positives.
2. **Run the safe test tiers** (no AD, no CA, no changes):
   ```powershell
   Invoke-Pester -Path .\Tests -ExcludeTag Lab
   ```
   New behavior needs tests. Each script has a `Tests/<name>.Tests.ps1` suite with four tiers:
   **Unit** (pure helpers, extracted from the script by AST so the real code runs), **Static**
   (parse + help binding), **Guard** (parameter/mode validation that throws before any external
   contact), and an opt-in **Lab** tier of live operations.
3. **Lab-tier tests must be surgical.** Track every created object by exact identity and remove
   it in teardown; scope any safety-net sweep to a run-unique `PESTER-<hex>` prefix behind a
   structural guard (an unset prefix must never widen a delete filter — `AfterAll` runs even when
   `BeforeAll` throws). Never delete by a broad wildcard. Run them only against a lab you own:
   ```powershell
   $cfg = New-PesterContainer -Path .\Tests\<name>.Tests.ps1 -Data @{ RunLab = $true }
   Invoke-Pester -Container $cfg
   ```

## Commit / PR notes

- Keep commits focused; explain the *why* in the message.
- Note anything you verified live (engine, environment) so reviewers know the coverage.
