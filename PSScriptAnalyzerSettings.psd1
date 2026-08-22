@{
    # Repo-wide PSScriptAnalyzer settings, consumed by CI (.github/workflows/ci.yml) and by
    # local runs: Invoke-ScriptAnalyzer -Path . -Recurse -Settings .\PSScriptAnalyzerSettings.psd1
    #
    # Excluded rules are DELIBERATE style decisions of these operator-facing scripts, not oversights:
    ExcludeRules = @(
        # These are interactive admin tools: colored console narration (progress, summaries,
        # ACL listings) is the intended UX, alongside structured PSCustomObject pipeline output.
        'PSAvoidUsingWriteHost',

        # Established names describe plural things (Submit-CertificateRequests,
        # Resolve-TemplateGrants, Get-TemplateAllowedAttributes); renaming would break users.
        'PSUseSingularNouns',

        # Sync-ADCSTemplate.ps1 declares SupportsShouldProcess at SCRIPT level and passes
        # $PSCmdlet into its internal helpers, which call $CallerCmdlet.ShouldProcess so that
        # -WhatIf/-Confirm are honored against the caller's state. The analyzer cannot follow
        # that propagation and would demand a redundant per-function attribute.
        'PSShouldProcess',
        'PSUseShouldProcessForStateChangingFunctions'
    )
}
