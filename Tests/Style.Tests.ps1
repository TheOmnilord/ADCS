#Requires -Version 5.1
<#
    Style lint for the operator-facing text: the rules of STYLE.md that a tool can check, as Pester
    Static tests. Scope: the comment-based help of every script, README.md, and the NEWEST section
    of CHANGELOG.md ([Unreleased] while a release is prepared, the latest release after it ships).
    Out of scope (see STYLE.md): code comments, tests, the PSScriptInfo block with its
    .RELEASENOTES lines, and older CHANGELOG sections.

    Rules checked:
      1. Every sentence has 25 words or fewer; every paragraph has 6 sentences or fewer.
      4. No banned word or phrase (figurative, ambiguous, or unapproved forms).
      6. Every accept-risk switch has a CAUTION line in its help and a CAUTION block in the README.
    The CHANGELOG checks read only the newest section, so a release that ships everything leaves
    nothing to lint there, and the next [Unreleased] section is linted as soon as it exists.

    How prose is read: code spans, URLs and bold markers are masked first, so a parameter name, a
    path or a function name never counts as several words and never triggers a banned word. Help
    .EXAMPLE command lines, fenced code, tables, headings, HTML and link references are skipped.
#>
param()

BeforeDiscovery {
    $script:StyleScripts = @(
        'Set-ADCSTemplateValidity.ps1',
        'Submit-CertificateRequests.ps1',
        'Sync-ADCSTemplate.ps1',
        'Add-CertificateEnrollmentPolicyServerOffline.ps1',
        'Add-CertificateEnrollmentPolicyServerToGpo.ps1'
    )
    $script:RiskSwitches = @{
        'Submit-CertificateRequests.ps1'                   = @('AllowUnprotectedOutputFolder', 'Force')
        'Sync-ADCSTemplate.ps1'                            = @('SkipAcl', 'AllowLinkedIssuancePolicy')
        'Add-CertificateEnrollmentPolicyServerOffline.ps1' = @('ReplaceExisting', 'Remove')
        'Add-CertificateEnrollmentPolicyServerToGpo.ps1'   = @('ReplaceExisting', 'Remove')
    }
}

Describe 'Style: operator-facing text follows STYLE.md' -Tag 'Static' {

    BeforeAll {
        $script:Root         = Split-Path -Path $PSScriptRoot -Parent
        $script:MaxWords     = 25
        $script:MaxSentences = 6

        # Banned words and phrases (STYLE.md rule 4), each with its plain replacement. Matched
        # case-insensitively on masked prose.
        $script:Banned = @(
            @{ Pattern = '\bplant(s|ed|ing)?\b';                 Say = 'create / place' }
            @{ Pattern = '\bswap(s|ped|ping|pable)?\b';          Say = 'replace / "can replace"' }
            @{ Pattern = '\bwindow\b';                            Say = 'the time between X and Y' }
            @{ Pattern = '\bcarr(y|ies|ied|ying)\b';              Say = 'apply to / send to' }
            @{ Pattern = '\bland(s|ed|ing)?\b';                  Say = 'is written to / goes to' }
            @{ Pattern = '\bwalk(s|ed|ing)?\b';                  Say = 'check each ... from ... to ...' }
            @{ Pattern = '\bdescend(s|ed|ing)?\b';               Say = 'go into / create inside' }
            @{ Pattern = '\bsilently\b';                          Say = 'without an error / without a message' }
            @{ Pattern = '\btolerat(e|es|ed|ing)\b';              Say = 'permit / allow' }
            @{ Pattern = '\bwedge[ds]?\b';                        Say = 'stop responding' }
            @{ Pattern = '\bhang(s|ing)?\b';                      Say = 'stop(s) responding' }
            @{ Pattern = '\bpops?\b';                             Say = 'open(s)' }
            @{ Pattern = '(?<!by )\bhand(s|ed)?\b(?!-)';          Say = 'give(s)' }
            @{ Pattern = '\bpick(s|ed|ing)? up\b';                Say = 'load / read / collect' }
            @{ Pattern = '\bgates? on\b';                         Say = 'uses ... to decide' }
            @{ Pattern = '\bup front\b';                          Say = 'first / before ...' }
            @{ Pattern = '\bload-bearing\b';                      Say = 'required' }
            @{ Pattern = '\bon the fly\b';                        Say = 'while it runs' }
            @{ Pattern = '\bunder the hood\b';                    Say = 'internally' }
            @{ Pattern = '\be\.g\.';                              Say = 'for example' }
            @{ Pattern = '\bi\.e\.';                              Say = 'that is' }
            @{ Pattern = '\betc\.';                               Say = 'and so on' }
            @{ Pattern = '\bvia\b';                               Say = 'through / by / with' }
            @{ Pattern = '\band/or\b';                            Say = 'X, Y, or both' }
            @{ Pattern = '\bin order to\b';                       Say = 'to' }
            @{ Pattern = '\bprior to\b';                          Say = 'before' }
            @{ Pattern = '\butili[sz]e[sd]?\b';                   Say = 'use' }
            @{ Pattern = '\bleverag(e|es|ed|ing)\b';              Say = 'use' }
            @{ Pattern = '\bensur(e|es|ed|ing)\b';                Say = 'make sure' }
            @{ Pattern = '\bplease\b';                            Say = '(omit)' }
            @{ Pattern = '\bvery\b';                              Say = '(omit, or give the value)' }
            @{ Pattern = '\b(don|doesn|isn|aren|wasn|weren|can|won|wouldn|shouldn|couldn|hasn|haven|didn)[''\u2019]t\b'; Say = 'the full form (does not, cannot, ...)' }
            @{ Pattern = '\b(it|that|there|what|who)[''\u2019]s\b';     Say = 'the full form (it is, ...)' }
            @{ Pattern = '\b(you|we|they)[''\u2019]re\b';               Say = 'the full form' }
            @{ Pattern = '\blet[''\u2019]s\b';                          Say = '(omit)' }
        )

        function ConvertTo-MaskedProse {
            # Code spans, URLs and bold markers never count as words and never trigger a banned word.
            param([string]$Text)
            $t = [regex]::Replace($Text, '`[^`]*`', ' CODE ')
            $t = [regex]::Replace($t, 'https?://\S+', ' URL ')
            $t = $t -replace '\*\*', ''
            $t
        }

        function Split-Sentence {
            # Sentences of a paragraph: a break after . ! or ? (optionally closed by ) " or ]) that is
            # followed by whitespace and the start of a new sentence.
            param([string]$Paragraph)
            $p = (($Paragraph -replace '\s+', ' ')).Trim()
            if (-not $p) { return @() }
            $parts = [regex]::Split($p, '(?<=[.!?][)"''\]]*)\s+(?=[A-Z0-9"(\[`*_-])')
            foreach ($s in $parts) { $s = $s.Trim(); if ($s) { $s } }
        }

        function Get-WordCount {
            # A token counts as a word only when it holds a letter or a digit: a dash or a lone
            # punctuation mark between words is not a word.
            param([string]$Sentence)
            @(($Sentence.Trim() -split '\s+') | Where-Object { $_ -match '[\p{L}\p{N}]' }).Count
        }

        function Get-HelpParagraph {
            # Paragraphs of the comment-based help block (the first "<#" block that is not
            # PSScriptInfo): section, first line number, and joined text. .EXAMPLE command lines and
            # .LINK entries are skipped; a "- " bullet starts its own paragraph.
            param([string]$Path)
            $lines = @(Get-Content -LiteralPath $Path -Encoding UTF8)   # explicit: 5.1 reads a BOM-less file as ANSI otherwise
            $start = -1; $end = -1
            for ($i = 0; $i -lt $lines.Count; $i++) {
                if ($start -lt 0) { if ($lines[$i] -match '^\s*<#\s*$') { $start = $i }; continue }
                if ($lines[$i] -match '^\s*#>') { $end = $i; break }
            }
            if ($start -lt 0 -or $end -lt 0) { throw "No comment-based help block found in $Path" }
            $paras = New-Object System.Collections.Generic.List[object]
            $section = ''; $buf = ''; $bufLine = 0; $skipCommand = $false
            for ($i = $start + 1; $i -lt $end; $i++) {
                $line = $lines[$i]
                if ($line -match '^\s*\.([A-Z]+)\b') {
                    if ($buf) { $paras.Add([pscustomobject]@{ Section = $section; Line = $bufLine; Text = $buf }); $buf = '' }
                    $section = $Matches[1]; $skipCommand = ($section -eq 'EXAMPLE')
                    continue
                }
                if ($line -match '^\s*$') {
                    if ($buf) { $paras.Add([pscustomobject]@{ Section = $section; Line = $bufLine; Text = $buf }); $buf = '' }
                    continue
                }
                if ($section -eq 'LINK') { continue }
                if ($section -eq 'EXAMPLE' -and ($skipCommand -or $line -match '^\s*(\.\\|PS>|PS |\$|#|[A-Za-z]:\\)')) { $skipCommand = $false; continue }
                if ($line -match '^\s*[-*]\s+') {   # a "- " or "* " bullet starts its own paragraph
                    if ($buf) { $paras.Add([pscustomobject]@{ Section = $section; Line = $bufLine; Text = $buf }) }
                    $buf = ($line -replace '^\s*[-*]\s+', '').Trim(); $bufLine = $i + 1
                    continue
                }
                if ($buf) { $buf += ' ' + $line.Trim() } else { $buf = $line.Trim(); $bufLine = $i + 1 }
            }
            if ($buf) { $paras.Add([pscustomobject]@{ Section = $section; Line = $bufLine; Text = $buf }) }
            $paras
        }

        function Get-MarkdownParagraph {
            # Paragraphs of a Markdown file between two line numbers. Fenced code, tables, headings,
            # HTML and link-reference lines are skipped; a list item or a blockquote line starts its
            # own paragraph.
            param([string]$Path, [int]$FromLine = 1, [int]$ToLine = [int]::MaxValue)
            $lines = @(Get-Content -LiteralPath $Path -Encoding UTF8)   # explicit: 5.1 reads a BOM-less file as ANSI otherwise
            $paras = New-Object System.Collections.Generic.List[object]
            $buf = ''; $bufLine = 0; $inFence = $false
            $last = [math]::Min($ToLine, $lines.Count)
            for ($i = $FromLine - 1; $i -lt $last; $i++) {
                $line = $lines[$i]
                if ($line -match '^\s*```') {
                    $inFence = -not $inFence
                    if ($buf) { $paras.Add([pscustomobject]@{ Section = 'md'; Line = $bufLine; Text = $buf }); $buf = '' }
                    continue
                }
                if ($inFence) { continue }
                if ($line -match '^\s*$' -or $line -match '^\s*(#|\||<|\[[^\]]+\]:\s*http)') {
                    if ($buf) { $paras.Add([pscustomobject]@{ Section = 'md'; Line = $bufLine; Text = $buf }); $buf = '' }
                    continue
                }
                $content = $line -replace '^\s*>\s?', ''
                if ($content -match '^\s*([-*+]|\d+\.)\s+') {
                    if ($buf) { $paras.Add([pscustomobject]@{ Section = 'md'; Line = $bufLine; Text = $buf }) }
                    $buf = ($content -replace '^\s*([-*+]|\d+\.)\s+', '').Trim(); $bufLine = $i + 1
                    continue
                }
                if ($buf) { $buf += ' ' + $content.Trim() } else { $buf = $content.Trim(); $bufLine = $i + 1 }
            }
            if ($buf) { $paras.Add([pscustomobject]@{ Section = 'md'; Line = $bufLine; Text = $buf }) }
            $paras
        }

        function Get-NewestSectionRange {
            # First and last line of the NEWEST section of CHANGELOG.md: [Unreleased] while a release
            # is prepared, the latest release after it ships. Older sections are history, not linted.
            param([string]$Path)
            $lines = @(Get-Content -LiteralPath $Path -Encoding UTF8)   # explicit: 5.1 reads a BOM-less file as ANSI otherwise
            $from = -1; $to = $lines.Count
            for ($i = 0; $i -lt $lines.Count; $i++) {
                if ($from -lt 0) { if ($lines[$i] -match '^## \[') { $from = $i + 1 }; continue }
                if ($lines[$i] -match '^## \[') { $to = $i; break }
            }
            if ($from -lt 0) { throw "CHANGELOG.md has no '## [' section" }
            @($from, $to)
        }

        function Get-LongSentence {
            param([string]$Label, [object[]]$Paragraphs)
            foreach ($p in $Paragraphs) {
                foreach ($s in (Split-Sentence -Paragraph (ConvertTo-MaskedProse -Text $p.Text))) {
                    $n = Get-WordCount -Sentence $s
                    if ($n -gt $script:MaxWords) {
                        $head = if ($s.Length -gt 90) { $s.Substring(0, 90) + '...' } else { $s }
                        "$Label`:$($p.Line) [$n words] $head"
                    }
                }
            }
        }

        function Get-LongParagraph {
            param([string]$Label, [object[]]$Paragraphs)
            foreach ($p in $Paragraphs) {
                $n = @(Split-Sentence -Paragraph (ConvertTo-MaskedProse -Text $p.Text)).Count
                if ($n -gt $script:MaxSentences) { "$Label`:$($p.Line) [$n sentences]" }
            }
        }

        function Get-BannedHit {
            param([string]$Label, [object[]]$Paragraphs)
            foreach ($p in $Paragraphs) {
                $masked = ConvertTo-MaskedProse -Text $p.Text
                foreach ($b in $script:Banned) {
                    $m = [regex]::Match($masked, $b.Pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    if ($m.Success) { "$Label`:$($p.Line) '$($m.Value)' -> $($b.Say)" }
                }
            }
        }

        function Get-HelpParameterBlock {
            # The text of one .PARAMETER block (up to the next section header).
            param([string]$Path, [string]$Name)
            $lines = @(Get-Content -LiteralPath $Path -Encoding UTF8)   # explicit: 5.1 reads a BOM-less file as ANSI otherwise
            $out = New-Object System.Text.StringBuilder
            $in = $false
            foreach ($line in $lines) {
                if ($line -match "^\s*\.PARAMETER\s+$([regex]::Escape($Name))\s*$") { $in = $true; continue }
                if ($in -and $line -match '^\s*\.([A-Z]+)\b') { break }
                if ($in) { [void]$out.AppendLine($line) }
            }
            $out.ToString()
        }

        function Get-CautionBlock {
            # Every blockquote block of a Markdown file that starts with **CAUTION**, joined to one line.
            param([string]$Path)
            $lines = @(Get-Content -LiteralPath $Path -Encoding UTF8)   # explicit: 5.1 reads a BOM-less file as ANSI otherwise
            $blocks = New-Object System.Collections.Generic.List[string]
            $buf = ''
            foreach ($line in $lines) {
                if ($line -match '^\s*>') { $buf += ' ' + ($line -replace '^\s*>\s?', '') }
                elseif ($buf) { $blocks.Add($buf.Trim()); $buf = '' }
            }
            if ($buf) { $blocks.Add($buf.Trim()) }
            @($blocks | Where-Object { $_ -match '^\*\*CAUTION\*\*' })
        }
    }

    Context 'Comment-based help' {
        foreach ($s in $script:StyleScripts) {
            It "$s`: every sentence has 25 words or fewer" -TestCases @{ Script = $s } {
                param($Script)
                $v = @(Get-LongSentence -Label $Script -Paragraphs (Get-HelpParagraph -Path (Join-Path $script:Root $Script)))
                $v.Count | Should -Be 0 -Because ("these sentences are too long:`n" + ($v -join "`n"))
            }
            It "$s`: every paragraph has 6 sentences or fewer" -TestCases @{ Script = $s } {
                param($Script)
                $v = @(Get-LongParagraph -Label $Script -Paragraphs (Get-HelpParagraph -Path (Join-Path $script:Root $Script)))
                $v.Count | Should -Be 0 -Because ("these paragraphs are too long:`n" + ($v -join "`n"))
            }
            It "$s`: uses no banned word or phrase" -TestCases @{ Script = $s } {
                param($Script)
                $v = @(Get-BannedHit -Label $Script -Paragraphs (Get-HelpParagraph -Path (Join-Path $script:Root $Script)))
                $v.Count | Should -Be 0 -Because ("replace these (STYLE.md rule 4):`n" + ($v -join "`n"))
            }
        }
        foreach ($s in $script:RiskSwitches.Keys) {
            foreach ($sw in $script:RiskSwitches[$s]) {
                It "$s`: -$sw has a CAUTION line in its help" -TestCases @{ Script = $s; Switch = $sw } {
                    param($Script, $Switch)
                    $block = Get-HelpParameterBlock -Path (Join-Path $script:Root $Script) -Name $Switch
                    $block | Should -Not -BeNullOrEmpty -Because ".PARAMETER $Switch must exist"
                    $block | Should -Match '(?m)^\s*CAUTION:' -Because "-$Switch accepts a risk the script refuses by default (STYLE.md rule 6)"
                }
            }
        }
    }

    Context 'README.md' {
        BeforeAll { $script:ReadmeParas = @(Get-MarkdownParagraph -Path (Join-Path $script:Root 'README.md')) }
        It 'every sentence has 25 words or fewer' {
            $v = @(Get-LongSentence -Label 'README.md' -Paragraphs $script:ReadmeParas)
            $v.Count | Should -Be 0 -Because ("these sentences are too long:`n" + ($v -join "`n"))
        }
        It 'every paragraph has 6 sentences or fewer' {
            $v = @(Get-LongParagraph -Label 'README.md' -Paragraphs $script:ReadmeParas)
            $v.Count | Should -Be 0 -Because ("these paragraphs are too long:`n" + ($v -join "`n"))
        }
        It 'uses no banned word or phrase' {
            $v = @(Get-BannedHit -Label 'README.md' -Paragraphs $script:ReadmeParas)
            $v.Count | Should -Be 0 -Because ("replace these (STYLE.md rule 4):`n" + ($v -join "`n"))
        }
        foreach ($s in $script:RiskSwitches.Keys) {
            foreach ($sw in $script:RiskSwitches[$s]) {
                It "has a CAUTION block that names -$sw ($s)" -TestCases @{ Switch = $sw } {
                    param($Switch)
                    $blocks = @(Get-CautionBlock -Path (Join-Path $script:Root 'README.md'))
                    @($blocks | Where-Object { $_ -match "-$Switch\b" }).Count | Should -BeGreaterThan 0 -Because "a '> **CAUTION**' block must name -$Switch (STYLE.md rule 6)"
                }
            }
        }
    }

    Context 'CHANGELOG.md newest section' {
        BeforeAll {
            $path = Join-Path $script:Root 'CHANGELOG.md'
            $range = Get-NewestSectionRange -Path $path
            $script:ChangelogParas = @(Get-MarkdownParagraph -Path $path -FromLine $range[0] -ToLine $range[1])
        }
        It 'every sentence has 25 words or fewer' {
            $v = @(Get-LongSentence -Label 'CHANGELOG.md' -Paragraphs $script:ChangelogParas)
            $v.Count | Should -Be 0 -Because ("these sentences are too long:`n" + ($v -join "`n"))
        }
        It 'every paragraph has 6 sentences or fewer' {
            $v = @(Get-LongParagraph -Label 'CHANGELOG.md' -Paragraphs $script:ChangelogParas)
            $v.Count | Should -Be 0 -Because ("these paragraphs are too long:`n" + ($v -join "`n"))
        }
        It 'uses no banned word or phrase' {
            $v = @(Get-BannedHit -Label 'CHANGELOG.md' -Paragraphs $script:ChangelogParas)
            $v.Count | Should -Be 0 -Because ("replace these (STYLE.md rule 4):`n" + ($v -join "`n"))
        }
    }
}
