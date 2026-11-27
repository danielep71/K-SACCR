<#
.SYNOPSIS
Functions of the automatic quick Excel check (#713) that need no Excel: the
guard against forks, the report of a run, the files of its artifact and the
upsert of its pull-request comment.
.DESCRIPTION
Dot-sourced by tools\Run-SACCR-QuickCheck.ps1, which the workflow
.github\workflows\excel-quick.yml runs on the self-hosted runner, and
by tools\test_quick_check_script.py. The file only defines functions; it runs
under Windows PowerShell 5.1 and PowerShell 7.
#>

Set-StrictMode -Version Latest

$script:QuickMarkerPrefix = "<!-- excel-quick-check:"

function Test-QuickRunAllowed {
    # A run starts for a branch of this repository, or by hand: never for the
    # branch of a fork, whose code would run on the self-hosted runner.
    param([string]$EventName, [string]$HeadRepository, [string]$Repository)
    if ($EventName -eq "workflow_dispatch") {
        return $true
    }
    if ($EventName -ne "pull_request") {
        return $false
    }
    if ([string]::IsNullOrWhiteSpace($HeadRepository) -or [string]::IsNullOrWhiteSpace($Repository)) {
        return $false
    }
    return $HeadRepository -ieq $Repository
}

function Get-QuickCheckMode {
    # The check a run makes (#801): "smoke" on a pull request, and on a run by
    # hand unless its mode is "full" or "screenshots" (#616); the full quick
    # check and the screenshots run only by hand.
    param([string]$EventName, [string]$Mode)
    if ($EventName -eq "workflow_dispatch" -and ($Mode -eq "full" -or $Mode -eq "screenshots")) {
        return $Mode
    }
    return "smoke"
}

function Get-QuickBundle {
    # The bundle folder named by the "Bundle:" line of a verification summary,
    # or $null.
    param([string]$Text)
    $match = [regex]::Match($Text, '(?m)^Bundle: (.+?)\s*$')
    if (-not $match.Success) {
        return $null
    }
    return $match.Groups[1].Value
}

function Hide-QuickPaths {
    # Replace the machine's own names in a text meant for a public pull request
    # or artifact: the paths given, longest first, the user and the computer.
    param([string]$Text, [hashtable]$Replacements)
    $result = $Text
    $ordered = @($Replacements.Keys | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                 Sort-Object -Property Length -Descending)
    foreach ($key in $ordered) {
        $replacement = ([string]$Replacements[$key]).Replace('$', '$$')
        $result = [regex]::Replace($result, [regex]::Escape([string]$key), $replacement,
                                   [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    }
    return $result
}

function Get-QuickMode {
    # The mode of a run, from its summary: "smoke", "screenshots" (the smoke
    # check followed by the documentation screenshots, #616) or "full".
    param([string]$Summary)
    if ($Summary -match '(?m)^Mode: smoke.*documentation screenshots') {
        return "screenshots"
    }
    if ($Summary -match '(?m)^Mode: smoke') {
        return "smoke"
    }
    return "full"
}

function Get-QuickMarker {
    # The marker of the comment of one commit and one mode (#828): a smoke
    # check of a pull request never replaces the report of a full quick check
    # run by hand on the same commit, nor the other way round.
    param([string]$Sha, [string]$Mode = "full")
    return $script:QuickMarkerPrefix + $Sha + ":" + $Mode + " -->"
}

function New-QuickReport {
    # The pull-request comment of one run: what it checked, its verdict, the
    # first lines of the summary and the rest folded.
    # Mode: the mode the runner chose (full, smoke, screenshots); without it, the mode the summary
    # names. A run that wrote no summary keeps the mode it was started in (#828, Codex on #830).
    param([string]$Summary, [string]$Sha, [bool]$Passed, [string]$RunUrl, [string]$ArtifactName, [string]$Mode = "")
    $verdict = "FAIL"
    if ($Passed) {
        $verdict = "PASS"
    }
    $short = $Sha
    if ($Sha.Length -ge 7) {
        $short = $Sha.Substring(0, 7)
    }
    # The title follows the summary when it names its mode; the marker follows the mode of the run.
    $summaryMode = Get-QuickMode $Summary
    $markerMode = $Mode
    if ([string]::IsNullOrEmpty($markerMode)) {
        $markerMode = $summaryMode
    }
    if ($Summary -notmatch '(?m)^Mode: ') {
        $summaryMode = $markerMode
    }
    $smoke = $summaryMode -ne "full"
    $kind = "quick"
    $what = "A pull-request check on the self-hosted Windows runner, not a release certification."
    $folded = "Test_RunAll, protection checks and sheet review"
    if ($smoke) {
        $kind = "smoke"
        $what = ("A smoke check of the pull request on the self-hosted Windows runner: the build, the compile and the " +
                 "calculation suites. Every suite, the sheet review and the protection checks run in the full quick " +
                 "check, by hand, and in certification before a release.")
        $folded = "Suites"
    }
    $lines = @($Summary -split "\r?\n")
    $split = [array]::IndexOf($lines, "Test_RunAll (summary.log):")
    if ($split -lt 0) {
        $split = [array]::IndexOf($lines, "Suites (summary.log):")
    }
    $head = $lines
    $rest = @()
    if ($split -gt 0) {
        $head = @($lines[0..($split - 1)])
        $rest = @($lines[$split..($lines.Count - 1)])
    }
    $head = @($head | Where-Object { $_ -notmatch '^Bundle: ' })
    $fence = '```'
    $out = New-Object System.Collections.Generic.List[string]
    $out.Add((Get-QuickMarker -Sha $Sha -Mode $markerMode))
    $out.Add("### Excel $kind check of ``$short``: $verdict")
    $out.Add("")
    $out.Add($what + " It does not replace the maintainer's verification before a VBA or template change is merged, " +
             "unless the maintainer says so here.")
    $out.Add("")
    $out.Add($fence + "text")
    $head | ForEach-Object { $out.Add($_) }
    $out.Add($fence)
    if ($rest.Count -gt 0) {
        $out.Add("")
        $out.Add("<details><summary>$folded</summary>")
        $out.Add("")
        $out.Add($fence + "text")
        $rest | ForEach-Object { $out.Add($_) }
        $out.Add($fence)
        $out.Add("</details>")
    }
    $out.Add("")
    $links = "[Workflow run]($RunUrl)"
    if (-not [string]::IsNullOrWhiteSpace($ArtifactName)) {
        $contents = "the sheet pictures"
        if ($smoke) {
            $contents = "the text results"
        }
        $links += "; $contents are in the run's artifact ``$ArtifactName``, kept for 14 days"
    }
    $out.Add($links + ".")
    $text = ($out -join "`n") + "`n"
    # A comment holds 65,536 characters.
    if ($text.Length -gt 60000) {
        $text = $text.Substring(0, 60000) + "`n... (cut; the whole summary is in the artifact)`n"
    }
    return $text
}

function Get-QuickArtifactFiles {
    # The files of a bundle that go into the artifact: the text results, the
    # benchmark's log when one ran (#771), the sheet pictures and the
    # documentation screenshots (#616), never the workbooks.
    param([string]$Bundle)
    $found = New-Object System.Collections.Generic.List[object]
    foreach ($name in @("verify-summary.txt", "verify.log", "summary.log", "sheet-review.txt", "protection-checks.txt",
                       "benchmark.log", "screenshots.txt")) {
        $path = Join-Path $Bundle $name
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            $found.Add([pscustomobject]@{ Source = $path; Relative = $name; Text = $true })
        }
    }
    foreach ($pictures in @("sheets", "screenshots")) {
        $folder = Join-Path $Bundle $pictures
        if (Test-Path -LiteralPath $folder -PathType Container) {
            foreach ($file in @(Get-ChildItem -LiteralPath $folder -Filter "*.png" -File | Sort-Object Name)) {
                $found.Add([pscustomobject]@{ Source = $file.FullName; Relative = ($pictures + "\" + $file.Name); Text = $false })
            }
        }
    }
    return $found.ToArray()
}

function Copy-QuickArtifacts {
    # Copy those files into the artifact folder; the text files go through
    # Hide-QuickPaths. Returns the number of files copied.
    param([string]$Bundle, [string]$Destination, [hashtable]$Replacements)
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    $count = 0
    foreach ($item in @(Get-QuickArtifactFiles -Bundle $Bundle)) {
        $target = Join-Path $Destination $item.Relative
        $folder = Split-Path -Parent $target
        if (-not (Test-Path -LiteralPath $folder)) {
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
        }
        if ($item.Text) {
            $text = [System.IO.File]::ReadAllText($item.Source)
            [System.IO.File]::WriteAllText($target, (Hide-QuickPaths -Text $text -Replacements $Replacements), $utf8NoBom)
        }
        else {
            Copy-Item -LiteralPath $item.Source -Destination $target -Force
        }
        $count++
    }
    return $count
}

function Find-QuickReportComment {
    # The id of the comment that carries the marker of this commit, or $null.
    # Comments: objects with id and body, as the GitHub API returns them.
    param($Comments, [string]$Marker)
    foreach ($comment in @($Comments)) {
        # Windows PowerShell 5.1 turns an empty JSON array (a pull request with no comment yet) into
        # an empty string, which has no body.
        if ($null -eq $comment -or $comment -is [string]) {
            continue
        }
        $body = $comment.PSObject.Properties["body"]
        if ($null -ne $body -and ([string]$body.Value).Contains($Marker)) {
            return $comment.id
        }
    }
    return $null
}

function Remove-OldQuickBundles {
    # Delete the saccr-verify-* and saccr-smoke-* folders of old runs next to the checkout, so
    # the runner's disk does not fill up. Only direct children whose name is
    # exactly that of a quick bundle, and only when older than the given days.
    param([string]$Parent, [int]$KeepDays, [datetime]$Now = (Get-Date))
    $removed = New-Object System.Collections.Generic.List[string]
    foreach ($folder in @(Get-ChildItem -LiteralPath $Parent -Directory -ErrorAction SilentlyContinue)) {
        if ($folder.Name -notmatch '^saccr-(verify|smoke)-[0-9a-f]{7}(-[0-9]{8}-[0-9]{6})?$') {
            continue
        }
        if ($folder.LastWriteTime -lt $Now.AddDays(-$KeepDays)) {
            Remove-Item -LiteralPath $folder.FullName -Recurse -Force
            $removed.Add($folder.Name)
        }
    }
    return $removed.ToArray()
}

function Invoke-QuickApi {
    # One call to the GitHub API with the job's token.
    param([string]$Method, [string]$Uri, [string]$Token, $Body = $null)
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $headers = @{ Authorization = "Bearer " + $Token; Accept = "application/vnd.github+json"; "X-GitHub-Api-Version" = "2022-11-28" }
    if ($null -eq $Body) {
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers
    }
    $json = $Body | ConvertTo-Json -Compress
    return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -ContentType "application/json; charset=utf-8" `
        -Body ([System.Text.Encoding]::UTF8.GetBytes($json))
}

function Select-QuickPullNumber {
    # The number of the first pull request of a list, as the GitHub API returns
    # it; $null when there is none. Windows PowerShell 5.1 turns an empty JSON
    # array (a branch with no open pull request) into an empty string, which has
    # no number: a run by hand on the release branch stopped on it.
    param($Pulls)
    foreach ($pull in @($Pulls)) {
        if ($null -eq $pull -or $pull -is [string]) {
            continue
        }
        $number = $pull.PSObject.Properties["number"]
        if ($null -ne $number) {
            return [int]$number.Value
        }
    }
    return $null
}

function Find-QuickPullNumber {
    # The open pull request of a branch of this repository, for a run started
    # by hand; $null when there is none.
    param([string]$Repository, [string]$Branch, [string]$Token)
    $owner = $Repository.Split("/")[0]
    $uri = "https://api.github.com/repos/$Repository/pulls?state=open&head=" + [uri]::EscapeDataString($owner + ":" + $Branch)
    return Select-QuickPullNumber -Pulls (Invoke-QuickApi -Method "GET" -Uri $uri -Token $Token)
}

function Publish-QuickReport {
    # One comment per commit and mode: created, or updated when the marker of
    # the report, its first line, is already on the pull request (a re-run of
    # the same mode on the same commit, #828).
    param([string]$Repository, [int]$PullNumber, [string]$Body, [string]$Sha, [string]$Token)
    $base = "https://api.github.com/repos/$Repository"
    $existing = @()
    foreach ($page in 1..5) {
        $batch = @(Invoke-QuickApi -Method "GET" -Uri ("$base/issues/$PullNumber/comments?per_page=100&page=$page") -Token $Token)
        $existing += $batch
        if ($batch.Count -lt 100) {
            break
        }
    }
    $marker = @($Body -split "\r?\n")[0]
    if (-not $marker.StartsWith($script:QuickMarkerPrefix)) {
        $marker = Get-QuickMarker -Sha $Sha
    }
    $id = Find-QuickReportComment -Comments $existing -Marker $marker
    if ($null -ne $id) {
        [void](Invoke-QuickApi -Method "PATCH" -Uri "$base/issues/comments/$id" -Token $Token -Body @{ body = $Body })
        return "updated"
    }
    [void](Invoke-QuickApi -Method "POST" -Uri "$base/issues/$PullNumber/comments" -Token $Token -Body @{ body = $Body })
    return "created"
}
