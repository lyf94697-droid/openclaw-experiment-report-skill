[CmdletBinding()]
param(
  [string]$RoadmapPath,

  [string]$OutputDir,

  [int]$Top = 8,

  [ValidateSet("text", "json", "markdown")]
  [string]$Format = "text"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-RepoRoot {
  return (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}

function Normalize-Text {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return ""
  }

  return ([string]$Text).ToLowerInvariant() -replace "[^a-z0-9]+", " "
}

function Get-Tokens {
  param(
    [AllowNull()]
    [string]$Text
  )

  $stopWords = @(
    "the", "and", "for", "with", "into", "that", "this", "from", "more", "less",
    "when", "then", "than", "such", "keep", "should", "can", "now", "not",
    "profile", "profiles", "report", "reports", "document", "documents",
    "pipeline", "work", "working", "support", "supporting"
  )

  return @(
    (Normalize-Text -Text $Text).Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries) |
      Where-Object { $_.Length -ge 4 -and $_ -notin $stopWords } |
      Select-Object -Unique
  )
}

function Get-RepositoryEvidenceText {
  param(
    [Parameter(Mandatory = $true)]
    [string]$RepoRoot
  )

  $parts = New-Object System.Collections.Generic.List[string]
  foreach ($path in @(
      (Join-Path $RepoRoot "README.md"),
      (Join-Path $RepoRoot "CHANGELOG.md"),
      (Join-Path $RepoRoot "scripts\run-smoke-tests.ps1")
    )) {
    if (Test-Path -LiteralPath $path -PathType Leaf) {
      [void]$parts.Add((Get-Content -LiteralPath $path -Raw -Encoding UTF8))
    }
  }

  $relativePaths = @(
    Get-ChildItem -LiteralPath $RepoRoot -Recurse -File |
      Where-Object {
        $_.FullName -notmatch "\\.git\\" -and
        $_.FullName -notmatch "\\tests-output\\" -and
        $_.FullName -notmatch "\\local-inputs\\" -and
        $_.FullName -notmatch "\\outputs\\"
      } |
      ForEach-Object { $_.FullName.Substring($RepoRoot.Length).TrimStart("\") }
  )
  [void]$parts.Add(($relativePaths -join [Environment]::NewLine))

  return ($parts -join [Environment]::NewLine)
}

function New-Candidate {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Source,

    [Parameter(Mandatory = $true)]
    [string]$RoadmapItem,

    [string]$SuggestedNextStep,

    [string]$SuggestedSmokeCoverage
  )

  [pscustomobject]@{
    source = $Source
    roadmapItem = $RoadmapItem
    suggestedNextStep = $SuggestedNextStep
    suggestedSmokeCoverage = $SuggestedSmokeCoverage
  }
}

function Get-RoadmapBulletCandidates {
  param(
    [AllowEmptyString()]
    [AllowEmptyCollection()]
    [Parameter(Mandatory = $true)]
    [string[]]$Lines
  )

  $candidates = New-Object System.Collections.Generic.List[object]
  $heading = "Roadmap"
  $inImplementedBlock = $false
  $interestingHeading = $false

  foreach ($line in $Lines) {
    if ($line -match "^(#{2,3})\s+(.+)$") {
      $heading = $Matches[2].Trim()
      $inImplementedBlock = $false
      $interestingHeading = $heading -match "Template Fit|Image Planning|Validation|Prompt Assets|Phase 3|Priority"
      continue
    }

    if ($line -match "^\s*Implemented") {
      $inImplementedBlock = $true
      continue
    }

    if ($line -match "^\s*(Status|Definition of done|Continuing goal|Non-Goals)") {
      $inImplementedBlock = $false
    }

    if ($inImplementedBlock -or -not $interestingHeading) {
      continue
    }

    if ($line -match "^\s*-\s+(.+?)\s*$") {
      $item = $Matches[1].Trim()
      if ($item -match "turning the repository|arbitrary unstructured|shipping a GUI") {
        continue
      }

      $nextStep = "Add the smallest script, fixture, or profile-metadata change that proves this roadmap item can move through the existing pipeline."
      $smokeCoverage = "Add a focused assertion to scripts/run-smoke-tests.ps1 and keep it fixture-driven."

      if ($item -match "template|field|blank|compatibility") {
        $nextStep = "Add one template-fit diagnostic fixture or expose one clearer diagnostic field."
        $smokeCoverage = "Assert the new diagnostic code or summary field in the template-fit smoke section."
      } elseif ($item -match "image|layout|2x2|group") {
        $nextStep = "Add one image-planning or grouped-layout fixture before broadening layout behavior."
        $smokeCoverage = "Assert generated image-map layout metadata and final docx image/caption counts."
      } elseif ($item -match "validation|risk|threshold|checks") {
        $nextStep = "Move one validation threshold or remediation hint into explicit profile metadata/output."
        $smokeCoverage = "Assert the machine-readable finding code, category, and summary propagation."
      } elseif ($item -match "prompt|example|fixture|preset") {
        $nextStep = "Add or refresh one prompt/profile sample and make its generated inputs inspectable."
        $smokeCoverage = "Assert the generated prompt, metadata, requirements, or markdown index."
      }

      [void]$candidates.Add((New-Candidate -Source $heading -RoadmapItem $item -SuggestedNextStep $nextStep -SuggestedSmokeCoverage $smokeCoverage))
    }
  }

  return $candidates.ToArray()
}

function Get-ProfileExpansionCandidates {
  param(
    [Parameter(Mandatory = $true)]
    [string]$RepoRoot
  )

  $candidates = New-Object System.Collections.Generic.List[object]
  $orderedProfiles = @("meeting-minutes", "monthly-report")
  foreach ($profileName in $orderedProfiles) {
    $builtInPath = Join-Path $RepoRoot ("profiles\{0}.json" -f $profileName)
    $presetPath = Join-Path $RepoRoot ("examples\profile-presets\{0}.json" -f $profileName)
    if (Test-Path -LiteralPath $builtInPath -PathType Leaf) {
      continue
    }

    $nextStep = if (Test-Path -LiteralPath $presetPath -PathType Leaf) {
      "Run the preset through sample generation, then add the smallest profile-specific pipeline fixture before considering promotion to profiles/."
    } else {
      "Scaffold a profile preset first, then validate it with generated prompt/metadata/requirements samples."
    }

    [void]$candidates.Add((New-Candidate `
        -Source "Recommended Profile Order" `
        -RoadmapItem ("{0} is listed in the recommended expansion order but is not a built-in profile." -f $profileName) `
        -SuggestedNextStep $nextStep `
        -SuggestedSmokeCoverage "Use validate-report-profiles.ps1, run-profile-preset-samples.ps1, and one focused pipeline assertion in run-smoke-tests.ps1."))
  }

  return $candidates.ToArray()
}

function Add-CandidateScoring {
  param(
    [Parameter(Mandatory = $true)]
    [object[]]$Candidates,

    [Parameter(Mandatory = $true)]
    [string]$EvidenceText
  )

  $normalizedEvidence = Normalize-Text -Text $EvidenceText
  $results = New-Object System.Collections.Generic.List[object]

  foreach ($candidate in $Candidates) {
    $item = [string]$candidate.roadmapItem
    $tokens = @(Get-Tokens -Text $item)
    $matchedTokens = @($tokens | Where-Object { $normalizedEvidence -match ("(^|\s)" + [regex]::Escape($_) + "(\s|$)") })
    $coverageSignals = 0
    $smallnessSignals = 0

    if ($item -match "fixture|smoke|test|validation|diagnostic|example|preset|metadata|json|summary|prompt") {
      $coverageSignals += 3
    }
    if ($item -match "profile|threshold|caption|field|template|image|layout") {
      $coverageSignals += 2
    }
    if ($item -match "fixture|example|diagnostic|threshold|metadata|prompt|preset") {
      $smallnessSignals += 3
    }
    if ($item -match "richer|more layout strategies|GUI|arbitrary") {
      $smallnessSignals -= 2
    }

    $implementedSignal = if ($tokens.Count -eq 0) { 0 } else { [math]::Round(($matchedTokens.Count / [double]$tokens.Count), 2) }
    $status = if ($implementedSignal -ge 0.75) {
      "likely-covered-or-partial"
    } else {
      "open"
    }
    $score = [int](($coverageSignals * 3) + ($smallnessSignals * 2) - [math]::Min(6, [int]($implementedSignal * 6)))

    [void]$results.Add([pscustomobject]@{
      status = $status
      priorityScore = $score
      smokeCoverable = ($coverageSignals -ge 3)
      smallChangeLikely = ($smallnessSignals -ge 2)
      source = [string]$candidate.source
      roadmapItem = $item
      suggestedNextStep = [string]$candidate.suggestedNextStep
      suggestedSmokeCoverage = [string]$candidate.suggestedSmokeCoverage
      implementedSignal = $implementedSignal
      matchedEvidenceTokens = @($matchedTokens)
    })
  }

  return @($results | Sort-Object -Property @{ Expression = "status"; Descending = $false }, @{ Expression = "priorityScore"; Descending = $true }, "roadmapItem")
}

function Write-MarkdownReport {
  param(
    [Parameter(Mandatory = $true)]
    [object]$Report,

    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $lines = New-Object System.Collections.Generic.List[string]
  [void]$lines.Add("# Roadmap Daily Triage")
  [void]$lines.Add("")
  [void]$lines.Add(("Generated at: {0}" -f [string]$Report.generatedAt))
  [void]$lines.Add(("Roadmap: {0}" -f [string]$Report.roadmapPath))
  [void]$lines.Add(("Open candidate count: {0}" -f [int]$Report.summary.openCandidateCount))
  [void]$lines.Add("")
  [void]$lines.Add("## Top Smoke-Coverable Candidates")
  [void]$lines.Add("")
  [void]$lines.Add("| Score | Source | Roadmap item | Suggested next step | Smoke coverage |")
  [void]$lines.Add("| ---: | --- | --- | --- | --- |")

  foreach ($candidate in @($Report.topCandidates)) {
    $cells = @(
      [string]$candidate.priorityScore,
      ([string]$candidate.source).Replace("|", "\|"),
      ([string]$candidate.roadmapItem).Replace("|", "\|"),
      ([string]$candidate.suggestedNextStep).Replace("|", "\|"),
      ([string]$candidate.suggestedSmokeCoverage).Replace("|", "\|")
    )
    [void]$lines.Add(("| {0} |" -f ($cells -join " | ")))
  }

  [void]$lines.Add("")
  [void]$lines.Add("## Notes")
  [void]$lines.Add("")
  [void]$lines.Add("- Candidates marked `open` have weaker evidence in README, CHANGELOG, scripts, or smoke tests.")
  [void]$lines.Add("- Prefer the first candidate that can be changed with a small fixture, script output field, or profile metadata edit.")
  [void]$lines.Add("- Keep the implementation small enough to validate in `scripts/run-smoke-tests.ps1`.")

  $parent = Split-Path -Parent $Path
  if (-not [string]::IsNullOrWhiteSpace($parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
  }
  [System.IO.File]::WriteAllText($Path, (($lines -join [Environment]::NewLine) + [Environment]::NewLine), (New-Object System.Text.UTF8Encoding($true)))
}

$repoRoot = Get-RepoRoot
$resolvedRoadmapPath = if ([string]::IsNullOrWhiteSpace($RoadmapPath)) {
  Join-Path $repoRoot "ROADMAP.md"
} else {
  (Resolve-Path -LiteralPath $RoadmapPath).Path
}

if (-not (Test-Path -LiteralPath $resolvedRoadmapPath -PathType Leaf)) {
  throw "Roadmap file was not found: $resolvedRoadmapPath"
}

if ($Top -lt 1) {
  throw "Top must be at least 1."
}

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
  $OutputDir = Join-Path $repoRoot "tests-output\roadmap-triage"
}

$resolvedOutputDir = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null

$roadmapLines = Get-Content -LiteralPath $resolvedRoadmapPath -Encoding UTF8
$evidenceText = Get-RepositoryEvidenceText -RepoRoot $repoRoot
$rawCandidates = @(
  Get-RoadmapBulletCandidates -Lines $roadmapLines
  Get-ProfileExpansionCandidates -RepoRoot $repoRoot
)
$scoredCandidates = @(Add-CandidateScoring -Candidates $rawCandidates -EvidenceText $evidenceText)
$openCandidates = @($scoredCandidates | Where-Object { [string]$_.status -eq "open" })
$topCandidates = @($openCandidates | Sort-Object -Property @{ Expression = "priorityScore"; Descending = $true }, "roadmapItem" | Select-Object -First $Top)

$jsonPath = Join-Path $resolvedOutputDir "roadmap-triage.json"
$markdownPath = Join-Path $resolvedOutputDir "roadmap-triage.md"

$report = [pscustomobject]@{
  generatedAt = (Get-Date).ToString("o")
  roadmapPath = $resolvedRoadmapPath
  outputDir = $resolvedOutputDir
  jsonPath = $jsonPath
  markdownPath = $markdownPath
  summary = [pscustomobject]@{
    candidateCount = $scoredCandidates.Count
    openCandidateCount = $openCandidates.Count
    topCandidateCount = $topCandidates.Count
  }
  topCandidates = @($topCandidates)
  candidates = @($scoredCandidates)
}

Write-MarkdownReport -Report $report -Path $markdownPath
[System.IO.File]::WriteAllText($jsonPath, ($report | ConvertTo-Json -Depth 8), (New-Object System.Text.UTF8Encoding($true)))

if ($Format -eq "json") {
  Write-Output ($report | ConvertTo-Json -Depth 8)
} elseif ($Format -eq "markdown") {
  Write-Output (Get-Content -LiteralPath $markdownPath -Raw -Encoding UTF8)
} else {
  Write-Output ("Roadmap triage written to: {0}" -f $markdownPath)
  foreach ($candidate in @($topCandidates)) {
    Write-Output ("[{0}] {1}" -f ([int]$candidate.priorityScore), ([string]$candidate.roadmapItem))
  }
}
