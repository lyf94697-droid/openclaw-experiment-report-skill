[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Assert-True {
  param(
    [Parameter(Mandatory = $true)]
    [bool]$Condition,

    [Parameter(Mandatory = $true)]
    [string]$Message
  )

  if (-not $Condition) {
    throw $Message
  }
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path

$requiredFiles = @(
  "README.md",
  "docs\README.md",
  "docs\usage-flow.md",
  "docs\template-filling.md",
  "docs\csdn-reference-policy.md",
  "docs\screenshot-evidence.md",
  "examples\README.md",
  "examples\cases\README.md",
  "examples\cases\network-dos\README.md",
  "examples\cases\network-dos\prompt.md",
  "examples\cases\network-dos\metadata.json",
  "examples\cases\network-dos\requirements.json",
  "examples\cases\network-dos\image-specs.json",
  "examples\cases\os-process-scheduling\README.md",
  "examples\cases\os-process-scheduling\prompt.md",
  "examples\cases\os-process-scheduling\metadata.json",
  "examples\cases\os-process-scheduling\requirements.json",
  "examples\cases\course-design-student-management\README.md",
  "examples\cases\course-design-student-management\prompt.md",
  "examples\cases\course-design-student-management\metadata.json",
  "examples\cases\course-design-student-management\requirements.json"
)

foreach ($relativePath in $requiredFiles) {
  $path = Join-Path $repoRoot $relativePath
  Assert-True -Condition (Test-Path -LiteralPath $path -PathType Leaf) -Message "Missing required project-readiness file: $relativePath"
}

$readme = Get-Content -LiteralPath (Join-Path $repoRoot "README.md") -Raw -Encoding UTF8
foreach ($requiredReadmeMarker in @(
    "project-readiness:usage-tutorial",
    "project-readiness:input-output",
    "project-readiness:scenarios",
    "project-readiness:limitations"
  )) {
  $requiredReadmeMarkerPattern = [regex]::Escape($requiredReadmeMarker)
  $hasRequiredReadmeMarker = $readme -match $requiredReadmeMarkerPattern
  Assert-True -Condition $hasRequiredReadmeMarker -Message "README is missing marker: $requiredReadmeMarker"
}

$docsReadme = Get-Content -LiteralPath (Join-Path $repoRoot "docs\README.md") -Raw -Encoding UTF8
foreach ($requiredDocLink in @("usage-flow.md", "template-filling.md", "csdn-reference-policy.md", "screenshot-evidence.md")) {
  $requiredDocLinkPattern = [regex]::Escape($requiredDocLink)
  $hasRequiredDocLink = $docsReadme -match $requiredDocLinkPattern
  Assert-True -Condition $hasRequiredDocLink -Message "docs/README.md is missing link: $requiredDocLink"
}

$examplesReadme = Get-Content -LiteralPath (Join-Path $repoRoot "examples\README.md") -Raw -Encoding UTF8
foreach ($requiredCase in @("network-dos", "os-process-scheduling", "course-design-student-management")) {
  $requiredCasePattern = [regex]::Escape($requiredCase)
  $hasRequiredCase = $examplesReadme -match $requiredCasePattern
  Assert-True -Condition $hasRequiredCase -Message "examples/README.md is missing case: $requiredCase"
}

$csdnDoc = Get-Content -LiteralPath (Join-Path $repoRoot "docs\csdn-reference-policy.md") -Raw -Encoding UTF8
$csdnDocHasCopyMarker = $csdnDoc -match "project-readiness:do-not-copy"
$csdnDocHasFactMarker = $csdnDoc -match "project-readiness:fact-source"
Assert-True -Condition $csdnDocHasCopyMarker -Message "CSDN policy doc should include the do-not-copy marker."
Assert-True -Condition $csdnDocHasFactMarker -Message "CSDN policy doc should distinguish facts from wording."

$screenshotDoc = Get-Content -LiteralPath (Join-Path $repoRoot "docs\screenshot-evidence.md") -Raw -Encoding UTF8
$screenshotDocMentionsSpecs = $screenshotDoc -match "image-specs"
$screenshotDocMentionsCaptions = $screenshotDoc -match "project-readiness:caption-policy"
Assert-True -Condition $screenshotDocMentionsSpecs -Message "Screenshot evidence doc should mention image-specs."
Assert-True -Condition $screenshotDocMentionsCaptions -Message "Screenshot evidence doc should mention captions."

$jsonFiles = Get-ChildItem -LiteralPath (Join-Path $repoRoot "examples\cases") -Recurse -Filter "*.json" -File
foreach ($jsonFile in $jsonFiles) {
  try {
    Get-Content -LiteralPath $jsonFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json | Out-Null
  } catch {
    throw "Invalid JSON in $($jsonFile.FullName): $($_.Exception.Message)"
  }
}

Write-Output "Project readiness check passed."
