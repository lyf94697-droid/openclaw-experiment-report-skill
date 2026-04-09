[CmdletBinding()]
param(
  [string[]]$ReferenceUrls,

  [string]$CourseName,

  [string]$ExperimentName,

  [string]$PromptPath,

  [string]$PromptText,

  [string[]]$ReferenceTextPaths,

  [string]$StudentName,

  [string]$StudentId,

  [string]$ClassName,

  [string]$TeacherName,

  [string]$ExperimentProperty,

  [string]$ExperimentDate,

  [string]$ExperimentLocation,

  [string]$ReportProfileName = "experiment-report",

  [string]$ReportProfilePath,

  [string[]]$RequiredKeywords,

  [string]$OutputDir,

  [string]$SummaryPath,

  [string]$OpenClawCmd = $env:OPENCLAW_CMD,

  [string]$BrowserProfile = $env:OPENCLAW_BROWSER_PROFILE,

  [int]$ReferenceMaxChars = 30000,

  [ValidateSet("standard", "full")]
  [string]$DetailLevel = "full"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "report-defaults.ps1")
. (Join-Path $PSScriptRoot "report-profiles.ps1")

function Resolve-AbsolutePathIfProvided {
  param(
    [AllowNull()]
    [string]$Path
  )

  if ([string]::IsNullOrWhiteSpace($Path)) {
    return $null
  }

  return (Resolve-Path -LiteralPath $Path).Path
}

function Get-NonEmptyList {
  param(
    [AllowNull()]
    [string[]]$Values
  )

  return @($Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

if (-not [string]::IsNullOrWhiteSpace($PromptPath) -and -not [string]::IsNullOrWhiteSpace($PromptText)) {
  throw "Provide at most one of -PromptPath or -PromptText."
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$reportProfile = Get-ReportProfile -ProfileName $ReportProfileName -ProfilePath $ReportProfilePath -RepoRoot $repoRoot
$promptLabels = Get-ReportProfilePromptLabels -Profile $reportProfile
$documentLabel = [string]$promptLabels.documentLabel
$courseNameLabel = [string]$promptLabels.courseNameLabel
$titleNameLabel = [string]$promptLabels.titleNameLabel

if ([string]::IsNullOrWhiteSpace($ExperimentProperty)) {
  $ExperimentProperty = [string]$reportProfile.defaultExperimentProperty
}

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
  $OutputDir = Join-Path $repoRoot ("tests-output\report-inputs-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
}

$resolvedOutputDir = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null

$resolvedSummaryPath = if ([string]::IsNullOrWhiteSpace($SummaryPath)) {
  Join-Path $resolvedOutputDir "report-inputs-summary.json"
} else {
  [System.IO.Path]::GetFullPath($SummaryPath)
}
$summaryParent = Split-Path -Parent $resolvedSummaryPath
if (-not [string]::IsNullOrWhiteSpace($summaryParent)) {
  New-Item -ItemType Directory -Path $summaryParent -Force | Out-Null
}

$resolvedPromptPath = Resolve-AbsolutePathIfProvided -Path $PromptPath
$referenceUrlList = Get-NonEmptyList -Values $ReferenceUrls
$referenceTextPathList = Get-NonEmptyList -Values $ReferenceTextPaths
$inferredExperimentName = Resolve-InferredExperimentName `
  -PromptText $PromptText `
  -PromptPath $resolvedPromptPath `
  -ReferenceTextPaths $referenceTextPathList `
  -ReferenceUrls $referenceUrlList

$fetchedReferenceTextPathList = @()
$effectiveReferenceTextPathList = $referenceTextPathList
if ([string]::IsNullOrWhiteSpace($ExperimentName) -and [string]::IsNullOrWhiteSpace($inferredExperimentName) -and $referenceUrlList.Count -gt 0) {
  $fetchedReferenceDir = Join-Path $resolvedOutputDir "references"
  New-Item -ItemType Directory -Path $fetchedReferenceDir -Force | Out-Null

  for ($referenceIndex = 0; $referenceIndex -lt $referenceUrlList.Count; $referenceIndex++) {
    $referenceUrl = [string]$referenceUrlList[$referenceIndex]
    $fetchedReferencePath = Join-Path $fetchedReferenceDir ("reference-{0:D2}.txt" -f ($referenceIndex + 1))
    $fetchedReference = & (Join-Path $repoRoot "scripts\fetch-web-article.ps1") `
      -Url $referenceUrl `
      -BrowserProfile $BrowserProfile `
      -OpenClawCmd $OpenClawCmd `
      -MaxChars $ReferenceMaxChars
    $fetchedReferenceText = (($fetchedReference | Out-String).Trim() + [Environment]::NewLine)
    [System.IO.File]::WriteAllText($fetchedReferencePath, $fetchedReferenceText, (New-Object System.Text.UTF8Encoding($true)))
    $fetchedReferenceTextPathList += $fetchedReferencePath

    if ([string]::IsNullOrWhiteSpace($inferredExperimentName)) {
      $inferredExperimentName = Get-ExperimentNameCandidateFromText -Text $fetchedReferenceText
    }
  }

  $effectiveReferenceTextPathList = @($referenceTextPathList + $fetchedReferenceTextPathList)
}

$resolvedNames = Resolve-ExperimentReportNames `
  -CourseName $CourseName `
  -ExperimentName $ExperimentName `
  -InferredExperimentName $inferredExperimentName `
  -ReportProfileName ([string]$reportProfile.name) `
  -ReportProfilePath ([string]$reportProfile.resolvedProfilePath)
$resolvedCourseName = [string]$resolvedNames.courseName
$resolvedExperimentName = [string]$resolvedNames.experimentName

if ([string]::IsNullOrWhiteSpace($resolvedCourseName) -or [string]::IsNullOrWhiteSpace($resolvedExperimentName)) {
  throw "$courseNameLabel and $titleNameLabel are required unless $titleNameLabel can be inferred from PromptText / PromptPath / ReferenceTextPaths / ReferenceUrls. After you set them once for $documentLabel, later runs can omit them."
}

$basePromptText = if (-not [string]::IsNullOrWhiteSpace($resolvedPromptPath)) {
  Get-Content -LiteralPath $resolvedPromptPath -Raw -Encoding UTF8
} elseif (-not [string]::IsNullOrWhiteSpace($PromptText)) {
  $PromptText
} else {
  New-ReportProfileAutoPromptText `
    -ResolvedCourseName $resolvedCourseName `
    -ResolvedExperimentName $resolvedExperimentName `
    -Profile $reportProfile `
    -DetailLevel $DetailLevel
}

$promptPathOut = Join-Path $resolvedOutputDir "prompt.txt"
[System.IO.File]::WriteAllText($promptPathOut, $basePromptText, (New-Object System.Text.UTF8Encoding($true)))

$metadataPathOut = Join-Path $resolvedOutputDir "metadata.auto.json"
$autoMetadataJson = New-ReportProfileAutoMetadataJson `
  -ResolvedCourseName $resolvedCourseName `
  -ResolvedExperimentName $resolvedExperimentName `
  -Profile $reportProfile `
  -ResolvedStudentName $StudentName `
  -ResolvedStudentId $StudentId `
  -ResolvedClassName $ClassName `
  -ResolvedTeacherName $TeacherName `
  -ResolvedExperimentProperty $ExperimentProperty `
  -ResolvedExperimentDate $ExperimentDate `
  -ResolvedExperimentLocation $ExperimentLocation
[System.IO.File]::WriteAllText($metadataPathOut, $autoMetadataJson, (New-Object System.Text.UTF8Encoding($true)))

$requirementsPathOut = Join-Path $resolvedOutputDir "requirements.auto.json"
$autoRequirementsJson = New-ReportProfileAutoRequirementsJson `
  -ResolvedCourseName $resolvedCourseName `
  -ResolvedExperimentName $resolvedExperimentName `
  -Profile $reportProfile `
  -ExtraKeywords $RequiredKeywords `
  -DetailLevel $DetailLevel
[System.IO.File]::WriteAllText($requirementsPathOut, $autoRequirementsJson, (New-Object System.Text.UTF8Encoding($true)))

$savedDefaultsPath = Save-ExperimentReportDefaults `
  -CourseName $resolvedCourseName `
  -ExperimentName $resolvedExperimentName `
  -DefaultsPath ([string]$resolvedNames.defaultsPath) `
  -ReportProfileName ([string]$reportProfile.name) `
  -ReportProfilePath ([string]$reportProfile.resolvedProfilePath)

$summary = [pscustomobject]@{
  outputDir = $resolvedOutputDir
  reportProfileName = [string]$reportProfile.name
  reportProfilePath = [string]$reportProfile.resolvedProfilePath
  reportProfileDisplayName = $documentLabel
  courseName = $resolvedCourseName
  experimentName = $resolvedExperimentName
  requestedCourseName = $CourseName
  requestedExperimentName = $ExperimentName
  inferredExperimentName = [string]$resolvedNames.inferredExperimentName
  defaultsPath = $savedDefaultsPath
  usedStoredCourseName = [bool]$resolvedNames.usedStoredCourseName
  usedStoredExperimentName = [bool]$resolvedNames.usedStoredExperimentName
  usedInferredExperimentName = [bool]$resolvedNames.usedInferredExperimentName
  detailLevel = $DetailLevel
  promptPath = $promptPathOut
  metadataPath = $metadataPathOut
  requirementsPath = $requirementsPathOut
  referenceUrls = $referenceUrlList
  referenceTextPaths = $effectiveReferenceTextPathList
  fetchedReferenceTextPaths = $fetchedReferenceTextPathList
}
[System.IO.File]::WriteAllText($resolvedSummaryPath, ($summary | ConvertTo-Json -Depth 6), (New-Object System.Text.UTF8Encoding($true)))

Write-Output ("Prompt path: {0}" -f $promptPathOut)
Write-Output ("Metadata path: {0}" -f $metadataPathOut)
Write-Output ("Requirements path: {0}" -f $requirementsPathOut)
Write-Output ("Summary path: {0}" -f $resolvedSummaryPath)
