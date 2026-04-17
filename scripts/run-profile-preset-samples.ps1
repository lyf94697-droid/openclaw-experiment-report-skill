[CmdletBinding()]
param(
  [string]$PresetDir,

  [string[]]$ProfilePath,

  [string]$OutputDir,

  [ValidateSet("standard", "full")]
  [string]$DetailLevel = "full",

  [switch]$UseCurrentDefaults,

  [ValidateSet("text", "json")]
  [string]$Format = "text"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "report-profiles.ps1")

function Get-ProfilePresetPaths {
  param(
    [AllowNull()]
    [string]$PresetDirectory,

    [AllowNull()]
    [string[]]$RequestedProfilePaths
  )

  if ($null -ne $RequestedProfilePaths -and $RequestedProfilePaths.Count -gt 0) {
    return @(
      $RequestedProfilePaths |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { (Resolve-Path -LiteralPath $_).Path }
    )
  }

  if ([string]::IsNullOrWhiteSpace($PresetDirectory)) {
    $repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
    $PresetDirectory = Join-Path $repoRoot "examples\profile-presets"
  }

  $resolvedPresetDirectory = (Resolve-Path -LiteralPath $PresetDirectory).Path
  return @(
    Get-ChildItem -LiteralPath $resolvedPresetDirectory -Filter "*.json" -File |
      Where-Object { $_.Name -ne "report-profile.schema.json" } |
      Sort-Object -Property Name |
      ForEach-Object { $_.FullName }
  )
}

function Get-SampleValuesForProfile {
  param(
    [Parameter(Mandatory = $true)]
    [psobject]$Profile
  )

  $profileName = [string]$Profile.name
  switch ($profileName) {
    "weekly-report" {
      return [pscustomobject]@{
        courseName = "campus-guide-miniapp"
        experimentName = "week-6-iteration-report"
        studentName = "sample-user"
        studentId = "20261234"
        className = "software-engineering-2302"
        teacherName = "sample-reviewer"
        experimentProperty = "project-weekly-report"
        experimentDate = "week 6"
        experimentLocation = "GitHub + Feishu + local development"
        requiredKeywords = @("iteration-progress", "delivery", "next-plan")
      }
    }
    "meeting-minutes" {
      return [pscustomobject]@{
        courseName = "campus-guide-miniapp"
        experimentName = "week-6-milestone-review-minutes"
        studentName = "sample-user"
        studentId = "20261234"
        className = "software-engineering-2302-team"
        teacherName = "sample-host"
        experimentProperty = "review-meeting-minutes"
        experimentDate = "2026-04-17"
        experimentLocation = "online meeting + Feishu document"
        requiredKeywords = @("milestone", "decision", "action-item")
      }
    }
    default {
      $displayName = [string]$Profile.name
      return [pscustomobject]@{
        courseName = "sample-project"
        experimentName = ("{0}-sample" -f $displayName)
        studentName = "sample-user"
        studentId = "20260000"
        className = "sample-team"
        teacherName = "sample-reviewer"
        experimentProperty = $displayName
        experimentDate = "2026-04-17"
        experimentLocation = "local sample environment"
        requiredKeywords = @($displayName)
      }
    }
  }
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$profilePresetPaths = Get-ProfilePresetPaths -PresetDirectory $PresetDir -RequestedProfilePaths $ProfilePath
if ($profilePresetPaths.Count -eq 0) {
  throw "No profile preset JSON files were found."
}

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
  $OutputDir = Join-Path $repoRoot ("tests-output\profile-preset-samples-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
}

$resolvedOutputDir = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null

$originalAgentsHome = $env:AGENTS_HOME
$isolatedAgentsHome = Join-Path $resolvedOutputDir ".agents-home"
if (-not $UseCurrentDefaults) {
  New-Item -ItemType Directory -Path $isolatedAgentsHome -Force | Out-Null
  $env:AGENTS_HOME = $isolatedAgentsHome
}

$generated = New-Object System.Collections.Generic.List[object]

try {
  foreach ($profilePresetPath in $profilePresetPaths) {
    $resolvedProfilePresetPath = (Resolve-Path -LiteralPath $profilePresetPath).Path
    $profile = Get-ReportProfile -ProfilePath $resolvedProfilePresetPath -RepoRoot $repoRoot
    $sampleValues = Get-SampleValuesForProfile -Profile $profile
    $profileOutputDir = Join-Path $resolvedOutputDir ([string]$profile.name)

    $inputArgs = @{
      ReportProfilePath = $resolvedProfilePresetPath
      CourseName = [string]$sampleValues.courseName
      ExperimentName = [string]$sampleValues.experimentName
      StudentName = [string]$sampleValues.studentName
      StudentId = [string]$sampleValues.studentId
      ClassName = [string]$sampleValues.className
      TeacherName = [string]$sampleValues.teacherName
      ExperimentProperty = [string]$sampleValues.experimentProperty
      ExperimentDate = [string]$sampleValues.experimentDate
      ExperimentLocation = [string]$sampleValues.experimentLocation
      RequiredKeywords = @($sampleValues.requiredKeywords)
      OutputDir = $profileOutputDir
      DetailLevel = $DetailLevel
    }

    & (Join-Path $PSScriptRoot "generate-report-inputs.ps1") @inputArgs | Out-Null

    $inputSummaryPath = Join-Path $profileOutputDir "report-inputs-summary.json"
    $inputSummary = (Get-Content -LiteralPath $inputSummaryPath -Raw -Encoding UTF8) | ConvertFrom-Json

    [void]$generated.Add([pscustomobject]@{
      reportProfileName = [string]$inputSummary.reportProfileName
      reportProfileDisplayName = [string]$inputSummary.reportProfileDisplayName
      profilePath = [string]$inputSummary.reportProfilePath
      outputDir = [string]$inputSummary.outputDir
      promptPath = [string]$inputSummary.promptPath
      metadataPath = [string]$inputSummary.metadataPath
      requirementsPath = [string]$inputSummary.requirementsPath
      summaryPath = $inputSummaryPath
      courseName = [string]$inputSummary.courseName
      experimentName = [string]$inputSummary.experimentName
    })
  }
} finally {
  if (-not $UseCurrentDefaults) {
    $env:AGENTS_HOME = $originalAgentsHome
  }
}

$defaultsIsolated = (-not [bool]$UseCurrentDefaults)
$defaultsHome = if ([bool]$UseCurrentDefaults) { [string]$originalAgentsHome } else { $isolatedAgentsHome }
$generatedItems = $generated.ToArray()

$summary = [pscustomobject]@{
  outputDir = $resolvedOutputDir
  detailLevel = $DetailLevel
  defaultsIsolated = $defaultsIsolated
  defaultsHome = $defaultsHome
  generatedCount = $generated.Count
  generated = $generatedItems
}

$summaryPath = Join-Path $resolvedOutputDir "profile-preset-samples-summary.json"
[System.IO.File]::WriteAllText($summaryPath, ($summary | ConvertTo-Json -Depth 8), (New-Object System.Text.UTF8Encoding($true)))
Add-Member -InputObject $summary -MemberType NoteProperty -Name summaryPath -Value $summaryPath -Force

if ($Format -eq "json") {
  Write-Output ($summary | ConvertTo-Json -Depth 8)
} else {
  Write-Output ("Generated profile preset samples: {0}" -f $generated.Count)
  Write-Output ("Output directory: {0}" -f $resolvedOutputDir)
  Write-Output ("Summary path: {0}" -f $summaryPath)
  foreach ($item in $generated) {
    Write-Output ("- {0}: {1}" -f ([string]$item.reportProfileName), ([string]$item.outputDir))
  }
}
