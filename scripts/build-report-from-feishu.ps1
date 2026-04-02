[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$TemplatePath,

  [string]$ReportPath,

  [string[]]$ReferenceUrls,

  [string[]]$ReferenceTextPaths,

  [string]$PromptPath,

  [string]$PromptText,

  [string]$CourseName,

  [string]$ExperimentName,

  [string]$MetadataPath,

  [string]$MetadataJson,

  [string]$StudentName,

  [string]$StudentId,

  [string]$ClassName,

  [string]$TeacherName,

  [string]$ExperimentProperty,

  [string]$ExperimentDate,

  [string]$ExperimentLocation,

  [string]$RequirementsPath,

  [string]$RequirementsJson,

  [string[]]$RequiredKeywords,

  [string[]]$ImagePaths,

  [string]$ImageSpecsPath,

  [string]$ImageSpecsJson,

  [string]$OutputDir,

  [string]$ArtifactsDir,

  [string]$FinalDocxPath,

  [string]$ReportOutPath,

  [string]$SummaryPath,

  [string]$OpenClawCmd = $env:OPENCLAW_CMD,

  [string]$BrowserProfile = $env:OPENCLAW_BROWSER_PROFILE,

  [int]$ReferenceMaxChars = 30000,

  [string]$SessionKey = "agent:gpt:main",

  [switch]$SkipSessionReset,

  [ValidateSet("auto", "default", "compact", "school")]
  [string]$StyleProfile = "auto",

  [string]$StyleProfilePath,

  [ValidateSet("standard", "full")]
  [string]$DetailLevel = "full"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "report-defaults.ps1")

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

function Ensure-ParentDirectory {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $parent = Split-Path -Parent $Path
  if (-not [string]::IsNullOrWhiteSpace($parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
  }
}

function Get-NonEmptyList {
  param(
    [AllowNull()]
    [string[]]$Values
  )

  return @($Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$resolvedTemplatePath = (Resolve-Path -LiteralPath $TemplatePath).Path
$resolvedReportPath = Resolve-AbsolutePathIfProvided -Path $ReportPath
$resolvedPromptPath = Resolve-AbsolutePathIfProvided -Path $PromptPath
$resolvedMetadataPath = Resolve-AbsolutePathIfProvided -Path $MetadataPath
$resolvedRequirementsPath = Resolve-AbsolutePathIfProvided -Path $RequirementsPath
$resolvedImageSpecsPath = Resolve-AbsolutePathIfProvided -Path $ImageSpecsPath
$resolvedStyleProfilePath = Resolve-AbsolutePathIfProvided -Path $StyleProfilePath

$generationInputsProvided = (-not [string]::IsNullOrWhiteSpace($PromptText)) -or `
  (-not [string]::IsNullOrWhiteSpace($resolvedPromptPath)) -or `
  (@(Get-NonEmptyList -Values $ReferenceUrls).Count -gt 0) -or `
  (@(Get-NonEmptyList -Values $ReferenceTextPaths).Count -gt 0)

$resolvedNames = Resolve-ExperimentReportNames -CourseName $CourseName -ExperimentName $ExperimentName
$resolvedCourseName = [string]$resolvedNames.courseName
$resolvedExperimentName = [string]$resolvedNames.experimentName

if (-not [string]::IsNullOrWhiteSpace($resolvedReportPath) -and $generationInputsProvided) {
  throw "Provide either -ReportPath or generation inputs such as -ReferenceUrls / -ReferenceTextPaths / -PromptText, but not both."
}

if ([string]::IsNullOrWhiteSpace($resolvedReportPath)) {
  if ([string]::IsNullOrWhiteSpace($resolvedCourseName) -or [string]::IsNullOrWhiteSpace($resolvedExperimentName)) {
    throw "CourseName and ExperimentName are required on the first generated run. After you set them once, later runs can omit them."
  }
}

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
  $OutputDir = Join-Path $repoRoot ("tests-output\feishu-build-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
}

$resolvedOutputDir = [System.IO.Path]::GetFullPath($OutputDir)
$resolvedArtifactsDir = if ([string]::IsNullOrWhiteSpace($ArtifactsDir)) {
  Join-Path $resolvedOutputDir "artifacts"
} else {
  [System.IO.Path]::GetFullPath($ArtifactsDir)
}
$resolvedFinalDocxPath = if ([string]::IsNullOrWhiteSpace($FinalDocxPath)) { $null } else { [System.IO.Path]::GetFullPath($FinalDocxPath) }
$resolvedReportOutPath = if ([string]::IsNullOrWhiteSpace($ReportOutPath)) { Join-Path $resolvedOutputDir "report.txt" } else { [System.IO.Path]::GetFullPath($ReportOutPath) }
$resolvedSummaryPath = if ([string]::IsNullOrWhiteSpace($SummaryPath)) { Join-Path $resolvedOutputDir "feishu-build-summary.json" } else { [System.IO.Path]::GetFullPath($SummaryPath) }

New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null
New-Item -ItemType Directory -Path $resolvedArtifactsDir -Force | Out-Null
Ensure-ParentDirectory -Path $resolvedReportOutPath
Ensure-ParentDirectory -Path $resolvedSummaryPath
if (-not [string]::IsNullOrWhiteSpace($resolvedFinalDocxPath)) {
  Ensure-ParentDirectory -Path $resolvedFinalDocxPath
}

$referenceUrlList = Get-NonEmptyList -Values $ReferenceUrls
$referenceTextPathList = Get-NonEmptyList -Values $ReferenceTextPaths
$wrapperMode = $null
$innerSummaryPath = $null
$innerSummary = $null
$sourceReportPath = $null

if (-not [string]::IsNullOrWhiteSpace($resolvedReportPath)) {
  $wrapperMode = "local-report"
  $buildParams = @{
    TemplatePath = $resolvedTemplatePath
    ReportPath = $resolvedReportPath
    OutputDir = $resolvedArtifactsDir
    StyleFinalDocx = $true
    StyleProfile = $StyleProfile
  }

  if (-not [string]::IsNullOrWhiteSpace($resolvedMetadataPath)) {
    $buildParams.MetadataPath = $resolvedMetadataPath
  } elseif (-not [string]::IsNullOrWhiteSpace($MetadataJson)) {
    $buildParams.MetadataJson = $MetadataJson
  }

  if (-not [string]::IsNullOrWhiteSpace($resolvedRequirementsPath)) {
    $buildParams.RequirementsPath = $resolvedRequirementsPath
  } elseif (-not [string]::IsNullOrWhiteSpace($RequirementsJson)) {
    $buildParams.RequirementsJson = $RequirementsJson
  }

  if (-not [string]::IsNullOrWhiteSpace($resolvedImageSpecsPath)) {
    $buildParams.ImageSpecsPath = $resolvedImageSpecsPath
  } elseif (-not [string]::IsNullOrWhiteSpace($ImageSpecsJson)) {
    $buildParams.ImageSpecsJson = $ImageSpecsJson
  } elseif ($null -ne $ImagePaths -and $ImagePaths.Count -gt 0) {
    $buildParams.ImagePaths = $ImagePaths
  }

  if (-not [string]::IsNullOrWhiteSpace($resolvedStyleProfilePath)) {
    $buildParams.StyleProfilePath = $resolvedStyleProfilePath
  }

  & (Join-Path $repoRoot "scripts\build-report.ps1") @buildParams | Out-Null
  $innerSummaryPath = Join-Path $resolvedArtifactsDir "summary.json"
  $innerSummary = (Get-Content -LiteralPath $innerSummaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
  $sourceReportPath = $resolvedReportPath
} else {
  $wrapperMode = "generated-report"
  $buildParams = @{
    TemplatePath = $resolvedTemplatePath
    CourseName = $resolvedCourseName
    ExperimentName = $resolvedExperimentName
    OutputDir = $resolvedArtifactsDir
    StyleProfile = $StyleProfile
    DetailLevel = $DetailLevel
    OpenClawCmd = $OpenClawCmd
    BrowserProfile = $BrowserProfile
    ReferenceMaxChars = $ReferenceMaxChars
    SessionKey = $SessionKey
  }

  if (-not [string]::IsNullOrWhiteSpace($resolvedPromptPath)) {
    $buildParams.PromptPath = $resolvedPromptPath
  } elseif (-not [string]::IsNullOrWhiteSpace($PromptText)) {
    $buildParams.PromptText = $PromptText
  }

  if ($referenceUrlList.Count -gt 0) {
    $buildParams.ReferenceUrls = $referenceUrlList
  }
  if ($referenceTextPathList.Count -gt 0) {
    $buildParams.ReferenceTextPaths = $referenceTextPathList
  }

  if (-not [string]::IsNullOrWhiteSpace($resolvedMetadataPath)) {
    $buildParams.MetadataPath = $resolvedMetadataPath
  } elseif (-not [string]::IsNullOrWhiteSpace($MetadataJson)) {
    $buildParams.MetadataJson = $MetadataJson
  }

  if (-not [string]::IsNullOrWhiteSpace($StudentName)) { $buildParams.StudentName = $StudentName }
  if (-not [string]::IsNullOrWhiteSpace($StudentId)) { $buildParams.StudentId = $StudentId }
  if (-not [string]::IsNullOrWhiteSpace($ClassName)) { $buildParams.ClassName = $ClassName }
  if (-not [string]::IsNullOrWhiteSpace($TeacherName)) { $buildParams.TeacherName = $TeacherName }
  if (-not [string]::IsNullOrWhiteSpace($ExperimentProperty)) { $buildParams.ExperimentProperty = $ExperimentProperty }
  if (-not [string]::IsNullOrWhiteSpace($ExperimentDate)) { $buildParams.ExperimentDate = $ExperimentDate }
  if (-not [string]::IsNullOrWhiteSpace($ExperimentLocation)) { $buildParams.ExperimentLocation = $ExperimentLocation }

  if (-not [string]::IsNullOrWhiteSpace($resolvedRequirementsPath)) {
    $buildParams.RequirementsPath = $resolvedRequirementsPath
  } elseif (-not [string]::IsNullOrWhiteSpace($RequirementsJson)) {
    $buildParams.RequirementsJson = $RequirementsJson
  }
  if ($null -ne $RequiredKeywords -and $RequiredKeywords.Count -gt 0) {
    $buildParams.RequiredKeywords = $RequiredKeywords
  }

  if (-not [string]::IsNullOrWhiteSpace($resolvedImageSpecsPath)) {
    $buildParams.ImageSpecsPath = $resolvedImageSpecsPath
  } elseif (-not [string]::IsNullOrWhiteSpace($ImageSpecsJson)) {
    $buildParams.ImageSpecsJson = $ImageSpecsJson
  } elseif ($null -ne $ImagePaths -and $ImagePaths.Count -gt 0) {
    $buildParams.ImagePaths = $ImagePaths
  }

  if (-not [string]::IsNullOrWhiteSpace($resolvedStyleProfilePath)) {
    $buildParams.StyleProfilePath = $resolvedStyleProfilePath
  }

  if ($SkipSessionReset) {
    $buildParams.SkipSessionReset = $true
  }

  & (Join-Path $repoRoot "scripts\build-report-from-url.ps1") @buildParams | Out-Null
  $innerSummaryPath = Join-Path $resolvedArtifactsDir "url-build-summary.json"
  $innerSummary = (Get-Content -LiteralPath $innerSummaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
  $sourceReportPath = [string]$innerSummary.cleanedReportPath
  if ([string]::IsNullOrWhiteSpace($sourceReportPath)) {
    $sourceReportPath = [string]$innerSummary.rawReportPath
  }
}

$generatedFinalDocxPath = [string]$innerSummary.finalDocxPath
if ([string]::IsNullOrWhiteSpace($generatedFinalDocxPath)) {
  throw "The inner build did not produce a final docx path."
}

$resolvedFinalDocxDestination = if ([string]::IsNullOrWhiteSpace($resolvedFinalDocxPath)) {
  Join-Path $resolvedOutputDir ([System.IO.Path]::GetFileName($generatedFinalDocxPath))
} else {
  $resolvedFinalDocxPath
}
Ensure-ParentDirectory -Path $resolvedFinalDocxDestination

if (-not [string]::Equals($generatedFinalDocxPath, $resolvedFinalDocxDestination, [System.StringComparison]::OrdinalIgnoreCase)) {
  Copy-Item -LiteralPath $generatedFinalDocxPath -Destination $resolvedFinalDocxDestination -Force
}

$resolvedCopiedReportPath = $null
if (-not [string]::IsNullOrWhiteSpace($sourceReportPath) -and (Test-Path -LiteralPath $sourceReportPath)) {
  if (-not [string]::Equals($sourceReportPath, $resolvedReportOutPath, [System.StringComparison]::OrdinalIgnoreCase)) {
    Copy-Item -LiteralPath $sourceReportPath -Destination $resolvedReportOutPath -Force
    $resolvedCopiedReportPath = $resolvedReportOutPath
  } else {
    $resolvedCopiedReportPath = $sourceReportPath
  }
}

$wrapperSummary = [pscustomobject]@{
  outputDir = $resolvedOutputDir
  artifactsDir = $resolvedArtifactsDir
  templatePath = $resolvedTemplatePath
  mode = $wrapperMode
  courseName = $resolvedCourseName
  experimentName = $resolvedExperimentName
  requestedCourseName = $CourseName
  requestedExperimentName = $ExperimentName
  defaultsPath = $(if (-not [string]::IsNullOrWhiteSpace($resolvedCourseName) -and -not [string]::IsNullOrWhiteSpace($resolvedExperimentName)) { Save-ExperimentReportDefaults -CourseName $resolvedCourseName -ExperimentName $resolvedExperimentName -DefaultsPath ([string]$resolvedNames.defaultsPath) } else { [string]$resolvedNames.defaultsPath })
  usedStoredCourseName = [bool]$resolvedNames.usedStoredCourseName
  usedStoredExperimentName = [bool]$resolvedNames.usedStoredExperimentName
  detailLevel = $DetailLevel
  referenceMaxChars = $ReferenceMaxChars
  reportPath = $resolvedCopiedReportPath
  finalDocxPath = $resolvedFinalDocxDestination
  summaryPath = $resolvedSummaryPath
  innerSummaryPath = $innerSummaryPath
  referenceUrls = $referenceUrlList
  referenceTextPaths = $referenceTextPathList
  styleProfile = $StyleProfile
  styleProfilePath = $resolvedStyleProfilePath
}
[System.IO.File]::WriteAllText($resolvedSummaryPath, ($wrapperSummary | ConvertTo-Json -Depth 6), (New-Object System.Text.UTF8Encoding($true)))

if (-not [string]::IsNullOrWhiteSpace($resolvedCopiedReportPath)) {
  Write-Output ("Report path: {0}" -f $resolvedCopiedReportPath)
}
Write-Output ("Artifacts dir: {0}" -f $resolvedArtifactsDir)
Write-Output ("Final docx path: {0}" -f $resolvedFinalDocxDestination)
Write-Output ("Summary path: {0}" -f $resolvedSummaryPath)
