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

  [string]$ReportProfileName = "experiment-report",

  [string]$ReportProfilePath,

  [string]$RequirementsPath,

  [string]$RequirementsJson,

  [string[]]$RequiredKeywords,

  [string[]]$ImagePaths,

  [string]$ImageArchiveDir,

  [string]$ImageSpecsPath,

  [string]$ImageSpecsJson,

  [string]$ImagePlanOutPath,

  [string]$OutputDir,

  [string]$ArtifactsDir,

  [string]$FinalDocxPath,

  [string]$TemplateFrameDocxPath,

  [switch]$CreateTemplateFrameDocx,

  [string]$ReportOutPath,

  [string]$SummaryPath,

  [string]$OpenClawCmd = $env:OPENCLAW_CMD,

  [string]$BrowserProfile = $env:OPENCLAW_BROWSER_PROFILE,

  [string]$PreGeneratedReportPath,

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

function Get-SafeFileName {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Name
  )

  $safeName = $Name
  foreach ($invalidChar in [System.IO.Path]::GetInvalidFileNameChars()) {
    $safeName = $safeName.Replace([string]$invalidChar, "_")
  }

  if ([string]::IsNullOrWhiteSpace($safeName)) {
    return "image"
  }

  return $safeName
}

function Copy-InputImagesToArchive {
  param(
    [AllowNull()]
    [string[]]$Paths,

    [Parameter(Mandatory = $true)]
    [string]$ArchiveDir
  )

  $imagePathList = Get-NonEmptyList -Values $Paths
  if ($imagePathList.Count -eq 0) {
    return @()
  }

  New-Item -ItemType Directory -Path $ArchiveDir -Force | Out-Null

  $archivedPaths = New-Object System.Collections.Generic.List[string]
  $index = 0
  foreach ($imagePath in $imagePathList) {
    $index++
    try {
      $resolvedImagePath = (Resolve-Path -LiteralPath $imagePath -ErrorAction Stop).Path
    } catch {
      throw "Image path '$imagePath' could not be resolved. Pass a readable local file path from the [media attached ...] attachment hint."
    }
    if (-not (Test-Path -LiteralPath $resolvedImagePath -PathType Leaf)) {
      throw "Image path '$imagePath' resolved to '$resolvedImagePath', but it is not a file."
    }

    $extension = [System.IO.Path]::GetExtension($resolvedImagePath)
    if ([string]::IsNullOrWhiteSpace($extension)) {
      $extension = ".png"
    }

    $baseName = Get-SafeFileName -Name ([System.IO.Path]::GetFileNameWithoutExtension($resolvedImagePath))
    $destinationName = ("{0:D2}-{1}{2}" -f $index, $baseName, $extension)
    $destinationPath = Join-Path $ArchiveDir $destinationName
    Copy-Item -LiteralPath $resolvedImagePath -Destination $destinationPath -Force
    $archivedPaths.Add((Resolve-Path -LiteralPath $destinationPath).Path) | Out-Null
  }

  return @($archivedPaths.ToArray())
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$reportProfile = Get-ReportProfile -ProfileName $ReportProfileName -ProfilePath $ReportProfilePath -RepoRoot $repoRoot
$labels = Get-ReportProfileLabels -Profile $reportProfile
$documentLabel = Get-ReportProfileDisplayName -Profile $reportProfile -Fallback "报告"
$courseNameLabel = if ($labels.Contains("CourseName") -and -not [string]::IsNullOrWhiteSpace([string]$labels["CourseName"])) { [string]$labels["CourseName"] } else { "课程名称" }
$titleNameLabel = if ($labels.Contains("ExperimentName") -and -not [string]::IsNullOrWhiteSpace([string]$labels["ExperimentName"])) { [string]$labels["ExperimentName"] } else { "题目名称" }
$resolvedTemplatePath = (Resolve-Path -LiteralPath $TemplatePath).Path
$resolvedReportPath = Resolve-AbsolutePathIfProvided -Path $ReportPath
$resolvedPromptPath = Resolve-AbsolutePathIfProvided -Path $PromptPath
$resolvedMetadataPath = Resolve-AbsolutePathIfProvided -Path $MetadataPath
$resolvedRequirementsPath = Resolve-AbsolutePathIfProvided -Path $RequirementsPath
$resolvedPreGeneratedReportPath = Resolve-AbsolutePathIfProvided -Path $PreGeneratedReportPath
$resolvedReportProfilePath = Resolve-AbsolutePathIfProvided -Path $ReportProfilePath
$resolvedImageSpecsPath = Resolve-AbsolutePathIfProvided -Path $ImageSpecsPath
$resolvedImagePlanOutPath = if ([string]::IsNullOrWhiteSpace($ImagePlanOutPath)) { $null } else { [System.IO.Path]::GetFullPath($ImagePlanOutPath) }
$resolvedStyleProfilePath = Resolve-AbsolutePathIfProvided -Path $StyleProfilePath

$generationInputsProvided = (-not [string]::IsNullOrWhiteSpace($PromptText)) -or `
  (-not [string]::IsNullOrWhiteSpace($resolvedPromptPath)) -or `
  (@(Get-NonEmptyList -Values $ReferenceUrls).Count -gt 0) -or `
  (@(Get-NonEmptyList -Values $ReferenceTextPaths).Count -gt 0)

$referenceUrlList = Get-NonEmptyList -Values $ReferenceUrls
$referenceTextPathList = Get-NonEmptyList -Values $ReferenceTextPaths
$inferredExperimentName = Resolve-InferredExperimentName `
  -PromptText $PromptText `
  -PromptPath $resolvedPromptPath `
  -ReferenceTextPaths $referenceTextPathList `
  -ReferenceUrls $referenceUrlList
$resolvedNames = Resolve-ExperimentReportNames `
  -CourseName $CourseName `
  -ExperimentName $ExperimentName `
  -InferredExperimentName $inferredExperimentName `
  -ReportProfileName ([string]$reportProfile.name) `
  -ReportProfilePath ([string]$reportProfile.resolvedProfilePath)
$resolvedCourseName = [string]$resolvedNames.courseName
$resolvedExperimentName = [string]$resolvedNames.experimentName

if (-not [string]::IsNullOrWhiteSpace($resolvedReportPath) -and $generationInputsProvided) {
  throw "Provide either -ReportPath or generation inputs such as -ReferenceUrls / -ReferenceTextPaths / -PromptText, but not both."
}
if (-not [string]::IsNullOrWhiteSpace($resolvedReportPath) -and -not [string]::IsNullOrWhiteSpace($resolvedPreGeneratedReportPath)) {
  throw "Provide either -ReportPath or -PreGeneratedReportPath, but not both."
}

if ([string]::IsNullOrWhiteSpace($resolvedReportPath)) {
  if ([string]::IsNullOrWhiteSpace($resolvedCourseName)) {
    throw "$courseNameLabel is required on the first generated run. $titleNameLabel can be inferred from PromptText / PromptPath / ReferenceTextPaths / ReferenceUrls or reused from saved defaults."
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
$resolvedTemplateFrameDocxPath = if ([string]::IsNullOrWhiteSpace($TemplateFrameDocxPath)) { $null } else { [System.IO.Path]::GetFullPath($TemplateFrameDocxPath) }
$resolvedReportOutPath = if ([string]::IsNullOrWhiteSpace($ReportOutPath)) { Join-Path $resolvedOutputDir "report.txt" } else { [System.IO.Path]::GetFullPath($ReportOutPath) }
$resolvedSummaryPath = if ([string]::IsNullOrWhiteSpace($SummaryPath)) { Join-Path $resolvedOutputDir "feishu-build-summary.json" } else { [System.IO.Path]::GetFullPath($SummaryPath) }

New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null
New-Item -ItemType Directory -Path $resolvedArtifactsDir -Force | Out-Null
Ensure-ParentDirectory -Path $resolvedReportOutPath
Ensure-ParentDirectory -Path $resolvedSummaryPath
if (-not [string]::IsNullOrWhiteSpace($resolvedFinalDocxPath)) {
  Ensure-ParentDirectory -Path $resolvedFinalDocxPath
}
if (-not [string]::IsNullOrWhiteSpace($resolvedTemplateFrameDocxPath)) {
  Ensure-ParentDirectory -Path $resolvedTemplateFrameDocxPath
}

$resolvedImageArchiveDir = if ([string]::IsNullOrWhiteSpace($ImageArchiveDir)) {
  Join-Path $resolvedOutputDir "images"
} else {
  [System.IO.Path]::GetFullPath($ImageArchiveDir)
}
$archivedImagePaths = @()
if ([string]::IsNullOrWhiteSpace($resolvedImageSpecsPath) -and [string]::IsNullOrWhiteSpace($ImageSpecsJson)) {
  $archivedImagePaths = @(Copy-InputImagesToArchive -Paths $ImagePaths -ArchiveDir $resolvedImageArchiveDir)
}

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
  } elseif ($archivedImagePaths.Count -gt 0) {
    $buildParams.ImagePaths = $archivedImagePaths
  } elseif ($null -ne $ImagePaths -and @($ImagePaths).Count -gt 0) {
    $buildParams.ImagePaths = $ImagePaths
  }

  if (-not [string]::IsNullOrWhiteSpace($resolvedStyleProfilePath)) {
    $buildParams.StyleProfilePath = $resolvedStyleProfilePath
  }
  if (-not [string]::IsNullOrWhiteSpace($resolvedImagePlanOutPath)) {
    $buildParams.ImagePlanOutPath = $resolvedImagePlanOutPath
  }
  if (-not [string]::IsNullOrWhiteSpace($ReportProfileName)) {
    $buildParams.ReportProfileName = [string]$reportProfile.name
  }
  if (-not [string]::IsNullOrWhiteSpace([string]$reportProfile.resolvedProfilePath)) {
    $buildParams.ReportProfilePath = [string]$reportProfile.resolvedProfilePath
  }
  if ($CreateTemplateFrameDocx -or (-not [string]::IsNullOrWhiteSpace($resolvedTemplateFrameDocxPath))) {
    $buildParams.CreateTemplateFrameDocx = $true
  }

  & (Join-Path $repoRoot "scripts\build-report.ps1") @buildParams | Out-Null
  $innerSummaryPath = Join-Path $resolvedArtifactsDir "summary.json"
  $innerSummary = (Get-Content -LiteralPath $innerSummaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
  $sourceReportPath = $resolvedReportPath
} else {
  $wrapperMode = "generated-report"
  $preparedInputsSummaryPath = Join-Path $resolvedArtifactsDir "report-inputs-summary.json"
  $inputParams = @{
    OutputDir = $resolvedArtifactsDir
    SummaryPath = $preparedInputsSummaryPath
    DetailLevel = $DetailLevel
    OpenClawCmd = $OpenClawCmd
    BrowserProfile = $BrowserProfile
    ReferenceMaxChars = $ReferenceMaxChars
  }

  if (-not [string]::IsNullOrWhiteSpace($resolvedPromptPath)) {
    $inputParams.PromptPath = $resolvedPromptPath
  } elseif (-not [string]::IsNullOrWhiteSpace($PromptText)) {
    $inputParams.PromptText = $PromptText
  }

  if (@($referenceUrlList).Count -gt 0) {
    $inputParams.ReferenceUrls = $referenceUrlList
  }
  if (@($referenceTextPathList).Count -gt 0) {
    $inputParams.ReferenceTextPaths = $referenceTextPathList
  }

  if (-not [string]::IsNullOrWhiteSpace($CourseName)) { $inputParams.CourseName = $CourseName }
  if (-not [string]::IsNullOrWhiteSpace($ExperimentName)) { $inputParams.ExperimentName = $ExperimentName }
  if (-not [string]::IsNullOrWhiteSpace($StudentName)) { $inputParams.StudentName = $StudentName }
  if (-not [string]::IsNullOrWhiteSpace($StudentId)) { $inputParams.StudentId = $StudentId }
  if (-not [string]::IsNullOrWhiteSpace($ClassName)) { $inputParams.ClassName = $ClassName }
  if (-not [string]::IsNullOrWhiteSpace($TeacherName)) { $inputParams.TeacherName = $TeacherName }
  if (-not [string]::IsNullOrWhiteSpace($ExperimentProperty)) { $inputParams.ExperimentProperty = $ExperimentProperty }
  if (-not [string]::IsNullOrWhiteSpace($ExperimentDate)) { $inputParams.ExperimentDate = $ExperimentDate }
  if (-not [string]::IsNullOrWhiteSpace($ExperimentLocation)) { $inputParams.ExperimentLocation = $ExperimentLocation }
  if ($null -ne $RequiredKeywords -and @($RequiredKeywords).Count -gt 0) { $inputParams.RequiredKeywords = $RequiredKeywords }
  if (-not [string]::IsNullOrWhiteSpace($ReportProfileName)) { $inputParams.ReportProfileName = [string]$reportProfile.name }
  if (-not [string]::IsNullOrWhiteSpace([string]$reportProfile.resolvedProfilePath)) { $inputParams.ReportProfilePath = [string]$reportProfile.resolvedProfilePath }

  & (Join-Path $repoRoot "scripts\generate-report-inputs.ps1") @inputParams | Out-Null

  $buildParams = @{
    TemplatePath = $resolvedTemplatePath
    OutputDir = $resolvedArtifactsDir
    StyleProfile = $StyleProfile
    DetailLevel = $DetailLevel
    OpenClawCmd = $OpenClawCmd
    BrowserProfile = $BrowserProfile
    ReferenceMaxChars = $ReferenceMaxChars
    SessionKey = $SessionKey
    PreparedInputsSummaryPath = $preparedInputsSummaryPath
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
  } elseif ($archivedImagePaths.Count -gt 0) {
    $buildParams.ImagePaths = $archivedImagePaths
  } elseif ($null -ne $ImagePaths -and @($ImagePaths).Count -gt 0) {
    $buildParams.ImagePaths = $ImagePaths
  }

  if (-not [string]::IsNullOrWhiteSpace($resolvedStyleProfilePath)) {
    $buildParams.StyleProfilePath = $resolvedStyleProfilePath
  }
  if (-not [string]::IsNullOrWhiteSpace($resolvedImagePlanOutPath)) {
    $buildParams.ImagePlanOutPath = $resolvedImagePlanOutPath
  }
  if (-not [string]::IsNullOrWhiteSpace($resolvedPreGeneratedReportPath)) {
    $buildParams.PreGeneratedReportPath = $resolvedPreGeneratedReportPath
  }
  if ($CreateTemplateFrameDocx -or (-not [string]::IsNullOrWhiteSpace($resolvedTemplateFrameDocxPath))) {
    $buildParams.CreateTemplateFrameDocx = $true
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
  if ([string]::IsNullOrWhiteSpace($resolvedExperimentName) -and -not [string]::IsNullOrWhiteSpace([string]$innerSummary.experimentName)) {
    $resolvedExperimentName = [string]$innerSummary.experimentName
  }
  if ([string]::IsNullOrWhiteSpace($inferredExperimentName) -and -not [string]::IsNullOrWhiteSpace([string]$innerSummary.inferredExperimentName)) {
    $inferredExperimentName = [string]$innerSummary.inferredExperimentName
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

$generatedTemplateFrameDocxPath = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "templateFrameDocxPath") { [string]$innerSummary.templateFrameDocxPath } else { $null })
$resolvedTemplateFrameDocxDestination = $null
if ($CreateTemplateFrameDocx -or (-not [string]::IsNullOrWhiteSpace($resolvedTemplateFrameDocxPath))) {
  if ([string]::IsNullOrWhiteSpace($generatedTemplateFrameDocxPath) -or -not (Test-Path -LiteralPath $generatedTemplateFrameDocxPath -PathType Leaf)) {
    throw "Template-frame docx was requested, but the inner build did not produce templateFrameDocxPath."
  }

  $resolvedTemplateFrameDocxDestination = if (-not [string]::IsNullOrWhiteSpace($resolvedTemplateFrameDocxPath)) {
    $resolvedTemplateFrameDocxPath
  } else {
    Join-Path $resolvedOutputDir (([System.IO.Path]::GetFileNameWithoutExtension($resolvedFinalDocxDestination)) + ".template-frame.docx")
  }
  Ensure-ParentDirectory -Path $resolvedTemplateFrameDocxDestination

  if (-not [string]::Equals($generatedTemplateFrameDocxPath, $resolvedTemplateFrameDocxDestination, [System.StringComparison]::OrdinalIgnoreCase)) {
    Copy-Item -LiteralPath $generatedTemplateFrameDocxPath -Destination $resolvedTemplateFrameDocxDestination -Force
  }
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

$innerUsedStoredExperimentName = $false
$innerUsedInferredExperimentName = $false
if ($null -ne $innerSummary) {
  if ($null -ne $innerSummary.PSObject.Properties["usedStoredExperimentName"]) {
    $innerUsedStoredExperimentName = [bool]$innerSummary.usedStoredExperimentName
  }
  if ($null -ne $innerSummary.PSObject.Properties["usedInferredExperimentName"]) {
    $innerUsedInferredExperimentName = [bool]$innerSummary.usedInferredExperimentName
  }
}
$generationMode = if ($wrapperMode -eq "local-report") {
  "none"
} elseif ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "generationMode" -and -not [string]::IsNullOrWhiteSpace([string]$innerSummary.generationMode)) {
  [string]$innerSummary.generationMode
} elseif (-not [string]::IsNullOrWhiteSpace($resolvedPreGeneratedReportPath)) {
  "replay"
} else {
  "live"
}
$buildReportInputMode = if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "buildReportInputMode" -and -not [string]::IsNullOrWhiteSpace([string]$innerSummary.buildReportInputMode)) {
  [string]$innerSummary.buildReportInputMode
} elseif ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "reportInputMode" -and -not [string]::IsNullOrWhiteSpace([string]$innerSummary.reportInputMode)) {
  [string]$innerSummary.reportInputMode
} else {
  $null
}
$buildMetadataInputMode = if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "buildMetadataInputMode" -and -not [string]::IsNullOrWhiteSpace([string]$innerSummary.buildMetadataInputMode)) {
  [string]$innerSummary.buildMetadataInputMode
} elseif ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "metadataInputMode" -and -not [string]::IsNullOrWhiteSpace([string]$innerSummary.metadataInputMode)) {
  [string]$innerSummary.metadataInputMode
} else {
  $null
}
$buildRequirementsInputMode = if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "buildRequirementsInputMode" -and -not [string]::IsNullOrWhiteSpace([string]$innerSummary.buildRequirementsInputMode)) {
  [string]$innerSummary.buildRequirementsInputMode
} elseif ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "requirementsInputMode" -and -not [string]::IsNullOrWhiteSpace([string]$innerSummary.requirementsInputMode)) {
  [string]$innerSummary.requirementsInputMode
} else {
  $null
}
$buildImageInputMode = if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "buildImageInputMode" -and -not [string]::IsNullOrWhiteSpace([string]$innerSummary.buildImageInputMode)) {
  [string]$innerSummary.buildImageInputMode
} elseif ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "imageInputMode" -and -not [string]::IsNullOrWhiteSpace([string]$innerSummary.imageInputMode)) {
  [string]$innerSummary.imageInputMode
} else {
  $null
}
$innerValidationErrorCodes = @()
$innerValidationWarningCodes = @()
$innerValidationWarningSummary = @()
if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "validationErrorCodes") {
  $innerValidationErrorCodes = @($innerSummary.validationErrorCodes)
}
if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "validationWarningCodes") {
  $innerValidationWarningCodes = @($innerSummary.validationWarningCodes)
}
if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "validationWarningSummary") {
  $innerValidationWarningSummary = @($innerSummary.validationWarningSummary)
}
$pipelineTracePath = Join-Path $resolvedOutputDir "pipeline-trace.json"
$pipelineTraceMarkdownPath = Join-Path $resolvedOutputDir "pipeline-trace.md"
$pipelineTrace = [pscustomobject]@{
  wrapper = [pscustomobject]@{
    script = "build-report-from-feishu.ps1"
    mode = $wrapperMode
    generationMode = $generationMode
    reportProfileName = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "reportProfileName") { [string]$innerSummary.reportProfileName } else { [string]$reportProfile.name })
    reportProfileDisplayName = $documentLabel
    detailLevel = $DetailLevel
  }
  build = [pscustomobject]@{
    summaryPath = $innerSummaryPath
    reportInputMode = $buildReportInputMode
    metadataInputMode = $buildMetadataInputMode
    requirementsInputMode = $buildRequirementsInputMode
    imageInputMode = $buildImageInputMode
    imagePlanPath = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "imagePlanPath") { [string]$innerSummary.imagePlanPath } else { $null })
    templateFrameDocxPath = $generatedTemplateFrameDocxPath
    layoutCheckPath = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "layoutCheckPath") { [string]$innerSummary.layoutCheckPath } else { $null })
    layoutCheckPassed = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "layoutCheckPassed") { $innerSummary.layoutCheckPassed } else { $null })
    validationPath = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "validationPath") { [string]$innerSummary.validationPath } else { $null })
    validationPassed = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "validationPassed") { $innerSummary.validationPassed } else { $null })
    validationErrorCount = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "validationErrorCount") { $innerSummary.validationErrorCount } else { $null })
    validationWarningCount = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "validationWarningCount") { $innerSummary.validationWarningCount } else { $null })
    validationPaginationRiskCount = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "validationPaginationRiskCount") { $innerSummary.validationPaginationRiskCount } else { $null })
    validationStructuralIssueCount = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "validationStructuralIssueCount") { $innerSummary.validationStructuralIssueCount } else { $null })
    validationWarningCodes = $innerValidationWarningCodes
  }
  artifacts = [pscustomobject]@{
    reportPath = $resolvedCopiedReportPath
    finalDocxPath = $resolvedFinalDocxDestination
    templateFrameDocxPath = $resolvedTemplateFrameDocxDestination
    summaryPath = $resolvedSummaryPath
    innerSummaryPath = $innerSummaryPath
    preGeneratedReportPath = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "preGeneratedReportPath" -and -not [string]::IsNullOrWhiteSpace([string]$innerSummary.preGeneratedReportPath)) { [string]$innerSummary.preGeneratedReportPath } else { $resolvedPreGeneratedReportPath })
  }
}
[System.IO.File]::WriteAllText($pipelineTracePath, ($pipelineTrace | ConvertTo-Json -Depth 8), (New-Object System.Text.UTF8Encoding($true)))
$pipelineTraceMarkdownLines = New-Object System.Collections.Generic.List[string]
[void]$pipelineTraceMarkdownLines.Add("# Pipeline Trace")
[void]$pipelineTraceMarkdownLines.Add("")
[void]$pipelineTraceMarkdownLines.Add("## Wrapper")
[void]$pipelineTraceMarkdownLines.Add("")
[void]$pipelineTraceMarkdownLines.Add("- Script: build-report-from-feishu.ps1")
[void]$pipelineTraceMarkdownLines.Add("- Mode: $wrapperMode")
[void]$pipelineTraceMarkdownLines.Add("- Generation mode: $generationMode")
[void]$pipelineTraceMarkdownLines.Add("- Report profile: $($pipelineTrace.wrapper.reportProfileName)")
[void]$pipelineTraceMarkdownLines.Add("- Display name: $documentLabel")
[void]$pipelineTraceMarkdownLines.Add("- Detail level: $DetailLevel")
[void]$pipelineTraceMarkdownLines.Add("")
[void]$pipelineTraceMarkdownLines.Add("## Build")
[void]$pipelineTraceMarkdownLines.Add("")
[void]$pipelineTraceMarkdownLines.Add("- Summary path: $innerSummaryPath")
[void]$pipelineTraceMarkdownLines.Add("- Report input mode: $buildReportInputMode")
[void]$pipelineTraceMarkdownLines.Add("- Metadata input mode: $buildMetadataInputMode")
[void]$pipelineTraceMarkdownLines.Add("- Requirements input mode: $buildRequirementsInputMode")
[void]$pipelineTraceMarkdownLines.Add("- Image input mode: $buildImageInputMode")
[void]$pipelineTraceMarkdownLines.Add("- Image plan path: $($pipelineTrace.build.imagePlanPath)")
[void]$pipelineTraceMarkdownLines.Add("- Template-frame docx path: $generatedTemplateFrameDocxPath")
[void]$pipelineTraceMarkdownLines.Add("- Layout check path: $($pipelineTrace.build.layoutCheckPath)")
[void]$pipelineTraceMarkdownLines.Add("- Layout check passed: $($pipelineTrace.build.layoutCheckPassed)")
[void]$pipelineTraceMarkdownLines.Add("- Validation path: $($pipelineTrace.build.validationPath)")
[void]$pipelineTraceMarkdownLines.Add("- Validation passed: $($pipelineTrace.build.validationPassed)")
[void]$pipelineTraceMarkdownLines.Add("- Validation warnings: $($pipelineTrace.build.validationWarningCount)")
[void]$pipelineTraceMarkdownLines.Add("- Pagination risks: $($pipelineTrace.build.validationPaginationRiskCount)")
[void]$pipelineTraceMarkdownLines.Add("")
[void]$pipelineTraceMarkdownLines.Add("## Artifacts")
[void]$pipelineTraceMarkdownLines.Add("")
[void]$pipelineTraceMarkdownLines.Add("- Report path: $resolvedCopiedReportPath")
[void]$pipelineTraceMarkdownLines.Add("- Final docx path: $resolvedFinalDocxDestination")
[void]$pipelineTraceMarkdownLines.Add("- Template-frame docx path: $resolvedTemplateFrameDocxDestination")
[void]$pipelineTraceMarkdownLines.Add("- Wrapper summary path: $resolvedSummaryPath")
[void]$pipelineTraceMarkdownLines.Add("- Inner summary path: $innerSummaryPath")
[void]$pipelineTraceMarkdownLines.Add("- Pre-generated report path: $($pipelineTrace.artifacts.preGeneratedReportPath)")
[System.IO.File]::WriteAllText($pipelineTraceMarkdownPath, ($pipelineTraceMarkdownLines -join [Environment]::NewLine), (New-Object System.Text.UTF8Encoding($true)))

$wrapperSummary = [pscustomobject]@{
  outputDir = $resolvedOutputDir
  artifactsDir = $resolvedArtifactsDir
  templatePath = $resolvedTemplatePath
  mode = $wrapperMode
  generationMode = $generationMode
  reportProfileName = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "reportProfileName") { [string]$innerSummary.reportProfileName } else { [string]$reportProfile.name })
  reportProfilePath = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "reportProfilePath") { [string]$innerSummary.reportProfilePath } else { [string]$reportProfile.resolvedProfilePath })
  reportProfileDisplayName = $documentLabel
  courseName = $resolvedCourseName
  experimentName = $resolvedExperimentName
  requestedCourseName = $CourseName
  requestedExperimentName = $ExperimentName
  inferredExperimentName = $inferredExperimentName
  defaultsPath = $(if (-not [string]::IsNullOrWhiteSpace($resolvedCourseName) -and -not [string]::IsNullOrWhiteSpace($resolvedExperimentName)) {
      Save-ExperimentReportDefaults `
        -CourseName $resolvedCourseName `
        -ExperimentName $resolvedExperimentName `
        -DefaultsPath ([string]$resolvedNames.defaultsPath) `
        -ReportProfileName ([string]$reportProfile.name) `
        -ReportProfilePath ([string]$reportProfile.resolvedProfilePath)
    } else {
      [string]$resolvedNames.defaultsPath
    })
  usedStoredCourseName = [bool]$resolvedNames.usedStoredCourseName
  usedStoredExperimentName = $([bool]$resolvedNames.usedStoredExperimentName -or $innerUsedStoredExperimentName)
  usedInferredExperimentName = $([bool]$resolvedNames.usedInferredExperimentName -or $innerUsedInferredExperimentName)
  detailLevel = $DetailLevel
  referenceMaxChars = $ReferenceMaxChars
  reportPath = $resolvedCopiedReportPath
  finalDocxPath = $resolvedFinalDocxDestination
  templateFrameDocxPath = $resolvedTemplateFrameDocxDestination
  buildTemplateFrameDocxPath = $generatedTemplateFrameDocxPath
  imageArchiveDir = $(if ($archivedImagePaths.Count -gt 0) { $resolvedImageArchiveDir } else { $null })
  archivedImagePaths = $archivedImagePaths
  summaryPath = $resolvedSummaryPath
  innerSummaryPath = $innerSummaryPath
  pipelineTracePath = $pipelineTracePath
  pipelineTraceMarkdownPath = $pipelineTraceMarkdownPath
  buildReportInputMode = $buildReportInputMode
  buildMetadataInputMode = $buildMetadataInputMode
  buildRequirementsInputMode = $buildRequirementsInputMode
  buildImageInputMode = $buildImageInputMode
  imagePlanPath = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "imagePlanPath") { [string]$innerSummary.imagePlanPath } else { $null })
  imagePlanLowConfidenceCount = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "imagePlanLowConfidenceCount") { $innerSummary.imagePlanLowConfidenceCount } else { $null })
  imagePlanNeedsReview = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "imagePlanNeedsReview") { $innerSummary.imagePlanNeedsReview } else { $null })
  layoutCheckPath = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "layoutCheckPath") { [string]$innerSummary.layoutCheckPath } else { $null })
  layoutCheckPassed = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "layoutCheckPassed") { $innerSummary.layoutCheckPassed } else { $null })
  layoutCheckMessage = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "layoutCheckMessage") { [string]$innerSummary.layoutCheckMessage } else { $null })
  layoutCheckErrorCount = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "layoutCheckErrorCount") { $innerSummary.layoutCheckErrorCount } else { $null })
  layoutCheckWarningCount = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "layoutCheckWarningCount") { $innerSummary.layoutCheckWarningCount } else { $null })
  validationPath = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "validationPath") { [string]$innerSummary.validationPath } else { $null })
  validationPassed = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "validationPassed") { $innerSummary.validationPassed } else { $null })
  validationErrorCount = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "validationErrorCount") { $innerSummary.validationErrorCount } else { $null })
  validationWarningCount = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "validationWarningCount") { $innerSummary.validationWarningCount } else { $null })
  validationPaginationRiskCount = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "validationPaginationRiskCount") { $innerSummary.validationPaginationRiskCount } else { $null })
  validationStructuralIssueCount = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "validationStructuralIssueCount") { $innerSummary.validationStructuralIssueCount } else { $null })
  validationFindingCountsByCode = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "validationFindingCountsByCode") { $innerSummary.validationFindingCountsByCode } else { $null })
  validationFindingCountsByCategory = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "validationFindingCountsByCategory") { $innerSummary.validationFindingCountsByCategory } else { $null })
  validationErrorCodes = $innerValidationErrorCodes
  validationWarningCodes = $innerValidationWarningCodes
  validationWarningSummary = $innerValidationWarningSummary
  referenceUrls = $referenceUrlList
  referenceTextPaths = $referenceTextPathList
  styleProfile = $StyleProfile
  styleProfilePath = $resolvedStyleProfilePath
  preGeneratedReportPath = $(if ($null -ne $innerSummary -and $innerSummary.PSObject.Properties.Name -contains "preGeneratedReportPath" -and -not [string]::IsNullOrWhiteSpace([string]$innerSummary.preGeneratedReportPath)) { [string]$innerSummary.preGeneratedReportPath } else { $resolvedPreGeneratedReportPath })
}
[System.IO.File]::WriteAllText($resolvedSummaryPath, ($wrapperSummary | ConvertTo-Json -Depth 8), (New-Object System.Text.UTF8Encoding($true)))

if (-not [string]::IsNullOrWhiteSpace($resolvedCopiedReportPath)) {
  Write-Output ("Report path: {0}" -f $resolvedCopiedReportPath)
}
Write-Output ("Artifacts dir: {0}" -f $resolvedArtifactsDir)
Write-Output ("Final docx path: {0}" -f $resolvedFinalDocxDestination)
if (-not [string]::IsNullOrWhiteSpace($resolvedTemplateFrameDocxDestination)) {
  Write-Output ("Template-frame docx path: {0}" -f $resolvedTemplateFrameDocxDestination)
}
Write-Output ("Summary path: {0}" -f $resolvedSummaryPath)
