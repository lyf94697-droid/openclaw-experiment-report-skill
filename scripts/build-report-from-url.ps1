[CmdletBinding()]
param(
  [string[]]$ReferenceUrls,

  [string]$CourseName,

  [string]$ExperimentName,

  [string]$PromptPath,

  [string]$PromptText,

  [string[]]$ReferenceTextPaths,

  [string]$TemplatePath,

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

  [string]$ImageSpecsPath,

  [string]$ImageSpecsJson,

  [string]$ImagePlanOutPath,

  [string]$OutputDir,

  [string]$FinalDocxPath,

  [string]$OpenClawCmd = $env:OPENCLAW_CMD,

  [string]$BrowserProfile = $env:OPENCLAW_BROWSER_PROFILE,

  [string]$SessionKey = "agent:gpt:main",

  [switch]$SkipSessionReset,

  [ValidateSet("auto", "default", "compact", "school")]
  [string]$StyleProfile = "auto",

  [string]$StyleProfilePath,

  [int]$ReferenceMaxChars = 30000,

  [ValidateSet("standard", "full")]
  [string]$DetailLevel = "full",

  [string]$PreparedInputsSummaryPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "report-defaults.ps1")
. (Join-Path $PSScriptRoot "report-profiles.ps1")

function Read-MetadataDocument {
  param(
    [AllowNull()]
    [string]$ResolvedMetadataPath,

    [AllowNull()]
    [string]$InlineMetadataJson
  )

  if (-not [string]::IsNullOrWhiteSpace($ResolvedMetadataPath)) {
    return (Get-Content -LiteralPath $ResolvedMetadataPath -Raw -Encoding UTF8 | ConvertFrom-Json)
  }

  if (-not [string]::IsNullOrWhiteSpace($InlineMetadataJson)) {
    return ($InlineMetadataJson | ConvertFrom-Json)
  }

  return $null
}

function Normalize-GeneratedReportText {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Text,

    [Parameter(Mandatory = $true)]
    [hashtable]$Labels
  )

  $normalized = $Text.Trim()
  if ([string]::IsNullOrWhiteSpace($normalized)) {
    return ""
  }

  $knownHeadingCount = 0
  foreach ($heading in @($Labels.Purpose, $Labels.Environment, $Labels.Theory, $Labels.Steps, $Labels.Results, $Labels.Analysis, $Labels.Summary)) {
    if ($normalized.Contains($heading)) {
      $knownHeadingCount++
    }
  }

  if ($knownHeadingCount -lt 3) {
    return ($normalized + [Environment]::NewLine)
  }

  $mojibakeHintChars = @([char]0x7039, [char]0x95C2, [char]0x940E)
  $lines = $normalized -split "\r?\n"
  $cleanedLines = New-Object System.Collections.Generic.List[string]

  foreach ($line in $lines) {
    if (-not [string]::IsNullOrWhiteSpace($line) -and $line.Length -ge 10) {
      $firstChar = $line[0]
      if ($mojibakeHintChars -contains $firstChar) {
        break
      }
    }

    [void]$cleanedLines.Add($line)
  }

  return ((($cleanedLines -join [Environment]::NewLine).TrimEnd()) + [Environment]::NewLine)
}

function ConvertTo-SafeFilenameSegment {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return ""
  }

  $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
  $builder = New-Object System.Text.StringBuilder
  foreach ($character in $Text.ToCharArray()) {
    if ($invalidChars -contains $character) {
      [void]$builder.Append('_')
    } else {
      [void]$builder.Append($character)
    }
  }

  return ($builder.ToString().Trim())
}

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

if (-not [string]::IsNullOrWhiteSpace($PromptPath) -and -not [string]::IsNullOrWhiteSpace($PromptText)) {
  throw "Provide at most one of -PromptPath or -PromptText."
}

if (-not [string]::IsNullOrWhiteSpace($MetadataPath) -and -not [string]::IsNullOrWhiteSpace($MetadataJson)) {
  throw "Provide at most one of -MetadataPath or -MetadataJson."
}

if (-not [string]::IsNullOrWhiteSpace($RequirementsPath) -and -not [string]::IsNullOrWhiteSpace($RequirementsJson)) {
  throw "Provide at most one of -RequirementsPath or -RequirementsJson."
}

$generationInputsProvided = (-not [string]::IsNullOrWhiteSpace($PromptText)) -or `
  (-not [string]::IsNullOrWhiteSpace($PromptPath)) -or `
  (@($ReferenceUrls | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count -gt 0) -or `
  (@($ReferenceTextPaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count -gt 0) -or `
  (-not [string]::IsNullOrWhiteSpace($CourseName)) -or `
  (-not [string]::IsNullOrWhiteSpace($ExperimentName)) -or `
  (-not [string]::IsNullOrWhiteSpace($StudentName)) -or `
  (-not [string]::IsNullOrWhiteSpace($StudentId)) -or `
  (-not [string]::IsNullOrWhiteSpace($ClassName)) -or `
  (-not [string]::IsNullOrWhiteSpace($TeacherName)) -or `
  (-not [string]::IsNullOrWhiteSpace($ExperimentProperty)) -or `
  (-not [string]::IsNullOrWhiteSpace($ExperimentDate)) -or `
  (-not [string]::IsNullOrWhiteSpace($ExperimentLocation)) -or `
  (@($RequiredKeywords | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count -gt 0)

if (-not [string]::IsNullOrWhiteSpace($PreparedInputsSummaryPath) -and $generationInputsProvided) {
  throw "Provide either -PreparedInputsSummaryPath or generation inputs such as -ReferenceUrls / -PromptText / -CourseName, but not both."
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$preparedInputsContext = Resolve-PreparedInputsSummaryContext `
  -PreparedInputsSummaryPath $PreparedInputsSummaryPath `
  -ReportProfileName $ReportProfileName `
  -ReportProfilePath $ReportProfilePath `
  -ReportProfileNameProvided:$PSBoundParameters.ContainsKey("ReportProfileName") `
  -ReportProfilePathProvided:$PSBoundParameters.ContainsKey("ReportProfilePath") `
  -DetailLevel $DetailLevel `
  -DetailLevelProvided:$PSBoundParameters.ContainsKey("DetailLevel")
$resolvedPreparedInputsSummaryPath = [string]$preparedInputsContext.resolvedPreparedInputsSummaryPath
$reportProfile = Get-ReportProfile -ProfileName ([string]$preparedInputsContext.reportProfileName) -ProfilePath ([string]$preparedInputsContext.reportProfilePath) -RepoRoot $repoRoot
$labels = Get-ReportProfileLabels -Profile $reportProfile
$documentLabel = Get-ReportProfileDisplayName -Profile $reportProfile -Fallback "报告"
$courseNameLabel = if ($labels.Contains("CourseName") -and -not [string]::IsNullOrWhiteSpace([string]$labels["CourseName"])) { [string]$labels["CourseName"] } else { "课程名称" }
$titleNameLabel = if ($labels.Contains("ExperimentName") -and -not [string]::IsNullOrWhiteSpace([string]$labels["ExperimentName"])) { [string]$labels["ExperimentName"] } else { "题目名称" }
if ([string]::IsNullOrWhiteSpace($ExperimentProperty)) {
  $ExperimentProperty = [string]$reportProfile.defaultExperimentProperty
}

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
  $OutputDir = Join-Path $repoRoot ("tests-output\url-build-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
}

$resolvedOutputDir = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null

$resolvedTemplatePath = Resolve-AbsolutePathIfProvided -Path $TemplatePath
$resolvedMetadataPath = Resolve-AbsolutePathIfProvided -Path $MetadataPath
$resolvedRequirementsPath = Resolve-AbsolutePathIfProvided -Path $RequirementsPath
$resolvedStyleProfilePath = Resolve-AbsolutePathIfProvided -Path $StyleProfilePath
$resolvedImagePlanOutPath = if ([string]::IsNullOrWhiteSpace($ImagePlanOutPath)) { $null } else { [System.IO.Path]::GetFullPath($ImagePlanOutPath) }
$resolvedFinalDocxPath = if ([string]::IsNullOrWhiteSpace($FinalDocxPath)) { $null } else { [System.IO.Path]::GetFullPath($FinalDocxPath) }
$inputSummaryPath = if ([string]::IsNullOrWhiteSpace($resolvedPreparedInputsSummaryPath)) {
  Join-Path $resolvedOutputDir "report-inputs-summary.json"
} else {
  $resolvedPreparedInputsSummaryPath
}

if ([string]::IsNullOrWhiteSpace($resolvedPreparedInputsSummaryPath)) {
  $inputParams = @{
    OutputDir = $resolvedOutputDir
    SummaryPath = $inputSummaryPath
    OpenClawCmd = $OpenClawCmd
    BrowserProfile = $BrowserProfile
    ReferenceMaxChars = $ReferenceMaxChars
    DetailLevel = $DetailLevel
    ReportProfileName = [string]$reportProfile.name
    ReportProfilePath = [string]$reportProfile.resolvedProfilePath
  }

  if (-not [string]::IsNullOrWhiteSpace($PromptPath)) { $inputParams.PromptPath = $PromptPath }
  if (-not [string]::IsNullOrWhiteSpace($PromptText)) { $inputParams.PromptText = $PromptText }
  if ($null -ne $ReferenceUrls -and @($ReferenceUrls).Count -gt 0) { $inputParams.ReferenceUrls = $ReferenceUrls }
  if ($null -ne $ReferenceTextPaths -and @($ReferenceTextPaths).Count -gt 0) { $inputParams.ReferenceTextPaths = $ReferenceTextPaths }
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

  & (Join-Path $repoRoot "scripts\generate-report-inputs.ps1") @inputParams | Out-Null
}

$inputSummary = (Get-Content -LiteralPath $inputSummaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
$resolvedCourseName = [string]$inputSummary.courseName
$resolvedExperimentName = [string]$inputSummary.experimentName
$effectiveDetailLevel = [string]$preparedInputsContext.detailLevel
$promptPathOut = [string]$inputSummary.promptPath
$effectiveReferenceUrlList = @($inputSummary.referenceUrls | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
$effectiveReferenceTextPathList = @($inputSummary.referenceTextPaths | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
$fetchedReferenceTextPathList = @($inputSummary.fetchedReferenceTextPaths | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
$savedDefaultsPath = [string]$inputSummary.defaultsPath

$effectiveMetadataPath = if (-not [string]::IsNullOrWhiteSpace($resolvedMetadataPath)) {
  $resolvedMetadataPath
} elseif (-not [string]::IsNullOrWhiteSpace($MetadataJson)) {
  $null
} else {
  [string]$inputSummary.metadataPath
}

$effectiveRequirementsPath = if (-not [string]::IsNullOrWhiteSpace($resolvedRequirementsPath)) {
  $resolvedRequirementsPath
} elseif (-not [string]::IsNullOrWhiteSpace($RequirementsJson)) {
  $null
} else {
  [string]$inputSummary.requirementsPath
}

$metadataDocument = Read-MetadataDocument -ResolvedMetadataPath $effectiveMetadataPath -InlineMetadataJson $MetadataJson
$studentIdForFilename = if (-not [string]::IsNullOrWhiteSpace($StudentId)) {
  $StudentId
} elseif ($null -ne $metadataDocument) {
  [string]$metadataDocument.($labels.StudentId)
} else {
  $null
}
$studentNameForFilename = if (-not [string]::IsNullOrWhiteSpace($StudentName)) {
  $StudentName
} elseif ($null -ne $metadataDocument) {
  [string]$metadataDocument.($labels.Name)
} else {
  $null
}

$rawReportPath = Join-Path $resolvedOutputDir "report.raw.txt"
$cleanedReportPath = Join-Path $resolvedOutputDir "report.cleaned.txt"
& (Join-Path $repoRoot "scripts\generate-report-chat.ps1") `
  -PromptPath $promptPathOut `
  -ReferenceTextPaths $effectiveReferenceTextPathList `
  -ReferenceUrls $effectiveReferenceUrlList `
  -BrowserProfile $BrowserProfile `
  -OpenClawCmd $OpenClawCmd `
  -ReferenceMaxChars $ReferenceMaxChars `
  -SessionKey $SessionKey `
  -OutFile $rawReportPath `
  $(if ($SkipSessionReset) { "-SkipSessionReset" }) | Out-Null

$rawReportText = Get-Content -LiteralPath $rawReportPath -Raw -Encoding UTF8
$cleanedReportText = Normalize-GeneratedReportText -Text $rawReportText -Labels $labels
[System.IO.File]::WriteAllText($cleanedReportPath, $cleanedReportText, (New-Object System.Text.UTF8Encoding($true)))

$buildSummaryPath = $null
$buildSummary = $null
$finalDocxPath = $null
if (-not [string]::IsNullOrWhiteSpace($resolvedTemplatePath)) {
  $buildParams = @{
    TemplatePath = $resolvedTemplatePath
    ReportPath = $cleanedReportPath
    OutputDir = $resolvedOutputDir
    StyleFinalDocx = $true
    StyleProfile = $StyleProfile
  }

  if (-not [string]::IsNullOrWhiteSpace($effectiveMetadataPath)) {
    $buildParams.MetadataPath = $effectiveMetadataPath
  } elseif (-not [string]::IsNullOrWhiteSpace($MetadataJson)) {
    $buildParams.MetadataJson = $MetadataJson
  }

  if (-not [string]::IsNullOrWhiteSpace($effectiveRequirementsPath)) {
    $buildParams.RequirementsPath = $effectiveRequirementsPath
  } elseif (-not [string]::IsNullOrWhiteSpace($RequirementsJson)) {
    $buildParams.RequirementsJson = $RequirementsJson
  }

  if (-not [string]::IsNullOrWhiteSpace($ImageSpecsPath)) {
    $buildParams.ImageSpecsPath = $ImageSpecsPath
  } elseif (-not [string]::IsNullOrWhiteSpace($ImageSpecsJson)) {
    $buildParams.ImageSpecsJson = $ImageSpecsJson
  } elseif ($null -ne $ImagePaths -and @($ImagePaths).Count -gt 0) {
    $buildParams.ImagePaths = $ImagePaths
  }

  if (-not [string]::IsNullOrWhiteSpace($resolvedStyleProfilePath)) {
    $buildParams.StyleProfilePath = $resolvedStyleProfilePath
  }
  if (-not [string]::IsNullOrWhiteSpace([string]$reportProfile.name)) {
    $buildParams.ReportProfileName = [string]$reportProfile.name
  }
  if (-not [string]::IsNullOrWhiteSpace([string]$reportProfile.resolvedProfilePath)) {
    $buildParams.ReportProfilePath = [string]$reportProfile.resolvedProfilePath
  }
  if (-not [string]::IsNullOrWhiteSpace($resolvedImagePlanOutPath)) {
    $buildParams.ImagePlanOutPath = $resolvedImagePlanOutPath
  }

  & (Join-Path $repoRoot "scripts\build-report.ps1") @buildParams | Out-Null
  $buildSummaryPath = Join-Path $resolvedOutputDir "summary.json"
  $buildSummary = (Get-Content -LiteralPath $buildSummaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
  $generatedFinalDocxPath = [string]$buildSummary.finalDocxPath
  if ([string]::IsNullOrWhiteSpace($generatedFinalDocxPath)) {
    throw "build-report.ps1 did not produce a finalDocxPath in $buildSummaryPath"
  }

  if (-not [string]::IsNullOrWhiteSpace($resolvedFinalDocxPath)) {
    $finalDocxPath = $resolvedFinalDocxPath
  } else {
    $fileNameParts = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($studentIdForFilename)) {
      [void]$fileNameParts.Add((ConvertTo-SafeFilenameSegment -Text $studentIdForFilename))
    }
    if (-not [string]::IsNullOrWhiteSpace($studentNameForFilename)) {
      [void]$fileNameParts.Add((ConvertTo-SafeFilenameSegment -Text $studentNameForFilename))
    }
    [void]$fileNameParts.Add((ConvertTo-SafeFilenameSegment -Text $resolvedExperimentName))
    $suffix = $labels.FinalEdition
    $finalDocxPath = Join-Path $resolvedOutputDir (($fileNameParts -join "-") + "-$suffix.docx")
  }

  $finalDocxParent = Split-Path -Parent $finalDocxPath
  if (-not [string]::IsNullOrWhiteSpace($finalDocxParent)) {
    New-Item -ItemType Directory -Path $finalDocxParent -Force | Out-Null
  }

  if (-not [string]::Equals($generatedFinalDocxPath, $finalDocxPath, [System.StringComparison]::OrdinalIgnoreCase)) {
    Copy-Item -LiteralPath $generatedFinalDocxPath -Destination $finalDocxPath -Force
  } else {
    $finalDocxPath = $generatedFinalDocxPath
  }
}

$wrapperSummaryPath = Join-Path $resolvedOutputDir "url-build-summary.json"
$wrapperSummary = [pscustomobject]@{
  outputDir = $resolvedOutputDir
  reportProfileName = [string]$reportProfile.name
  reportProfilePath = [string]$reportProfile.resolvedProfilePath
  reportProfileDisplayName = $documentLabel
  courseName = $resolvedCourseName
  experimentName = $resolvedExperimentName
  requestedCourseName = $(if ($inputSummary.PSObject.Properties.Name -contains "requestedCourseName") { [string]$inputSummary.requestedCourseName } else { $CourseName })
  requestedExperimentName = $(if ($inputSummary.PSObject.Properties.Name -contains "requestedExperimentName") { [string]$inputSummary.requestedExperimentName } else { $ExperimentName })
  inferredExperimentName = $(if ($inputSummary.PSObject.Properties.Name -contains "inferredExperimentName") { [string]$inputSummary.inferredExperimentName } else { $null })
  defaultsPath = $savedDefaultsPath
  usedStoredCourseName = $(if ($inputSummary.PSObject.Properties.Name -contains "usedStoredCourseName") { [bool]$inputSummary.usedStoredCourseName } else { $false })
  usedStoredExperimentName = $(if ($inputSummary.PSObject.Properties.Name -contains "usedStoredExperimentName") { [bool]$inputSummary.usedStoredExperimentName } else { $false })
  usedInferredExperimentName = $(if ($inputSummary.PSObject.Properties.Name -contains "usedInferredExperimentName") { [bool]$inputSummary.usedInferredExperimentName } else { $false })
  detailLevel = $effectiveDetailLevel
  promptPath = $promptPathOut
  requestedReferenceUrls = $(if ($inputSummary.PSObject.Properties.Name -contains "requestedReferenceUrls") { @($inputSummary.requestedReferenceUrls) } else { @() })
  referenceUrls = $effectiveReferenceUrlList
  referenceTextPaths = $effectiveReferenceTextPathList
  fetchedReferenceTextPaths = $fetchedReferenceTextPathList
  reportInputsSummaryPath = $inputSummaryPath
  rawReportPath = $rawReportPath
  cleanedReportPath = $cleanedReportPath
  metadataPath = $effectiveMetadataPath
  requirementsPath = $effectiveRequirementsPath
  styleProfile = $StyleProfile
  styleProfilePath = $resolvedStyleProfilePath
  buildSummaryPath = $buildSummaryPath
  imagePlanPath = $(if ($null -ne $buildSummary -and $buildSummary.PSObject.Properties.Name -contains "imagePlanPath") { [string]$buildSummary.imagePlanPath } else { $null })
  imagePlanLowConfidenceCount = $(if ($null -ne $buildSummary -and $buildSummary.PSObject.Properties.Name -contains "imagePlanLowConfidenceCount") { $buildSummary.imagePlanLowConfidenceCount } else { $null })
  imagePlanNeedsReview = $(if ($null -ne $buildSummary -and $buildSummary.PSObject.Properties.Name -contains "imagePlanNeedsReview") { $buildSummary.imagePlanNeedsReview } else { $null })
  layoutCheckPath = $(if ($null -ne $buildSummary -and $buildSummary.PSObject.Properties.Name -contains "layoutCheckPath") { [string]$buildSummary.layoutCheckPath } else { $null })
  layoutCheckPassed = $(if ($null -ne $buildSummary -and $buildSummary.PSObject.Properties.Name -contains "layoutCheckPassed") { $buildSummary.layoutCheckPassed } else { $null })
  layoutCheckMessage = $(if ($null -ne $buildSummary -and $buildSummary.PSObject.Properties.Name -contains "layoutCheckMessage") { [string]$buildSummary.layoutCheckMessage } else { $null })
  layoutCheckErrorCount = $(if ($null -ne $buildSummary -and $buildSummary.PSObject.Properties.Name -contains "layoutCheckErrorCount") { $buildSummary.layoutCheckErrorCount } else { $null })
  layoutCheckWarningCount = $(if ($null -ne $buildSummary -and $buildSummary.PSObject.Properties.Name -contains "layoutCheckWarningCount") { $buildSummary.layoutCheckWarningCount } else { $null })
  finalDocxPath = $finalDocxPath
}
[System.IO.File]::WriteAllText($wrapperSummaryPath, ($wrapperSummary | ConvertTo-Json -Depth 6), (New-Object System.Text.UTF8Encoding($true)))

Write-Output ("Prompt path: {0}" -f $promptPathOut)
Write-Output ("Raw report path: {0}" -f $rawReportPath)
Write-Output ("Cleaned report path: {0}" -f $cleanedReportPath)
if (-not [string]::IsNullOrWhiteSpace($effectiveMetadataPath)) {
  Write-Output ("Metadata path: {0}" -f $effectiveMetadataPath)
}
if (-not [string]::IsNullOrWhiteSpace($effectiveRequirementsPath)) {
  Write-Output ("Requirements path: {0}" -f $effectiveRequirementsPath)
}
if (-not [string]::IsNullOrWhiteSpace($finalDocxPath)) {
  Write-Output ("Final docx path: {0}" -f $finalDocxPath)
}
Write-Output ("Wrapper summary path: {0}" -f $wrapperSummaryPath)
