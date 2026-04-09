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
  [string]$DetailLevel = "full"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "report-defaults.ps1")
. (Join-Path $PSScriptRoot "report-profiles.ps1")

function New-AutoPromptText {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ResolvedCourseName,

    [Parameter(Mandatory = $true)]
    [string]$ResolvedExperimentName,

    [Parameter(Mandatory = $true)]
    [psobject]$Profile,

    [Parameter(Mandatory = $true)]
    [hashtable]$Labels,

    [ValidateSet("standard", "full")]
    [string]$DetailLevel = "full"
  )

  $requiredHeadings = (Get-ReportProfileRequiredHeadings -Profile $Profile) -join ", "
  $detailProfile = Get-ReportProfileDetailProfile -Profile $Profile -DetailLevel $DetailLevel
  $detailRequirements = @(
    @($detailProfile.promptGuidance | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }) |
      ForEach-Object { "- {0}" -f [string]$_ }
  ) -join [Environment]::NewLine

  return @"
Write a formal Chinese university lab report body based on the reference webpages.

Course name: $ResolvedCourseName
Experiment name: $ResolvedExperimentName

Requirements:
- The report must begin by explicitly writing the course name and experiment name in Chinese.
- The report must include these Chinese headings: $requiredHeadings.
- Use the webpages as procedural reference for background, theory, steps, and verification ideas, but do not copy them verbatim.
- If the webpages are tutorial pages or lab guides, rewrite them into a submit-ready lab report in natural Chinese.
- If the webpages do not provide real screenshots, exact measured values, packet captures, teacher comments, or error logs, do not fabricate them. Write the result section as validation-oriented results instead.
$($detailRequirements.Trim())
- Return only the final Chinese report body.
"@
}

function New-AutoRequirementsJson {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ResolvedCourseName,

    [Parameter(Mandatory = $true)]
    [string]$ResolvedExperimentName,

    [Parameter(Mandatory = $true)]
    [psobject]$Profile,

    [Parameter(Mandatory = $true)]
    [hashtable]$Labels,

    [string[]]$ExtraKeywords,

    [ValidateSet("standard", "full")]
    [string]$DetailLevel = "full"
  )

  $keywordList = New-Object System.Collections.Generic.List[string]
  [void]$keywordList.Add($ResolvedCourseName)
  [void]$keywordList.Add($ResolvedExperimentName)

  foreach ($keyword in @($ExtraKeywords)) {
    if (-not [string]::IsNullOrWhiteSpace($keyword) -and -not $keywordList.Contains($keyword)) {
      [void]$keywordList.Add($keyword)
    }
  }

  $detailProfile = Get-ReportProfileDetailProfile -Profile $Profile -DetailLevel $DetailLevel
  $sectionRequirements = foreach ($sectionField in (Get-ReportProfileSectionFields -Profile $Profile)) {
    $minChars = 0
    if ($null -ne $sectionField.minChars -and $sectionField.minChars.PSObject.Properties.Name -contains $DetailLevel) {
      $minChars = [int]$sectionField.minChars.$DetailLevel
    }

    [pscustomobject]@{
      name = [string]$sectionField.heading
      aliases = @($sectionField.aliases | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
      minChars = $minChars
    }
  }

  $requirements = [pscustomobject]@{
    courseName = $ResolvedCourseName
    experimentName = $ResolvedExperimentName
    minChars = [int]$detailProfile.minChars
    sections = @($sectionRequirements)
    requiredKeywords = @($keywordList)
    forbiddenPatterns = @($Profile.forbiddenPatterns)
  }

  return ($requirements | ConvertTo-Json -Depth 6)
}

function New-AutoMetadataJson {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ResolvedCourseName,

    [Parameter(Mandatory = $true)]
    [string]$ResolvedExperimentName,

    [Parameter(Mandatory = $true)]
    [psobject]$Profile,

    [Parameter(Mandatory = $true)]
    [hashtable]$Labels,

    [AllowNull()]
    [string]$ResolvedStudentName,

    [AllowNull()]
    [string]$ResolvedStudentId,

    [AllowNull()]
    [string]$ResolvedClassName,

    [AllowNull()]
    [string]$ResolvedTeacherName,

    [AllowNull()]
    [string]$ResolvedExperimentProperty,

    [AllowNull()]
    [string]$ResolvedExperimentDate,

    [AllowNull()]
    [string]$ResolvedExperimentLocation
  )

  $metadataValues = @{
    Name = $ResolvedStudentName
    StudentId = $ResolvedStudentId
    ClassName = $ResolvedClassName
    TeacherName = $ResolvedTeacherName
    CourseName = $ResolvedCourseName
    ExperimentName = $ResolvedExperimentName
    ExperimentProperty = $ResolvedExperimentProperty
    ExperimentDate = $ResolvedExperimentDate
    ExperimentLocation = $ResolvedExperimentLocation
  }

  $metadata = [ordered]@{}
  foreach ($field in @($Profile.metadataFields)) {
    $key = [string]$field.key
    if (-not [string]::IsNullOrWhiteSpace($key) -and $metadataValues.ContainsKey($key)) {
      $metadata[[string]$field.label] = $metadataValues[$key]
    }
  }
  if ($Labels.Contains("Date")) {
    $metadata[[string]$Labels["Date"]] = $ResolvedExperimentDate
  }

  return ($metadata | ConvertTo-Json -Depth 4)
}

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

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$reportProfile = Get-ReportProfile -ProfileName $ReportProfileName -ProfilePath $ReportProfilePath -RepoRoot $repoRoot
$labels = Get-ReportProfileLabels -Profile $reportProfile
if ([string]::IsNullOrWhiteSpace($ExperimentProperty)) {
  $ExperimentProperty = [string]$reportProfile.defaultExperimentProperty
}

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
  $OutputDir = Join-Path $repoRoot ("tests-output\url-build-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
}

$resolvedOutputDir = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null

$resolvedPromptPath = Resolve-AbsolutePathIfProvided -Path $PromptPath
$referenceUrlList = @($ReferenceUrls | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
$referenceTextPathList = @($ReferenceTextPaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
$inferredExperimentName = Resolve-InferredExperimentName `
  -PromptText $PromptText `
  -PromptPath $resolvedPromptPath `
  -ReferenceTextPaths $referenceTextPathList `
  -ReferenceUrls $referenceUrlList

$fetchedReferenceTextPathList = @()
$effectiveReferenceUrlList = $referenceUrlList
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

  $effectiveReferenceUrlList = @()
  $effectiveReferenceTextPathList = @($referenceTextPathList + $fetchedReferenceTextPathList)
}

$resolvedNames = Resolve-ExperimentReportNames -CourseName $CourseName -ExperimentName $ExperimentName -InferredExperimentName $inferredExperimentName
$resolvedCourseName = [string]$resolvedNames.courseName
$resolvedExperimentName = [string]$resolvedNames.experimentName
if ([string]::IsNullOrWhiteSpace($resolvedCourseName) -or [string]::IsNullOrWhiteSpace($resolvedExperimentName)) {
  throw "CourseName and ExperimentName are required unless ExperimentName can be inferred from PromptText / PromptPath / ReferenceTextPaths / ReferenceUrls. After you set them once, later runs can omit them."
}

$resolvedTemplatePath = Resolve-AbsolutePathIfProvided -Path $TemplatePath
$resolvedMetadataPath = Resolve-AbsolutePathIfProvided -Path $MetadataPath
$resolvedRequirementsPath = Resolve-AbsolutePathIfProvided -Path $RequirementsPath
$resolvedStyleProfilePath = Resolve-AbsolutePathIfProvided -Path $StyleProfilePath
$resolvedImagePlanOutPath = if ([string]::IsNullOrWhiteSpace($ImagePlanOutPath)) { $null } else { [System.IO.Path]::GetFullPath($ImagePlanOutPath) }
$resolvedFinalDocxPath = if ([string]::IsNullOrWhiteSpace($FinalDocxPath)) { $null } else { [System.IO.Path]::GetFullPath($FinalDocxPath) }

$basePromptText = if (-not [string]::IsNullOrWhiteSpace($resolvedPromptPath)) {
  Get-Content -LiteralPath $resolvedPromptPath -Raw -Encoding UTF8
} elseif (-not [string]::IsNullOrWhiteSpace($PromptText)) {
  $PromptText
} else {
  New-AutoPromptText -ResolvedCourseName $resolvedCourseName -ResolvedExperimentName $resolvedExperimentName -Profile $reportProfile -Labels $labels -DetailLevel $DetailLevel
}

$promptPathOut = Join-Path $resolvedOutputDir "prompt.txt"
[System.IO.File]::WriteAllText($promptPathOut, $basePromptText, (New-Object System.Text.UTF8Encoding($true)))

$effectiveMetadataPath = $resolvedMetadataPath
if ([string]::IsNullOrWhiteSpace($effectiveMetadataPath) -and [string]::IsNullOrWhiteSpace($MetadataJson)) {
  $effectiveMetadataPath = Join-Path $resolvedOutputDir "metadata.auto.json"
  $autoMetadataJson = New-AutoMetadataJson `
    -ResolvedCourseName $resolvedCourseName `
    -ResolvedExperimentName $resolvedExperimentName `
    -Profile $reportProfile `
    -Labels $labels `
    -ResolvedStudentName $StudentName `
    -ResolvedStudentId $StudentId `
    -ResolvedClassName $ClassName `
    -ResolvedTeacherName $TeacherName `
    -ResolvedExperimentProperty $ExperimentProperty `
    -ResolvedExperimentDate $ExperimentDate `
    -ResolvedExperimentLocation $ExperimentLocation
  [System.IO.File]::WriteAllText($effectiveMetadataPath, $autoMetadataJson, (New-Object System.Text.UTF8Encoding($true)))
}

$effectiveRequirementsPath = $resolvedRequirementsPath
if ([string]::IsNullOrWhiteSpace($effectiveRequirementsPath) -and [string]::IsNullOrWhiteSpace($RequirementsJson)) {
  $effectiveRequirementsPath = Join-Path $resolvedOutputDir "requirements.auto.json"
  $autoRequirementsJson = New-AutoRequirementsJson `
    -ResolvedCourseName $resolvedCourseName `
    -ResolvedExperimentName $resolvedExperimentName `
    -Profile $reportProfile `
    -Labels $labels `
    -ExtraKeywords $RequiredKeywords `
    -DetailLevel $DetailLevel
  [System.IO.File]::WriteAllText($effectiveRequirementsPath, $autoRequirementsJson, (New-Object System.Text.UTF8Encoding($true)))
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
  if (-not [string]::IsNullOrWhiteSpace($ReportProfileName)) {
    $buildParams.ReportProfileName = $ReportProfileName
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

$savedDefaultsPath = Save-ExperimentReportDefaults -CourseName $resolvedCourseName -ExperimentName $resolvedExperimentName -DefaultsPath ([string]$resolvedNames.defaultsPath)

$wrapperSummaryPath = Join-Path $resolvedOutputDir "url-build-summary.json"
$wrapperSummary = [pscustomobject]@{
  outputDir = $resolvedOutputDir
  reportProfileName = [string]$reportProfile.name
  reportProfilePath = [string]$reportProfile.resolvedProfilePath
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
  referenceUrls = $referenceUrlList
  referenceTextPaths = $effectiveReferenceTextPathList
  fetchedReferenceTextPaths = $fetchedReferenceTextPathList
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
