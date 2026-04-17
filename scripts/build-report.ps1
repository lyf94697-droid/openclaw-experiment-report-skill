[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$TemplatePath,

  [Parameter(Mandatory = $true)]
  [string]$ReportPath,

  [string]$MetadataPath,

  [string]$MetadataJson,

  [string]$ImageSpecsPath,

  [string]$ImageSpecsJson,

  [string[]]$ImagePaths,

  [string]$ReportProfileName = "experiment-report",

  [string]$ReportProfilePath,

  [string]$RequirementsPath,

  [string]$RequirementsJson,

  [string]$OutputDir,

  [string]$FieldMapOutPath,

  [string]$FilledDocxOutPath,

  [string]$ImagePlanOutPath,

  [string]$ImageMapOutPath,

  [string]$FilledDocxWithImagesOutPath,

  [string]$StyledDocxOutPath,

  [string]$TemplateFrameDocxOutPath,

  [switch]$StyleFinalDocx,

  [switch]$CreateTemplateFrameDocx,

  [ValidateSet("auto", "default", "compact", "school")]
  [string]$StyleProfile = "auto",

  [string]$StyleProfilePath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "report-profiles.ps1")

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

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$resolvedTemplatePath = (Resolve-Path -LiteralPath $TemplatePath).Path
$resolvedReportPath = (Resolve-Path -LiteralPath $ReportPath).Path

$resolvedMetadataPath = $null
if (-not [string]::IsNullOrWhiteSpace($MetadataPath)) {
  $resolvedMetadataPath = (Resolve-Path -LiteralPath $MetadataPath).Path
}

$resolvedRequirementsPath = $null
if (-not [string]::IsNullOrWhiteSpace($RequirementsPath)) {
  $resolvedRequirementsPath = (Resolve-Path -LiteralPath $RequirementsPath).Path
}

$reportProfile = Get-ReportProfile -ProfileName $ReportProfileName -ProfilePath $ReportProfilePath -RepoRoot $repoRoot
$resolvedReportProfilePath = [string]$reportProfile.resolvedProfilePath
$effectiveStyleProfile = if ($PSBoundParameters.ContainsKey("StyleProfile")) {
  $StyleProfile
} else {
  Get-ReportProfileDefaultStyleProfile -Profile $reportProfile
}
$metadataInputMode = if (-not [string]::IsNullOrWhiteSpace($resolvedMetadataPath)) {
  "path"
} elseif (-not [string]::IsNullOrWhiteSpace($MetadataJson)) {
  "inline"
} else {
  "none"
}
$requirementsInputMode = if (-not [string]::IsNullOrWhiteSpace($resolvedRequirementsPath)) {
  "path"
} elseif (-not [string]::IsNullOrWhiteSpace($RequirementsJson)) {
  "inline"
} else {
  "none"
}

$imageInputModes = 0
if (-not [string]::IsNullOrWhiteSpace($ImageSpecsPath)) { $imageInputModes++ }
if (-not [string]::IsNullOrWhiteSpace($ImageSpecsJson)) { $imageInputModes++ }
if ($null -ne $ImagePaths -and @($ImagePaths).Count -gt 0) { $imageInputModes++ }
$imageInputsProvided = ($imageInputModes -gt 0)
if ($imageInputModes -gt 1) {
  throw "Provide zero or one of -ImageSpecsPath, -ImageSpecsJson, or -ImagePaths."
}
$imageInputMode = if (-not [string]::IsNullOrWhiteSpace($ImageSpecsPath)) {
  "specs-path"
} elseif (-not [string]::IsNullOrWhiteSpace($ImageSpecsJson)) {
  "specs-json"
} elseif ($null -ne $ImagePaths -and @($ImagePaths).Count -gt 0) {
  "image-paths"
} else {
  "none"
}

$styleOutputRequested = $StyleFinalDocx -or (-not [string]::IsNullOrWhiteSpace($StyledDocxOutPath))

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
  $OutputDir = Join-Path $repoRoot ("tests-output\build-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
}

$resolvedOutputDir = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null

$resolvedFieldMapOutPath = if ([string]::IsNullOrWhiteSpace($FieldMapOutPath)) {
  Join-Path $resolvedOutputDir "generated-field-map.json"
} else {
  [System.IO.Path]::GetFullPath($FieldMapOutPath)
}
Ensure-ParentDirectory -Path $resolvedFieldMapOutPath

$resolvedFilledDocxOutPath = if ([string]::IsNullOrWhiteSpace($FilledDocxOutPath)) {
  Join-Path $resolvedOutputDir (([System.IO.Path]::GetFileNameWithoutExtension($resolvedTemplatePath)) + ".filled.docx")
} else {
  [System.IO.Path]::GetFullPath($FilledDocxOutPath)
}
Ensure-ParentDirectory -Path $resolvedFilledDocxOutPath

$resolvedImagePlanOutPath = $null
$resolvedImageMapOutPath = $null
$resolvedFilledDocxWithImagesOutPath = $null
$resolvedStyledDocxOutPath = $null
$validationPath = $null
$filledOutlinePath = $null
$filledWithImagesOutlinePath = $null
$styledOutlinePath = $null
$resolvedTemplateFrameDocxOutPath = $null
$styleResult = $null
$summaryPath = Join-Path $resolvedOutputDir "summary.json"
$layoutCheckPath = Join-Path $resolvedOutputDir "layout-check.json"
$layoutCheckResult = $null
$expectedLayoutImageCount = -1
$expectedLayoutCaptionCount = -1
$imagePlanLowConfidenceCount = $null
$imagePlanNeedsReview = $null

$validationResult = $null
if (-not [string]::IsNullOrWhiteSpace($resolvedRequirementsPath) -or -not [string]::IsNullOrWhiteSpace($RequirementsJson)) {
  $validationPath = Join-Path $resolvedOutputDir "validation.json"
  $validationParams = @{
    Path = $resolvedReportPath
    Format = "json"
  }
  if (-not [string]::IsNullOrWhiteSpace($ReportProfileName)) {
    $validationParams.ReportProfileName = $ReportProfileName
  }
  if (-not [string]::IsNullOrWhiteSpace($resolvedReportProfilePath)) {
    $validationParams.ReportProfilePath = $resolvedReportProfilePath
  }
  if (-not [string]::IsNullOrWhiteSpace($resolvedRequirementsPath)) {
    $validationParams.RequirementsPath = $resolvedRequirementsPath
  } else {
    $validationParams.RequirementsJson = $RequirementsJson
  }

  $validationJson = & (Join-Path $repoRoot "scripts\validate-report-draft.ps1") @validationParams | Out-String
  [System.IO.File]::WriteAllText($validationPath, $validationJson, (New-Object System.Text.UTF8Encoding($true)))
  $validationResult = $validationJson | ConvertFrom-Json
}

$fieldMapParams = @{
  TemplatePath = $resolvedTemplatePath
  ReportPath = $resolvedReportPath
  ReportProfileName = $ReportProfileName
  ReportProfilePath = $resolvedReportProfilePath
  Format = "json"
  OutFile = $resolvedFieldMapOutPath
}
if (-not [string]::IsNullOrWhiteSpace($resolvedMetadataPath)) {
  $fieldMapParams.MetadataPath = $resolvedMetadataPath
} elseif (-not [string]::IsNullOrWhiteSpace($MetadataJson)) {
  $fieldMapParams.MetadataJson = $MetadataJson
}

& (Join-Path $repoRoot "scripts\generate-docx-field-map.ps1") @fieldMapParams | Out-Null
& (Join-Path $repoRoot "scripts\apply-docx-field-map.ps1") -TemplatePath $resolvedTemplatePath -MappingPath $resolvedFieldMapOutPath -OutPath $resolvedFilledDocxOutPath -Overwrite | Out-Null

$filledOutlinePath = Join-Path $resolvedOutputDir "filled-template-outline.md"
$filledOutline = & (Join-Path $repoRoot "scripts\extract-docx-template.ps1") -Path $resolvedFilledDocxOutPath -Format markdown | Out-String
[System.IO.File]::WriteAllText($filledOutlinePath, $filledOutline, (New-Object System.Text.UTF8Encoding($true)))

if ($imageInputsProvided) {
  $resolvedImagePlanOutPath = if ([string]::IsNullOrWhiteSpace($ImagePlanOutPath)) {
    Join-Path $resolvedOutputDir "image-placement-plan.md"
  } else {
    [System.IO.Path]::GetFullPath($ImagePlanOutPath)
  }
  Ensure-ParentDirectory -Path $resolvedImagePlanOutPath

  $resolvedImageMapOutPath = if ([string]::IsNullOrWhiteSpace($ImageMapOutPath)) {
    Join-Path $resolvedOutputDir "generated-image-map.json"
  } else {
    [System.IO.Path]::GetFullPath($ImageMapOutPath)
  }
  Ensure-ParentDirectory -Path $resolvedImageMapOutPath

  $resolvedFilledDocxWithImagesOutPath = if ([string]::IsNullOrWhiteSpace($FilledDocxWithImagesOutPath)) {
    Join-Path $resolvedOutputDir (([System.IO.Path]::GetFileNameWithoutExtension($resolvedFilledDocxOutPath)) + ".images.docx")
  } else {
    [System.IO.Path]::GetFullPath($FilledDocxWithImagesOutPath)
  }
  Ensure-ParentDirectory -Path $resolvedFilledDocxWithImagesOutPath

  $imageInputParams = @{
    DocxPath = $resolvedFilledDocxOutPath
    ReportProfileName = $ReportProfileName
    ReportProfilePath = $resolvedReportProfilePath
  }
  if (-not [string]::IsNullOrWhiteSpace($ImageSpecsPath)) {
    $imageInputParams.ImageSpecsPath = (Resolve-Path -LiteralPath $ImageSpecsPath).Path
  } elseif (-not [string]::IsNullOrWhiteSpace($ImageSpecsJson)) {
    $imageInputParams.ImageSpecsJson = $ImageSpecsJson
  } else {
    $imageInputParams.ImagePaths = $ImagePaths
  }

  $imagePlanJsonParams = $imageInputParams.Clone()
  $imagePlanJsonParams.Format = "json"
  $imagePlanJsonParams.PlanOnly = $true
  $imagePlanResult = ((& (Join-Path $repoRoot "scripts\generate-docx-image-map.ps1") @imagePlanJsonParams) | Out-String) | ConvertFrom-Json
  $imagePlanEntries = if ($null -ne $imagePlanResult -and $imagePlanResult.PSObject.Properties.Name -contains "plan") {
    @($imagePlanResult.plan)
  } else {
    @()
  }
  $imagePlanLowConfidenceCount = @($imagePlanEntries | Where-Object { [string]$_.confidence -eq "low" }).Count
  $imagePlanNeedsReview = ($imagePlanLowConfidenceCount -gt 0)

  $imagePlanMarkdownParams = $imageInputParams.Clone()
  $imagePlanMarkdownParams.Format = "markdown"
  $imagePlanMarkdownParams.PlanOnly = $true
  $imagePlanMarkdownParams.OutFile = $resolvedImagePlanOutPath
  & (Join-Path $repoRoot "scripts\generate-docx-image-map.ps1") @imagePlanMarkdownParams | Out-Null

  $imageMapParams = $imageInputParams.Clone()
  $imageMapParams.Format = "json"
  $imageMapParams.OutFile = $resolvedImageMapOutPath
  & (Join-Path $repoRoot "scripts\generate-docx-image-map.ps1") @imageMapParams | Out-Null
  $generatedImageMap = (Get-Content -LiteralPath $resolvedImageMapOutPath -Raw -Encoding UTF8) | ConvertFrom-Json
  if ($null -ne $generatedImageMap -and $generatedImageMap.PSObject.Properties.Name -contains "images") {
    $expectedLayoutImageCount = @($generatedImageMap.images).Count
    $expectedLayoutCaptionCount = @($generatedImageMap.images | Where-Object {
        $_.PSObject.Properties.Name -contains "caption" -and -not [string]::IsNullOrWhiteSpace([string]$_.caption)
      }).Count
  }

  & (Join-Path $repoRoot "scripts\insert-docx-images.ps1") `
    -DocxPath $resolvedFilledDocxOutPath `
    -MappingPath $resolvedImageMapOutPath `
    -ReportProfileName $ReportProfileName `
    -ReportProfilePath $resolvedReportProfilePath `
    -OutPath $resolvedFilledDocxWithImagesOutPath `
    -Overwrite | Out-Null

  $filledWithImagesOutlinePath = Join-Path $resolvedOutputDir "filled-template-with-images-outline.md"
  $filledWithImagesOutline = & (Join-Path $repoRoot "scripts\extract-docx-template.ps1") -Path $resolvedFilledDocxWithImagesOutPath -Format markdown | Out-String
  [System.IO.File]::WriteAllText($filledWithImagesOutlinePath, $filledWithImagesOutline, (New-Object System.Text.UTF8Encoding($true)))
}

if ($styleOutputRequested) {
  $styleInputPath = if ($null -ne $resolvedFilledDocxWithImagesOutPath) { $resolvedFilledDocxWithImagesOutPath } else { $resolvedFilledDocxOutPath }
  $resolvedStyledDocxOutPath = if ([string]::IsNullOrWhiteSpace($StyledDocxOutPath)) {
    Join-Path $resolvedOutputDir (([System.IO.Path]::GetFileNameWithoutExtension($styleInputPath)) + ".styled.docx")
  } else {
    [System.IO.Path]::GetFullPath($StyledDocxOutPath)
  }

  $styleParams = @{
    DocxPath = $styleInputPath
    OutPath = $resolvedStyledDocxOutPath
    Overwrite = $true
    Profile = $effectiveStyleProfile
    ReportProfileName = [string]$reportProfile.name
    ReportProfilePath = [string]$reportProfile.resolvedProfilePath
  }
  if (-not [string]::IsNullOrWhiteSpace($StyleProfilePath)) {
    $styleParams.ProfilePath = (Resolve-Path -LiteralPath $StyleProfilePath).Path
  }

  $styleResult = & (Join-Path $repoRoot "scripts\format-docx-report-style.ps1") @styleParams

  $styledOutlinePath = Join-Path $resolvedOutputDir "styled-template-outline.md"
  $styledOutline = & (Join-Path $repoRoot "scripts\extract-docx-template.ps1") -Path $resolvedStyledDocxOutPath -Format markdown | Out-String
  [System.IO.File]::WriteAllText($styledOutlinePath, $styledOutline, (New-Object System.Text.UTF8Encoding($true)))
}

$finalDocxPath = if ($null -ne $resolvedStyledDocxOutPath) {
  $resolvedStyledDocxOutPath
} elseif ($null -ne $resolvedFilledDocxWithImagesOutPath) {
  $resolvedFilledDocxWithImagesOutPath
} else {
  $resolvedFilledDocxOutPath
}

if ($CreateTemplateFrameDocx -or (-not [string]::IsNullOrWhiteSpace($TemplateFrameDocxOutPath))) {
  $resolvedTemplateFrameDocxOutPath = if ([string]::IsNullOrWhiteSpace($TemplateFrameDocxOutPath)) {
    Join-Path $resolvedOutputDir (([System.IO.Path]::GetFileNameWithoutExtension($finalDocxPath)) + ".template-frame.docx")
  } else {
    [System.IO.Path]::GetFullPath($TemplateFrameDocxOutPath)
  }
  Ensure-ParentDirectory -Path $resolvedTemplateFrameDocxOutPath
  & (Join-Path $repoRoot "scripts\convert-docx-template-frame.ps1") `
    -DocxPath $finalDocxPath `
    -OutPath $resolvedTemplateFrameDocxOutPath `
    -Overwrite | Out-Null
}

$layoutCheckParams = @{
  DocxPath = $finalDocxPath
  Format = "json"
  OutFile = $layoutCheckPath
}
if (-not [string]::IsNullOrWhiteSpace($ReportProfileName)) {
  $layoutCheckParams.ReportProfileName = $ReportProfileName
}
if (-not [string]::IsNullOrWhiteSpace($resolvedReportProfilePath)) {
  $layoutCheckParams.ReportProfilePath = $resolvedReportProfilePath
}
if ($expectedLayoutImageCount -ge 0) {
  $layoutCheckParams.ExpectedImageCount = $expectedLayoutImageCount
}
if ($expectedLayoutCaptionCount -ge 0) {
  $layoutCheckParams.ExpectedCaptionCount = $expectedLayoutCaptionCount
}
& (Join-Path $repoRoot "scripts\check-docx-layout.ps1") @layoutCheckParams | Out-Null
$layoutCheckResult = (Get-Content -LiteralPath $layoutCheckPath -Raw -Encoding UTF8) | ConvertFrom-Json

$validationWarningSummary = @()
$validationErrorCodes = @()
$validationWarningCodes = @()
if ($null -ne $validationResult -and $validationResult.PSObject.Properties.Name -contains "findings") {
  $validationWarningSummary = @(
    $validationResult.findings |
      Where-Object { [string]$_.severity -eq "warning" } |
      ForEach-Object {
        [pscustomobject]@{
          severity = [string]$_.severity
          code = [string]$_.code
          category = $(if ($_.PSObject.Properties.Name -contains "category") { [string]$_.category } else { $null })
          message = [string]$_.message
          remediation = $(if ($_.PSObject.Properties.Name -contains "remediation") { [string]$_.remediation } else { $null })
        }
      }
  )
}
if ($null -ne $validationResult -and $validationResult.summary.PSObject.Properties.Name -contains "errorCodes") {
  $validationErrorCodes = @($validationResult.summary.errorCodes)
}
if ($null -ne $validationResult -and $validationResult.summary.PSObject.Properties.Name -contains "warningCodes") {
  $validationWarningCodes = @($validationResult.summary.warningCodes)
}

$summary = [pscustomobject]@{
  outputDir = $resolvedOutputDir
  reportProfileName = [string]$reportProfile.name
  reportProfilePath = $resolvedReportProfilePath
  templatePath = $resolvedTemplatePath
  reportPath = $resolvedReportPath
  reportInputMode = "path"
  metadataPath = $resolvedMetadataPath
  metadataInputMode = $metadataInputMode
  requirementsInputMode = $requirementsInputMode
  imageInputMode = $imageInputMode
  fieldMapPath = $resolvedFieldMapOutPath
  filledDocxPath = $resolvedFilledDocxOutPath
  filledOutlinePath = $filledOutlinePath
  imagePlanPath = $resolvedImagePlanOutPath
  imagePlanLowConfidenceCount = $imagePlanLowConfidenceCount
  imagePlanNeedsReview = $imagePlanNeedsReview
  imageMapPath = $resolvedImageMapOutPath
  filledDocxWithImagesPath = $resolvedFilledDocxWithImagesOutPath
  filledWithImagesOutlinePath = $filledWithImagesOutlinePath
  styledDocxPath = $resolvedStyledDocxOutPath
  styledOutlinePath = $styledOutlinePath
  templateFrameDocxPath = $resolvedTemplateFrameDocxOutPath
  layoutCheckPath = $layoutCheckPath
  layoutCheckPassed = $(if ($null -ne $layoutCheckResult) { [bool]$layoutCheckResult.passed } else { $null })
  layoutCheckMessage = $(if ($null -ne $layoutCheckResult -and $layoutCheckResult.PSObject.Properties.Name -contains "message") { [string]$layoutCheckResult.message } else { $null })
  layoutCheckErrorCount = $(if ($null -ne $layoutCheckResult) { [int]$layoutCheckResult.summary.errorCount } else { $null })
  layoutCheckWarningCount = $(if ($null -ne $layoutCheckResult) { [int]$layoutCheckResult.summary.warningCount } else { $null })
  expectedLayoutImageCount = $(if ($expectedLayoutImageCount -ge 0) { $expectedLayoutImageCount } else { $null })
  expectedLayoutCaptionCount = $(if ($expectedLayoutCaptionCount -ge 0) { $expectedLayoutCaptionCount } else { $null })
  requestedStyleProfile = $(if ($styleOutputRequested) { $effectiveStyleProfile } else { $null })
  styleProfilePath = $(if ($null -ne $styleResult) { [string]$styleResult.profilePath } elseif (-not [string]::IsNullOrWhiteSpace($StyleProfilePath)) { (Resolve-Path -LiteralPath $StyleProfilePath).Path } else { $null })
  styleProfile = $(if ($null -ne $styleResult) { [string]$styleResult.styleProfile } else { $null })
  resolvedStyleProfile = $(if ($null -ne $styleResult) { [string]$styleResult.resolvedProfile } else { $null })
  styleProfileReason = $(if ($null -ne $styleResult) { [string]$styleResult.profileReason } else { $null })
  appliedStyleSettings = $(if ($null -ne $styleResult) { $styleResult.appliedSettings } else { $null })
  finalDocxPath = $finalDocxPath
  validationPath = $validationPath
  validationPassed = $(if ($null -ne $validationResult) { [bool]$validationResult.passed } else { $null })
  validationErrorCount = $(if ($null -ne $validationResult) { [int]$validationResult.summary.errorCount } else { $null })
  validationWarningCount = $(if ($null -ne $validationResult) { [int]$validationResult.summary.warningCount } else { $null })
  validationPaginationRiskCount = $(if ($null -ne $validationResult -and $validationResult.summary.PSObject.Properties.Name -contains "paginationRiskCount") { [int]$validationResult.summary.paginationRiskCount } else { $null })
  validationPaginationRiskThresholds = $(if ($null -ne $validationResult -and $validationResult.summary.PSObject.Properties.Name -contains "paginationRiskThresholds") { $validationResult.summary.paginationRiskThresholds } else { $null })
  validationPaginationRiskRemediations = $(if ($null -ne $validationResult -and $validationResult.summary.PSObject.Properties.Name -contains "paginationRiskRemediations") { $validationResult.summary.paginationRiskRemediations } else { $null })
  validationStructuralIssueCount = $(if ($null -ne $validationResult -and $validationResult.summary.PSObject.Properties.Name -contains "structuralIssueCount") { [int]$validationResult.summary.structuralIssueCount } else { $null })
  validationFindingCountsByCode = $(if ($null -ne $validationResult -and $validationResult.summary.PSObject.Properties.Name -contains "findingCountsByCode") { $validationResult.summary.findingCountsByCode } else { $null })
  validationFindingCountsByCategory = $(if ($null -ne $validationResult -and $validationResult.summary.PSObject.Properties.Name -contains "findingCountsByCategory") { $validationResult.summary.findingCountsByCategory } else { $null })
  validationErrorCodes = $validationErrorCodes
  validationWarningCodes = $validationWarningCodes
  validationWarningSummary = $validationWarningSummary
}
[System.IO.File]::WriteAllText($summaryPath, ($summary | ConvertTo-Json -Depth 8), (New-Object System.Text.UTF8Encoding($true)))

Write-Output ("Field-map path: {0}" -f $resolvedFieldMapOutPath)
Write-Output ("Filled docx path: {0}" -f $resolvedFilledDocxOutPath)
if ($null -ne $resolvedFilledDocxWithImagesOutPath) {
  Write-Output ("Image-plan path: {0}" -f $resolvedImagePlanOutPath)
  Write-Output ("Image-map path: {0}" -f $resolvedImageMapOutPath)
  Write-Output ("Filled docx with images path: {0}" -f $resolvedFilledDocxWithImagesOutPath)
}
if ($null -ne $resolvedStyledDocxOutPath) {
  Write-Output ("Styled docx path: {0}" -f $resolvedStyledDocxOutPath)
}
if ($null -ne $resolvedTemplateFrameDocxOutPath) {
  Write-Output ("Template-frame docx path: {0}" -f $resolvedTemplateFrameDocxOutPath)
}
Write-Output ("Final docx path: {0}" -f $finalDocxPath)
Write-Output ("Layout check path: {0}" -f $layoutCheckPath)
Write-Output ("Summary path: {0}" -f $summaryPath)

if ($null -ne $validationResult -and -not $validationResult.passed) {
  throw "Report validation failed. See $validationPath"
}
