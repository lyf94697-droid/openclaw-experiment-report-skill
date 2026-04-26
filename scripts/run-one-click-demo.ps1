[CmdletBinding()]
param(
  [string]$OutputDir,

  [ValidateSet("fast", "full")]
  [string]$PipelineMode = "full",

  [ValidateSet("auto", "default", "compact", "school", "excellent")]
  [string]$StyleProfile = "auto"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
  $OutputDir = Join-Path $repoRoot ("tests-output\one-click-demo-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
}

$resolvedOutputDir = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null

$templatePath = (Resolve-Path (Join-Path $repoRoot "examples\report-templates\experiment-report-template.docx")).Path
$reportPath = (Resolve-Path (Join-Path $repoRoot "examples\demo-one-click\report.txt")).Path
$metadataPath = (Resolve-Path (Join-Path $repoRoot "examples\demo-one-click\metadata.json")).Path
$requirementsPath = (Resolve-Path (Join-Path $repoRoot "examples\demo-one-click\requirements.json")).Path
$imageSpecsTemplatePath = (Resolve-Path (Join-Path $repoRoot "examples\docx-image-specs-row.json")).Path
$imageSpecs = (Get-Content -LiteralPath $imageSpecsTemplatePath -Raw -Encoding UTF8) | ConvertFrom-Json
$demoAssetPaths = @(
  (Resolve-Path (Join-Path $repoRoot "demo\assets\step-network-config.png")).Path,
  (Resolve-Path (Join-Path $repoRoot "demo\assets\step-ipconfig.png")).Path,
  (Resolve-Path (Join-Path $repoRoot "demo\assets\result-ping.png")).Path,
  (Resolve-Path (Join-Path $repoRoot "demo\assets\result-arp.png")).Path
)

for ($index = 0; $index -lt @($imageSpecs.images).Count; $index++) {
  $imageSpecs.images[$index].path = $demoAssetPaths[$index]
}

$imageSpecsJson = $imageSpecs | ConvertTo-Json -Depth 8

& (Join-Path $repoRoot "scripts\build-report.ps1") `
  -TemplatePath $templatePath `
  -ReportPath $reportPath `
  -MetadataPath $metadataPath `
  -RequirementsPath $requirementsPath `
  -ImageSpecsJson $imageSpecsJson `
  -OutputDir $resolvedOutputDir `
  -StyleFinalDocx `
  -PipelineMode $PipelineMode `
  -StyleProfile $StyleProfile | Out-Null

$summaryPath = Join-Path $resolvedOutputDir "summary.json"
$summary = (Get-Content -LiteralPath $summaryPath -Raw -Encoding UTF8) | ConvertFrom-Json

if (-not (Test-Path -LiteralPath ([string]$summary.finalDocxPath) -PathType Leaf)) {
  throw "One-click demo did not produce a final docx."
}

if ($summary.PSObject.Properties.Name -contains "layoutCheckPassed" -and -not [bool]$summary.layoutCheckPassed) {
  throw "One-click demo layout check failed. See $($summary.layoutCheckPath)"
}

Write-Output ("Demo output dir: {0}" -f $resolvedOutputDir)
Write-Output ("Final docx path: {0}" -f $summary.finalDocxPath)
if ($summary.PSObject.Properties.Name -contains "imagePlanPath" -and -not [string]::IsNullOrWhiteSpace([string]$summary.imagePlanPath)) {
  Write-Output ("Image-plan path: {0}" -f $summary.imagePlanPath)
}
if ($summary.PSObject.Properties.Name -contains "layoutCheckPath" -and -not [string]::IsNullOrWhiteSpace([string]$summary.layoutCheckPath)) {
  Write-Output ("Layout check path: {0}" -f $summary.layoutCheckPath)
}
Write-Output ("Summary path: {0}" -f $summaryPath)
