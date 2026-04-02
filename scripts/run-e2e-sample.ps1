[CmdletBinding()]
param(
  [string]$OpenClawCmd = $env:OPENCLAW_CMD,

  [string]$Agent = "gpt",

  [string]$PromptPath,

  [string]$RequirementsPath,

  [string]$OutputDir,

  [string[]]$ReferenceTextPaths,

  [string[]]$ReferenceUrls,

  [string]$BrowserProfile = $env:OPENCLAW_BROWSER_PROFILE,

  [int]$ReferenceMaxChars = 30000,

  [string]$TemplatePath,

  [string]$MetadataPath,

  [string]$FieldMapOutPath,

  [string]$FilledDocxOutPath,

  [string]$ImageSpecsPath,

  [string]$ImageSpecsJson,

  [string[]]$ImagePaths,

  [string]$ImageMapOutPath,

  [string]$FilledDocxWithImagesOutPath,

  [string]$StyledDocxOutPath,

  [AllowEmptyString()]
  [string]$SkillCommand = "/experiment-report",

  [string]$SessionKey,

  [ValidateSet("guided-chat", "native-agent")]
  [string]$Mode = "guided-chat",

  [ValidateSet("minimal", "low", "medium", "high")]
  [string]$Thinking = "medium",

  [int]$TimeoutSeconds = 240,

  [switch]$SkipInstall,

  [switch]$StyleFinalDocx,

  [ValidateSet("auto", "default", "compact", "school")]
  [string]$StyleProfile = "default",

  [string]$StyleProfilePath,

  [switch]$SkipSessionReset
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-OpenClawCommand {
  param(
    [AllowNull()]
    [string]$Candidate
  )

  if (-not [string]::IsNullOrWhiteSpace($Candidate)) {
    return (Resolve-Path -LiteralPath $Candidate).Path
  }

  foreach ($name in @("openclaw.cmd", "openclaw")) {
    $command = Get-Command $name -ErrorAction SilentlyContinue
    if ($null -ne $command -and $command.Source) {
      return $command.Source
    }
  }

  throw "Could not resolve openclaw CLI. Set -OpenClawCmd or OPENCLAW_CMD."
}

function Try-ParseJson {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return $null
  }

  $trimmed = $Text.Trim()
  if (-not ($trimmed.StartsWith('{') -or $trimmed.StartsWith('['))) {
    return $null
  }

  try {
    return $trimmed | ConvertFrom-Json
  } catch {
    return $null
  }
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
if ([string]::IsNullOrWhiteSpace($PromptPath)) {
  $PromptPath = Join-Path $repoRoot "examples\e2e-sample-prompt.txt"
}
if ([string]::IsNullOrWhiteSpace($RequirementsPath)) {
  $RequirementsPath = Join-Path $repoRoot "examples\e2e-sample-requirements.json"
}
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
  $OutputDir = Join-Path $repoRoot ("tests-output\e2e-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
}
if ([string]::IsNullOrWhiteSpace($SessionKey)) {
  $SessionKey = "agent:{0}:main" -f $Agent
}

$resolvedOpenClaw = $null
if ($Mode -eq "native-agent") {
  $resolvedOpenClaw = Resolve-OpenClawCommand -Candidate $OpenClawCmd
}
$resolvedPromptPath = (Resolve-Path -LiteralPath $PromptPath).Path
$resolvedRequirementsPath = (Resolve-Path -LiteralPath $RequirementsPath).Path
$resolvedOutputDir = [System.IO.Path]::GetFullPath($OutputDir)
$resolvedTemplatePath = $null
$resolvedMetadataPath = $null
if (-not [string]::IsNullOrWhiteSpace($TemplatePath)) {
  $resolvedTemplatePath = (Resolve-Path -LiteralPath $TemplatePath).Path
}
if (-not [string]::IsNullOrWhiteSpace($MetadataPath)) {
  $resolvedMetadataPath = (Resolve-Path -LiteralPath $MetadataPath).Path
}
$styleOutputRequested = $StyleFinalDocx -or (-not [string]::IsNullOrWhiteSpace($StyledDocxOutPath))

$imageInputsProvided = (-not [string]::IsNullOrWhiteSpace($ImageSpecsPath)) -or (-not [string]::IsNullOrWhiteSpace($ImageSpecsJson)) -or ($null -ne $ImagePaths -and $ImagePaths.Count -gt 0)
if ($imageInputsProvided -and $null -eq $resolvedTemplatePath) {
  throw "TemplatePath is required when images should be embedded into the final docx."
}
if ($styleOutputRequested -and $null -eq $resolvedTemplatePath) {
  throw "TemplatePath is required when the final docx should be style-formatted."
}

New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null

if (-not $SkipInstall) {
  & (Join-Path $repoRoot "scripts\install-skill.ps1") -Force | Out-Null
}
if (-not $SkipSessionReset) {
  & (Join-Path $repoRoot "scripts\reset-openclaw-session.ps1") -SessionKey $SessionKey | Out-Null
}

$promptText = Get-Content -LiteralPath $resolvedPromptPath -Raw -Encoding UTF8
if (-not [string]::IsNullOrWhiteSpace($SkillCommand)) {
  $trimmedPrompt = $promptText.TrimStart()
  if (-not $trimmedPrompt.StartsWith($SkillCommand, [System.StringComparison]::OrdinalIgnoreCase)) {
    $promptText = $SkillCommand + [Environment]::NewLine + $promptText
  }
}
$agentOutputPath = Join-Path $resolvedOutputDir "agent-output.txt"
$agentJsonPath = Join-Path $resolvedOutputDir "agent-result.json"
$promptOutPath = Join-Path $resolvedOutputDir "prompt.txt"
$reportPath = Join-Path $resolvedOutputDir "report.txt"
$validationPath = Join-Path $resolvedOutputDir "validation.json"
$summaryPath = Join-Path $resolvedOutputDir "summary.json"
$resolvedFieldMapOutPath = $null
$resolvedFilledDocxOutPath = $null
$filledOutlinePath = $null
$resolvedImageMapOutPath = $null
$resolvedFilledDocxWithImagesOutPath = $null
$filledWithImagesOutlinePath = $null
$resolvedStyledDocxOutPath = $null
$styleResult = $null
$preparedPromptResult = & (Join-Path $repoRoot "scripts\prepare-report-prompt.ps1") `
  -PromptText $promptText `
  -OutFile $promptOutPath `
  -ReferenceTextPaths $ReferenceTextPaths `
  -ReferenceUrls $ReferenceUrls `
  -BrowserProfile $BrowserProfile `
  -OpenClawCmd $OpenClawCmd `
  -ReferenceMaxChars $ReferenceMaxChars
$promptText = Get-Content -LiteralPath $promptOutPath -Raw -Encoding UTF8

$responseFormat = 'plain-text'
$reportText = $null
$skillNames = @()
$provider = $null
$model = $null
$skillActive = $false

if ($Mode -eq "guided-chat") {
  $guidedOutput = & (Join-Path $repoRoot "scripts\generate-report-chat.ps1") -PromptPath $promptOutPath -SessionKey $SessionKey -OutFile $reportPath $(if ($SkipSessionReset) { '-SkipSessionReset' })
  $reportText = (Get-Content -LiteralPath $reportPath -Raw -Encoding UTF8).Trim()
  [System.IO.File]::WriteAllText($agentOutputPath, ($guidedOutput | Out-String), (New-Object System.Text.UTF8Encoding($true)))
  $responseFormat = 'gateway-chat'
  $skillActive = $true
} else {
  $agentRawOutput = & $resolvedOpenClaw --no-color agent --agent $Agent --message $promptText --json --timeout $TimeoutSeconds --thinking $Thinking | Out-String
  [System.IO.File]::WriteAllText($agentOutputPath, $agentRawOutput, (New-Object System.Text.UTF8Encoding($true)))

  $agentResult = Try-ParseJson -Text $agentRawOutput
  if ($null -ne $agentResult) {
    $responseFormat = 'json'
    [System.IO.File]::WriteAllText($agentJsonPath, ($agentResult | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding($true)))

    if ($null -ne $agentResult.result.payloads) {
      $payloadTexts = @($agentResult.result.payloads | ForEach-Object { $_.text } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
      if ($payloadTexts.Count -gt 0) {
        $reportText = ($payloadTexts -join ([Environment]::NewLine + [Environment]::NewLine)).Trim()
      }
    }

    if ($null -ne $agentResult.result.meta.agentMeta) {
      $provider = $agentResult.result.meta.agentMeta.provider
      $model = $agentResult.result.meta.agentMeta.model
    }

    if ($null -ne $agentResult.result.meta.systemPromptReport.skills.entries) {
      $skillNames = @($agentResult.result.meta.systemPromptReport.skills.entries | ForEach-Object { $_.name })
      $skillActive = [bool]($skillNames -contains 'experiment-report')
    }
  }

  if ([string]::IsNullOrWhiteSpace($reportText)) {
    $reportText = $agentRawOutput.Trim()
  }
  if ([string]::IsNullOrWhiteSpace($reportText)) {
    throw "Agent returned no report text."
  }
  [System.IO.File]::WriteAllText($reportPath, $reportText, (New-Object System.Text.UTF8Encoding($true)))
}

$validationJson = & (Join-Path $repoRoot "scripts\validate-report-draft.ps1") -Path $reportPath -RequirementsPath $resolvedRequirementsPath -Format json | Out-String
[System.IO.File]::WriteAllText($validationPath, $validationJson, (New-Object System.Text.UTF8Encoding($true)))
$validationResult = $validationJson | ConvertFrom-Json

if ($null -ne $resolvedTemplatePath) {
  if ([string]::IsNullOrWhiteSpace($FieldMapOutPath)) {
    $resolvedFieldMapOutPath = Join-Path $resolvedOutputDir "generated-field-map.json"
  } else {
    $resolvedFieldMapOutPath = [System.IO.Path]::GetFullPath($FieldMapOutPath)
  }

  if ([string]::IsNullOrWhiteSpace($FilledDocxOutPath)) {
    $resolvedFilledDocxOutPath = Join-Path $resolvedOutputDir (([System.IO.Path]::GetFileNameWithoutExtension($resolvedTemplatePath)) + ".filled.docx")
  } else {
    $resolvedFilledDocxOutPath = [System.IO.Path]::GetFullPath($FilledDocxOutPath)
  }

  $fieldMapParams = @{
    TemplatePath = $resolvedTemplatePath
    ReportPath = $reportPath
    Format = "json"
    OutFile = $resolvedFieldMapOutPath
  }
  if ($null -ne $resolvedMetadataPath) {
    $fieldMapParams.MetadataPath = $resolvedMetadataPath
  }

  & (Join-Path $repoRoot "scripts\generate-docx-field-map.ps1") @fieldMapParams | Out-Null
  & (Join-Path $repoRoot "scripts\apply-docx-field-map.ps1") -TemplatePath $resolvedTemplatePath -MappingPath $resolvedFieldMapOutPath -OutPath $resolvedFilledDocxOutPath -Overwrite | Out-Null

  $filledOutlinePath = Join-Path $resolvedOutputDir "filled-template-outline.md"
  $filledOutline = & (Join-Path $repoRoot "scripts\extract-docx-template.ps1") -Path $resolvedFilledDocxOutPath -Format markdown | Out-String
  [System.IO.File]::WriteAllText($filledOutlinePath, $filledOutline, (New-Object System.Text.UTF8Encoding($true)))

  if ($imageInputsProvided) {
    if ([string]::IsNullOrWhiteSpace($ImageMapOutPath)) {
      $resolvedImageMapOutPath = Join-Path $resolvedOutputDir "generated-image-map.json"
    } else {
      $resolvedImageMapOutPath = [System.IO.Path]::GetFullPath($ImageMapOutPath)
    }

    if ([string]::IsNullOrWhiteSpace($FilledDocxWithImagesOutPath)) {
      $resolvedFilledDocxWithImagesOutPath = Join-Path $resolvedOutputDir (([System.IO.Path]::GetFileNameWithoutExtension($resolvedFilledDocxOutPath)) + ".images.docx")
    } else {
      $resolvedFilledDocxWithImagesOutPath = [System.IO.Path]::GetFullPath($FilledDocxWithImagesOutPath)
    }

    $imageMapParams = @{
      DocxPath = $resolvedFilledDocxOutPath
      Format = "json"
      OutFile = $resolvedImageMapOutPath
    }
    if (-not [string]::IsNullOrWhiteSpace($ImageSpecsPath)) {
      $imageMapParams.ImageSpecsPath = $ImageSpecsPath
    } elseif (-not [string]::IsNullOrWhiteSpace($ImageSpecsJson)) {
      $imageMapParams.ImageSpecsJson = $ImageSpecsJson
    } else {
      $imageMapParams.ImagePaths = $ImagePaths
    }

    & (Join-Path $repoRoot "scripts\generate-docx-image-map.ps1") @imageMapParams | Out-Null
    & (Join-Path $repoRoot "scripts\insert-docx-images.ps1") -DocxPath $resolvedFilledDocxOutPath -MappingPath $resolvedImageMapOutPath -OutPath $resolvedFilledDocxWithImagesOutPath -Overwrite | Out-Null

    $filledWithImagesOutlinePath = Join-Path $resolvedOutputDir "filled-template-with-images-outline.md"
    $filledWithImagesOutline = & (Join-Path $repoRoot "scripts\extract-docx-template.ps1") -Path $resolvedFilledDocxWithImagesOutPath -Format markdown | Out-String
    [System.IO.File]::WriteAllText($filledWithImagesOutlinePath, $filledWithImagesOutline, (New-Object System.Text.UTF8Encoding($true)))
  }

  if ($styleOutputRequested) {
    $styleInputPath = if ($null -ne $resolvedFilledDocxWithImagesOutPath) { $resolvedFilledDocxWithImagesOutPath } else { $resolvedFilledDocxOutPath }
    if ([string]::IsNullOrWhiteSpace($StyledDocxOutPath)) {
      $resolvedStyledDocxOutPath = Join-Path $resolvedOutputDir (([System.IO.Path]::GetFileNameWithoutExtension($styleInputPath)) + ".styled.docx")
    } else {
      $resolvedStyledDocxOutPath = [System.IO.Path]::GetFullPath($StyledDocxOutPath)
    }

    $styleParams = @{
      DocxPath = $styleInputPath
      OutPath = $resolvedStyledDocxOutPath
      Overwrite = $true
      Profile = $StyleProfile
    }
    if (-not [string]::IsNullOrWhiteSpace($StyleProfilePath)) {
      $styleParams.ProfilePath = (Resolve-Path -LiteralPath $StyleProfilePath).Path
    }

    $styleResult = & (Join-Path $repoRoot "scripts\format-docx-report-style.ps1") @styleParams
  }
}

$summary = [pscustomobject]@{
  passed = [bool]$validationResult.passed
  mode = $Mode
  responseFormat = $responseFormat
  agent = $Agent
  skillCommand = $SkillCommand
  sessionKey = $SessionKey
  sessionReset = (-not $SkipSessionReset)
  provider = $provider
  model = $model
  skillNames = $skillNames
  skillActive = $skillActive
  outputDir = $resolvedOutputDir
  promptPath = $promptOutPath
  rawOutputPath = $agentOutputPath
  agentJsonPath = $(if ($responseFormat -eq 'json') { $agentJsonPath } else { $null })
  reportPath = $reportPath
  validationPath = $validationPath
  templatePath = $resolvedTemplatePath
  metadataPath = $resolvedMetadataPath
  fieldMapPath = $resolvedFieldMapOutPath
  filledDocxPath = $resolvedFilledDocxOutPath
  filledOutlinePath = $filledOutlinePath
  imageMapPath = $resolvedImageMapOutPath
  filledDocxWithImagesPath = $resolvedFilledDocxWithImagesOutPath
  filledWithImagesOutlinePath = $filledWithImagesOutlinePath
  styledDocxPath = $resolvedStyledDocxOutPath
  requestedStyleProfile = $(if ($styleOutputRequested) { $StyleProfile } else { $null })
  styleProfilePath = $(if ($null -ne $styleResult) { [string]$styleResult.profilePath } elseif (-not [string]::IsNullOrWhiteSpace($StyleProfilePath)) { (Resolve-Path -LiteralPath $StyleProfilePath).Path } else { $null })
  styleProfile = $(if ($null -ne $styleResult) { [string]$styleResult.styleProfile } else { $null })
  resolvedStyleProfile = $(if ($null -ne $styleResult) { [string]$styleResult.resolvedProfile } else { $null })
  styleProfileReason = $(if ($null -ne $styleResult) { [string]$styleResult.profileReason } else { $null })
  appliedStyleSettings = $(if ($null -ne $styleResult) { $styleResult.appliedSettings } else { $null })
  referenceCount = $(if ($null -ne $preparedPromptResult) { [int]$preparedPromptResult.referenceCount } else { 0 })
  referenceSources = $(if ($null -ne $preparedPromptResult) { @($preparedPromptResult.sources) } else { @() })
  charCount = $validationResult.summary.charCount
  errorCount = $validationResult.summary.errorCount
  warningCount = $validationResult.summary.warningCount
}
[System.IO.File]::WriteAllText($summaryPath, ($summary | ConvertTo-Json -Depth 5), (New-Object System.Text.UTF8Encoding($true)))

Write-Output ("E2E passed: {0}" -f $summary.passed)
Write-Output ("Response format: {0}" -f $summary.responseFormat)
if (-not [string]::IsNullOrWhiteSpace([string]$summary.provider) -or -not [string]::IsNullOrWhiteSpace([string]$summary.model)) {
  Write-Output ("Model: {0}/{1}" -f $summary.provider, $summary.model)
}
Write-Output ("Skills in prompt: {0}" -f (($skillNames -join ", ").Trim()))
Write-Output ("Skill active: {0}" -f $summary.skillActive)
Write-Output ("Report path: {0}" -f $reportPath)
Write-Output ("Validation path: {0}" -f $validationPath)
if ($null -ne $resolvedFilledDocxOutPath) {
  Write-Output ("Field-map path: {0}" -f $resolvedFieldMapOutPath)
  Write-Output ("Filled docx path: {0}" -f $resolvedFilledDocxOutPath)
  Write-Output ("Filled outline path: {0}" -f $filledOutlinePath)
  if ($null -ne $resolvedFilledDocxWithImagesOutPath) {
    Write-Output ("Image-map path: {0}" -f $resolvedImageMapOutPath)
    Write-Output ("Filled docx with images path: {0}" -f $resolvedFilledDocxWithImagesOutPath)
    Write-Output ("Filled with images outline path: {0}" -f $filledWithImagesOutlinePath)
  }
  if ($null -ne $resolvedStyledDocxOutPath) {
    Write-Output ("Styled docx path: {0}" -f $resolvedStyledDocxOutPath)
  }
}

if (-not $summary.passed) {
  throw "End-to-end sample validation failed. See $validationPath"
}

