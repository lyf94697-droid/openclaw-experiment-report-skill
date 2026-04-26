[CmdletBinding()]
param(
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

  [ValidateSet("fast", "full")]
  [string]$PipelineMode = "fast",

  [ValidateSet("auto", "default", "compact", "school", "excellent")]
  [string]$StyleProfile = "auto",

  [string]$StyleProfilePath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "report-defaults.ps1")
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

function Convert-TemplateToDocxIfNeeded {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir
  )

  $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
  if ($extension -eq ".docx") {
    return [pscustomobject]@{
      templatePath = $Path
      sourceTemplatePath = $Path
      status = "none"
      converter = "none"
      convertedTemplatePath = $null
    }
  }

  if ($extension -ne ".doc") {
    throw "Only .docx templates are supported directly; .doc templates can be converted with Word COM: $Path"
  }

  $convertedDir = Join-Path $OutputDir "converted-templates"
  New-Item -ItemType Directory -Path $convertedDir -Force | Out-Null
  $convertedPath = Join-Path $convertedDir (([System.IO.Path]::GetFileNameWithoutExtension($Path)) + ".docx")
  if (Test-Path -LiteralPath $convertedPath) {
    Remove-Item -LiteralPath $convertedPath -Force
  }

  $wpsErrorMessage = ""
  $wpsApp = $null
  $wpsDoc = $null
  try {
    $wpsApp = New-Object -ComObject KWPS.Application
    $wpsApp.Visible = $false
    $wpsDoc = $wpsApp.Documents.Open($Path)
    Start-Sleep -Milliseconds 300
    $wpsDoc.SaveAs($convertedPath, 16)
    Start-Sleep -Milliseconds 300
    if (Test-Path -LiteralPath $convertedPath -PathType Leaf) {
      return [pscustomobject]@{
        templatePath = (Resolve-Path -LiteralPath $convertedPath).Path
        sourceTemplatePath = $Path
        status = "converted"
        converter = "wps"
        convertedTemplatePath = (Resolve-Path -LiteralPath $convertedPath).Path
      }
    }
  } catch {
    $wpsErrorMessage = $_.Exception.Message
    if (Test-Path -LiteralPath $convertedPath) {
      Remove-Item -LiteralPath $convertedPath -Force -ErrorAction SilentlyContinue
    }
  } finally {
    if ($null -ne $wpsDoc) {
      try {
        $wpsDoc.Close($false)
      } catch {
        # Best-effort cleanup only.
      }
    }
    if ($null -ne $wpsApp) {
      try {
        $wpsApp.Quit()
      } catch {
        # Best-effort cleanup only.
      }
    }
  }

  $lastErrorMessage = ""
  for ($attempt = 1; $attempt -le 3; $attempt++) {
    $app = $null
    $doc = $null
    try {
      $app = New-Object -ComObject Word.Application
      $app.Visible = $false
      $app.DisplayAlerts = 0
      try {
        $app.AutomationSecurity = 3
      } catch {
        # Older Word automation hosts may not expose AutomationSecurity.
      }
      Start-Sleep -Milliseconds (300 * $attempt)
      $doc = $app.Documents.Open($Path, $false, $true, $false)
      Start-Sleep -Milliseconds (500 * $attempt)
      $doc.SaveAs2($convertedPath, 16)
      Start-Sleep -Milliseconds 300
      break
    } catch {
      $lastErrorMessage = $_.Exception.Message
      if (Test-Path -LiteralPath $convertedPath) {
        Remove-Item -LiteralPath $convertedPath -Force -ErrorAction SilentlyContinue
      }
      if ($attempt -eq 3) {
        throw "Failed to convert .doc template to .docx. WPS error: $wpsErrorMessage. Word COM error after 3 attempts: $lastErrorMessage"
      }
      Start-Sleep -Milliseconds (900 * $attempt)
    } finally {
      if ($null -ne $doc) {
        try {
          $doc.Close($false)
        } catch {
          # Best-effort cleanup only.
        }
      }
      if ($null -ne $app) {
        try {
          $app.Quit()
        } catch {
          # Best-effort cleanup only.
        }
      }
    }
  }

  if (-not (Test-Path -LiteralPath $convertedPath -PathType Leaf)) {
    throw "Word COM did not produce the converted template: $convertedPath"
  }

  return [pscustomobject]@{
    templatePath = (Resolve-Path -LiteralPath $convertedPath).Path
    sourceTemplatePath = $Path
    status = "converted"
    converter = "word"
    convertedTemplatePath = (Resolve-Path -LiteralPath $convertedPath).Path
  }
}

function Get-OptionalTextContent {
  param(
    [AllowNull()]
    [string]$Path,

    [AllowNull()]
    [string]$InlineText
  )

  if (-not [string]::IsNullOrWhiteSpace($Path)) {
    return Get-Content -LiteralPath $Path -Raw -Encoding UTF8
  }

  if (-not [string]::IsNullOrWhiteSpace($InlineText)) {
    return $InlineText
  }

  return ""
}

function Get-JsonObjectOrNull {
  param(
    [AllowNull()]
    [string]$JsonText
  )

  if ([string]::IsNullOrWhiteSpace($JsonText)) {
    return $null
  }

  try {
    return $JsonText | ConvertFrom-Json
  } catch {
    return $null
  }
}

function Get-ImageInputItems {
  param(
    [Parameter(Mandatory = $true)]
    [string]$InputMode,

    [AllowNull()]
    [string]$SpecsPath,

    [AllowNull()]
    [string]$SpecsJson,

    [AllowNull()]
    [string[]]$Paths
  )

  if ([string]::Equals($InputMode, "image-paths", [System.StringComparison]::OrdinalIgnoreCase)) {
    return @(@($Paths | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object {
          [pscustomobject]@{
            path = [string]$_
          }
        }))
  }

  $jsonText = Get-OptionalTextContent -Path $SpecsPath -InlineText $SpecsJson
  $rootObject = Get-JsonObjectOrNull -JsonText $jsonText
  if ($null -eq $rootObject) {
    return @()
  }

  if ($rootObject -is [System.Collections.IEnumerable] -and $rootObject -isnot [string]) {
    return @($rootObject)
  }

  if ($rootObject.PSObject.Properties.Name -contains "images") {
    return @($rootObject.images)
  }

  return @($rootObject)
}

function Get-ImageItemValue {
  param(
    [Parameter(Mandatory = $true)]
    [object]$Item,

    [Parameter(Mandatory = $true)]
    [string[]]$Keys
  )

  foreach ($key in $Keys) {
    if ($Item -is [System.Collections.IDictionary]) {
      if ($Item.Contains($key) -and -not [string]::IsNullOrWhiteSpace([string]$Item[$key])) {
        return ([string]$Item[$key]).Trim()
      }
      continue
    }

    $property = $Item.PSObject.Properties[$key]
    if ($null -ne $property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
      return ([string]$property.Value).Trim()
    }
  }

  return $null
}

function Test-IsFlowchartSignal {
  param(
    [AllowNull()]
    [string]$Text
  )

  return (-not [string]::IsNullOrWhiteSpace($Text) -and $Text -match '(?i)(流程图|flowchart|flow-chart)')
}

function Test-ImageItemsContainFlowchart {
  param(
    [AllowEmptyCollection()]
    [object[]]$Items
  )

  foreach ($item in @($Items)) {
    foreach ($signal in @(
        (Get-ImageItemValue -Item $item -Keys @("caption", "title", "figureCaption")),
        (Get-ImageItemValue -Item $item -Keys @("section", "sectionName", "heading")),
        (Get-ImageItemValue -Item $item -Keys @("path", "imagePath", "file"))
      )) {
      if (Test-IsFlowchartSignal -Text $signal) {
        return $true
      }
    }
  }

  return $false
}

function Test-ImageItemsContainExplicitCaption {
  param(
    [AllowEmptyCollection()]
    [object[]]$Items
  )

  foreach ($item in @($Items)) {
    if (-not [string]::IsNullOrWhiteSpace((Get-ImageItemValue -Item $item -Keys @("caption", "title", "figureCaption")))) {
      return $true
    }
  }

  return $false
}

function Get-CourseDesignFlowchartTitle {
  param(
    [AllowNull()]
    [string]$MetadataPath,

    [AllowNull()]
    [string]$MetadataJson,

    [Parameter(Mandatory = $true)]
    [string]$ReportPath
  )

  $metadataText = Get-OptionalTextContent -Path $MetadataPath -InlineText $MetadataJson
  $metadataRoot = Get-JsonObjectOrNull -JsonText $metadataText
  if ($null -ne $metadataRoot) {
    foreach ($key in @("课题名称", "项目名称", "题目", "标题", "实验名称")) {
      $value = Get-ImageItemValue -Item $metadataRoot -Keys @($key)
      if (-not [string]::IsNullOrWhiteSpace($value)) {
        return ("{0}流程图" -f $value)
      }
    }
  }

  $reportText = Get-Content -LiteralPath $ReportPath -Raw -Encoding UTF8
  foreach ($pattern in @(
      '课题名称[:：]\s*(?<value>.+)',
      '项目名称[:：]\s*(?<value>.+)',
      '题目[:：]\s*(?<value>.+)'
    )) {
    if ($reportText -match $pattern) {
      return ("{0}流程图" -f $matches["value"].Trim())
    }
  }

  return "课程设计实现流程图"
}

function Get-CourseDesignFlowchartSteps {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ReportPath,

    [AllowNull()]
    [string]$RequirementsPath,

    [AllowNull()]
    [string]$RequirementsJson
  )

  $reportText = Get-Content -LiteralPath $ReportPath -Raw -Encoding UTF8
  $requirementsText = Get-OptionalTextContent -Path $RequirementsPath -InlineText $RequirementsJson
  $combinedText = ($reportText + [Environment]::NewLine + $requirementsText)
  $rootTitle = "课程设计系统"
  foreach ($pattern in @(
      '课题名称[:：]\s*(?<value>.+)',
      '项目名称[:：]\s*(?<value>.+)',
      '题目[:：]\s*(?<value>.+)'
    )) {
    if ($reportText -match $pattern) {
      $rootTitle = $matches["value"].Trim()
      break
    }
  }

  $frontendModules = New-Object System.Collections.Generic.List[string]
  if ($combinedText -match '分类|目录|导航|首页|列表') {
    $frontendModules.Add("分类浏览") | Out-Null
  }
  if ($combinedText -match '搜索|查询|检索') {
    $frontendModules.Add("信息检索") | Out-Null
  }
  if ($combinedText -match '详情|展示|说明|结果') {
    $frontendModules.Add("详情展示") | Out-Null
  }
  foreach ($fallbackLabel in @("页面展示", "交互入口", "结果详情")) {
    if ($frontendModules.Count -ge 3) {
      break
    }
    if ($frontendModules -notcontains $fallbackLabel) {
      $frontendModules.Add($fallbackLabel) | Out-Null
    }
  }

  $backendModules = New-Object System.Collections.Generic.List[string]
  if ($combinedText -match '逻辑|算法|调度|推荐|接口|服务端|业务') {
    $backendModules.Add("业务处理") | Out-Null
  }
  if ($combinedText -match '数据库|数据表|SQL|MySQL|SQLite|ER图|存储') {
    $backendModules.Add("数据管理") | Out-Null
  }
  if ($combinedText -match '收藏|订单|成绩|权限|日志|状态|审核') {
    $backendModules.Add("状态维护") | Out-Null
  }
  $backendModules.Add("测试验证") | Out-Null
  foreach ($fallbackLabel in @("异常处理", "日志管理", "配置维护")) {
    if ($backendModules.Count -ge 4) {
      break
    }
    if ($backendModules -notcontains $fallbackLabel) {
      $backendModules.Add($fallbackLabel) | Out-Null
    }
  }

  return @(
    "@TREE $rootTitle",
    ("@GROUP 前台模块|{0}" -f ((@($frontendModules | Select-Object -Unique -First 3)) -join "|")),
    ("@GROUP 后台模块|{0}" -f ((@($backendModules | Select-Object -Unique -First 4)) -join "|"))
  )
}

function Try-NewCourseDesignAutoFlowchart {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ReportPath,

    [AllowNull()]
    [string]$MetadataPath,

    [AllowNull()]
    [string]$MetadataJson,

    [AllowNull()]
    [string]$RequirementsPath,

    [AllowNull()]
    [string]$RequirementsJson,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [Parameter(Mandatory = $true)]
    [string]$RepoRoot,

    [Parameter(Mandatory = $true)]
    [string]$ImageInputMode,

    [AllowNull()]
    [string]$ImageSpecsPath,

    [AllowNull()]
    [string]$ImageSpecsJson,

    [AllowNull()]
    [string[]]$ImagePaths
  )

  $reportText = Get-Content -LiteralPath $ReportPath -Raw -Encoding UTF8
  $requirementsText = Get-OptionalTextContent -Path $RequirementsPath -InlineText $RequirementsJson
  if (($reportText + [Environment]::NewLine + $requirementsText) -match '(?i)(不要流程图|不需要流程图|no flowchart)') {
    return $null
  }

  $existingItems = @(Get-ImageInputItems -InputMode $ImageInputMode -SpecsPath $ImageSpecsPath -SpecsJson $ImageSpecsJson -Paths $ImagePaths)
  if (Test-ImageItemsContainFlowchart -Items $existingItems) {
    return $null
  }

  $rendererPath = Join-Path $RepoRoot "scripts\render-vertical-lab-flowchart.py"
  if (-not (Test-Path -LiteralPath $rendererPath)) {
    return $null
  }

  $artifactsDir = Join-Path $OutputDir "artifacts"
  New-Item -ItemType Directory -Path $artifactsDir -Force | Out-Null

  $stepsPath = Join-Path $artifactsDir "course-design-auto-flowchart.steps.txt"
  $flowchartPath = Join-Path $artifactsDir "course-design-auto-flowchart.png"
  $mergedSpecsPath = Join-Path $artifactsDir "course-design-auto-image-specs.json"
  $flowchartSteps = @(Get-CourseDesignFlowchartSteps -ReportPath $ReportPath -RequirementsPath $RequirementsPath -RequirementsJson $RequirementsJson)
  $flowchartTitle = Get-CourseDesignFlowchartTitle -MetadataPath $MetadataPath -MetadataJson $MetadataJson -ReportPath $ReportPath

  [System.IO.File]::WriteAllLines($stepsPath, $flowchartSteps, (New-Object System.Text.UTF8Encoding($true)))

  $rendered = $false
  foreach ($pythonOption in @(
      @{ command = "python"; prefix = @() },
      @{ command = "py"; prefix = @("-3") }
    )) {
    if ($null -eq (Get-Command $pythonOption.command -ErrorAction SilentlyContinue)) {
      continue
    }

    try {
      & $pythonOption.command @($pythonOption.prefix + @($rendererPath, "--out", $flowchartPath, "--title", $flowchartTitle, "--steps-file", $stepsPath))
      if (Test-Path -LiteralPath $flowchartPath) {
        $rendered = $true
        break
      }
    } catch {
      if (Test-Path -LiteralPath $flowchartPath) {
        Remove-Item -LiteralPath $flowchartPath -Force -ErrorAction SilentlyContinue
      }
    }
  }

  if (-not $rendered) {
    Write-Warning "Skipped course-design auto flowchart because the renderer could not run successfully."
    return $null
  }

  $flowchartCaption = if (Test-ImageItemsContainExplicitCaption -Items $existingItems) {
    "系统总体设计图"
  } else {
    "图1 系统总体设计图"
  }

  $mergedImages = @(
    [ordered]@{
      path = $flowchartPath
      section = "方案设计与实现"
      caption = $flowchartCaption
      widthCm = 15.8
    }
  ) + @($existingItems)

  $mergedSpecsRoot = [ordered]@{
    images = $mergedImages
  }
  [System.IO.File]::WriteAllText($mergedSpecsPath, ($mergedSpecsRoot | ConvertTo-Json -Depth 8), (New-Object System.Text.UTF8Encoding($true)))

  return [pscustomobject]@{
    flowchartPath = $flowchartPath
    imageSpecsPath = $mergedSpecsPath
  }
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$reportProfile = Get-ReportProfile -ProfileName $ReportProfileName -ProfilePath $ReportProfilePath -RepoRoot $repoRoot
$resolvedReportProfilePath = [string]$reportProfile.resolvedProfilePath
$resolvedTemplatePath = Resolve-ExperimentReportTemplatePath `
  -TemplatePath $TemplatePath `
  -ReportProfileName ([string]$reportProfile.name) `
  -ReportProfilePath $resolvedReportProfilePath `
  -RepoRoot $repoRoot
$sourceTemplatePath = $resolvedTemplatePath
$templateConversion = $null
$resolvedReportPath = (Resolve-Path -LiteralPath $ReportPath).Path

$resolvedMetadataPath = $null
if (-not [string]::IsNullOrWhiteSpace($MetadataPath)) {
  $resolvedMetadataPath = (Resolve-Path -LiteralPath $MetadataPath).Path
}

$resolvedRequirementsPath = $null
if (-not [string]::IsNullOrWhiteSpace($RequirementsPath)) {
  $resolvedRequirementsPath = (Resolve-Path -LiteralPath $RequirementsPath).Path
}

$effectiveStyleProfile = if ($PSBoundParameters.ContainsKey("StyleProfile")) {
  $StyleProfile
} else {
  Get-ReportProfileDefaultStyleProfile -Profile $reportProfile
}
$templateFrameDefaulted = ((-not [bool]$CreateTemplateFrameDocx) -and [string]::IsNullOrWhiteSpace($TemplateFrameDocxOutPath) -and (Test-ExperimentReportTemplateFrameDefault -ReportProfileName ([string]$reportProfile.name) -ReportProfilePath $resolvedReportProfilePath))
$shouldCreateTemplateFrameDocx = ([bool]$CreateTemplateFrameDocx) -or (-not [string]::IsNullOrWhiteSpace($TemplateFrameDocxOutPath)) -or $templateFrameDefaulted
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
$runFullPipeline = [string]::Equals($PipelineMode, "full", [System.StringComparison]::OrdinalIgnoreCase)
$shouldRunValidation = $runFullPipeline -or ($requirementsInputMode -ne "none")
$shouldGenerateDebugOutlines = $runFullPipeline

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
  $OutputDir = Join-Path $repoRoot ("tests-output\build-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
}

$resolvedOutputDir = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null

$effectiveImageInputMode = $imageInputMode
$effectiveImageInputsProvided = $imageInputsProvided
$effectiveImageSpecsPath = $ImageSpecsPath
$effectiveImageSpecsJson = $ImageSpecsJson
$effectiveImagePaths = $ImagePaths
$autoCourseDesignFlowchart = $null
if ([string]::Equals([string]$reportProfile.name, "course-design-report", [System.StringComparison]::OrdinalIgnoreCase)) {
  $autoCourseDesignFlowchart = Try-NewCourseDesignAutoFlowchart `
    -ReportPath $resolvedReportPath `
    -MetadataPath $resolvedMetadataPath `
    -MetadataJson $MetadataJson `
    -RequirementsPath $resolvedRequirementsPath `
    -RequirementsJson $RequirementsJson `
    -OutputDir $resolvedOutputDir `
    -RepoRoot $repoRoot `
    -ImageInputMode $imageInputMode `
    -ImageSpecsPath $ImageSpecsPath `
    -ImageSpecsJson $ImageSpecsJson `
    -ImagePaths $ImagePaths
  if ($null -ne $autoCourseDesignFlowchart) {
    $effectiveImageInputMode = "specs-path"
    $effectiveImageInputsProvided = $true
    $effectiveImageSpecsPath = [string]$autoCourseDesignFlowchart.imageSpecsPath
    $effectiveImageSpecsJson = $null
    $effectiveImagePaths = @()
  }
}

$shouldGenerateImagePlan = $effectiveImageInputsProvided -or (-not [string]::IsNullOrWhiteSpace($ImagePlanOutPath))
$shouldRunLayoutCheck = $runFullPipeline -or $effectiveImageInputsProvided

$templateConversion = Convert-TemplateToDocxIfNeeded -Path $resolvedTemplatePath -OutputDir $resolvedOutputDir
$resolvedTemplatePath = [string]$templateConversion.templatePath

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
$resolvedCourseDesignTablesDocxOutPath = $null
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
$courseDesignTablesResult = $null
$expectedLayoutImageCount = -1
$expectedLayoutCaptionCount = -1
$imagePlanLowConfidenceCount = $null
$imagePlanNeedsReview = $null

$validationResult = $null
if ($shouldRunValidation) {
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
  } elseif (-not [string]::IsNullOrWhiteSpace($RequirementsJson)) {
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

if ($shouldGenerateDebugOutlines) {
  $filledOutlinePath = Join-Path $resolvedOutputDir "filled-template-outline.md"
  $filledOutline = & (Join-Path $repoRoot "scripts\extract-docx-template.ps1") -Path $resolvedFilledDocxOutPath -Format markdown | Out-String
  [System.IO.File]::WriteAllText($filledOutlinePath, $filledOutline, (New-Object System.Text.UTF8Encoding($true)))
}

if ($effectiveImageInputsProvided) {
  if ($shouldGenerateImagePlan) {
    $resolvedImagePlanOutPath = if ([string]::IsNullOrWhiteSpace($ImagePlanOutPath)) {
      Join-Path $resolvedOutputDir "image-placement-plan.md"
    } else {
      [System.IO.Path]::GetFullPath($ImagePlanOutPath)
    }
    Ensure-ParentDirectory -Path $resolvedImagePlanOutPath
  }

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
  if ([string]::Equals($effectiveImageInputMode, "specs-path", [System.StringComparison]::OrdinalIgnoreCase)) {
    $imageInputParams.ImageSpecsPath = (Resolve-Path -LiteralPath $effectiveImageSpecsPath).Path
  } elseif ([string]::Equals($effectiveImageInputMode, "specs-json", [System.StringComparison]::OrdinalIgnoreCase)) {
    $imageInputParams.ImageSpecsJson = $effectiveImageSpecsJson
  } else {
    $imageInputParams.ImagePaths = $effectiveImagePaths
  }

  if ($shouldGenerateImagePlan) {
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
  }

  $imageMapParams = $imageInputParams.Clone()
  $imageMapParams.Format = "json"
  $imageMapParams.OutFile = $resolvedImageMapOutPath
  & (Join-Path $repoRoot "scripts\generate-docx-image-map.ps1") @imageMapParams | Out-Null
  $generatedImageMap = (Get-Content -LiteralPath $resolvedImageMapOutPath -Raw -Encoding UTF8) | ConvertFrom-Json
  if ($shouldRunLayoutCheck -and $null -ne $generatedImageMap -and $generatedImageMap.PSObject.Properties.Name -contains "images") {
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

  if ($shouldGenerateDebugOutlines) {
    $filledWithImagesOutlinePath = Join-Path $resolvedOutputDir "filled-template-with-images-outline.md"
    $filledWithImagesOutline = & (Join-Path $repoRoot "scripts\extract-docx-template.ps1") -Path $resolvedFilledDocxWithImagesOutPath -Format markdown | Out-String
    [System.IO.File]::WriteAllText($filledWithImagesOutlinePath, $filledWithImagesOutline, (New-Object System.Text.UTF8Encoding($true)))
  }
}

if ([string]::Equals([string]$reportProfile.name, "course-design-report", [System.StringComparison]::OrdinalIgnoreCase)) {
  $courseDesignTablesInputPath = if ($null -ne $resolvedFilledDocxWithImagesOutPath) { $resolvedFilledDocxWithImagesOutPath } else { $resolvedFilledDocxOutPath }
  $resolvedCourseDesignTablesDocxOutPath = Join-Path $resolvedOutputDir (([System.IO.Path]::GetFileNameWithoutExtension($courseDesignTablesInputPath)) + ".course-tables.docx")
  Ensure-ParentDirectory -Path $resolvedCourseDesignTablesDocxOutPath
  $courseDesignTablesResult = & (Join-Path $repoRoot "scripts\insert-course-design-tables.ps1") `
    -DocxPath $courseDesignTablesInputPath `
    -OutPath $resolvedCourseDesignTablesDocxOutPath `
    -Overwrite
}

if ($styleOutputRequested) {
  $styleInputPath = if ($null -ne $resolvedCourseDesignTablesDocxOutPath) { $resolvedCourseDesignTablesDocxOutPath } elseif ($null -ne $resolvedFilledDocxWithImagesOutPath) { $resolvedFilledDocxWithImagesOutPath } else { $resolvedFilledDocxOutPath }
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

  if ($shouldGenerateDebugOutlines) {
    $styledOutlinePath = Join-Path $resolvedOutputDir "styled-template-outline.md"
    $styledOutline = & (Join-Path $repoRoot "scripts\extract-docx-template.ps1") -Path $resolvedStyledDocxOutPath -Format markdown | Out-String
    [System.IO.File]::WriteAllText($styledOutlinePath, $styledOutline, (New-Object System.Text.UTF8Encoding($true)))
  }
}

$finalDocxPath = if ($null -ne $resolvedStyledDocxOutPath) {
  $resolvedStyledDocxOutPath
} elseif ($null -ne $resolvedCourseDesignTablesDocxOutPath) {
  $resolvedCourseDesignTablesDocxOutPath
} elseif ($null -ne $resolvedFilledDocxWithImagesOutPath) {
  $resolvedFilledDocxWithImagesOutPath
} else {
  $resolvedFilledDocxOutPath
}

if ($shouldCreateTemplateFrameDocx) {
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

if ($shouldRunLayoutCheck) {
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
}

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
  pipelineMode = $PipelineMode
  reportProfileName = [string]$reportProfile.name
  reportProfilePath = $resolvedReportProfilePath
  templatePath = $resolvedTemplatePath
  sourceTemplatePath = $sourceTemplatePath
  templatePathDefaulted = (-not $PSBoundParameters.ContainsKey("TemplatePath") -or [string]::IsNullOrWhiteSpace($TemplatePath))
  templateFrameDefaulted = $templateFrameDefaulted
  fixedExperimentReportStyle = (Test-IsExperimentReportProfile -ReportProfileName ([string]$reportProfile.name) -ReportProfilePath $resolvedReportProfilePath)
  templateConversionStatus = $(if ($null -ne $templateConversion) { [string]$templateConversion.status } else { "none" })
  templateConversionConverter = $(if ($null -ne $templateConversion) { [string]$templateConversion.converter } else { "none" })
  convertedTemplatePath = $(if ($null -ne $templateConversion) { [string]$templateConversion.convertedTemplatePath } else { $null })
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
  courseDesignTablesDocxPath = $resolvedCourseDesignTablesDocxOutPath
  courseDesignTablesInserted = $(if ($null -ne $courseDesignTablesResult -and $courseDesignTablesResult.PSObject.Properties.Name -contains "inserted") { [bool]$courseDesignTablesResult.inserted } else { $null })
  courseDesignTablesCount = $(if ($null -ne $courseDesignTablesResult -and $courseDesignTablesResult.PSObject.Properties.Name -contains "tableCount") { [int]$courseDesignTablesResult.tableCount } else { $null })
  styledDocxPath = $resolvedStyledDocxOutPath
  styledOutlinePath = $styledOutlinePath
  templateFrameDocxPath = $resolvedTemplateFrameDocxOutPath
  layoutCheckPath = $(if ($shouldRunLayoutCheck) { $layoutCheckPath } else { $null })
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
  validationPath = $(if ($shouldRunValidation) { $validationPath } else { $null })
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
  if (-not [string]::IsNullOrWhiteSpace($resolvedImagePlanOutPath)) {
    Write-Output ("Image-plan path: {0}" -f $resolvedImagePlanOutPath)
  }
  Write-Output ("Image-map path: {0}" -f $resolvedImageMapOutPath)
  Write-Output ("Filled docx with images path: {0}" -f $resolvedFilledDocxWithImagesOutPath)
}
if ($null -ne $resolvedCourseDesignTablesDocxOutPath) {
  Write-Output ("Course-design tables docx path: {0}" -f $resolvedCourseDesignTablesDocxOutPath)
}
if ($null -ne $resolvedStyledDocxOutPath) {
  Write-Output ("Styled docx path: {0}" -f $resolvedStyledDocxOutPath)
}
if ($null -ne $resolvedTemplateFrameDocxOutPath) {
  Write-Output ("Template-frame docx path: {0}" -f $resolvedTemplateFrameDocxOutPath)
}
Write-Output ("Final docx path: {0}" -f $finalDocxPath)
if ($shouldRunLayoutCheck) {
  Write-Output ("Layout check path: {0}" -f $layoutCheckPath)
}
Write-Output ("Summary path: {0}" -f $summaryPath)

if ($null -ne $validationResult -and -not $validationResult.passed) {
  throw "Report validation failed. See $validationPath"
}
