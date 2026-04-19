[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string[]]$Path,

  [string]$OutputDir,

  [ValidateSet("auto", "wps", "word")]
  [string]$Converter = "auto",

  [switch]$ConvertPdf,

  [switch]$NoExtractOutline,

  [switch]$Overwrite
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-SafeFileNameBase {
  param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [int]$Index
  )

  $baseName = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)
  foreach ($invalidChar in [System.IO.Path]::GetInvalidFileNameChars()) {
    $baseName = $baseName.Replace([string]$invalidChar, "-")
  }
  $baseName = ($baseName -replace '\s+', '-').Trim('-')
  if ([string]::IsNullOrWhiteSpace($baseName)) {
    $baseName = "template"
  }
  return ("{0:D2}-{1}" -f $Index, $baseName)
}

function Test-ComObjectAvailable {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ProgId
  )

  try {
    $app = New-Object -ComObject $ProgId
    try {
      if ($app.PSObject.Methods.Name -contains "Quit") {
        $app.Quit()
      }
    } catch {
      # Best-effort cleanup only.
    }
    return $true
  } catch {
    return $false
  }
}

function Convert-WithWps {
  param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$TargetPath
  )

  $app = New-Object -ComObject KWPS.Application
  $app.Visible = $false
  try {
    $doc = $app.Documents.Open($SourcePath)
    try {
      Start-Sleep -Milliseconds 300
      $doc.SaveAs($TargetPath, 16)
      Start-Sleep -Milliseconds 300
      $tableCount = $doc.Tables.Count
      $inlineShapeCount = $doc.InlineShapes.Count
      $shapeCount = $doc.Shapes.Count
      return [pscustomobject]@{
        status = "converted"
        converter = "wps"
        message = ""
        tableCount = $tableCount
        inlineShapeCount = $inlineShapeCount
        shapeCount = $shapeCount
      }
    } finally {
      try {
        $doc.Close($false)
      } catch {
        # WPS may reject close calls while background conversion settles.
      }
    }
  } finally {
    try {
      $app.Quit()
    } catch {
      # Best-effort cleanup only.
    }
  }
}

function Convert-WithWord {
  param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$TargetPath
  )

  $app = New-Object -ComObject Word.Application
  $app.Visible = $false
  $app.DisplayAlerts = 0
  try {
    $doc = $app.Documents.Open($SourcePath, $false, $true, $false)
    try {
      Start-Sleep -Milliseconds 500
      $doc.SaveAs2($TargetPath, 16)
      Start-Sleep -Milliseconds 500
      $pageCount = $doc.ComputeStatistics(2)
      $tableCount = $doc.Tables.Count
      $inlineShapeCount = $doc.InlineShapes.Count
      $shapeCount = $doc.Shapes.Count
      return [pscustomobject]@{
        status = "converted"
        converter = "word"
        message = ""
        pageCount = $pageCount
        tableCount = $tableCount
        inlineShapeCount = $inlineShapeCount
        shapeCount = $shapeCount
      }
    } finally {
      try {
        $doc.Close($false)
      } catch {
        # Word can reject close calls if conversion is still settling.
      }
    }
  } finally {
    try {
      $app.Quit()
    } catch {
      # Best-effort cleanup only.
    }
  }
}

function Copy-DocxReference {
  param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$TargetPath
  )

  Copy-Item -LiteralPath $SourcePath -Destination $TargetPath -Force
  return [pscustomobject]@{
    status = "copied"
    converter = "none"
    message = ""
    tableCount = $null
    inlineShapeCount = $null
    shapeCount = $null
  }
}

function Copy-PdfReference {
  param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath
  )

  return [pscustomobject]@{
    status = "copied"
    converter = "none"
    message = "PDF reference copied only. Pass -ConvertPdf to let Word COM convert it to docx for outline extraction."
    pageCount = $null
    tableCount = $null
    inlineShapeCount = $null
    shapeCount = $null
  }
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
  $OutputDir = Join-Path $repoRoot ("tests-output\real-template-references-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
}

$resolvedOutputDir = [System.IO.Path]::GetFullPath($OutputDir)
if ((Test-Path -LiteralPath $resolvedOutputDir) -and (-not $Overwrite)) {
  throw "OutputDir already exists: $resolvedOutputDir. Pass -Overwrite to replace or choose another OutputDir."
}

New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null
$sourceCopiesDir = Join-Path $resolvedOutputDir "source-copies"
$convertedDir = Join-Path $resolvedOutputDir "converted-docx"
$outlinesDir = Join-Path $resolvedOutputDir "outlines"
New-Item -ItemType Directory -Path $sourceCopiesDir, $convertedDir, $outlinesDir -Force | Out-Null

$requestedExtensions = @(
  foreach ($rawPath in $Path) {
    $resolvedPath = (Resolve-Path -LiteralPath $rawPath).Path
    (Get-Item -LiteralPath $resolvedPath).Extension.ToLowerInvariant()
  }
)
$hasDocReferences = $requestedExtensions -contains ".doc"
$hasPdfReferences = $requestedExtensions -contains ".pdf"
$needsWps = $hasDocReferences -and ($Converter -eq "auto" -or $Converter -eq "wps")

$wpsAvailable = if ($needsWps) { Test-ComObjectAvailable -ProgId "KWPS.Application" } else { $false }
$needsWordForDoc = $hasDocReferences -and ($Converter -eq "word" -or ($Converter -eq "auto" -and -not $wpsAvailable))
$needsWord = $needsWordForDoc -or ($hasPdfReferences -and $ConvertPdf)
$wordAvailable = if ($needsWord) { Test-ComObjectAvailable -ProgId "Word.Application" } else { $false }

$results = New-Object System.Collections.Generic.List[object]
$index = 0
foreach ($rawPath in $Path) {
  $index++
  $resolvedPath = (Resolve-Path -LiteralPath $rawPath).Path
  $sourceItem = Get-Item -LiteralPath $resolvedPath
  $safeBaseName = Get-SafeFileNameBase -SourcePath $resolvedPath -Index $index
  $sourceCopyPath = Join-Path $sourceCopiesDir ($safeBaseName + $sourceItem.Extension)
  $convertedDocxPath = Join-Path $convertedDir ($safeBaseName + ".docx")
  $outlinePath = Join-Path $outlinesDir ($safeBaseName + ".outline.md")

  Copy-Item -LiteralPath $resolvedPath -Destination $sourceCopyPath -Force
  Unblock-File -LiteralPath $sourceCopyPath -ErrorAction SilentlyContinue

  $extension = $sourceItem.Extension.ToLowerInvariant()
  $conversion = $null
  try {
    if ($extension -eq ".docx") {
      $conversion = Copy-DocxReference -SourcePath $sourceCopyPath -TargetPath $convertedDocxPath
    } elseif ($extension -eq ".doc") {
      if (($Converter -eq "auto" -or $Converter -eq "wps") -and $wpsAvailable) {
        $conversion = Convert-WithWps -SourcePath $sourceCopyPath -TargetPath $convertedDocxPath
      } elseif ($Converter -eq "wps") {
        throw "WPS COM is not available on this machine."
      } elseif (($Converter -eq "auto" -or $Converter -eq "word") -and $wordAvailable) {
        $conversion = Convert-WithWord -SourcePath $sourceCopyPath -TargetPath $convertedDocxPath
      } else {
        throw "No supported Word/WPS converter is available on this machine."
      }
    } elseif ($extension -eq ".pdf") {
      if (-not $ConvertPdf) {
        $conversion = Copy-PdfReference -SourcePath $sourceCopyPath
        $convertedDocxPath = ""
      } elseif ($Converter -eq "wps") {
        throw "PDF conversion uses Word COM; rerun with -Converter auto or -Converter word."
      } elseif (-not $wordAvailable) {
        throw "Word COM is required to convert PDF references into docx."
      } else {
        $conversion = Convert-WithWord -SourcePath $sourceCopyPath -TargetPath $convertedDocxPath
      }
    } else {
      throw "Unsupported reference extension: $extension"
    }
  } catch {
    $conversion = [pscustomobject]@{
      status = "failed"
      converter = $Converter
      message = $_.Exception.Message
      tableCount = $null
      inlineShapeCount = $null
      shapeCount = $null
    }
    $convertedDocxPath = ""
  }

  if (
    -not $NoExtractOutline `
      -and $conversion.status -ne "failed" `
      -and -not [string]::IsNullOrWhiteSpace($convertedDocxPath) `
      -and (Test-Path -LiteralPath $convertedDocxPath -PathType Leaf)
  ) {
    & (Join-Path $repoRoot "scripts\extract-docx-template.ps1") -Path $convertedDocxPath -Format markdown |
      Set-Content -LiteralPath $outlinePath -Encoding UTF8
  } else {
    $outlinePath = ""
  }

  [void]$results.Add([pscustomobject]@{
      sourcePath = $resolvedPath
      sourceCopyPath = $sourceCopyPath
      convertedDocxPath = $convertedDocxPath
      outlinePath = $outlinePath
      status = [string]$conversion.status
      converter = [string]$conversion.converter
      message = [string]$conversion.message
      pageCount = if ($conversion.PSObject.Properties.Name -contains "pageCount") { $conversion.pageCount } else { $null }
      tableCount = $conversion.tableCount
      inlineShapeCount = $conversion.inlineShapeCount
      shapeCount = $conversion.shapeCount
    })
}

$summary = [pscustomobject]@{
  outputDir = $resolvedOutputDir
  sourceCopiesDir = $sourceCopiesDir
  convertedDocxDir = $convertedDir
  outlinesDir = $outlinesDir
  requestedConverter = $Converter
  wpsAvailable = $wpsAvailable
  wordAvailable = $wordAvailable
  importedCount = $results.Count
  failedCount = @($results | Where-Object { $_.status -eq "failed" }).Count
  imported = $results.ToArray()
}

$summaryPath = Join-Path $resolvedOutputDir "import-summary.json"
[System.IO.File]::WriteAllText(
  $summaryPath,
  (($summary | ConvertTo-Json -Depth 8) + [Environment]::NewLine),
  (New-Object System.Text.UTF8Encoding($true))
)

$summary
