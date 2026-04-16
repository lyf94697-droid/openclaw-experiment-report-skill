[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$DocxPath,

  [Parameter(Mandatory = $true)]
  [string]$OutPath,

  [ValidateRange(0, 720)]
  [int]$MetadataCellMarginTwips = 180,

  [ValidateRange(0, 720)]
  [int]$BodyCellMarginTwips = 260,

  [switch]$Overwrite
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

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

function New-WAttr {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$Document,

    [Parameter(Mandatory = $true)]
    [string]$Name,

    [Parameter(Mandatory = $true)]
    [string]$Value
  )

  $attr = $Document.CreateAttribute(
    "w",
    $Name,
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  )
  $attr.Value = $Value
  return $attr
}

function Add-BorderAttrs {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$Document,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlElement]$Element
  )

  foreach ($pair in @(
      @("val", "single"),
      @("sz", "4"),
      @("space", "0"),
      @("color", "000000")
    )) {
    [void]$Element.Attributes.Append((New-WAttr -Document $Document -Name $pair[0] -Value $pair[1]))
  }
}

function New-TemplateBorders {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$Document,

    [Parameter(Mandatory = $true)]
    [string]$Name
  )

  $wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  $borders = $Document.CreateElement("w", $Name, $wNs)
  foreach ($side in @("top", "left", "bottom", "right", "insideH", "insideV")) {
    $element = $Document.CreateElement("w", $side, $wNs)
    Add-BorderAttrs -Document $Document -Element $element
    [void]$borders.AppendChild($element)
  }
  return $borders
}

function Set-CellMargins {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$Document,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlElement]$Cell,

    [Parameter(Mandatory = $true)]
    [int]$MarginTwips
  )

  $wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  $cellProperties = $Cell.SelectSingleNode(
    "./w:tcPr",
    $script:namespaceManagerForCellMargins
  )
  if ($null -eq $cellProperties) {
    $cellProperties = $Document.CreateElement("w", "tcPr", $wNs)
    [void]$Cell.PrependChild($cellProperties)
  }

  foreach ($existingMargins in @($cellProperties.SelectNodes("w:tcMar", $script:namespaceManagerForCellMargins))) {
    [void]$cellProperties.RemoveChild($existingMargins)
  }

  $cellMargins = $Document.CreateElement("w", "tcMar", $wNs)
  foreach ($side in @("top", "left", "bottom", "right")) {
    $margin = $Document.CreateElement("w", $side, $wNs)
    [void]$margin.Attributes.Append((New-WAttr -Document $Document -Name "w" -Value ([string]$MarginTwips)))
    [void]$margin.Attributes.Append((New-WAttr -Document $Document -Name "type" -Value "dxa"))
    [void]$cellMargins.AppendChild($margin)
  }
  [void]$cellProperties.AppendChild($cellMargins)
}

function Get-CellGridSpan {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlElement]$Cell,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $gridSpan = $Cell.SelectSingleNode("./w:tcPr/w:gridSpan", $NamespaceManager)
  if ($null -ne $gridSpan) {
    $value = $gridSpan.GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    if (-not [string]::IsNullOrWhiteSpace($value)) {
      return [int]$value
    }
  }

  return 1
}

function Get-TableColumnSpan {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlElement]$Table,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $maxSpan = 1
  foreach ($row in @($Table.SelectNodes("./w:tr", $NamespaceManager))) {
    $span = 0
    foreach ($cell in @($row.SelectNodes("./w:tc", $NamespaceManager))) {
      $span += Get-CellGridSpan -Cell $cell -NamespaceManager $NamespaceManager
    }
    if ($span -gt $maxSpan) {
      $maxSpan = $span
    }
  }
  return $maxSpan
}

function Get-TableGridWidth {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlElement]$Table,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  $width = 0
  foreach ($gridColumn in @($Table.SelectNodes("./w:tblGrid/w:gridCol", $NamespaceManager))) {
    $value = $gridColumn.GetAttribute("w", $wNs)
    $parsed = 0
    if ([int]::TryParse($value, [ref]$parsed)) {
      $width += $parsed
    }
  }
  if ($width -gt 0) {
    return $width
  }

  foreach ($row in @($Table.SelectNodes("./w:tr", $NamespaceManager))) {
    $rowWidth = 0
    foreach ($cellWidth in @($row.SelectNodes("./w:tc/w:tcPr/w:tcW", $NamespaceManager))) {
      $value = $cellWidth.GetAttribute("w", $wNs)
      $parsed = 0
      if ([int]::TryParse($value, [ref]$parsed)) {
        $rowWidth += $parsed
      }
    }
    if ($rowWidth -gt 0) {
      return $rowWidth
    }
  }

  return 0
}

function Set-TableFixedWidth {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$Document,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlElement]$Table,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager,

    [Parameter(Mandatory = $true)]
    [int]$WidthTwips
  )

  if ($WidthTwips -le 0) {
    return
  }

  $wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  $tableProperties = $Table.SelectSingleNode("w:tblPr", $NamespaceManager)
  if ($null -eq $tableProperties) {
    $tableProperties = $Document.CreateElement("w", "tblPr", $wNs)
    [void]$Table.PrependChild($tableProperties)
  }

  foreach ($existingWidth in @($tableProperties.SelectNodes("w:tblW", $NamespaceManager))) {
    [void]$tableProperties.RemoveChild($existingWidth)
  }
  foreach ($existingLayout in @($tableProperties.SelectNodes("w:tblLayout", $NamespaceManager))) {
    [void]$tableProperties.RemoveChild($existingLayout)
  }

  $tableWidth = $Document.CreateElement("w", "tblW", $wNs)
  [void]$tableWidth.Attributes.Append((New-WAttr -Document $Document -Name "w" -Value ([string]$WidthTwips)))
  [void]$tableWidth.Attributes.Append((New-WAttr -Document $Document -Name "type" -Value "dxa"))

  $fixedLayout = $Document.CreateElement("w", "tblLayout", $wNs)
  [void]$fixedLayout.Attributes.Append((New-WAttr -Document $Document -Name "type" -Value "fixed"))

  [void]$tableProperties.PrependChild($fixedLayout)
  [void]$tableProperties.PrependChild($tableWidth)
}

function Write-ZipFromDirectory {
  param(
    [Parameter(Mandatory = $true)]
    [string]$SourceDirectory,

    [Parameter(Mandatory = $true)]
    [string]$DestinationPath
  )

  $fileStream = [System.IO.File]::Open($DestinationPath, [System.IO.FileMode]::CreateNew)
  try {
    $archive = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
      $rootFullName = (Get-Item -LiteralPath $SourceDirectory).FullName.TrimEnd("\") + "\"
      foreach ($file in Get-ChildItem -LiteralPath $SourceDirectory -Recurse -File) {
        $relativeName = $file.FullName.Substring($rootFullName.Length).Replace("\", "/")
        $entry = $archive.CreateEntry($relativeName, [System.IO.Compression.CompressionLevel]::Optimal)
        $entryStream = $entry.Open()
        try {
          $inputStream = [System.IO.File]::OpenRead($file.FullName)
          try {
            $inputStream.CopyTo($entryStream)
          } finally {
            $inputStream.Dispose()
          }
        } finally {
          $entryStream.Dispose()
        }
      }
    } finally {
      $archive.Dispose()
    }
  } finally {
    $fileStream.Dispose()
  }
}

$resolvedDocxPath = (Resolve-Path -LiteralPath $DocxPath).Path
$resolvedOutPath = [System.IO.Path]::GetFullPath($OutPath)
Ensure-ParentDirectory -Path $resolvedOutPath

if ((Test-Path -LiteralPath $resolvedOutPath) -and -not $Overwrite) {
  throw "Output already exists. Use -Overwrite to replace: $resolvedOutPath"
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("docx-template-frame-" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

try {
  $sourceZipPath = Join-Path $tempRoot "source.zip"
  $unzipDir = Join-Path $tempRoot "unzipped"
  Copy-Item -LiteralPath $resolvedDocxPath -Destination $sourceZipPath
  [System.IO.Compression.ZipFile]::ExtractToDirectory($sourceZipPath, $unzipDir)

  $documentXmlPath = Join-Path $unzipDir "word\document.xml"
  if (-not (Test-Path -LiteralPath $documentXmlPath)) {
    throw "The docx package is missing word/document.xml: $resolvedDocxPath"
  }

  [xml]$documentXml = Get-Content -LiteralPath $documentXmlPath -Raw -Encoding UTF8
  $wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  $namespaceManager = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
  $namespaceManager.AddNamespace("w", $wNs)
  $script:namespaceManagerForCellMargins = $namespaceManager

  $body = $documentXml.SelectSingleNode("//w:body", $namespaceManager)
  if ($null -eq $body) {
    throw "The docx package is missing w:body: $resolvedDocxPath"
  }

  $mainTable = $body.SelectSingleNode("w:tbl[1]", $namespaceManager)
  if ($null -eq $mainTable) {
    throw "The document has no top-level metadata table to extend: $resolvedDocxPath"
  }
  $mainTableWidthTwips = Get-TableGridWidth -Table $mainTable -NamespaceManager $namespaceManager
  Set-TableFixedWidth -Document $documentXml -Table $mainTable -NamespaceManager $namespaceManager -WidthTwips $mainTableWidthTwips

  foreach ($table in @($documentXml.SelectNodes("//w:tbl", $namespaceManager))) {
    $tableProperties = $table.SelectSingleNode("w:tblPr", $namespaceManager)
    if ($null -eq $tableProperties) {
      $tableProperties = $documentXml.CreateElement("w", "tblPr", $wNs)
      [void]$table.PrependChild($tableProperties)
    }

    foreach ($existingBorder in @($tableProperties.SelectNodes("w:tblBorders", $namespaceManager))) {
      [void]$tableProperties.RemoveChild($existingBorder)
    }
    [void]$tableProperties.AppendChild((New-TemplateBorders -Document $documentXml -Name "tblBorders"))
  }

  foreach ($cell in @($mainTable.SelectNodes("./w:tr/w:tc", $namespaceManager))) {
    Set-CellMargins -Document $documentXml -Cell $cell -MarginTwips $MetadataCellMarginTwips
  }

  $nodesToMove = New-Object System.Collections.Generic.List[System.Xml.XmlNode]
  $seenMainTable = $false
  foreach ($child in @($body.ChildNodes)) {
    if ($child -eq $mainTable) {
      $seenMainTable = $true
      continue
    }
    if (-not $seenMainTable) {
      continue
    }
    if ($child.LocalName -eq "sectPr") {
      continue
    }
    [void]$nodesToMove.Add($child)
  }

  $movedBlockCount = 0
  if ($nodesToMove.Count -gt 0) {
    $columnSpan = Get-TableColumnSpan -Table $mainTable -NamespaceManager $namespaceManager
    $tableRow = $documentXml.CreateElement("w", "tr", $wNs)
    $tableCell = $documentXml.CreateElement("w", "tc", $wNs)
    $cellProperties = $documentXml.CreateElement("w", "tcPr", $wNs)
    $cellWidth = $documentXml.CreateElement("w", "tcW", $wNs)
    $bodyCellWidthTwips = if ($mainTableWidthTwips -gt 0) { $mainTableWidthTwips } else { 8522 }
    [void]$cellWidth.Attributes.Append((New-WAttr -Document $documentXml -Name "w" -Value ([string]$bodyCellWidthTwips)))
    [void]$cellWidth.Attributes.Append((New-WAttr -Document $documentXml -Name "type" -Value "dxa"))
    [void]$cellProperties.AppendChild($cellWidth)

    $gridSpan = $documentXml.CreateElement("w", "gridSpan", $wNs)
    [void]$gridSpan.Attributes.Append((New-WAttr -Document $documentXml -Name "val" -Value ([string]$columnSpan)))
    [void]$cellProperties.AppendChild($gridSpan)

    [void]$tableCell.AppendChild($cellProperties)
    Set-CellMargins -Document $documentXml -Cell $tableCell -MarginTwips $BodyCellMarginTwips

    foreach ($node in $nodesToMove) {
      [void]$body.RemoveChild($node)
      [void]$tableCell.AppendChild($node)
      $movedBlockCount++
    }

    $lastCellChild = if ($tableCell.ChildNodes.Count -gt 0) {
      $tableCell.ChildNodes[$tableCell.ChildNodes.Count - 1]
    } else {
      $null
    }
    if ($null -eq $lastCellChild -or $lastCellChild.LocalName -ne "p") {
      $emptyParagraph = $documentXml.CreateElement("w", "p", $wNs)
      [void]$tableCell.AppendChild($emptyParagraph)
    }

    [void]$tableRow.AppendChild($tableCell)
    [void]$mainTable.AppendChild($tableRow)
  }

  $writerSettings = New-Object System.Xml.XmlWriterSettings
  $writerSettings.Encoding = New-Object System.Text.UTF8Encoding($false)
  $writerSettings.Indent = $false
  $writer = [System.Xml.XmlWriter]::Create($documentXmlPath, $writerSettings)
  try {
    $documentXml.Save($writer)
  } finally {
    $writer.Close()
  }

  if (Test-Path -LiteralPath $resolvedOutPath) {
    Remove-Item -LiteralPath $resolvedOutPath -Force
  }
  Write-ZipFromDirectory -SourceDirectory $unzipDir -DestinationPath $resolvedOutPath

  $result = [pscustomobject]@{
    sourceDocxPath = $resolvedDocxPath
    outputDocxPath = $resolvedOutPath
    movedBlockCount = $movedBlockCount
    mainTableRowCount = @($mainTable.SelectNodes("./w:tr", $namespaceManager)).Count
    topLevelTableCount = @($body.SelectNodes("./w:tbl", $namespaceManager)).Count
    tableBorderStyle = "single sz=4"
    tableWidthTwips = $mainTableWidthTwips
    metadataCellMarginTwips = $MetadataCellMarginTwips
    bodyCellMarginTwips = $BodyCellMarginTwips
  }

  Write-Output $result
} finally {
  Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}
