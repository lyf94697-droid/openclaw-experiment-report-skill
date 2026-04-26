[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$DocxPath,

  [Parameter(Mandatory = $true)]
  [string]$OutPath,

  [ValidateRange(0, 720)]
  [int]$MetadataCellMarginTwips = 108,

  [ValidateRange(0, 720)]
  [int]$BodyCellMarginTwips = 108,

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

function New-RAttr {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$Document,

    [Parameter(Mandatory = $true)]
    [string]$Name,

    [Parameter(Mandatory = $true)]
    [string]$Value
  )

  $attr = $Document.CreateAttribute(
    "r",
    $Name,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  )
  $attr.Value = $Value
  return $attr
}

function Set-WAttrValue {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlElement]$Element,

    [Parameter(Mandatory = $true)]
    [string]$Name,

    [Parameter(Mandatory = $true)]
    [string]$Value
  )

  $Element.SetAttribute($Name, "http://schemas.openxmlformats.org/wordprocessingml/2006/main", $Value)
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
      @("color", "auto")
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

function New-FrameOnlyBorders {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$Document,

    [Parameter(Mandatory = $true)]
    [string]$Name
  )

  $wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  $borders = $Document.CreateElement("w", $Name, $wNs)
  foreach ($side in @("top", "left", "bottom", "right")) {
    $element = $Document.CreateElement("w", $side, $wNs)
    Add-BorderAttrs -Document $Document -Element $element
    [void]$borders.AppendChild($element)
  }
  foreach ($side in @("insideH", "insideV")) {
    $element = $Document.CreateElement("w", $side, $wNs)
    [void]$element.Attributes.Append((New-WAttr -Document $Document -Name "val" -Value "nil"))
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
    [int]$MarginTwips,

    [int]$TopTwips = $MarginTwips,

    [int]$LeftTwips = $MarginTwips,

    [int]$BottomTwips = $MarginTwips,

    [int]$RightTwips = $MarginTwips
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
  foreach ($marginSpec in @(
      @{ Name = "top"; Value = $TopTwips },
      @{ Name = "left"; Value = $LeftTwips },
      @{ Name = "bottom"; Value = $BottomTwips },
      @{ Name = "right"; Value = $RightTwips }
    )) {
    $margin = $Document.CreateElement("w", $marginSpec.Name, $wNs)
    [void]$margin.Attributes.Append((New-WAttr -Document $Document -Name "w" -Value ([string]$marginSpec.Value)))
    [void]$margin.Attributes.Append((New-WAttr -Document $Document -Name "type" -Value "dxa"))
    [void]$cellMargins.AppendChild($margin)
  }
  [void]$cellProperties.AppendChild($cellMargins)
}

function Set-CellHorizontalBorderValues {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$Document,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlElement]$Cell,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager,

    [ValidateSet("single", "nil")]
    [string]$Top = "nil",

    [ValidateSet("single", "nil")]
    [string]$Bottom = "nil"
  )

  $wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  $cellProperties = $Cell.SelectSingleNode("./w:tcPr", $NamespaceManager)
  if ($null -eq $cellProperties) {
    $cellProperties = $Document.CreateElement("w", "tcPr", $wNs)
    [void]$Cell.PrependChild($cellProperties)
  }

  foreach ($existingBorders in @($cellProperties.SelectNodes("w:tcBorders", $NamespaceManager))) {
    [void]$cellProperties.RemoveChild($existingBorders)
  }

  $cellBorders = $Document.CreateElement("w", "tcBorders", $wNs)
  foreach ($borderSpec in @(
      @{ Name = "top"; Value = $Top },
      @{ Name = "bottom"; Value = $Bottom }
    )) {
    $border = $Document.CreateElement("w", $borderSpec.Name, $wNs)
    if ([string]$borderSpec.Value -eq "single") {
      Add-BorderAttrs -Document $Document -Element $border
    } else {
      [void]$border.Attributes.Append((New-WAttr -Document $Document -Name "val" -Value "nil"))
    }
    [void]$cellBorders.AppendChild($border)
  }
  [void]$cellProperties.AppendChild($cellBorders)
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

function Get-OrCreateTableProperties {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$Document,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlElement]$Table,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  $tableProperties = $Table.SelectSingleNode("./w:tblPr", $NamespaceManager)
  if ($null -eq $tableProperties) {
    $tableProperties = $Document.CreateElement("w", "tblPr", $wNs)
    [void]$Table.PrependChild($tableProperties)
  }

  return $tableProperties
}

function Set-TableWidth {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$Document,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlElement]$Table,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager,

    [Parameter(Mandatory = $true)]
    [string]$WidthValue,

    [Parameter(Mandatory = $true)]
    [string]$WidthType
  )

  $wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  $tableProperties = Get-OrCreateTableProperties -Document $Document -Table $Table -NamespaceManager $NamespaceManager
  foreach ($existingWidth in @($tableProperties.SelectNodes("w:tblW", $NamespaceManager))) {
    [void]$tableProperties.RemoveChild($existingWidth)
  }

  $tableWidth = $Document.CreateElement("w", "tblW", $wNs)
  [void]$tableWidth.Attributes.Append((New-WAttr -Document $Document -Name "w" -Value $WidthValue))
  [void]$tableWidth.Attributes.Append((New-WAttr -Document $Document -Name "type" -Value $WidthType))
  [void]$tableProperties.PrependChild($tableWidth)
}

function Set-TableLayout {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$Document,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlElement]$Table,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager,

    [Parameter(Mandatory = $true)]
    [ValidateSet("autofit", "fixed")]
    [string]$LayoutType
  )

  $wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  $tableProperties = Get-OrCreateTableProperties -Document $Document -Table $Table -NamespaceManager $NamespaceManager
  foreach ($existingLayout in @($tableProperties.SelectNodes("w:tblLayout", $NamespaceManager))) {
    [void]$tableProperties.RemoveChild($existingLayout)
  }

  $layout = $Document.CreateElement("w", "tblLayout", $wNs)
  [void]$layout.Attributes.Append((New-WAttr -Document $Document -Name "type" -Value $LayoutType))
  [void]$tableProperties.AppendChild($layout)
}

function Set-TableJustification {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$Document,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlElement]$Table,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager,

    [Parameter(Mandatory = $true)]
    [ValidateSet("left", "center", "right")]
    [string]$Value
  )

  $wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  $tableProperties = Get-OrCreateTableProperties -Document $Document -Table $Table -NamespaceManager $NamespaceManager
  foreach ($existingJc in @($tableProperties.SelectNodes("w:jc", $NamespaceManager))) {
    [void]$tableProperties.RemoveChild($existingJc)
  }

  $justification = $Document.CreateElement("w", "jc", $wNs)
  [void]$justification.Attributes.Append((New-WAttr -Document $Document -Name "val" -Value $Value))
  [void]$tableProperties.AppendChild($justification)
}

function Set-TableCellMargins {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$Document,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlElement]$Table,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager,

    [Parameter(Mandatory = $true)]
    [int]$TopTwips,

    [Parameter(Mandatory = $true)]
    [int]$LeftTwips,

    [Parameter(Mandatory = $true)]
    [int]$BottomTwips,

    [Parameter(Mandatory = $true)]
    [int]$RightTwips
  )

  $wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  $tableProperties = Get-OrCreateTableProperties -Document $Document -Table $Table -NamespaceManager $NamespaceManager
  foreach ($existingMargins in @($tableProperties.SelectNodes("w:tblCellMar", $NamespaceManager))) {
    [void]$tableProperties.RemoveChild($existingMargins)
  }

  $cellMargins = $Document.CreateElement("w", "tblCellMar", $wNs)
  foreach ($marginSpec in @(
      @{ Name = "top"; Value = $TopTwips },
      @{ Name = "left"; Value = $LeftTwips },
      @{ Name = "bottom"; Value = $BottomTwips },
      @{ Name = "right"; Value = $RightTwips }
    )) {
    $margin = $Document.CreateElement("w", $marginSpec.Name, $wNs)
    [void]$margin.Attributes.Append((New-WAttr -Document $Document -Name "w" -Value ([string]$marginSpec.Value)))
    [void]$margin.Attributes.Append((New-WAttr -Document $Document -Name "type" -Value "dxa"))
    [void]$cellMargins.AppendChild($margin)
  }

  [void]$tableProperties.AppendChild($cellMargins)
}

function Set-TableBorders {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$Document,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlElement]$Table,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $tableProperties = Get-OrCreateTableProperties -Document $Document -Table $Table -NamespaceManager $NamespaceManager
  foreach ($existingBorder in @($tableProperties.SelectNodes("w:tblBorders", $NamespaceManager))) {
    [void]$tableProperties.RemoveChild($existingBorder)
  }

  [void]$tableProperties.AppendChild((New-TemplateBorders -Document $Document -Name "tblBorders"))
}

function Set-TableGridColumns {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$Document,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlElement]$Table,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager,

    [Parameter(Mandatory = $true)]
    [int[]]$ColumnWidths
  )

  $wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  foreach ($existingGrid in @($Table.SelectNodes("./w:tblGrid", $NamespaceManager))) {
    [void]$Table.RemoveChild($existingGrid)
  }

  $tableGrid = $Document.CreateElement("w", "tblGrid", $wNs)
  foreach ($width in $ColumnWidths) {
    $gridColumn = $Document.CreateElement("w", "gridCol", $wNs)
    [void]$gridColumn.Attributes.Append((New-WAttr -Document $Document -Name "w" -Value ([string]$width)))
    [void]$tableGrid.AppendChild($gridColumn)
  }

  $insertAfter = $Table.SelectSingleNode("./w:tblPr", $NamespaceManager)
  if ($null -ne $insertAfter) {
    [void]$Table.InsertAfter($tableGrid, $insertAfter)
  } elseif ($Table.HasChildNodes) {
    [void]$Table.InsertBefore($tableGrid, $Table.FirstChild)
  } else {
    [void]$Table.AppendChild($tableGrid)
  }
}

function Remove-CellMargins {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlElement]$Cell,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $cellProperties = $Cell.SelectSingleNode("./w:tcPr", $NamespaceManager)
  if ($null -eq $cellProperties) {
    return
  }

  foreach ($existingMargins in @($cellProperties.SelectNodes("w:tcMar", $NamespaceManager))) {
    [void]$cellProperties.RemoveChild($existingMargins)
  }
}

function Set-CellWidthAndSpan {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$Document,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlElement]$Cell,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager,

    [Parameter(Mandatory = $true)]
    [string]$WidthValue,

    [Parameter(Mandatory = $true)]
    [string]$WidthType,

    [Parameter(Mandatory = $true)]
    [int]$GridSpan
  )

  $wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  $cellProperties = $Cell.SelectSingleNode("./w:tcPr", $NamespaceManager)
  if ($null -eq $cellProperties) {
    $cellProperties = $Document.CreateElement("w", "tcPr", $wNs)
    [void]$Cell.PrependChild($cellProperties)
  }

  foreach ($existingWidth in @($cellProperties.SelectNodes("w:tcW", $NamespaceManager))) {
    [void]$cellProperties.RemoveChild($existingWidth)
  }
  foreach ($existingSpan in @($cellProperties.SelectNodes("w:gridSpan", $NamespaceManager))) {
    [void]$cellProperties.RemoveChild($existingSpan)
  }

  $cellWidth = $Document.CreateElement("w", "tcW", $wNs)
  [void]$cellWidth.Attributes.Append((New-WAttr -Document $Document -Name "w" -Value $WidthValue))
  [void]$cellWidth.Attributes.Append((New-WAttr -Document $Document -Name "type" -Value $WidthType))
  [void]$cellProperties.AppendChild($cellWidth)

  if ($GridSpan -gt 1) {
    $gridSpanElement = $Document.CreateElement("w", "gridSpan", $wNs)
    [void]$gridSpanElement.Attributes.Append((New-WAttr -Document $Document -Name "val" -Value ([string]$GridSpan)))
    [void]$cellProperties.AppendChild($gridSpanElement)
  }
}

function Set-MetadataRowsToTemplateStandard {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$Document,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlElement]$Table,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager,

    [Parameter(Mandatory = $true)]
    [int]$MetadataCellMarginTwips
  )

  Set-TableWidth -Document $Document -Table $Table -NamespaceManager $NamespaceManager -WidthValue "10080" -WidthType "dxa"
  Set-TableJustification -Document $Document -Table $Table -NamespaceManager $NamespaceManager -Value "center"
  Set-TableLayout -Document $Document -Table $Table -NamespaceManager $NamespaceManager -LayoutType "fixed"
  Set-TableCellMargins -Document $Document -Table $Table -NamespaceManager $NamespaceManager -TopTwips 0 -LeftTwips $MetadataCellMarginTwips -BottomTwips 0 -RightTwips $MetadataCellMarginTwips
  Set-TableBorders -Document $Document -Table $Table -NamespaceManager $NamespaceManager
  Set-TableGridColumns -Document $Document -Table $Table -NamespaceManager $NamespaceManager -ColumnWidths @(3535, 426, 1277, 1686, 3156)

  $rowSpecs = New-Object object[] 4
  $rowSpecs[0] = @(
    @{ Width = "3535"; Type = "dxa"; Span = 1 },
    @{ Width = "3389"; Type = "dxa"; Span = 3 },
    @{ Width = "3156"; Type = "dxa"; Span = 1 }
  )
  $rowSpecs[1] = @(
    @{ Width = "3961"; Type = "dxa"; Span = 2 },
    @{ Width = "6119"; Type = "dxa"; Span = 3 }
  )
  $rowSpecs[2] = @(
    @{ Width = "10080"; Type = "dxa"; Span = 5 }
  )
  $rowSpecs[3] = @(
    @{ Width = "5238"; Type = "dxa"; Span = 3 },
    @{ Width = "4842"; Type = "dxa"; Span = 2 }
  )

  $rows = @($Table.SelectNodes("./w:tr", $NamespaceManager))
  if ($rows.Count -lt $rowSpecs.Count) {
    return
  }

  for ($rowIndex = 0; $rowIndex -lt $rowSpecs.Count; $rowIndex++) {
    $cells = @($rows[$rowIndex].SelectNodes("./w:tc", $NamespaceManager))
    $cellSpecs = @($rowSpecs[$rowIndex])
    if ($cells.Count -ne $cellSpecs.Count) {
      return
    }
  }

  for ($rowIndex = 0; $rowIndex -lt $rowSpecs.Count; $rowIndex++) {
    $cells = @($rows[$rowIndex].SelectNodes("./w:tc", $NamespaceManager))
    $cellSpecs = @($rowSpecs[$rowIndex])
    for ($cellIndex = 0; $cellIndex -lt $cellSpecs.Count; $cellIndex++) {
      $cellSpec = $cellSpecs[$cellIndex]
      $widthValue = [string]($cellSpec["Width"])
      $widthType = [string]($cellSpec["Type"])
      $gridSpanValue = [int]($cellSpec["Span"])
      Set-CellWidthAndSpan -Document $Document -Cell $cells[$cellIndex] -NamespaceManager $NamespaceManager -WidthValue $widthValue -WidthType $widthType -GridSpan $gridSpanValue
      Remove-CellMargins -Cell $cells[$cellIndex] -NamespaceManager $NamespaceManager
    }
  }
}

function Ensure-ParagraphAfterTable {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$Document,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlElement]$Body,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlElement]$Table,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $nextNode = $Table.NextSibling
  while ($null -ne $nextNode -and $nextNode.NodeType -ne [System.Xml.XmlNodeType]::Element) {
    $nextNode = $nextNode.NextSibling
  }

  if ($null -ne $nextNode -and $nextNode.LocalName -eq "p") {
    return
  }

  $wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  $emptyParagraph = $Document.CreateElement("w", "p", $wNs)
  if ($null -ne $nextNode) {
    [void]$Body.InsertBefore($emptyParagraph, $nextNode)
  } else {
    [void]$Body.AppendChild($emptyParagraph)
  }
}

function Remove-ChildNodesByXPath {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Parent,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager,

    [Parameter(Mandatory = $true)]
    [string]$XPath
  )

  foreach ($node in @($Parent.SelectNodes($XPath, $NamespaceManager))) {
    [void]$Parent.RemoveChild($node)
  }
}

function Get-ParagraphNodeText {
  param(
    [AllowNull()]
    [System.Xml.XmlNode]$Paragraph,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  if ($null -eq $Paragraph -or $Paragraph.LocalName -ne "p") {
    return ""
  }

  return ((@($Paragraph.SelectNodes(".//w:t", $NamespaceManager)) | ForEach-Object { $_.InnerText }) -join "").Trim()
}

function Test-IsCommandLikeText {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return $false
  }

  $trimmed = $Text.Trim()
  if ($trimmed.Length -gt 180) {
    return $false
  }

  return (
    $trimmed -match '^(?:PS\s+[^>]+>|[A-Za-z]:\\[^>]*>|>\s*)\s*\S+' -or
    $trimmed -match '(?i)^(?:ipconfig|ping|arp|tracert|netstat|nslookup|route|netsh|net\s+|cd\s+|dir\b|java\b|javac\b|gradle\b|adb\b|git\b|powershell\b|cmd\b|gcc\b|g\+\+\b|clang\b|clang\+\+\b|make\b|cmake\b|\.\/\S+)(?:\s|$)'
  )
}

function Test-IsCodeLikeText {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return $false
  }

  $trimmed = $Text.Trim()
  if ($trimmed.Length -gt 240 -or (Test-IsCommandLikeText -Text $trimmed)) {
    return $false
  }

  $matchesExplicitCodePattern = (
    $trimmed -match '^(?:#\s*(?:define|include)\b|if\s*\(|else(?:\s+if\b.*)?\s*\{?|while\s*\(|for\s*\(|switch\s*\(|return\b.*;|break;|continue;|\{|\}|}\s*else(?:\s+if\b.*)?\s*\{?)' -or
    $trimmed -match '^[A-Za-z_][A-Za-z0-9_\->\.\[\]]*\s*=\s*.+;$' -or
    $trimmed -match '^[A-Za-z_][A-Za-z0-9_\s\*\->\.\[\]\(\)]*\([^)]*\)\s*\{?$'
  )
  if ($matchesExplicitCodePattern) {
    return $true
  }

  $hasCjkCharacters = $false
  foreach ($character in $trimmed.ToCharArray()) {
    $codePoint = [int][char]$character
    if (
      (($codePoint -ge 0x3400) -and ($codePoint -le 0x4DBF)) -or
      (($codePoint -ge 0x4E00) -and ($codePoint -le 0x9FFF)) -or
      (($codePoint -ge 0xF900) -and ($codePoint -le 0xFAFF))
    ) {
      $hasCjkCharacters = $true
      break
    }
  }

  return (
    -not $hasCjkCharacters -and
    ($trimmed -match ';$' -or $trimmed -match '\->') -and
    $trimmed -match '[A-Za-z_]'
  )
}

function Get-NodeGroupMode {
  param(
    [AllowNull()]
    [System.Xml.XmlNode]$Node,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  if ($null -eq $Node -or $Node.LocalName -ne "p") {
    return "single"
  }

  $text = Get-ParagraphNodeText -Paragraph $Node -NamespaceManager $NamespaceManager
  if ((Test-IsCommandLikeText -Text $text) -or (Test-IsCodeLikeText -Text $text)) {
    return "code"
  }

  return "single"
}

function Test-IsDrawingParagraphNode {
  param(
    [AllowNull()]
    [System.Xml.XmlNode]$Node,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  return ($null -ne $Node -and $Node.LocalName -eq "p" -and $null -ne $Node.SelectSingleNode(".//w:drawing", $NamespaceManager))
}

function Test-IsFigureCaptionNode {
  param(
    [AllowNull()]
    [System.Xml.XmlNode]$Node,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $text = Get-ParagraphNodeText -Paragraph $Node -NamespaceManager $NamespaceManager
  return ($text -match '^图\s*\d+')
}

function Get-TemplateFrameBodyNodeGroups {
  param(
    [Parameter(Mandatory = $true)]
    [System.Collections.Generic.List[System.Xml.XmlNode]]$Nodes,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $groups = New-Object System.Collections.Generic.List[object]
  $index = 0
  while ($index -lt $Nodes.Count) {
    $group = New-Object System.Collections.Generic.List[System.Xml.XmlNode]
    $node = $Nodes[$index]
    [void]$group.Add($node)

    if (
      (Test-IsDrawingParagraphNode -Node $node -NamespaceManager $NamespaceManager) -and
      ($index + 1 -lt $Nodes.Count) -and
      (Test-IsFigureCaptionNode -Node $Nodes[$index + 1] -NamespaceManager $NamespaceManager)
    ) {
      $index++
      [void]$group.Add($Nodes[$index])
    }

    [void]$groups.Add($group.ToArray())
    $index++
  }

  return $groups
}

function Ensure-ClosedPageFrameBorder {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$Document,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager,

    [Parameter(Mandatory = $true)]
    [int]$TableWidthTwips
  )

  $wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  foreach ($sectionProperties in @($Document.SelectNodes("//w:sectPr", $NamespaceManager))) {
    $pageSize = $sectionProperties.SelectSingleNode("w:pgSz", $NamespaceManager)
    $pageMargins = $sectionProperties.SelectSingleNode("w:pgMar", $NamespaceManager)
    if ($null -ne $pageSize -and $null -ne $pageMargins -and $TableWidthTwips -gt 0) {
      $pageWidthText = $pageSize.GetAttribute("w", $wNs)
      $pageWidthTwips = 0
      if ([int]::TryParse($pageWidthText, [ref]$pageWidthTwips) -and $pageWidthTwips -gt $TableWidthTwips) {
        $sideMarginTwips = [int][Math]::Floor(($pageWidthTwips - $TableWidthTwips) / 2)
        [void](Set-WAttrValue -Element $pageMargins -Name "left" -Value ([string]$sideMarginTwips))
        [void](Set-WAttrValue -Element $pageMargins -Name "right" -Value ([string]$sideMarginTwips))
      }
    }

    foreach ($existingPageBorders in @($sectionProperties.SelectNodes("w:pgBorders", $NamespaceManager))) {
      [void]$sectionProperties.RemoveChild($existingPageBorders)
    }

    $pageBorders = $Document.CreateElement("w", "pgBorders", $wNs)
    [void]$pageBorders.Attributes.Append((New-WAttr -Document $Document -Name "offsetFrom" -Value "text"))
    [void]$pageBorders.Attributes.Append((New-WAttr -Document $Document -Name "display" -Value "notFirstPage"))

    foreach ($side in @("top", "left", "bottom", "right")) {
      $border = $Document.CreateElement("w", $side, $wNs)
      Add-BorderAttrs -Document $Document -Element $border
      [void]$pageBorders.AppendChild($border)
    }

    if ($sectionProperties.HasChildNodes) {
      [void]$sectionProperties.InsertBefore($pageBorders, $sectionProperties.FirstChild)
    } else {
      [void]$sectionProperties.AppendChild($pageBorders)
    }
  }
}

function Get-NextRelationshipId {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$RelationshipsXml
  )

  $maxRelationshipNumber = 0
  foreach ($relationship in @($RelationshipsXml.SelectNodes("//*[local-name()='Relationship']"))) {
    $id = [string]$relationship.GetAttribute("Id")
    if ($id -match '^rId(\d+)$') {
      $number = [int]$Matches[1]
      if ($number -gt $maxRelationshipNumber) {
        $maxRelationshipNumber = $number
      }
    }
  }

  return ("rId{0}" -f ($maxRelationshipNumber + 1))
}

function Write-XmlDocument {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$Document,

    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $settings = New-Object System.Xml.XmlWriterSettings
  $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
  $settings.Indent = $false
  $writer = [System.Xml.XmlWriter]::Create($Path, $settings)
  try {
    $Document.Save($writer)
  } finally {
    $writer.Close()
  }
}

function Ensure-FirstPageBottomFrameFooter {
  param(
    [Parameter(Mandatory = $true)]
    [string]$PackageDirectory,

    [Parameter(Mandatory = $true)]
    [xml]$Document,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlElement]$SectionProperties,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  $wordDir = Join-Path $PackageDirectory "word"
  $relsPath = Join-Path $wordDir "_rels\document.xml.rels"
  $contentTypesPath = Join-Path $PackageDirectory "[Content_Types].xml"
  $pageMargins = $SectionProperties.SelectSingleNode("w:pgMar", $NamespaceManager)
  $pageSize = $SectionProperties.SelectSingleNode("w:pgSz", $NamespaceManager)
  $frameLeftTwips = 913
  $frameWidthTwips = 10080
  $frameTopTwips = 2880
  $frameHeightTwips = 12270
  if ($null -ne $pageMargins) {
    $bottomMargin = $pageMargins.GetAttribute("bottom", $wNs)
    if (-not [string]::IsNullOrWhiteSpace($bottomMargin)) {
      # Align the first-page footer line with the text bottom so it closes the table frame.
      [void](Set-WAttrValue -Element $pageMargins -Name "footer" -Value $bottomMargin)
    }

    $leftMargin = $pageMargins.GetAttribute("left", $wNs)
    $rightMargin = $pageMargins.GetAttribute("right", $wNs)
    $pageWidthText = if ($null -ne $pageSize) { $pageSize.GetAttribute("w", $wNs) } else { "" }
    $pageHeightText = if ($null -ne $pageSize) { $pageSize.GetAttribute("h", $wNs) } else { "" }
    $pageWidthTwips = 0
    $pageHeightTwips = 0
    $leftMarginTwips = 0
    $rightMarginTwips = 0
    $bottomMarginTwips = 0
    if (
      [int]::TryParse($pageWidthText, [ref]$pageWidthTwips) -and
      [int]::TryParse($pageHeightText, [ref]$pageHeightTwips) -and
      [int]::TryParse($leftMargin, [ref]$leftMarginTwips) -and
      [int]::TryParse($rightMargin, [ref]$rightMarginTwips) -and
      [int]::TryParse($bottomMargin, [ref]$bottomMarginTwips)
    ) {
      $frameLeftTwips = $leftMarginTwips
      $calculatedWidth = $pageWidthTwips - $leftMarginTwips - $rightMarginTwips
      if ($calculatedWidth -gt 0) {
        $frameWidthTwips = $calculatedWidth
      }
      $frameBottomTwips = $pageHeightTwips - $bottomMarginTwips - 240
      if ($frameBottomTwips -gt $frameTopTwips) {
        $frameHeightTwips = $frameBottomTwips - $frameTopTwips
      }
    }
  }
  $pointCulture = [System.Globalization.CultureInfo]::InvariantCulture
  $frameLeftPoints = [string]::Format($pointCulture, "{0:0.###}pt", ($frameLeftTwips / 20.0))
  $frameTopPoints = [string]::Format($pointCulture, "{0:0.###}pt", ($frameTopTwips / 20.0))
  $frameWidthPoints = [string]::Format($pointCulture, "{0:0.###}pt", ($frameWidthTwips / 20.0))
  $frameHeightPoints = [string]::Format($pointCulture, "{0:0.###}pt", ($frameHeightTwips / 20.0))
  $frameRightPoints = [string]::Format($pointCulture, "{0:0.###}pt", (($frameLeftTwips + $frameWidthTwips) / 20.0))
  $frameBottomPoints = [string]::Format($pointCulture, "{0:0.###}pt", (($frameTopTwips + $frameHeightTwips) / 20.0))

  $footerIndex = 1
  do {
    $footerFileName = "footer$footerIndex.xml"
    $footerPath = Join-Path $wordDir $footerFileName
    $footerIndex++
  } while (Test-Path -LiteralPath $footerPath)

  $footerXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
  <w:p>
    <w:pPr>
      <w:spacing w:before="0" w:after="0" />
    </w:pPr>
    <w:r>
      <w:pict>
        <v:line id="FirstPageFrameLeft" o:allowincell="f" strokecolor="#000000" strokeweight="0.5pt" style="position:absolute;margin-left:$frameLeftPoints;margin-top:$frameTopPoints;width:0;height:$frameHeightPoints;z-index:251654144;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;mso-wrap-style:none" from="0,0" to="0,$frameHeightPoints" />
        <v:line id="FirstPageFrameRight" o:allowincell="f" strokecolor="#000000" strokeweight="0.5pt" style="position:absolute;margin-left:$frameRightPoints;margin-top:$frameTopPoints;width:0;height:$frameHeightPoints;z-index:251654145;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;mso-wrap-style:none" from="0,0" to="0,$frameHeightPoints" />
        <v:line id="FirstPageFrameBottom" o:allowincell="f" strokecolor="#000000" strokeweight="0.5pt" style="position:absolute;margin-left:$frameLeftPoints;margin-top:$frameBottomPoints;width:$frameWidthPoints;height:0;z-index:251654146;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;mso-wrap-style:none" from="0,0" to="$frameWidthPoints,0" />
      </w:pict>
    </w:r>
  </w:p>
  <w:p>
    <w:pPr>
      <w:pBdr>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="auto" />
      </w:pBdr>
      <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto" />
    </w:pPr>
  </w:p>
</w:ftr>
"@
  [System.IO.File]::WriteAllText($footerPath, $footerXml, (New-Object System.Text.UTF8Encoding($false)))

  [xml]$relationshipsXml = Get-Content -LiteralPath $relsPath -Raw -Encoding UTF8
  $relsNs = "http://schemas.openxmlformats.org/package/2006/relationships"
  $relationshipId = Get-NextRelationshipId -RelationshipsXml $relationshipsXml
  $relationship = $relationshipsXml.CreateElement("Relationship", $relsNs)
  $relationship.SetAttribute("Id", $relationshipId)
  $relationship.SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer")
  $relationship.SetAttribute("Target", $footerFileName)
  [void]$relationshipsXml.DocumentElement.AppendChild($relationship)
  Write-XmlDocument -Document $relationshipsXml -Path $relsPath

  [xml]$contentTypesXml = Get-Content -LiteralPath $contentTypesPath -Raw -Encoding UTF8
  $ctNs = "http://schemas.openxmlformats.org/package/2006/content-types"
  $overridePartName = "/word/$footerFileName"
  $existingOverride = @($contentTypesXml.SelectNodes("//*[local-name()='Override']") | Where-Object { $_.GetAttribute("PartName") -eq $overridePartName })
  if ($existingOverride.Count -eq 0) {
    $override = $contentTypesXml.CreateElement("Override", $ctNs)
    $override.SetAttribute("PartName", $overridePartName)
    $override.SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")
    [void]$contentTypesXml.DocumentElement.AppendChild($override)
    Write-XmlDocument -Document $contentTypesXml -Path $contentTypesPath
  }

  foreach ($existingFirstFooter in @($SectionProperties.SelectNodes("w:footerReference[@w:type='first']", $NamespaceManager))) {
    [void]$SectionProperties.RemoveChild($existingFirstFooter)
  }

  $footerReference = $Document.CreateElement("w", "footerReference", $wNs)
  [void]$footerReference.Attributes.Append((New-WAttr -Document $Document -Name "type" -Value "first"))
  [void]$footerReference.Attributes.Append((New-RAttr -Document $Document -Name "id" -Value $relationshipId))
  [void]$SectionProperties.InsertBefore($footerReference, $SectionProperties.FirstChild)

  if ($null -eq $SectionProperties.SelectSingleNode("w:titlePg", $NamespaceManager)) {
    $titlePage = $Document.CreateElement("w", "titlePg", $wNs)
    [void]$SectionProperties.InsertAfter($titlePage, $footerReference)
  }
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

  Set-MetadataRowsToTemplateStandard -Document $documentXml -Table $mainTable -NamespaceManager $namespaceManager -MetadataCellMarginTwips $MetadataCellMarginTwips
  $columnSpan = Get-TableColumnSpan -Table $mainTable -NamespaceManager $namespaceManager
  $mainTableWidthTwips = Get-TableGridWidth -Table $mainTable -NamespaceManager $namespaceManager

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
    $nodeGroups = @(Get-TemplateFrameBodyNodeGroups -Nodes $nodesToMove -NamespaceManager $namespaceManager)
    for ($groupIndex = 0; $groupIndex -lt $nodeGroups.Count; $groupIndex++) {
      $tableRow = $documentXml.CreateElement("w", "tr", $wNs)
      $tableCell = $documentXml.CreateElement("w", "tc", $wNs)
      Set-CellWidthAndSpan -Document $documentXml -Cell $tableCell -NamespaceManager $namespaceManager -WidthValue ([string]$mainTableWidthTwips) -WidthType "dxa" -GridSpan $columnSpan

      $topMargin = if ($groupIndex -eq 0) { $BodyCellMarginTwips } else { 0 }
      $bottomMargin = if ($groupIndex -eq ($nodeGroups.Count - 1)) { $BodyCellMarginTwips } else { 0 }
      Set-CellMargins -Document $documentXml -Cell $tableCell -MarginTwips $BodyCellMarginTwips -TopTwips $topMargin -BottomTwips $bottomMargin

      $topBorder = if ($groupIndex -eq 0) { "single" } else { "nil" }
      $bottomBorder = "nil"
      Set-CellHorizontalBorderValues -Document $documentXml -Cell $tableCell -NamespaceManager $namespaceManager -Top $topBorder -Bottom $bottomBorder

      foreach ($node in @($nodeGroups[$groupIndex])) {
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
  }

  Ensure-ParagraphAfterTable -Document $documentXml -Body $body -Table $mainTable -NamespaceManager $namespaceManager
  Ensure-ClosedPageFrameBorder -Document $documentXml -NamespaceManager $namespaceManager -TableWidthTwips $mainTableWidthTwips
  $firstSectionProperties = $documentXml.SelectSingleNode("//w:sectPr[1]", $namespaceManager)
  if ($null -ne $firstSectionProperties) {
    Ensure-FirstPageBottomFrameFooter -PackageDirectory $unzipDir -Document $documentXml -SectionProperties $firstSectionProperties -NamespaceManager $namespaceManager
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
