[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$TemplatePath,

  [string]$MappingPath,

  [string]$FieldsJson,

  [string]$OutPath,

  [switch]$Overwrite
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$fullWidthColon = [string][char]0xFF1A
$labelPattern = '^(?<label>[^:\uFF1A]{1,60})[:\uFF1A]\s*(?<rest>.*)$'
$sectionHeadingPattern = '^(?:(?:\d+|[一二三四五六七八九十])[.\u3001]?\s*)?(?:\u5B9E\u9A8C\u76EE\u7684|\u5B9E\u9A8C\u73AF\u5883|\u5B9E\u9A8C\u8BBE\u5907\u4E0E\u73AF\u5883|\u5B9E\u9A8C\u539F\u7406|\u5B9E\u9A8C\u539F\u7406\u6216\u4EFB\u52A1\u8981\u6C42|\u4EFB\u52A1\u8981\u6C42|\u5B9E\u9A8C\u6B65\u9AA4|\u5B9E\u9A8C\u7ED3\u679C|\u5B9E\u9A8C\u73B0\u8C61\u4E0E\u7ED3\u679C\u8BB0\u5F55|\u7ED3\u679C\u5206\u6790|\u95EE\u9898\u5206\u6790|\u5B9E\u9A8C\u603B\u7ED3|\u603B\u7ED3\u4E0E\u601D\u8003|\u5B9E\u9A8C\u5185\u5BB9|\u5B9E\u9A8C\u8FC7\u7A0B|\u5B9E\u9A8C\u7ED3\u8BBA|\u5B9E\u9A8C\u8981\u6C42|\u5B9E\u9A8C\u5668\u6750|\u5B9E\u9A8C\u4EEA\u5668|\u5B9E\u9A8C\u8BB0\u5F55|\u6CE8\u610F\u4E8B\u9879|\u5B9E\u9A8C\u5C0F\u7ED3)$'
$choiceMarkerPattern = '[①②③④⑤⑥⑦⑧⑨⑩]'

function Normalize-OpenXmlText {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return ""
  }

  return (($Text -replace "\s+", " ").Trim())
}

function Get-NodeText {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Node,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $parts = New-Object System.Collections.Generic.List[string]
  $textNodes = $Node.SelectNodes(".//w:t | .//w:instrText", $NamespaceManager)
  foreach ($textNode in $textNodes) {
    [void]$parts.Add($textNode.InnerText)
  }

  return Normalize-OpenXmlText -Text ($parts -join "")
}

function Normalize-FieldKey {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return ""
  }

  $normalized = $Text.ToLowerInvariant()
  $normalized = $normalized -replace '\s+', ''
  $normalized = $normalized -replace '[\p{P}\p{S}\uFF3F]+', ''
  return $normalized
}

function Is-PlaceholderLike {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return $true
  }

  return [bool]($Text -match '[_\uFF3F]{2,}|\.{3,}|\uFF08\s*\uFF09|\(\s*\)|\u25A1|\u25A0')
}

function Get-OptionFieldInfo {
  param(
    [AllowNull()]
    [string]$Text
  )

  $normalizedText = Normalize-OpenXmlText -Text $Text
  if ([string]::IsNullOrWhiteSpace($normalizedText) -or ($normalizedText -notmatch $labelPattern)) {
    return $null
  }

  $label = $matches["label"].Trim()
  $rest = $matches["rest"]
  if ([string]::IsNullOrWhiteSpace($rest)) {
    return $null
  }

  $trimmedRest = (($rest -replace '[:\uFF1A]\s*$', '').Trim())
  if ([string]::IsNullOrWhiteSpace($trimmedRest) -or ($trimmedRest -notmatch $choiceMarkerPattern)) {
    return $null
  }

  $choiceMatches = [regex]::Matches($trimmedRest, "(?<choice>$choiceMarkerPattern\s*.*?)(?=(?:\s*$choiceMarkerPattern)|$)")
  if ($choiceMatches.Count -lt 2) {
    return $null
  }

  $choices = New-Object System.Collections.Generic.List[string]
  foreach ($choiceMatch in $choiceMatches) {
    $choiceText = Normalize-OpenXmlText -Text ($choiceMatch.Groups["choice"].Value -replace '^[√✔☑]\s*', '')
    if (-not [string]::IsNullOrWhiteSpace($choiceText)) {
      [void]$choices.Add($choiceText)
    }
  }

  if ($choices.Count -lt 2) {
    return $null
  }

  return [pscustomobject]@{
    Label = $label
    Choices = @($choices)
    HasTrailingColon = [bool]($normalizedText -match '[:\uFF1A]\s*$')
  }
}

function Resolve-OptionFieldText {
  param(
    [AllowNull()]
    [string]$TemplateText,

    [AllowNull()]
    [string]$SelectionText
  )

  $optionField = Get-OptionFieldInfo -Text $TemplateText
  if ($null -eq $optionField -or [string]::IsNullOrWhiteSpace($SelectionText)) {
    return $null
  }

  $selectionCandidates = New-Object System.Collections.Generic.List[string]
  $selectionCandidates.Add((Normalize-OpenXmlText -Text $SelectionText)) | Out-Null

  $checkedChoiceMatch = [regex]::Match($SelectionText, "[√✔☑]\s*(?<choice>$choiceMarkerPattern\s*.*?)(?=(?:\s*$choiceMarkerPattern)|$)")
  if ($checkedChoiceMatch.Success) {
    $selectionCandidates.Add((Normalize-OpenXmlText -Text $checkedChoiceMatch.Groups["choice"].Value)) | Out-Null
  }

  if ($SelectionText -match $labelPattern) {
    $selectionCandidates.Add((Normalize-OpenXmlText -Text (($matches["rest"] -replace '[:\uFF1A]\s*$', '').Trim()))) | Out-Null
  }

  $normalizedCandidates = @(
    $selectionCandidates |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
    ForEach-Object { Normalize-FieldKey -Text $_ } |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
    Select-Object -Unique
  )

  if ($normalizedCandidates.Count -eq 0) {
    return $null
  }

  $selectedIndex = -1
  for ($choiceIndex = 0; $choiceIndex -lt $optionField.Choices.Count; $choiceIndex++) {
    $normalizedChoice = Normalize-FieldKey -Text ([string]$optionField.Choices[$choiceIndex])
    foreach ($normalizedCandidate in $normalizedCandidates) {
      if (
        ($normalizedCandidate -eq $normalizedChoice) -or
        $normalizedCandidate.EndsWith($normalizedChoice) -or
        $normalizedChoice.EndsWith($normalizedCandidate)
      ) {
        $selectedIndex = $choiceIndex
        break
      }
    }

    if ($selectedIndex -ge 0) {
      break
    }
  }

  if ($selectedIndex -lt 0) {
    return $null
  }

  $renderedChoices = for ($choiceIndex = 0; $choiceIndex -lt $optionField.Choices.Count; $choiceIndex++) {
    $choiceText = Normalize-OpenXmlText -Text ([string]$optionField.Choices[$choiceIndex])
    if ($choiceIndex -eq $selectedIndex) {
      "√$choiceText"
    } else {
      $choiceText
    }
  }

  $line = ("{0}{1} {2}" -f $optionField.Label, $fullWidthColon, ($renderedChoices -join "  "))
  if ($optionField.HasTrailingColon) {
    $line += $fullWidthColon
  }

  return $line
}

function Is-SectionHeadingLike {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return $false
  }

  return [bool]($Text -match $sectionHeadingPattern)
}

function Test-ParagraphHasBoldFormatting {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Paragraph,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  foreach ($boldNode in @($Paragraph.SelectNodes("./w:pPr/w:rPr/w:b | ./w:r/w:rPr/w:b", $NamespaceManager))) {
    if ($null -eq $boldNode) {
      continue
    }

    $valueAttribute = $boldNode.Attributes.GetNamedItem("w:val")
    if ($null -eq $valueAttribute -or [string]::IsNullOrWhiteSpace($valueAttribute.Value)) {
      return $true
    }

    if ($valueAttribute.Value -notin @("0", "false", "False")) {
      return $true
    }
  }

  return $false
}

function ConvertTo-PlainHashtable {
  param(
    [Parameter(Mandatory = $true)]
    [object]$InputObject
  )

  $table = @{}
  if ($InputObject -is [System.Collections.IDictionary]) {
    foreach ($key in $InputObject.Keys) {
      $table[[string]$key] = $InputObject[$key]
    }
    return $table
  }

  foreach ($property in $InputObject.PSObject.Properties) {
    $table[[string]$property.Name] = $property.Value
  }

  return $table
}

function Add-ParagraphTexts {
  param(
    [Parameter(Mandatory = $true)]
    [object]$Target,

    [AllowNull()]
    [object]$Value
  )

  if ($null -eq $Value) {
    return
  }

  if (($Value -is [System.Collections.IEnumerable]) -and ($Value -isnot [string])) {
    foreach ($item in $Value) {
      Add-ParagraphTexts -Target $Target -Value $item
    }
    return
  }

  $text = [string]$Value
  if ($text -match "\r?\n") {
    foreach ($line in ($text -split "\r?\n")) {
      $normalizedLine = Normalize-OpenXmlText -Text $line
      if (-not [string]::IsNullOrWhiteSpace($normalizedLine)) {
        [void]$Target.Add($normalizedLine)
      }
    }
    return
  }

  $normalizedText = Normalize-OpenXmlText -Text $text
  if (-not [string]::IsNullOrWhiteSpace($normalizedText)) {
    [void]$Target.Add($normalizedText)
  }
}

function Get-FillSpec {
  param(
    [AllowNull()]
    [object]$Value
  )

  $mode = "auto"
  $throughLocation = $null
  $paragraphs = New-Object System.Collections.Generic.List[string]

  if (($Value -is [System.Collections.IDictionary]) -or (($Value -isnot [string]) -and $null -ne $Value -and $null -ne $Value.PSObject -and @($Value.PSObject.Properties).Count -gt 0)) {
    $valueTable = ConvertTo-PlainHashtable -InputObject $Value
    if ($valueTable.ContainsKey("mode") -and -not [string]::IsNullOrWhiteSpace([string]$valueTable["mode"])) {
      $mode = ([string]$valueTable["mode"]).ToLowerInvariant()
    }

    if ($valueTable.ContainsKey("through") -and -not [string]::IsNullOrWhiteSpace([string]$valueTable["through"])) {
      $throughLocation = ([string]$valueTable["through"]).Trim().ToUpperInvariant()
    } elseif ($valueTable.ContainsKey("throughLocation") -and -not [string]::IsNullOrWhiteSpace([string]$valueTable["throughLocation"])) {
      $throughLocation = ([string]$valueTable["throughLocation"]).Trim().ToUpperInvariant()
    }

    if ($valueTable.ContainsKey("paragraphs")) {
      Add-ParagraphTexts -Target $paragraphs -Value $valueTable["paragraphs"]
    } elseif ($valueTable.ContainsKey("text")) {
      Add-ParagraphTexts -Target $paragraphs -Value $valueTable["text"]
    } elseif ($valueTable.ContainsKey("value")) {
      Add-ParagraphTexts -Target $paragraphs -Value $valueTable["value"]
    } else {
      Add-ParagraphTexts -Target $paragraphs -Value $Value
    }
  } else {
    Add-ParagraphTexts -Target $paragraphs -Value $Value
  }

  if ($paragraphs.Count -eq 0) {
    [void]$paragraphs.Add("")
  }

  return [pscustomobject]@{
    Mode = $mode
    Paragraphs = $paragraphs.ToArray()
    ThroughLocation = $throughLocation
  }
}

function Get-InlineText {
  param(
    [Parameter(Mandatory = $true)]
    [object]$FillSpec
  )

  return (($FillSpec.Paragraphs | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join " ")
}

function Parse-TableCellLocation {
  param(
    [AllowNull()]
    [string]$Location
  )

  if ([string]::IsNullOrWhiteSpace($Location) -or ($Location -notmatch '^(?i)T(?<table>\d+)R(?<row>\d+)C(?<cell>\d+)$')) {
    return $null
  }

  return [pscustomobject]@{
    Location = $Location.Trim().ToUpperInvariant()
    TableIndex = [int]$matches["table"]
    RowIndex = [int]$matches["row"]
    CellIndex = [int]$matches["cell"]
  }
}

function Get-NextParagraphSibling {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Node
  )

  $current = $Node.NextSibling
  while ($null -ne $current) {
    if ($current.LocalName -eq "p") {
      return $current
    }

    if ($current.LocalName -eq "tbl") {
      return $null
    }

    $current = $current.NextSibling
  }

  return $null
}

function Get-PreviousParagraphSibling {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Node
  )

  $current = $Node.PreviousSibling
  while ($null -ne $current) {
    if ($current.LocalName -eq "p") {
      return $current
    }

    if ($current.LocalName -eq "tbl") {
      return $null
    }

    $current = $current.PreviousSibling
  }

  return $null
}

function Get-TableInsertReferenceNode {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Table
  )

  $current = $Table.NextSibling
  while ($null -ne $current) {
    if ($current.LocalName -in @("p", "tbl", "sectPr")) {
      return $current
    }

    $current = $current.NextSibling
  }

  return $null
}

function Should-UseAfterHeadingFill {
  param(
    [AllowNull()]
    [string]$HeadingText,

    [Parameter(Mandatory = $true)]
    [object]$FillSpec,

    [AllowNull()]
    [System.Xml.XmlNode]$NextParagraph,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  if ($FillSpec.Mode -in @("inline", "replace")) {
    return $false
  }

  if ($null -eq $NextParagraph) {
    return $false
  }

  $nextText = Get-NodeText -Node $NextParagraph -NamespaceManager $NamespaceManager
  if (-not ([string]::IsNullOrWhiteSpace($nextText) -or (Is-PlaceholderLike -Text $nextText))) {
    return $false
  }

  if ($FillSpec.Mode -eq "after") {
    return $true
  }

  return (Is-SectionHeadingLike -Text $HeadingText)
}

function Get-MappingTables {
  param(
    [AllowNull()]
    [string]$PathToJson,

    [AllowNull()]
    [string]$InlineJson
  )

  if ([string]::IsNullOrWhiteSpace($PathToJson) -eq [string]::IsNullOrWhiteSpace($InlineJson)) {
    throw "Provide exactly one of -MappingPath or -FieldsJson."
  }

  if (-not [string]::IsNullOrWhiteSpace($PathToJson)) {
    $resolvedPath = Resolve-Path -LiteralPath $PathToJson
    $rootObject = (Get-Content -LiteralPath $resolvedPath.Path -Raw -Encoding UTF8) | ConvertFrom-Json
  } else {
    $rootObject = $InlineJson | ConvertFrom-Json
  }

  if ($null -eq $rootObject) {
    throw "Field mapping JSON is empty."
  }

  $mappingRoot = ConvertTo-PlainHashtable -InputObject $rootObject
  if ($mappingRoot.ContainsKey("fields")) {
    $mappingRoot = ConvertTo-PlainHashtable -InputObject $mappingRoot["fields"]
  } elseif ($mappingRoot.ContainsKey("fieldMap")) {
    $mappingRoot = ConvertTo-PlainHashtable -InputObject $mappingRoot["fieldMap"]
  }

  $locationMap = @{}
  $labelMap = @{}

  foreach ($entry in $mappingRoot.GetEnumerator()) {
    if ($null -eq $entry.Value) {
      continue
    }

    $value = $entry.Value
    $key = [string]$entry.Key

    if ($key -match '^(?i:P\d+|T\d+R\d+C\d+)$') {
      $locationMap[$key.ToUpperInvariant()] = $value
      continue
    }

    $normalizedKey = Normalize-FieldKey -Text $key
    if (-not [string]::IsNullOrWhiteSpace($normalizedKey)) {
      $labelMap[$normalizedKey] = $value
    }
  }

  return @{
    LocationMap = $locationMap
    LabelMap = $labelMap
  }
}

function Ensure-TextNode {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Node,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $document = $Node.OwnerDocument
  $namespaceUri = $NamespaceManager.LookupNamespace("w")

  if ($Node.LocalName -eq "p") {
    $paragraph = $Node
  } elseif ($Node.LocalName -eq "tc") {
    $paragraph = $Node.SelectSingleNode("./w:p[1]", $NamespaceManager)
    if ($null -eq $paragraph) {
      $paragraph = $document.CreateElement("w", "p", $namespaceUri)
      $Node.AppendChild($paragraph) | Out-Null
    }
  } else {
    $paragraph = $Node.SelectSingleNode(".//w:p[1]", $NamespaceManager)
    if ($null -eq $paragraph) {
      throw "Unsupported container without paragraph text nodes: $($Node.LocalName)"
    }
  }

  $textNode = $paragraph.SelectSingleNode(".//w:t[1]", $NamespaceManager)
  if ($null -ne $textNode) {
    return $textNode
  }

  $run = $paragraph.SelectSingleNode("./w:r[1]", $NamespaceManager)
  if ($null -eq $run) {
    $run = $document.CreateElement("w", "r", $namespaceUri)
    $paragraph.AppendChild($run) | Out-Null
  }

  $textNode = $document.CreateElement("w", "t", $namespaceUri)
  $run.AppendChild($textNode) | Out-Null
  return $textNode
}

function Set-NodeText {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Node,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager,

    [Parameter(Mandatory = $true)]
    [string]$Value
  )

  $textNodes = $Node.SelectNodes(".//w:t", $NamespaceManager)
  if ($null -eq $textNodes -or $textNodes.Count -eq 0) {
    $textNodes = @(Ensure-TextNode -Node $Node -NamespaceManager $NamespaceManager)
  }

  $textNodes[0].InnerText = $Value
  for ($i = 1; $i -lt $textNodes.Count; $i++) {
    $textNodes[$i].InnerText = ""
  }
}

function Get-RunPropertiesClone {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Paragraph,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $runProperties = $Paragraph.SelectSingleNode("./w:r/w:rPr[1]", $NamespaceManager)
  if ($null -eq $runProperties) {
    return $null
  }

  return $runProperties.CloneNode($true)
}

function Clear-ParagraphContent {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Paragraph
  )

  foreach ($child in @($Paragraph.ChildNodes)) {
    if ($child.LocalName -ne "pPr") {
      $Paragraph.RemoveChild($child) | Out-Null
    }
  }
}

function Add-TextRunToParagraph {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Paragraph,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager,

    [Parameter(Mandatory = $true)]
    [string]$Text,

    [AllowNull()]
    [System.Xml.XmlNode]$RunPropertiesTemplate
  )

  $document = $Paragraph.OwnerDocument
  $namespaceUri = $NamespaceManager.LookupNamespace("w")
  $run = $document.CreateElement("w", "r", $namespaceUri)

  if ($null -ne $RunPropertiesTemplate) {
    $run.AppendChild($RunPropertiesTemplate.CloneNode($true)) | Out-Null
  }

  $textNode = $document.CreateElement("w", "t", $namespaceUri)
  $textNode.InnerText = $Text
  $run.AppendChild($textNode) | Out-Null
  $Paragraph.AppendChild($run) | Out-Null
}

function Set-ParagraphText {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Paragraph,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager,

    [Parameter(Mandatory = $true)]
    [string]$Text
  )

  $runPropertiesTemplate = Get-RunPropertiesClone -Paragraph $Paragraph -NamespaceManager $NamespaceManager
  Clear-ParagraphContent -Paragraph $Paragraph
  Add-TextRunToParagraph -Paragraph $Paragraph -NamespaceManager $NamespaceManager -Text $Text -RunPropertiesTemplate $runPropertiesTemplate
}

function New-ParagraphLike {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$TemplateParagraph,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager,

    [Parameter(Mandatory = $true)]
    [string]$Text
  )

  $paragraph = $TemplateParagraph.CloneNode($true)
  Set-ParagraphText -Paragraph $paragraph -NamespaceManager $NamespaceManager -Text $Text
  return $paragraph
}

function Get-CellParagraphTemplateSet {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Cell,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $paragraphNodes = @($Cell.SelectNodes("./w:p", $NamespaceManager))
  if ($paragraphNodes.Count -eq 0) {
    $paragraphNodes = @()
  }

  $fallbackTemplate = $null
  $headingTemplate = $null
  $bodyTemplate = $null

  foreach ($paragraphNode in $paragraphNodes) {
    if ($null -eq $fallbackTemplate) {
      $fallbackTemplate = $paragraphNode.CloneNode($true)
    }

    $paragraphText = Get-NodeText -Node $paragraphNode -NamespaceManager $NamespaceManager
    $isHeading = Is-SectionHeadingLike -Text $paragraphText
    $isBlank = [string]::IsNullOrWhiteSpace($paragraphText)
    $hasBoldFormatting = Test-ParagraphHasBoldFormatting -Paragraph $paragraphNode -NamespaceManager $NamespaceManager

    if ($null -eq $headingTemplate -and ($isHeading -or ($hasBoldFormatting -and $paragraphText.Length -le 40))) {
      $headingTemplate = $paragraphNode.CloneNode($true)
    }

    if ($null -eq $bodyTemplate -and (($isBlank -and -not $hasBoldFormatting) -or (-not $isHeading -and -not $hasBoldFormatting))) {
      $bodyTemplate = $paragraphNode.CloneNode($true)
    }
  }

  if ($null -eq $fallbackTemplate) {
    $namespaceUri = $NamespaceManager.LookupNamespace("w")
    $fallbackTemplate = $Cell.OwnerDocument.CreateElement("w", "p", $namespaceUri)
  }

  if ($null -eq $headingTemplate) {
    $headingTemplate = $fallbackTemplate.CloneNode($true)
  }

  if ($null -eq $bodyTemplate) {
    foreach ($paragraphNode in $paragraphNodes) {
      $paragraphText = Get-NodeText -Node $paragraphNode -NamespaceManager $NamespaceManager
      if (-not (Is-SectionHeadingLike -Text $paragraphText)) {
        $bodyTemplate = $paragraphNode.CloneNode($true)
        break
      }
    }
  }

  if ($null -eq $bodyTemplate) {
    $bodyTemplate = $fallbackTemplate.CloneNode($true)
  }

  return [pscustomobject]@{
    FallbackTemplate = $fallbackTemplate
    HeadingTemplate = $headingTemplate
    BodyTemplate = $bodyTemplate
  }
}

function Apply-ParagraphBlock {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$TargetParagraph,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager,

    [Parameter(Mandatory = $true)]
    [string[]]$Paragraphs
  )

  Set-ParagraphText -Paragraph $TargetParagraph -NamespaceManager $NamespaceManager -Text $Paragraphs[0]
  $templateParagraph = $TargetParagraph.CloneNode($true)
  $insertAfter = $TargetParagraph
  $insertedCount = 0

  for ($i = 1; $i -lt $Paragraphs.Count; $i++) {
    $newParagraph = New-ParagraphLike -TemplateParagraph $templateParagraph -NamespaceManager $NamespaceManager -Text $Paragraphs[$i]
    $insertAfter.ParentNode.InsertAfter($newParagraph, $insertAfter) | Out-Null
    $insertAfter = $newParagraph
    $insertedCount++
  }

  return $insertedCount
}

function Get-BlockParagraphTemplate {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$AnchorNode,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $nextParagraph = Get-NextParagraphSibling -Node $AnchorNode
  if ($null -ne $nextParagraph) {
    return $nextParagraph.CloneNode($true)
  }

  $previousParagraph = Get-PreviousParagraphSibling -Node $AnchorNode
  if ($null -ne $previousParagraph) {
    return $previousParagraph.CloneNode($true)
  }

  $tableParagraph = $AnchorNode.SelectSingleNode(".//w:p[1]", $NamespaceManager)
  if ($null -ne $tableParagraph) {
    return $tableParagraph.CloneNode($true)
  }

  $namespaceUri = $NamespaceManager.LookupNamespace("w")
  return $AnchorNode.OwnerDocument.CreateElement("w", "p", $namespaceUri)
}

function Insert-ParagraphBlockBeforeNode {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$ParentNode,

    [AllowNull()]
    [System.Xml.XmlNode]$ReferenceNode,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$TemplateParagraph,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager,

    [Parameter(Mandatory = $true)]
    [string[]]$Paragraphs
  )

  $insertAfter = $null
  $insertedCount = 0

  foreach ($paragraphText in $Paragraphs) {
    $newParagraph = New-ParagraphLike -TemplateParagraph $TemplateParagraph -NamespaceManager $NamespaceManager -Text $paragraphText
    if ($null -eq $insertAfter) {
      if ($null -ne $ReferenceNode) {
        $ParentNode.InsertBefore($newParagraph, $ReferenceNode) | Out-Null
      } else {
        $ParentNode.AppendChild($newParagraph) | Out-Null
      }
    } else {
      $ParentNode.InsertAfter($newParagraph, $insertAfter) | Out-Null
    }

    $insertAfter = $newParagraph
    $insertedCount++
  }

  return $insertedCount
}

function Move-TableBodyRowsAfterTable {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Table,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager,

    [Parameter(Mandatory = $true)]
    [string]$StartLocation,

    [AllowNull()]
    [string]$ThroughLocation,

    [Parameter(Mandatory = $true)]
    [string[]]$Paragraphs
  )

  $startInfo = Parse-TableCellLocation -Location $StartLocation
  if ($null -eq $startInfo) {
    throw "Invalid after-table start location: $StartLocation"
  }

  $throughInfo = if ([string]::IsNullOrWhiteSpace($ThroughLocation)) {
    $startInfo
  } else {
    Parse-TableCellLocation -Location $ThroughLocation
  }

  if ($null -eq $throughInfo) {
    throw "Invalid after-table end location: $ThroughLocation"
  }

  if ($startInfo.TableIndex -ne $throughInfo.TableIndex) {
    throw "after-table fill must stay within one table: $StartLocation -> $ThroughLocation"
  }

  if ($throughInfo.RowIndex -lt $startInfo.RowIndex) {
    throw "after-table fill end row must not be above the start row: $StartLocation -> $ThroughLocation"
  }

  $rowNodes = @($Table.SelectNodes("./w:tr", $NamespaceManager))
  if ($startInfo.RowIndex -gt $rowNodes.Count -or $throughInfo.RowIndex -gt $rowNodes.Count) {
    throw "after-table fill row range is outside the table: $StartLocation -> $ThroughLocation"
  }

  if (($rowNodes.Count - ($throughInfo.RowIndex - $startInfo.RowIndex + 1)) -le 0) {
    throw "after-table fill would remove every row from the table: $StartLocation -> $ThroughLocation"
  }

  $insertReference = Get-TableInsertReferenceNode -Table $Table
  $nextParagraph = if (($null -ne $insertReference) -and ($insertReference.LocalName -eq "p")) { $insertReference } else { $null }
  $reuseNextParagraph = $false
  if ($null -ne $nextParagraph) {
    $nextParagraphText = Get-NodeText -Node $nextParagraph -NamespaceManager $NamespaceManager
    $reuseNextParagraph = ([string]::IsNullOrWhiteSpace($nextParagraphText) -or (Is-PlaceholderLike -Text $nextParagraphText))
  }

  $templateParagraph = Get-BlockParagraphTemplate -AnchorNode $Table -NamespaceManager $NamespaceManager

  $removedRowCount = 0
  for ($rowNumber = $throughInfo.RowIndex; $rowNumber -ge $startInfo.RowIndex; $rowNumber--) {
    $rowNode = $rowNodes[$rowNumber - 1]
    if (($null -ne $rowNode) -and ($rowNode.ParentNode -eq $Table)) {
      $Table.RemoveChild($rowNode) | Out-Null
      $removedRowCount++
    }
  }

  $insertedParagraphCount = if ($reuseNextParagraph) {
    Apply-ParagraphBlock -TargetParagraph $nextParagraph -NamespaceManager $NamespaceManager -Paragraphs $Paragraphs
  } else {
    Insert-ParagraphBlockBeforeNode -ParentNode $Table.ParentNode -ReferenceNode $insertReference -TemplateParagraph $templateParagraph -NamespaceManager $NamespaceManager -Paragraphs $Paragraphs
  }

  return [pscustomobject]@{
    InsertedParagraphCount = $insertedParagraphCount
    RemovedRowCount = $removedRowCount
  }
}

function Set-CellParagraphs {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Cell,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager,

    [Parameter(Mandatory = $true)]
    [string[]]$Paragraphs
  )

  $templateSet = Get-CellParagraphTemplateSet -Cell $Cell -NamespaceManager $NamespaceManager

  foreach ($child in @($Cell.ChildNodes)) {
    if ($child.LocalName -ne "tcPr") {
      $Cell.RemoveChild($child) | Out-Null
    }
  }

  foreach ($paragraphText in $Paragraphs) {
    $templateParagraph = if (($Paragraphs.Count -gt 1) -and (Is-SectionHeadingLike -Text $paragraphText)) {
      $templateSet.HeadingTemplate
    } elseif (($Paragraphs.Count -gt 1) -and -not [string]::IsNullOrWhiteSpace($paragraphText)) {
      $templateSet.BodyTemplate
    } else {
      $templateSet.FallbackTemplate
    }

    $newParagraph = New-ParagraphLike -TemplateParagraph $templateParagraph -NamespaceManager $NamespaceManager -Text $paragraphText
    $Cell.AppendChild($newParagraph) | Out-Null
  }
}

function Write-OpenXmlPackage {
  param(
    [Parameter(Mandatory = $true)]
    [string]$SourceDirectory,

    [Parameter(Mandatory = $true)]
    [string]$DestinationPath
  )

  $archive = [System.IO.Compression.ZipFile]::Open($DestinationPath, [System.IO.Compression.ZipArchiveMode]::Create)
  try {
    foreach ($file in Get-ChildItem -LiteralPath $SourceDirectory -Recurse -File) {
      $relativePath = $file.FullName.Substring($SourceDirectory.Length).TrimStart('\', '/')
      $entryName = $relativePath -replace '\\', '/'
      $entry = $archive.CreateEntry($entryName, [System.IO.Compression.CompressionLevel]::Optimal)
      $inputStream = [System.IO.File]::OpenRead($file.FullName)
      $outputStream = $entry.Open()
      try {
        $inputStream.CopyTo($outputStream)
      } finally {
        $outputStream.Dispose()
        $inputStream.Dispose()
      }
    }
  } finally {
    $archive.Dispose()
  }
}

$resolvedTemplatePath = (Resolve-Path -LiteralPath $TemplatePath).Path
if ([System.IO.Path]::GetExtension($resolvedTemplatePath).ToLowerInvariant() -ne ".docx") {
  throw "Only .docx templates are supported: $resolvedTemplatePath"
}

$mappingTables = Get-MappingTables -PathToJson $MappingPath -InlineJson $FieldsJson
$locationMap = $mappingTables.LocationMap
$labelMap = $mappingTables.LabelMap

if ([string]::IsNullOrWhiteSpace($OutPath)) {
  $directory = Split-Path -Parent $resolvedTemplatePath
  $fileName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedTemplatePath)
  $OutPath = Join-Path $directory ($fileName + ".filled.docx")
}

$resolvedOutPath = [System.IO.Path]::GetFullPath($OutPath)
if ((-not $Overwrite) -and (Test-Path -LiteralPath $resolvedOutPath)) {
  throw "Output file already exists: $resolvedOutPath. Re-run with -Overwrite to replace it."
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("openclaw-docx-fill-" + [System.Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

try {
  [System.IO.Compression.ZipFile]::ExtractToDirectory($resolvedTemplatePath, $tempRoot)

  $documentXmlPath = Join-Path $tempRoot "word\document.xml"
  if (-not (Test-Path -LiteralPath $documentXmlPath)) {
    throw "word/document.xml was not found in $resolvedTemplatePath"
  }

  [xml]$documentXml = [System.IO.File]::ReadAllText($documentXmlPath, (New-Object System.Text.UTF8Encoding($false)))
  $namespaceManager = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
  $namespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

  $body = $documentXml.SelectSingleNode("/w:document/w:body", $namespaceManager)
  if ($null -eq $body) {
    throw "Could not locate /w:document/w:body in $resolvedTemplatePath"
  }

  $directFillCount = 0
  $labelFillCount = 0
  $blockFillCount = 0
  $insertedParagraphCount = 0
  $skippedLockedCount = 0
  $removedTableRowCount = 0
  $paragraphIndex = 0
  $tableIndex = 0

  foreach ($child in @($body.ChildNodes)) {
    if ($child.LocalName -eq "p") {
      $paragraphIndex++
      $location = ("P{0}" -f $paragraphIndex).ToUpperInvariant()
      $currentText = Get-NodeText -Node $child -NamespaceManager $namespaceManager
      $nextParagraph = Get-NextParagraphSibling -Node $child

      if ($locationMap.ContainsKey($location)) {
        $fillSpec = Get-FillSpec -Value $locationMap[$location]
        if (($fillSpec.Mode -eq "after") -and -not (Should-UseAfterHeadingFill -HeadingText $currentText -FillSpec $fillSpec -NextParagraph $nextParagraph -NamespaceManager $namespaceManager)) {
          $skippedLockedCount++
          continue
        }

        if (Should-UseAfterHeadingFill -HeadingText $currentText -FillSpec $fillSpec -NextParagraph $nextParagraph -NamespaceManager $namespaceManager) {
          $insertedParagraphCount += Apply-ParagraphBlock -TargetParagraph $nextParagraph -NamespaceManager $namespaceManager -Paragraphs $fillSpec.Paragraphs
          $blockFillCount++
        } else {
          $insertedParagraphCount += Apply-ParagraphBlock -TargetParagraph $child -NamespaceManager $namespaceManager -Paragraphs $fillSpec.Paragraphs
          if ($fillSpec.Paragraphs.Count -gt 1) {
            $blockFillCount++
          }
        }

        $directFillCount++
        continue
      }

      $paragraphMatched = $false
      if ($currentText -match $labelPattern) {
        $label = $matches["label"].Trim()
        $rest = $matches["rest"]
        $labelKey = Normalize-FieldKey -Text $label
        if ($labelMap.ContainsKey($labelKey)) {
          $fillSpec = Get-FillSpec -Value $labelMap[$labelKey]
          $inlineText = Get-InlineText -FillSpec $fillSpec
          $resolvedOptionFieldText = Resolve-OptionFieldText -TemplateText $currentText -SelectionText $inlineText
          if (-not [string]::IsNullOrWhiteSpace($resolvedOptionFieldText)) {
            Set-ParagraphText -Paragraph $child -NamespaceManager $namespaceManager -Text $resolvedOptionFieldText
            $labelFillCount++
          } elseif ([string]::IsNullOrWhiteSpace($rest) -or (Is-PlaceholderLike -Text $rest)) {
            Set-ParagraphText -Paragraph $child -NamespaceManager $namespaceManager -Text ("{0}{1}{2}" -f $label, $fullWidthColon, $inlineText)
            $labelFillCount++
          } else {
            $skippedLockedCount++
          }
          $paragraphMatched = $true
        }
      }

      if (-not $paragraphMatched) {
        $paragraphKey = Normalize-FieldKey -Text $currentText
        if ((-not [string]::IsNullOrWhiteSpace($paragraphKey)) -and $labelMap.ContainsKey($paragraphKey) -and $currentText.Length -le 30) {
          $fillSpec = Get-FillSpec -Value $labelMap[$paragraphKey]
          $isSectionHeading = Is-SectionHeadingLike -Text $currentText

          if (Should-UseAfterHeadingFill -HeadingText $currentText -FillSpec $fillSpec -NextParagraph $nextParagraph -NamespaceManager $namespaceManager) {
            $insertedParagraphCount += Apply-ParagraphBlock -TargetParagraph $nextParagraph -NamespaceManager $namespaceManager -Paragraphs $fillSpec.Paragraphs
            $blockFillCount++
            $labelFillCount++
          } elseif ($isSectionHeading -and (($fillSpec.Mode -eq "after") -or ($fillSpec.Paragraphs.Count -gt 1))) {
            $skippedLockedCount++
          } else {
            $inlineText = Get-InlineText -FillSpec $fillSpec
            Set-ParagraphText -Paragraph $child -NamespaceManager $namespaceManager -Text ("{0}{1}{2}" -f $currentText.Trim(), $fullWidthColon, $inlineText)
            $labelFillCount++
          }
        }
      }

      continue
    }

    if ($child.LocalName -eq "tbl") {
      $tableIndex++
      $tableId = "T{0}" -f $tableIndex
      $rowNodes = @($child.SelectNodes("./w:tr", $namespaceManager))
      $tableHandledByAfterFill = $false

      for ($rowOffset = 0; $rowOffset -lt $rowNodes.Count -and -not $tableHandledByAfterFill; $rowOffset++) {
        $rowIndex = $rowOffset + 1
        $row = $rowNodes[$rowOffset]
        $cellInfos = New-Object System.Collections.Generic.List[object]
        $cellIndex = 0

        foreach ($cell in $row.SelectNodes("./w:tc", $namespaceManager)) {
          $cellIndex++
          $location = ("{0}R{1}C{2}" -f $tableId, $rowIndex, $cellIndex).ToUpperInvariant()
          $text = Get-NodeText -Node $cell -NamespaceManager $namespaceManager
          $cellInfos.Add([pscustomobject]@{
              Node = $cell
              Location = $location
              Text = $text
            }) | Out-Null
        }

        for ($i = 0; $i -lt $cellInfos.Count; $i++) {
          $cellInfo = $cellInfos[$i]
          if ($locationMap.ContainsKey($cellInfo.Location)) {
            $fillSpec = Get-FillSpec -Value $locationMap[$cellInfo.Location]
            if ($fillSpec.Mode -eq "after-table") {
              $tableBodyResult = Move-TableBodyRowsAfterTable -Table $child -NamespaceManager $namespaceManager -StartLocation $cellInfo.Location -ThroughLocation $fillSpec.ThroughLocation -Paragraphs $fillSpec.Paragraphs
              $directFillCount++
              $blockFillCount++
              $insertedParagraphCount += $tableBodyResult.InsertedParagraphCount
              $removedTableRowCount += $tableBodyResult.RemovedRowCount
              $tableHandledByAfterFill = $true
              break
            }

            Set-CellParagraphs -Cell $cellInfo.Node -NamespaceManager $namespaceManager -Paragraphs $fillSpec.Paragraphs
            if ($fillSpec.Paragraphs.Count -gt 1) {
              $blockFillCount++
            }
            $directFillCount++
            continue
          }

          $currentText = $cellInfo.Text
          $matchedSameCell = $false

          if ($currentText -match $labelPattern) {
            $label = $matches["label"].Trim()
            $rest = $matches["rest"]
            $labelKey = Normalize-FieldKey -Text $label
            if ($labelMap.ContainsKey($labelKey)) {
              $fillSpec = Get-FillSpec -Value $labelMap[$labelKey]
              $inlineText = Get-InlineText -FillSpec $fillSpec
              $resolvedOptionFieldText = Resolve-OptionFieldText -TemplateText $currentText -SelectionText $inlineText
              if (-not [string]::IsNullOrWhiteSpace($resolvedOptionFieldText)) {
                Set-CellParagraphs -Cell $cellInfo.Node -NamespaceManager $namespaceManager -Paragraphs @($resolvedOptionFieldText)
                $labelFillCount++
              } elseif ([string]::IsNullOrWhiteSpace($rest) -or (Is-PlaceholderLike -Text $rest)) {
                Set-CellParagraphs -Cell $cellInfo.Node -NamespaceManager $namespaceManager -Paragraphs @("{0}{1}{2}" -f $label, $fullWidthColon, $inlineText)
                $labelFillCount++
              } else {
                $skippedLockedCount++
              }
              $matchedSameCell = $true
            }
          }

          if ($matchedSameCell) {
            continue
          }

          $labelKey = Normalize-FieldKey -Text $currentText
          if ((-not [string]::IsNullOrWhiteSpace($labelKey)) -and $labelMap.ContainsKey($labelKey) -and ($i + 1) -lt $cellInfos.Count) {
            $nextCellInfo = $cellInfos[$i + 1]
            $nextText = Get-NodeText -Node $nextCellInfo.Node -NamespaceManager $namespaceManager
            if ([string]::IsNullOrWhiteSpace($nextText) -or (Is-PlaceholderLike -Text $nextText)) {
              if ($locationMap.ContainsKey($nextCellInfo.Location)) {
                $fillSpec = Get-FillSpec -Value $locationMap[$nextCellInfo.Location]
                if ($fillSpec.Mode -eq "after-table") {
                  $tableBodyResult = Move-TableBodyRowsAfterTable -Table $child -NamespaceManager $namespaceManager -StartLocation $nextCellInfo.Location -ThroughLocation $fillSpec.ThroughLocation -Paragraphs $fillSpec.Paragraphs
                  $directFillCount++
                  $blockFillCount++
                  $insertedParagraphCount += $tableBodyResult.InsertedParagraphCount
                  $removedTableRowCount += $tableBodyResult.RemovedRowCount
                  $tableHandledByAfterFill = $true
                  break
                }

                Set-CellParagraphs -Cell $nextCellInfo.Node -NamespaceManager $namespaceManager -Paragraphs $fillSpec.Paragraphs
                if ($fillSpec.Paragraphs.Count -gt 1) {
                  $blockFillCount++
                }
                $directFillCount++
              } else {
                $fillSpec = Get-FillSpec -Value $labelMap[$labelKey]
                Set-CellParagraphs -Cell $nextCellInfo.Node -NamespaceManager $namespaceManager -Paragraphs $fillSpec.Paragraphs
                if ($fillSpec.Paragraphs.Count -gt 1) {
                  $blockFillCount++
                }
                $labelFillCount++
              }
            } else {
              $skippedLockedCount++
            }
          }
        }
      }

      continue
    }
  }

  [System.IO.File]::WriteAllText($documentXmlPath, $documentXml.OuterXml, (New-Object System.Text.UTF8Encoding($false)))

  if (Test-Path -LiteralPath $resolvedOutPath) {
    Remove-Item -LiteralPath $resolvedOutPath -Force
  }
  Write-OpenXmlPackage -SourceDirectory $tempRoot -DestinationPath $resolvedOutPath

  [pscustomobject]@{
    templatePath = $resolvedTemplatePath
    outPath = $resolvedOutPath
    appliedFillCount = $directFillCount + $labelFillCount
    directFillCount = $directFillCount
    labelFillCount = $labelFillCount
    blockFillCount = $blockFillCount
    insertedParagraphCount = $insertedParagraphCount
    removedTableRowCount = $removedTableRowCount
    skippedLockedCount = $skippedLockedCount
  }
} finally {
  if (Test-Path -LiteralPath $tempRoot) {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force
  }
}

