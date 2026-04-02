[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$Path,

  [ValidateSet("markdown", "json")]
  [string]$Format = "markdown",

  [string]$OutFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$sectionHeadingPattern = '^(?:\d+[.\u3001]?\s*)?(?:\u5B9E\u9A8C\u76EE\u7684|\u5B9E\u9A8C\u73AF\u5883|\u5B9E\u9A8C\u8BBE\u5907\u4E0E\u73AF\u5883|\u5B9E\u9A8C\u539F\u7406|\u5B9E\u9A8C\u6B65\u9AA4|\u5B9E\u9A8C\u7ED3\u679C|\u5B9E\u9A8C\u73B0\u8C61\u4E0E\u7ED3\u679C\u8BB0\u5F55|\u7ED3\u679C\u5206\u6790|\u95EE\u9898\u5206\u6790|\u5B9E\u9A8C\u603B\u7ED3|\u603B\u7ED3\u4E0E\u601D\u8003|\u5B9E\u9A8C\u5185\u5BB9|\u5B9E\u9A8C\u8FC7\u7A0B|\u5B9E\u9A8C\u7ED3\u8BBA|\u5B9E\u9A8C\u8981\u6C42|\u5B9E\u9A8C\u5668\u6750|\u5B9E\u9A8C\u4EEA\u5668|\u5B9E\u9A8C\u8BB0\u5F55|\u6CE8\u610F\u4E8B\u9879)$'

function Get-ZipEntryText {
  param(
    [Parameter(Mandatory = $true)]
    [System.IO.Compression.ZipArchive]$Archive,

    [Parameter(Mandatory = $true)]
    [string]$EntryName
  )

  $entry = $Archive.GetEntry($EntryName)
  if ($null -eq $entry) {
    return $null
  }

  $stream = $entry.Open()
  $reader = New-Object System.IO.StreamReader($stream)
  try {
    return $reader.ReadToEnd()
  } finally {
    $reader.Dispose()
    $stream.Dispose()
  }
}

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

function Get-FieldReason {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return $null
  }

  if ($Text -match '[:\uFF1A]\s*$') {
    return "label-ending-colon"
  }

  if ($Text -match '[_\uFF3F]{2,}|\.{3,}|\uFF08\s*\uFF09|\(\s*\)|\u25A1|\u25A0') {
    return "placeholder-shape"
  }

  if ($Text.Length -le 30 -and $Text -match $sectionHeadingPattern) {
    return "common-report-section-heading"
  }

  if ($Text.Length -le 60 -and $Text -match '\u8BFE\u7A0B|\u5B9E\u9A8C|\u59D3\u540D|\u5B66\u53F7|\u73ED\u7EA7|\u65E5\u671F|\u6559\u5E08|\u5B66\u9662|\u4E13\u4E1A|\u6210\u7EE9|\u5730\u70B9|\u7EC4\u5458|\u6307\u5BFC') {
    return "common-report-field"
  }

  return $null
}

function Add-LikelyField {
  param(
    [System.Collections.Generic.List[object]]$Target,

    [Parameter(Mandatory = $true)]
    [string]$Location,

    [AllowNull()]
    [string]$Text,

    [AllowNull()]
    [string]$Reason
  )

  if ([string]::IsNullOrWhiteSpace($Reason)) {
    return
  }

  $Target.Add([pscustomobject]@{
      location = $Location
      text = $Text
      reason = $Reason
    }) | Out-Null
}

$resolvedPath = Resolve-Path -LiteralPath $Path
if ([System.IO.Path]::GetExtension($resolvedPath.Path).ToLowerInvariant() -ne ".docx") {
  throw "Only .docx templates are supported: $($resolvedPath.Path)"
}

$archive = [System.IO.Compression.ZipFile]::OpenRead($resolvedPath.Path)
try {
  $documentXmlText = Get-ZipEntryText -Archive $archive -EntryName "word/document.xml"
  if ([string]::IsNullOrWhiteSpace($documentXmlText)) {
    throw "word/document.xml was not found in $($resolvedPath.Path)"
  }

  [xml]$documentXml = $documentXmlText
  $namespaceManager = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
  $namespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

  $body = $documentXml.SelectSingleNode("/w:document/w:body", $namespaceManager)
  if ($null -eq $body) {
    throw "Could not locate /w:document/w:body in $($resolvedPath.Path)"
  }

  $blocks = New-Object System.Collections.Generic.List[object]
  $likelyFields = New-Object System.Collections.Generic.List[object]
  $paragraphCount = 0
  $tableCount = 0

  foreach ($child in $body.ChildNodes) {
    if ($child.LocalName -eq "p") {
      $paragraphCount++
      $text = Get-NodeText -Node $child -NamespaceManager $namespaceManager
      $location = "P$paragraphCount"
      $reason = Get-FieldReason -Text $text
      Add-LikelyField -Target $likelyFields -Location $location -Text $text -Reason $reason
      $blocks.Add([pscustomobject]@{
          type = "paragraph"
          location = $location
          text = $text
        }) | Out-Null
      continue
    }

    if ($child.LocalName -eq "tbl") {
      $tableCount++
      $tableId = "T$tableCount"
      $rowIndex = 0
      $rows = New-Object System.Collections.Generic.List[object]

      foreach ($row in $child.SelectNodes("./w:tr", $namespaceManager)) {
        $rowIndex++
        $cellIndex = 0
        $cells = New-Object System.Collections.Generic.List[object]
        $previousCellText = ""
        $previousCellReason = $null

        foreach ($cell in $row.SelectNodes("./w:tc", $namespaceManager)) {
          $cellIndex++
          $cellText = Get-NodeText -Node $cell -NamespaceManager $namespaceManager
          $location = "{0}R{1}C{2}" -f $tableId, $rowIndex, $cellIndex
          $reason = Get-FieldReason -Text $cellText
          if ([string]::IsNullOrWhiteSpace($cellText) -and -not [string]::IsNullOrWhiteSpace($previousCellText) -and $previousCellReason -in @("label-ending-colon", "common-report-field")) {
            $reason = "empty-cell-after-label"
          }

          Add-LikelyField -Target $likelyFields -Location $location -Text $cellText -Reason $reason

          $cells.Add([pscustomobject]@{
              location = $location
              text = $cellText
            }) | Out-Null

          $previousCellText = $cellText
          $previousCellReason = Get-FieldReason -Text $cellText
        }

        $rows.Add([pscustomobject]@{
            row = $rowIndex
            cells = $cells
          }) | Out-Null
      }

      $blocks.Add([pscustomobject]@{
          type = "table"
          location = $tableId
          rows = $rows
        }) | Out-Null
    }
  }

  $result = [pscustomobject]@{
    source = $resolvedPath.Path
    summary = [pscustomobject]@{
      paragraphCount = $paragraphCount
      tableCount = $tableCount
      likelyFieldCount = $likelyFields.Count
    }
    blocks = $blocks
    likelyFields = $likelyFields
  }

  if ($Format -eq "json") {
    $output = $result | ConvertTo-Json -Depth 8
  } else {
    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add("# DOCX Template Outline")
    [void]$lines.Add("")
    [void]$lines.Add("- Source: $($result.source)")
    [void]$lines.Add("- Paragraph count: $paragraphCount")
    [void]$lines.Add("- Table count: $tableCount")
    [void]$lines.Add("- Likely field count: $($likelyFields.Count)")
    [void]$lines.Add("")
    [void]$lines.Add("## Block Order")

    foreach ($block in $blocks) {
      if ($block.type -eq "paragraph") {
        $displayText = if ([string]::IsNullOrWhiteSpace($block.text)) { "[empty]" } else { $block.text }
        [void]$lines.Add("- $($block.location): $displayText")
        continue
      }

      [void]$lines.Add("- $($block.location): table")
      foreach ($row in $block.rows) {
        foreach ($cell in $row.cells) {
          $displayText = if ([string]::IsNullOrWhiteSpace($cell.text)) { "[empty]" } else { $cell.text }
          [void]$lines.Add("  - $($cell.location): $displayText")
        }
      }
    }

    [void]$lines.Add("")
    [void]$lines.Add("## Likely Fillable Fields")
    if ($likelyFields.Count -eq 0) {
      [void]$lines.Add("- None detected automatically")
    } else {
      foreach ($field in $likelyFields) {
        $displayText = if ([string]::IsNullOrWhiteSpace($field.text)) { "[empty]" } else { $field.text }
        [void]$lines.Add("- $($field.location): $displayText ($($field.reason))")
      }
    }

    $output = $lines -join [Environment]::NewLine
  }

  if ([string]::IsNullOrWhiteSpace($OutFile)) {
    Write-Output $output
  } else {
    [System.IO.File]::WriteAllText($OutFile, $output, (New-Object System.Text.UTF8Encoding($true)))
    Write-Output "Wrote template analysis to $OutFile"
  }
} finally {
  $archive.Dispose()
}

