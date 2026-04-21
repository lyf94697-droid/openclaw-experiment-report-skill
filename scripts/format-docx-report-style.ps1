[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$DocxPath,

  [string]$OutPath,

  [switch]$Overwrite,

  [ValidateSet("auto", "default", "compact", "school")]
  [string]$Profile = "auto",

  [string]$ProfilePath,

  [string]$ReportProfileName = "experiment-report",

  [string]$ReportProfilePath,

  [int]$BodyFirstLineTwips = 420,

  [int]$BodyLineTwips = 360,

  [int]$BodyAfterTwips = 0,

  [int]$HeadingBeforeTwips = 120,

  [int]$HeadingAfterTwips = 80,

  [int]$CaptionAfterTwips = 80,

  [int]$TitleAfterTwips = 160,

  [int]$ImageBeforeTwips = 80,

  [int]$ImageAfterTwips = 80,

  [int]$TitleFontHalfPoints = 32,

  [int]$HeadingFontHalfPoints = 28,

  [int]$BodyFontHalfPoints = 24,

  [int]$CaptionFontHalfPoints = 21,

  [int]$MetadataFontHalfPoints = 24,

  [int]$ListFontHalfPoints = 24,

  [int]$ListAfterTwips = 0,

  [int]$CommandBeforeTwips = 40,

  [int]$CommandAfterTwips = 40,

  [int]$CommandLineTwips = 240,

  [int]$CommandFontHalfPoints = 20
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

. (Join-Path $PSScriptRoot "report-profiles.ps1")

$script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$reportProfile = Get-ReportProfile -ProfileName $ReportProfileName -ProfilePath $ReportProfilePath -RepoRoot $script:RepoRoot
$wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
$sectionHeadingRules = @(
  @(Get-ReportProfileSectionRules -Profile $reportProfile) |
    ForEach-Object {
      [pscustomobject]@{ aliases = @($_.headingAliases) }
    }
)
foreach ($extraHeading in (Get-ReportProfileExtraSectionHeadings -Profile $reportProfile)) {
  $sectionHeadingRules += [pscustomobject]@{ aliases = @([string]$extraHeading) }
}
$metadataPrefixes = @(Get-ReportProfileMetadataPrefixes -Profile $reportProfile)
$styleSettingNames = @(
  "BodyFirstLineTwips",
  "BodyLineTwips",
  "BodyAfterTwips",
  "HeadingBeforeTwips",
  "HeadingAfterTwips",
  "CaptionAfterTwips",
  "TitleAfterTwips",
  "ImageBeforeTwips",
  "ImageAfterTwips",
  "TitleFontHalfPoints",
  "HeadingFontHalfPoints",
  "BodyFontHalfPoints",
  "CaptionFontHalfPoints",
  "MetadataFontHalfPoints",
  "ListFontHalfPoints",
  "ListAfterTwips",
  "CommandBeforeTwips",
  "CommandAfterTwips",
  "CommandLineTwips",
  "CommandFontHalfPoints"
)

function Get-StyleSettingNames {
  return $styleSettingNames
}

function Get-StyleProfileSettings {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ProfileName
  )

  switch ($ProfileName.ToLowerInvariant()) {
    "default" {
      return [ordered]@{
        BodyFirstLineTwips = 420
        BodyLineTwips = 360
        BodyAfterTwips = 0
        HeadingBeforeTwips = 120
        HeadingAfterTwips = 80
        CaptionAfterTwips = 80
        TitleAfterTwips = 160
        ImageBeforeTwips = 80
        ImageAfterTwips = 80
        TitleFontHalfPoints = 32
        HeadingFontHalfPoints = 28
        BodyFontHalfPoints = 24
        CaptionFontHalfPoints = 21
        MetadataFontHalfPoints = 24
        ListFontHalfPoints = 24
        ListAfterTwips = 0
        CommandBeforeTwips = 40
        CommandAfterTwips = 40
        CommandLineTwips = 240
        CommandFontHalfPoints = 20
      }
    }
    "compact" {
      return [ordered]@{
        BodyFirstLineTwips = 420
        BodyLineTwips = 320
        BodyAfterTwips = 0
        HeadingBeforeTwips = 80
        HeadingAfterTwips = 40
        CaptionAfterTwips = 80
        TitleAfterTwips = 80
        ImageBeforeTwips = 40
        ImageAfterTwips = 40
        TitleFontHalfPoints = 30
        HeadingFontHalfPoints = 30
        BodyFontHalfPoints = 24
        CaptionFontHalfPoints = 24
        MetadataFontHalfPoints = 24
        ListFontHalfPoints = 24
        ListAfterTwips = 0
        CommandBeforeTwips = 20
        CommandAfterTwips = 20
        CommandLineTwips = 240
        CommandFontHalfPoints = 20
      }
    }
    "school" {
      return [ordered]@{
        BodyFirstLineTwips = 420
        BodyLineTwips = 400
        BodyAfterTwips = 0
        HeadingBeforeTwips = 160
        HeadingAfterTwips = 100
        CaptionAfterTwips = 120
        TitleAfterTwips = 220
        ImageBeforeTwips = 100
        ImageAfterTwips = 100
        TitleFontHalfPoints = 32
        HeadingFontHalfPoints = 28
        BodyFontHalfPoints = 24
        CaptionFontHalfPoints = 21
        MetadataFontHalfPoints = 24
        ListFontHalfPoints = 24
        ListAfterTwips = 0
        CommandBeforeTwips = 60
        CommandAfterTwips = 60
        CommandLineTwips = 240
        CommandFontHalfPoints = 20
      }
    }
    default {
      throw "Unsupported style profile: $ProfileName"
    }
  }
}

function ConvertTo-ValidatedTwipsValue {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Name,

    [AllowNull()]
    [object]$Value
  )

  if ($null -eq $Value) {
    throw "Style setting '$Name' cannot be null."
  }

  try {
    $parsedValue = [int]$Value
  } catch {
    throw "Style setting '$Name' must be an integer, got '$Value'."
  }

  if ($parsedValue -lt 0) {
    throw "Style setting '$Name' must be greater than or equal to 0."
  }

  return $parsedValue
}

function Get-OptionalObjectPropertyValue {
  param(
    [AllowNull()]
    [object]$Object,

    [Parameter(Mandatory = $true)]
    [string]$Name
  )

  if ($null -eq $Object) {
    return $null
  }

  $property = $Object.PSObject.Properties[$Name]
  if ($null -eq $property) {
    return $null
  }

  return $property.Value
}

function Read-StyleProfileFile {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $resolvedProfilePath = (Resolve-Path -LiteralPath $Path).Path
  $rawProfileJson = Get-Content -LiteralPath $resolvedProfilePath -Raw -Encoding UTF8
  if ([string]::IsNullOrWhiteSpace($rawProfileJson)) {
    throw "Style profile file is empty: $resolvedProfilePath"
  }

  try {
    $profileDocument = $rawProfileJson | ConvertFrom-Json
  } catch {
    throw "Style profile file is not valid JSON: $resolvedProfilePath. $($_.Exception.Message)"
  }

  if ($null -eq $profileDocument) {
    throw "Style profile file did not produce a JSON object: $resolvedProfilePath"
  }

  $baseProfileValue = Get-OptionalObjectPropertyValue -Object $profileDocument -Name "baseProfile"
  $profileAliasValue = Get-OptionalObjectPropertyValue -Object $profileDocument -Name "profile"
  if (
    -not [string]::IsNullOrWhiteSpace([string]$baseProfileValue) -and
    -not [string]::IsNullOrWhiteSpace([string]$profileAliasValue) -and
    (-not [string]::Equals([string]$baseProfileValue, [string]$profileAliasValue, [System.StringComparison]::OrdinalIgnoreCase))
  ) {
    throw "Style profile file cannot define different values for 'baseProfile' and 'profile'."
  }

  $baseProfile = if (-not [string]::IsNullOrWhiteSpace([string]$baseProfileValue)) {
    [string]$baseProfileValue
  } elseif (-not [string]::IsNullOrWhiteSpace([string]$profileAliasValue)) {
    [string]$profileAliasValue
  } else {
    $null
  }

  if (-not [string]::IsNullOrWhiteSpace($baseProfile)) {
    $baseProfile = $baseProfile.ToLowerInvariant()
  }

  if (-not [string]::IsNullOrWhiteSpace($baseProfile) -and @("auto", "default", "compact", "school") -notcontains $baseProfile) {
    throw "Style profile file has unsupported baseProfile/profile value '$baseProfile'."
  }

  $settingsNode = Get-OptionalObjectPropertyValue -Object $profileDocument -Name "settings"
  if ($null -eq $settingsNode) {
    $settingsNode = $profileDocument
  }

  $overrides = [ordered]@{}
  foreach ($settingName in (Get-StyleSettingNames)) {
    $settingValue = Get-OptionalObjectPropertyValue -Object $settingsNode -Name $settingName
    if ($null -ne $settingValue) {
      $overrides[$settingName] = ConvertTo-ValidatedTwipsValue -Name $settingName -Value $settingValue
    }
  }

  return [pscustomobject]@{
    path = $resolvedProfilePath
    baseProfile = $baseProfile
    overrides = [pscustomobject]$overrides
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

function Get-ParagraphText {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Paragraph,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $parts = New-Object System.Collections.Generic.List[string]
  foreach ($textNode in @($Paragraph.SelectNodes(".//w:t | .//w:instrText", $NamespaceManager))) {
    [void]$parts.Add($textNode.InnerText)
  }
  return Normalize-OpenXmlText -Text ($parts -join "")
}

function Test-IsRemovableTrailingParagraph {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Paragraph,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $text = Get-ParagraphText -Paragraph $Paragraph -NamespaceManager $NamespaceManager
  if (-not [string]::IsNullOrWhiteSpace($text)) {
    return $false
  }

  $protectedContent = $Paragraph.SelectSingleNode(".//w:drawing | .//w:tbl | .//w:sectPr | .//w:br[@w:type='page']", $NamespaceManager)
  return ($null -eq $protectedContent)
}

function Remove-TrailingEmptyBodyParagraphs {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Body,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $removedCount = 0
  while ($Body.ChildNodes.Count -gt 1) {
    $candidateIndex = $Body.ChildNodes.Count - 1
    $lastChild = $Body.ChildNodes[$candidateIndex]
    if ($lastChild.LocalName -eq "sectPr") {
      $candidateIndex--
      if ($candidateIndex -lt 0) {
        break
      }
      $lastChild = $Body.ChildNodes[$candidateIndex]
    }

    if ($lastChild.LocalName -ne "p") {
      break
    }

    if (-not (Test-IsRemovableTrailingParagraph -Paragraph $lastChild -NamespaceManager $NamespaceManager)) {
      break
    }

    [void]$Body.RemoveChild($lastChild)
    $removedCount++
  }

  return $removedCount
}

function Get-NextMeaningfulBodyParagraphInfo {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Paragraph,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $candidate = $Paragraph.NextSibling
  while ($null -ne $candidate) {
    if ($candidate.LocalName -eq "p") {
      $text = Get-ParagraphText -Paragraph $candidate -NamespaceManager $NamespaceManager
      if (-not [string]::IsNullOrWhiteSpace($text)) {
        return [pscustomobject]@{
          Paragraph = $candidate
          Text = $text
        }
      }
    }

    $candidate = $candidate.NextSibling
  }

  return $null
}

function Remove-CourseDesignDuplicatePlaceholderParagraphs {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Body,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $removedCount = 0
  $paragraphs = @($Body.SelectNodes("./w:p", $NamespaceManager))
  foreach ($paragraph in $paragraphs) {
    if ($null -eq $paragraph.ParentNode) {
      continue
    }

    $text = Get-ParagraphText -Paragraph $paragraph -NamespaceManager $NamespaceManager
    if ([string]::IsNullOrWhiteSpace($text)) {
      continue
    }

    $nextInfo = Get-NextMeaningfulBodyParagraphInfo -Paragraph $paragraph -NamespaceManager $NamespaceManager
    if ($null -eq $nextInfo) {
      continue
    }

    $removeParagraph = $false
    switch ($text) {
      "问题分析（需求分析、可行性分析）" {
        $removeParagraph = [string]::Equals([string]$nextInfo.Text, "摘要：", [System.StringComparison]::Ordinal)
        break
      }
      "实现结果" {
        $removeParagraph = ([string]$nextInfo.Text -match '^[五5][、\.．]\s*实现结果$')
        break
      }
      "总结" {
        $removeParagraph = ([string]$nextInfo.Text -match '^[七7][、\.．]\s*(设计)?总结$')
        break
      }
    }

    if ($removeParagraph) {
      [void]$paragraph.ParentNode.RemoveChild($paragraph)
      $removedCount++
    }
  }

  return $removedCount
}

function Get-HeadingPattern {
  param(
    [Parameter(Mandatory = $true)]
    [string[]]$Aliases
  )

  $escapedAliases = foreach ($alias in $Aliases) {
    [regex]::Escape($alias)
  }

  return '^(?:第?[0-9一二三四五六七八九十]+(?:章|节)?[.\)\u3001]?\s*)?(?:' + ($escapedAliases -join '|') + ')\s*(?:[:\uFF1A])?$'
}

function Test-IsSectionHeading {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return $false
  }

  foreach ($rule in $sectionHeadingRules) {
    if ($Text -match (Get-HeadingPattern -Aliases $rule.aliases)) {
      return $true
    }
  }

  return $false
}

function Test-IsMetadataParagraph {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return $false
  }

  foreach ($prefix in $metadataPrefixes) {
    if ($Text -match ('^' + [regex]::Escape($prefix) + '\s*[:：]')) {
      return $true
    }
  }

  return $false
}

function Test-IsTitleParagraph {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return $false
  }

  $compact = ($Text -replace '\s+', '')
  return ($compact -match '实验报告(?:[（(][^）)]+[）)])?$')
}

function Test-IsCaptionParagraph {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return $false
  }

  return ($Text -match '^图\d+\s*')
}

function Test-IsStepListParagraph {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return $false
  }

  return (
    $Text -match '^\s*[\(（]?[0-9一二三四五六七八九十]+[\)）\u3001]\s*\S+' -or
    $Text -match '^\s*[0-9一二三四五六七八九十]+\.\s+\S+' -or
    $Text -match '^\s*步骤\s*[一二三四五六七八九十0-9]*\s*[\.:：\u3001]?\s*\S+'
  )
}

function Test-IsCommandParagraph {
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
    $trimmed -match '(?i)^(?:ipconfig|ping|arp|tracert|netstat|nslookup|route|netsh|net\s+|cd\s+|dir\b|java\b|javac\b|gradle\b|adb\b|git\b|powershell\b|cmd\b|gcc\b|g\+\+\b|clang\b|clang\+\+\b|make\b|cmake\b|\.\/\S+)(?:\s|$)' -or
    $trimmed -match '(?i)^(?:reply from|pinging|packets:|minimum =|maximum =|ipv4 address|subnet mask|default gateway|physical address|ethernet adapter)\b'
  )
}

function Test-IsCodeParagraph {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return $false
  }

  $trimmed = $Text.Trim()
  if ($trimmed.Length -gt 240 -or (Test-IsCommandParagraph -Text $trimmed)) {
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

  return (
    $trimmed -notmatch '[一-龥]' -and
    ($trimmed -match ';$' -or $trimmed -match '\->') -and
    $trimmed -match '[A-Za-z_]'
  )
}

function Get-OrCreateChildElement {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Parent,

    [Parameter(Mandatory = $true)]
    [string]$LocalName
  )

  $existing = $Parent.SelectSingleNode("./w:$LocalName", $script:namespaceManager)
  if ($null -ne $existing) {
    return $existing
  }

  $child = $Parent.OwnerDocument.CreateElement("w", $LocalName, $wordNamespace)
  $Parent.AppendChild($child) | Out-Null
  return $child
}

function Get-OrCreateParagraphProperties {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Paragraph
  )

  $existing = $Paragraph.SelectSingleNode("./w:pPr", $script:namespaceManager)
  if ($null -ne $existing) {
    return $existing
  }

  $pPr = $Paragraph.OwnerDocument.CreateElement("w", "pPr", $wordNamespace)
  if ($Paragraph.HasChildNodes) {
    $Paragraph.InsertBefore($pPr, $Paragraph.FirstChild) | Out-Null
  } else {
    $Paragraph.AppendChild($pPr) | Out-Null
  }
  return $pPr
}

function Get-OrCreateCellProperties {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Cell
  )

  $existing = $Cell.SelectSingleNode("./w:tcPr", $script:namespaceManager)
  if ($null -ne $existing) {
    return $existing
  }

  $tcPr = $Cell.OwnerDocument.CreateElement("w", "tcPr", $wordNamespace)
  if ($Cell.HasChildNodes) {
    $Cell.InsertBefore($tcPr, $Cell.FirstChild) | Out-Null
  } else {
    $Cell.AppendChild($tcPr) | Out-Null
  }
  return $tcPr
}

function Set-WordAttribute {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlElement]$Element,

    [Parameter(Mandatory = $true)]
    [string]$LocalName,

    [Parameter(Mandatory = $true)]
    [string]$Value
  )

  [void]$Element.SetAttribute($LocalName, $wordNamespace, $Value)
}

function Remove-WordChildElement {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Parent,

    [Parameter(Mandatory = $true)]
    [string]$LocalName
  )

  $existing = $Parent.SelectSingleNode("./w:$LocalName", $script:namespaceManager)
  if ($null -ne $existing) {
    $Parent.RemoveChild($existing) | Out-Null
    return $true
  }

  return $false
}

function Set-ParagraphJustification {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Paragraph,

    [Parameter(Mandatory = $true)]
    [string]$Value
  )

  $pPr = Get-OrCreateParagraphProperties -Paragraph $Paragraph
  $jc = Get-OrCreateChildElement -Parent $pPr -LocalName "jc"
  Set-WordAttribute -Element $jc -LocalName "val" -Value $Value
}

function Set-ParagraphIndent {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Paragraph,

    [AllowNull()]
    [int]$FirstLine
  )

  $pPr = Get-OrCreateParagraphProperties -Paragraph $Paragraph
  $ind = Get-OrCreateChildElement -Parent $pPr -LocalName "ind"
  $ind.RemoveAttribute("firstLine", $wordNamespace)
  $ind.RemoveAttribute("hanging", $wordNamespace)

  if ($null -ne $FirstLine) {
    Set-WordAttribute -Element $ind -LocalName "firstLine" -Value ([string]$FirstLine)
  }
}

function Set-ParagraphSpacing {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Paragraph,

    [AllowNull()]
    [int]$Before,

    [AllowNull()]
    [int]$After,

    [AllowNull()]
    [int]$Line
  )

  $pPr = Get-OrCreateParagraphProperties -Paragraph $Paragraph
  $spacing = Get-OrCreateChildElement -Parent $pPr -LocalName "spacing"
  if ($null -ne $Before) {
    Set-WordAttribute -Element $spacing -LocalName "before" -Value ([string]$Before)
  }
  if ($null -ne $After) {
    Set-WordAttribute -Element $spacing -LocalName "after" -Value ([string]$After)
  }
  if ($null -ne $Line) {
    Set-WordAttribute -Element $spacing -LocalName "line" -Value ([string]$Line)
    Set-WordAttribute -Element $spacing -LocalName "lineRule" -Value "auto"
  }
}

function Set-ParagraphPagination {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Paragraph,

    [bool]$KeepNext = $false,

    [bool]$KeepLines = $false
  )

  $pPr = Get-OrCreateParagraphProperties -Paragraph $Paragraph
  if ($KeepNext) {
    [void](Get-OrCreateChildElement -Parent $pPr -LocalName "keepNext")
  } else {
    [void](Remove-WordChildElement -Parent $pPr -LocalName "keepNext")
  }

  if ($KeepLines) {
    [void](Get-OrCreateChildElement -Parent $pPr -LocalName "keepLines")
  } else {
    [void](Remove-WordChildElement -Parent $pPr -LocalName "keepLines")
  }
}

function Set-ParagraphBold {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Paragraph
  )

  $pPr = Get-OrCreateParagraphProperties -Paragraph $Paragraph
  $rPr = Get-OrCreateChildElement -Parent $pPr -LocalName "rPr"
  [void](Get-OrCreateChildElement -Parent $rPr -LocalName "b")
  [void](Get-OrCreateChildElement -Parent $rPr -LocalName "bCs")

  foreach ($run in @($Paragraph.SelectNodes("./w:r", $script:namespaceManager))) {
    $runRPr = $run.SelectSingleNode("./w:rPr", $script:namespaceManager)
    if ($null -eq $runRPr) {
      $runRPr = $Paragraph.OwnerDocument.CreateElement("w", "rPr", $wordNamespace)
      if ($run.HasChildNodes) {
        $run.InsertBefore($runRPr, $run.FirstChild) | Out-Null
      } else {
        $run.AppendChild($runRPr) | Out-Null
      }
    }
    [void](Get-OrCreateChildElement -Parent $runRPr -LocalName "b")
    [void](Get-OrCreateChildElement -Parent $runRPr -LocalName "bCs")
  }
}

function Set-RunFont {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Paragraph,

    [Parameter(Mandatory = $true)]
    [string]$FontName,

    [Parameter(Mandatory = $true)]
    [int]$SizeHalfPoints
  )

  $runProperties = New-Object System.Collections.Generic.List[System.Xml.XmlNode]
  $pPr = Get-OrCreateParagraphProperties -Paragraph $Paragraph
  [void]$runProperties.Add((Get-OrCreateChildElement -Parent $pPr -LocalName "rPr"))

  foreach ($run in @($Paragraph.SelectNodes("./w:r", $script:namespaceManager))) {
    $runRPr = $run.SelectSingleNode("./w:rPr", $script:namespaceManager)
    if ($null -eq $runRPr) {
      $runRPr = $Paragraph.OwnerDocument.CreateElement("w", "rPr", $wordNamespace)
      if ($run.HasChildNodes) {
        $run.InsertBefore($runRPr, $run.FirstChild) | Out-Null
      } else {
        $run.AppendChild($runRPr) | Out-Null
      }
    }
    [void]$runProperties.Add($runRPr)
  }

  foreach ($rPr in $runProperties) {
    $rFonts = Get-OrCreateChildElement -Parent $rPr -LocalName "rFonts"
    Set-WordAttribute -Element $rFonts -LocalName "ascii" -Value $FontName
    Set-WordAttribute -Element $rFonts -LocalName "hAnsi" -Value $FontName
    Set-WordAttribute -Element $rFonts -LocalName "eastAsia" -Value $FontName
    Set-WordAttribute -Element $rFonts -LocalName "cs" -Value $FontName

    $sz = Get-OrCreateChildElement -Parent $rPr -LocalName "sz"
    Set-WordAttribute -Element $sz -LocalName "val" -Value ([string]$SizeHalfPoints)
    $szCs = Get-OrCreateChildElement -Parent $rPr -LocalName "szCs"
    Set-WordAttribute -Element $szCs -LocalName "val" -Value ([string]$SizeHalfPoints)
  }
}

function Set-RunFontSize {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Paragraph,

    [Parameter(Mandatory = $true)]
    [int]$SizeHalfPoints
  )

  $runProperties = New-Object System.Collections.Generic.List[System.Xml.XmlNode]
  $pPr = Get-OrCreateParagraphProperties -Paragraph $Paragraph
  [void]$runProperties.Add((Get-OrCreateChildElement -Parent $pPr -LocalName "rPr"))

  foreach ($run in @($Paragraph.SelectNodes("./w:r", $script:namespaceManager))) {
    $runRPr = $run.SelectSingleNode("./w:rPr", $script:namespaceManager)
    if ($null -eq $runRPr) {
      $runRPr = $Paragraph.OwnerDocument.CreateElement("w", "rPr", $wordNamespace)
      if ($run.HasChildNodes) {
        $run.InsertBefore($runRPr, $run.FirstChild) | Out-Null
      } else {
        $run.AppendChild($runRPr) | Out-Null
      }
    }
    [void]$runProperties.Add($runRPr)
  }

  foreach ($rPr in $runProperties) {
    $sz = Get-OrCreateChildElement -Parent $rPr -LocalName "sz"
    Set-WordAttribute -Element $sz -LocalName "val" -Value ([string]$SizeHalfPoints)
    $szCs = Get-OrCreateChildElement -Parent $rPr -LocalName "szCs"
    Set-WordAttribute -Element $szCs -LocalName "val" -Value ([string]$SizeHalfPoints)
  }
}

function Set-RunColor {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Paragraph,

    [Parameter(Mandatory = $true)]
    [string]$Color
  )

  $runProperties = New-Object System.Collections.Generic.List[System.Xml.XmlNode]
  $pPr = Get-OrCreateParagraphProperties -Paragraph $Paragraph
  [void]$runProperties.Add((Get-OrCreateChildElement -Parent $pPr -LocalName "rPr"))

  foreach ($run in @($Paragraph.SelectNodes("./w:r", $script:namespaceManager))) {
    $runRPr = $run.SelectSingleNode("./w:rPr", $script:namespaceManager)
    if ($null -eq $runRPr) {
      $runRPr = $Paragraph.OwnerDocument.CreateElement("w", "rPr", $wordNamespace)
      if ($run.HasChildNodes) {
        $run.InsertBefore($runRPr, $run.FirstChild) | Out-Null
      } else {
        $run.AppendChild($runRPr) | Out-Null
      }
    }
    [void]$runProperties.Add($runRPr)
  }

  foreach ($rPr in $runProperties) {
    $colorNode = Get-OrCreateChildElement -Parent $rPr -LocalName "color"
    Set-WordAttribute -Element $colorNode -LocalName "val" -Value $Color
  }
}

function Set-RunTypography {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Paragraph,

    [Parameter(Mandatory = $true)]
    [string]$FontName,

    [Parameter(Mandatory = $true)]
    [int]$SizeHalfPoints,

    [bool]$PreserveFontFamily = $false
  )

  if ($PreserveFontFamily) {
    Set-RunFontSize -Paragraph $Paragraph -SizeHalfPoints $SizeHalfPoints
  } else {
    Set-RunFont -Paragraph $Paragraph -FontName $FontName -SizeHalfPoints $SizeHalfPoints
  }
  Set-RunColor -Paragraph $Paragraph -Color "000000"
}

function Set-ParagraphShading {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Paragraph,

    [Parameter(Mandatory = $true)]
    [string]$Fill
  )

  $pPr = Get-OrCreateParagraphProperties -Paragraph $Paragraph
  $shd = Get-OrCreateChildElement -Parent $pPr -LocalName "shd"
  Set-WordAttribute -Element $shd -LocalName "val" -Value "clear"
  Set-WordAttribute -Element $shd -LocalName "color" -Value "auto"
  Set-WordAttribute -Element $shd -LocalName "fill" -Value $Fill
}

function Set-CellVerticalAlignmentTop {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Cell
  )

  $tcPr = Get-OrCreateCellProperties -Cell $Cell
  $vAlign = Get-OrCreateChildElement -Parent $tcPr -LocalName "vAlign"
  Set-WordAttribute -Element $vAlign -LocalName "val" -Value "top"
}

function Set-CellMargins {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Cell,

    [int]$TopTwips = 60,

    [int]$BottomTwips = 60,

    [int]$LeftTwips = 80,

    [int]$RightTwips = 80
  )

  $tcPr = Get-OrCreateCellProperties -Cell $Cell
  $tcMar = Get-OrCreateChildElement -Parent $tcPr -LocalName "tcMar"
  foreach ($marginSpec in @(
      @{ Name = "top"; Value = $TopTwips },
      @{ Name = "bottom"; Value = $BottomTwips },
      @{ Name = "left"; Value = $LeftTwips },
      @{ Name = "right"; Value = $RightTwips }
    )) {
    $margin = Get-OrCreateChildElement -Parent $tcMar -LocalName $marginSpec.Name
    Set-WordAttribute -Element $margin -LocalName "w" -Value ([string]$marginSpec.Value)
    Set-WordAttribute -Element $margin -LocalName "type" -Value "dxa"
  }
}

function Test-IsParagraphInTable {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Paragraph
  )

  return ($null -ne $Paragraph.SelectSingleNode("ancestor::w:tbl", $script:namespaceManager))
}

function Test-IsBodyTableRow {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Row
  )

  $paragraphs = @($Row.SelectNodes(".//w:p", $script:namespaceManager))
  if ($paragraphs.Count -eq 0) {
    return $false
  }

  $texts = New-Object System.Collections.Generic.List[string]
  $metadataParagraphCount = 0
  foreach ($paragraph in $paragraphs) {
    $text = Get-ParagraphText -Paragraph $paragraph -NamespaceManager $script:namespaceManager
    if ([string]::IsNullOrWhiteSpace($text)) {
      continue
    }

    [void]$texts.Add($text)
    if (Test-IsMetadataParagraph -Text $text) {
      $metadataParagraphCount++
    }
    if (Test-IsSectionHeading -Text $text) {
      return $true
    }
  }

  if ($null -ne $Row.SelectSingleNode(".//w:drawing", $script:namespaceManager)) {
    return $true
  }

  if ($texts.Count -eq 0) {
    return $false
  }

  $combined = Normalize-OpenXmlText -Text ($texts -join " ")
  if ($metadataParagraphCount -eq $texts.Count) {
    return $false
  }

  return ($combined.Length -ge 80)
}

function Normalize-TableRowLayout {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Row,

    [bool]$NormalizeCellMargins = $true
  )

  if (-not (Test-IsBodyTableRow -Row $Row)) {
    return $false
  }

  $trPr = $Row.SelectSingleNode("./w:trPr", $script:namespaceManager)
  if ($null -ne $trPr) {
    [void](Remove-WordChildElement -Parent $trPr -LocalName "cantSplit")
    [void](Remove-WordChildElement -Parent $trPr -LocalName "trHeight")
    if (-not $trPr.HasChildNodes) {
      $Row.RemoveChild($trPr) | Out-Null
    }
  }

  foreach ($cell in @($Row.SelectNodes("./w:tc", $script:namespaceManager))) {
    Set-CellVerticalAlignmentTop -Cell $cell
    if ($NormalizeCellMargins) {
      Set-CellMargins -Cell $cell
    }
  }

  return $true
}

function Get-TopLevelTableProfileSignal {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Table
  )

  $nonEmptyTextCount = 0
  $metadataParagraphCount = 0
  $rowCount = 0
  $hasSectionHeading = $false
  $hasDrawing = ($null -ne $Table.SelectSingleNode(".//w:drawing", $script:namespaceManager))

  foreach ($row in @($Table.SelectNodes("./w:tr", $script:namespaceManager))) {
    $rowCount++
    foreach ($paragraph in @($row.SelectNodes("./w:tc//w:p", $script:namespaceManager))) {
      $text = Get-ParagraphText -Paragraph $paragraph -NamespaceManager $script:namespaceManager
      if ([string]::IsNullOrWhiteSpace($text)) {
        continue
      }

      $nonEmptyTextCount++
      if (Test-IsMetadataParagraph -Text $text) {
        $metadataParagraphCount++
      }
      if (Test-IsSectionHeading -Text $text) {
        $hasSectionHeading = $true
      }
    }
  }

  return [pscustomobject]@{
    RowCount = $rowCount
    NonEmptyTextCount = $nonEmptyTextCount
    MetadataParagraphCount = $metadataParagraphCount
    HasSectionHeading = $hasSectionHeading
    HasDrawing = $hasDrawing
  }
}

function Test-IsCourseDesignCoverTitleText {
  param(
    [AllowNull()]
    [string]$Text
  )

  $compact = Normalize-OpenXmlText -Text $Text
  return ($compact -match '课程设计报告$')
}

function Test-IsCourseDesignCoverSubtitleText {
  param(
    [AllowNull()]
    [string]$Text
  )

  $compact = Normalize-OpenXmlText -Text $Text
  return ($compact -match '学年.*学期')
}

function Get-CourseDesignCoverElements {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Body
  )

  $titleParagraph = $null
  $subtitleParagraph = $null
  $coverTable = $null

  foreach ($child in @($Body.ChildNodes)) {
    if ($child.LocalName -eq "p") {
      $text = Get-ParagraphText -Paragraph $child -NamespaceManager $script:namespaceManager
      if ([string]::IsNullOrWhiteSpace($text)) {
        continue
      }

      if ($null -eq $titleParagraph -and (Test-IsCourseDesignCoverTitleText -Text $text)) {
        $titleParagraph = $child
        continue
      }

      if ($null -ne $titleParagraph -and $null -eq $subtitleParagraph -and (Test-IsCourseDesignCoverSubtitleText -Text $text)) {
        $subtitleParagraph = $child
        continue
      }

      if (Test-IsSectionHeading -Text $text) {
        break
      }

      continue
    }

    if ($child.LocalName -eq "tbl" -and $null -eq $coverTable) {
      $tableSignal = Get-TopLevelTableProfileSignal -Table $child
      if (
        -not $tableSignal.HasDrawing -and
        -not $tableSignal.HasSectionHeading -and
        $tableSignal.RowCount -ge 4 -and
        $tableSignal.MetadataParagraphCount -ge 4
      ) {
        $coverTable = $child
      }
      continue
    }
  }

  return [pscustomobject]@{
    TitleParagraph = $titleParagraph
    SubtitleParagraph = $subtitleParagraph
    CoverTable = $coverTable
  }
}

function Apply-CourseDesignCoverStyles {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Body
  )

  $coverElements = Get-CourseDesignCoverElements -Body $Body

  if ($null -ne $coverElements.TitleParagraph) {
    Set-ParagraphJustification -Paragraph $coverElements.TitleParagraph -Value "center"
    Set-ParagraphIndent -Paragraph $coverElements.TitleParagraph -FirstLine 0
    Set-ParagraphSpacing -Paragraph $coverElements.TitleParagraph -Before 0 -After 120 -Line 320
    Set-ParagraphBold -Paragraph $coverElements.TitleParagraph
    Set-RunTypography -Paragraph $coverElements.TitleParagraph -FontName "黑体" -SizeHalfPoints 52
  }

  if ($null -ne $coverElements.SubtitleParagraph) {
    Set-ParagraphJustification -Paragraph $coverElements.SubtitleParagraph -Value "center"
    Set-ParagraphIndent -Paragraph $coverElements.SubtitleParagraph -FirstLine 0
    Set-ParagraphSpacing -Paragraph $coverElements.SubtitleParagraph -Before 0 -After 160 -Line 320
    Set-ParagraphPagination -Paragraph $coverElements.SubtitleParagraph -KeepNext $false -KeepLines $false
    Set-RunTypography -Paragraph $coverElements.SubtitleParagraph -FontName "楷体_GB2312" -SizeHalfPoints 30
  }

  if ($null -eq $coverElements.CoverTable) {
    return
  }

  $rows = @($coverElements.CoverTable.SelectNodes("./w:tr", $script:namespaceManager))
  for ($rowIndex = 0; $rowIndex -lt $rows.Count; $rowIndex++) {
    $cells = @($rows[$rowIndex].SelectNodes("./w:tc", $script:namespaceManager))
    for ($cellIndex = 0; $cellIndex -lt $cells.Count; $cellIndex++) {
      foreach ($paragraph in @($cells[$cellIndex].SelectNodes(".//w:p", $script:namespaceManager))) {
        $text = Get-ParagraphText -Paragraph $paragraph -NamespaceManager $script:namespaceManager
        if ([string]::IsNullOrWhiteSpace($text)) {
          continue
        }

        Set-ParagraphJustification -Paragraph $paragraph -Value "center"
        Set-ParagraphIndent -Paragraph $paragraph -FirstLine 0
        Set-ParagraphSpacing -Paragraph $paragraph -Before 0 -After 0 -Line 320
        Set-ParagraphPagination -Paragraph $paragraph -KeepNext $false -KeepLines $false

        if ($cellIndex -eq 0) {
          Set-ParagraphBold -Paragraph $paragraph
          Set-RunTypography -Paragraph $paragraph -FontName "黑体" -SizeHalfPoints 32
        } else {
          Set-RunTypography -Paragraph $paragraph -FontName "楷体_GB2312" -SizeHalfPoints 32
        }
      }
    }
  }
}

function Resolve-StyleProfileDecision {
  param(
    [Parameter(Mandatory = $true)]
    [string]$RequestedProfile,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Body,

    [Parameter(Mandatory = $true)]
    [psobject]$ReportProfile
  )

  if ($RequestedProfile -ne "auto") {
    return [pscustomobject]@{
      RequestedProfile = $RequestedProfile
      ResolvedProfile = $RequestedProfile
      Reason = "Profile explicitly requested."
    }
  }

  $profileDefaultStyleProfile = Get-ReportProfileDefaultStyleProfile -Profile $ReportProfile
  if ($profileDefaultStyleProfile -ne "auto") {
    return [pscustomobject]@{
      RequestedProfile = $RequestedProfile
      ResolvedProfile = $profileDefaultStyleProfile
      Reason = "Resolved from report profile defaultStyleProfile."
    }
  }

  $metadataParagraphCountBeforeHeading = 0
  $firstPreHeadingTableSignal = $null

  foreach ($child in @($Body.ChildNodes)) {
    if ($child.LocalName -eq "p") {
      $text = Get-ParagraphText -Paragraph $child -NamespaceManager $script:namespaceManager
      if ([string]::IsNullOrWhiteSpace($text) -or (Test-IsTitleParagraph -Text $text)) {
        continue
      }

      if (Test-IsSectionHeading -Text $text) {
        break
      }

      if (Test-IsMetadataParagraph -Text $text) {
        $metadataParagraphCountBeforeHeading++
      }

      continue
    }

    if ($child.LocalName -eq "tbl" -and $null -eq $firstPreHeadingTableSignal) {
      $firstPreHeadingTableSignal = Get-TopLevelTableProfileSignal -Table $child
      continue
    }
  }

  if ($null -ne $firstPreHeadingTableSignal -and -not $firstPreHeadingTableSignal.HasDrawing -and -not $firstPreHeadingTableSignal.HasSectionHeading -and $firstPreHeadingTableSignal.RowCount -ge 4 -and $firstPreHeadingTableSignal.MetadataParagraphCount -ge 4) {
    return [pscustomobject]@{
      RequestedProfile = $RequestedProfile
      ResolvedProfile = "compact"
      Reason = "Detected a cover-style metadata table before the first report section."
    }
  }

  if ($null -eq $firstPreHeadingTableSignal -and $metadataParagraphCountBeforeHeading -ge 3) {
    return [pscustomobject]@{
      RequestedProfile = $RequestedProfile
      ResolvedProfile = "school"
      Reason = "Detected a paragraph-based cover area with multiple metadata lines before the first report section."
    }
  }

  return [pscustomobject]@{
    RequestedProfile = $RequestedProfile
    ResolvedProfile = "default"
    Reason = "No auto-profile cover-layout heuristic matched, so the default profile was used."
  }
}

function Write-OpenXmlPackage {
  param(
    [Parameter(Mandatory = $true)]
    [string]$SourceDirectory,

    [Parameter(Mandatory = $true)]
    [string]$DestinationPath
  )

  if (Test-Path -LiteralPath $DestinationPath) {
    Remove-Item -LiteralPath $DestinationPath -Force
  }

  $archive = [System.IO.Compression.ZipFile]::Open($DestinationPath, [System.IO.Compression.ZipArchiveMode]::Create)
  try {
    foreach ($file in Get-ChildItem -LiteralPath $SourceDirectory -Recurse -File) {
      $relativePath = $file.FullName.Substring($SourceDirectory.Length).TrimStart('\', '/') -replace '\\', '/'
      $entry = $archive.CreateEntry($relativePath)
      $entryStream = $entry.Open()
      try {
        $fileStream = [System.IO.File]::OpenRead($file.FullName)
        try {
          $fileStream.CopyTo($entryStream)
        } finally {
          $fileStream.Dispose()
        }
      } finally {
        $entryStream.Dispose()
      }
    }
  } finally {
    $archive.Dispose()
  }
}

$resolvedDocxPath = (Resolve-Path -LiteralPath $DocxPath).Path
if ([System.IO.Path]::GetExtension($resolvedDocxPath).ToLowerInvariant() -ne ".docx") {
  throw "Only .docx files are supported: $resolvedDocxPath"
}

if ([string]::IsNullOrWhiteSpace($OutPath)) {
  $directory = Split-Path -Parent $resolvedDocxPath
  $fileName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedDocxPath)
  $OutPath = Join-Path $directory ($fileName + ".styled.docx")
}

$resolvedOutPath = [System.IO.Path]::GetFullPath($OutPath)
if ((-not $Overwrite) -and (Test-Path -LiteralPath $resolvedOutPath)) {
  throw "Output file already exists: $resolvedOutPath. Re-run with -Overwrite to replace it."
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("openclaw-docx-style-" + [System.Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

try {
  [System.IO.Compression.ZipFile]::ExtractToDirectory($resolvedDocxPath, $tempRoot)

  $documentXmlPath = Join-Path $tempRoot "word\document.xml"
  if (-not (Test-Path -LiteralPath $documentXmlPath)) {
    throw "word/document.xml was not found in $resolvedDocxPath"
  }

  [xml]$documentXml = [System.IO.File]::ReadAllText($documentXmlPath, (New-Object System.Text.UTF8Encoding($false)))
  $script:namespaceManager = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
  $script:namespaceManager.AddNamespace("w", $wordNamespace)

  $body = $documentXml.SelectSingleNode("/w:document/w:body", $script:namespaceManager)
  if ($null -eq $body) {
    throw "Could not locate /w:document/w:body in $resolvedDocxPath"
  }

  $fileProfile = $null
  if (-not [string]::IsNullOrWhiteSpace($ProfilePath)) {
    $fileProfile = Read-StyleProfileFile -Path $ProfilePath
  }

  $requestedProfile = if ($PSBoundParameters.ContainsKey("Profile")) {
    $Profile
  } elseif ($null -ne $fileProfile -and -not [string]::IsNullOrWhiteSpace([string]$fileProfile.baseProfile)) {
    [string]$fileProfile.baseProfile
  } else {
    $Profile
  }

  $profileDecision = Resolve-StyleProfileDecision -RequestedProfile $requestedProfile -Body $body -ReportProfile $reportProfile
  $profileSettings = Get-StyleProfileSettings -ProfileName $profileDecision.ResolvedProfile
  $effectiveStyleSettings = [ordered]@{}
  foreach ($settingName in (Get-StyleSettingNames)) {
    $effectiveStyleSettings[$settingName] = if ($PSBoundParameters.ContainsKey($settingName)) {
      Get-Variable -Name $settingName -ValueOnly
    } elseif ($null -ne $fileProfile -and $null -ne $fileProfile.overrides.PSObject.Properties[$settingName]) {
      $fileProfile.overrides.$settingName
    } else {
      $profileSettings[$settingName]
    }
  }
  $styleSettings = [pscustomobject]$effectiveStyleSettings
  $isCourseDesignReport = [string]::Equals([string]$reportProfile.name, "course-design-report", [System.StringComparison]::OrdinalIgnoreCase)
  if ($isCourseDesignReport) {
    $styleSettings.BodyLineTwips = 320
    $styleSettings.BodyFontHalfPoints = 21
    $styleSettings.MetadataFontHalfPoints = 21
    $styleSettings.ListFontHalfPoints = 21
    $styleSettings.CaptionFontHalfPoints = 18
    $styleSettings.HeadingFontHalfPoints = 24
    $styleSettings.HeadingBeforeTwips = 80
    $styleSettings.HeadingAfterTwips = 40
    $styleSettings.CaptionAfterTwips = 40
  }
  $useTemplateLikeCompactStyle = ([string]$profileDecision.ResolvedProfile -eq "compact")
  $usePaginationHints = (-not $useTemplateLikeCompactStyle)

  $paragraphs = @($documentXml.SelectNodes("//w:p", $script:namespaceManager))
  $styledTitleCount = 0
  $styledHeadingCount = 0
  $styledBodyCount = 0
  $styledCaptionCount = 0
  $styledImageCount = 0
  $styledMetadataCount = 0
  $styledListCount = 0
  $styledCommandCount = 0
  $styledCodeCount = 0
  $styledTableParagraphCount = 0
  $normalizedBodyRowCount = 0

  foreach ($paragraph in $paragraphs) {
    $text = Get-ParagraphText -Paragraph $paragraph -NamespaceManager $script:namespaceManager
    $hasDrawing = ($null -ne $paragraph.SelectSingleNode(".//w:drawing", $script:namespaceManager))
    $isInTable = Test-IsParagraphInTable -Paragraph $paragraph

    if ($hasDrawing) {
      Set-ParagraphJustification -Paragraph $paragraph -Value "center"
      Set-ParagraphIndent -Paragraph $paragraph -FirstLine 0
      Set-ParagraphSpacing -Paragraph $paragraph -Before $styleSettings.ImageBeforeTwips -After $styleSettings.ImageAfterTwips -Line $null
      Set-ParagraphPagination -Paragraph $paragraph -KeepNext $usePaginationHints -KeepLines $usePaginationHints
      $styledImageCount++
      if ($isInTable) { $styledTableParagraphCount++ }
      continue
    }

    if ([string]::IsNullOrWhiteSpace($text)) {
      continue
    }

    if (Test-IsTitleParagraph -Text $text) {
      Set-ParagraphJustification -Paragraph $paragraph -Value "center"
      Set-ParagraphIndent -Paragraph $paragraph -FirstLine 0
      Set-ParagraphSpacing -Paragraph $paragraph -Before 0 -After $styleSettings.TitleAfterTwips -Line $styleSettings.BodyLineTwips
      Set-ParagraphBold -Paragraph $paragraph
      if (-not $useTemplateLikeCompactStyle) {
        Set-RunTypography -Paragraph $paragraph -FontName "黑体" -SizeHalfPoints $styleSettings.TitleFontHalfPoints
      }
      Set-ParagraphPagination -Paragraph $paragraph -KeepNext $usePaginationHints -KeepLines $usePaginationHints
      $styledTitleCount++
      if ($isInTable) { $styledTableParagraphCount++ }
      continue
    }

    if (Test-IsCaptionParagraph -Text $text) {
      Set-ParagraphJustification -Paragraph $paragraph -Value "center"
      Set-ParagraphIndent -Paragraph $paragraph -FirstLine 0
      Set-ParagraphSpacing -Paragraph $paragraph -Before 0 -After $styleSettings.CaptionAfterTwips -Line $(if ($useTemplateLikeCompactStyle) { $null } else { $styleSettings.BodyLineTwips })
      if (-not $useTemplateLikeCompactStyle) {
        Set-RunTypography -Paragraph $paragraph -FontName "宋体" -SizeHalfPoints $styleSettings.CaptionFontHalfPoints
      }
      Set-ParagraphPagination -Paragraph $paragraph -KeepNext $false -KeepLines $false
      $styledCaptionCount++
      if ($isInTable) { $styledTableParagraphCount++ }
      continue
    }

    if (Test-IsSectionHeading -Text $text) {
      Set-ParagraphJustification -Paragraph $paragraph -Value "left"
      Set-ParagraphIndent -Paragraph $paragraph -FirstLine 0
      Set-ParagraphSpacing -Paragraph $paragraph -Before $styleSettings.HeadingBeforeTwips -After $styleSettings.HeadingAfterTwips -Line $styleSettings.BodyLineTwips
      Set-ParagraphBold -Paragraph $paragraph
      if (-not $useTemplateLikeCompactStyle) {
        Set-RunTypography -Paragraph $paragraph -FontName "黑体" -SizeHalfPoints $styleSettings.HeadingFontHalfPoints
      }
      Set-ParagraphPagination -Paragraph $paragraph -KeepNext $usePaginationHints -KeepLines $usePaginationHints
      $styledHeadingCount++
      if ($isInTable) { $styledTableParagraphCount++ }
      continue
    }

    if (Test-IsMetadataParagraph -Text $text) {
      Set-ParagraphJustification -Paragraph $paragraph -Value "left"
      Set-ParagraphIndent -Paragraph $paragraph -FirstLine 0
      Set-ParagraphSpacing -Paragraph $paragraph -Before 0 -After 0 -Line $styleSettings.BodyLineTwips
      if (-not $useTemplateLikeCompactStyle) {
        Set-RunTypography -Paragraph $paragraph -FontName "宋体" -SizeHalfPoints $styleSettings.MetadataFontHalfPoints
      }
      $styledMetadataCount++
      if ($isInTable) { $styledTableParagraphCount++ }
      continue
    }

    if ((-not $isInTable) -and (Test-IsCommandParagraph -Text $text)) {
      Set-ParagraphJustification -Paragraph $paragraph -Value "left"
      Set-ParagraphIndent -Paragraph $paragraph -FirstLine 0
      Set-ParagraphSpacing -Paragraph $paragraph -Before $styleSettings.CommandBeforeTwips -After $styleSettings.CommandAfterTwips -Line $styleSettings.CommandLineTwips
      if (-not $useTemplateLikeCompactStyle) {
        Set-RunTypography -Paragraph $paragraph -FontName "Consolas" -SizeHalfPoints $styleSettings.CommandFontHalfPoints
      }
      Set-ParagraphShading -Paragraph $paragraph -Fill "F2F2F2"
      Set-ParagraphPagination -Paragraph $paragraph -KeepNext $false -KeepLines $false
      $styledCommandCount++
      if ($isInTable) { $styledTableParagraphCount++ }
      continue
    }

    if ((-not $isInTable) -and (Test-IsCodeParagraph -Text $text)) {
      Set-ParagraphJustification -Paragraph $paragraph -Value "left"
      Set-ParagraphIndent -Paragraph $paragraph -FirstLine 0
      Set-ParagraphSpacing -Paragraph $paragraph -Before 0 -After 0 -Line $styleSettings.CommandLineTwips
      if (-not $useTemplateLikeCompactStyle) {
        Set-RunTypography -Paragraph $paragraph -FontName "Consolas" -SizeHalfPoints $styleSettings.CommandFontHalfPoints
      }
      Set-ParagraphShading -Paragraph $paragraph -Fill "F8F8F8"
      Set-ParagraphPagination -Paragraph $paragraph -KeepNext $false -KeepLines $false
      $styledCodeCount++
      if ($isInTable) { $styledTableParagraphCount++ }
      continue
    }

    if (Test-IsStepListParagraph -Text $text) {
      Set-ParagraphJustification -Paragraph $paragraph -Value "left"
      Set-ParagraphIndent -Paragraph $paragraph -FirstLine 0
      Set-ParagraphSpacing -Paragraph $paragraph -Before 0 -After $styleSettings.ListAfterTwips -Line $styleSettings.BodyLineTwips
      if (-not $useTemplateLikeCompactStyle) {
        Set-RunTypography -Paragraph $paragraph -FontName "宋体" -SizeHalfPoints $styleSettings.ListFontHalfPoints
      }
      Set-ParagraphPagination -Paragraph $paragraph -KeepNext $false -KeepLines $false
      $styledListCount++
      if ($isInTable) { $styledTableParagraphCount++ }
      continue
    }

    Set-ParagraphJustification -Paragraph $paragraph -Value "left"
    Set-ParagraphIndent -Paragraph $paragraph -FirstLine $(if ($isInTable) { 0 } else { $styleSettings.BodyFirstLineTwips })
    Set-ParagraphSpacing -Paragraph $paragraph -Before 0 -After $styleSettings.BodyAfterTwips -Line $styleSettings.BodyLineTwips
    if (-not $useTemplateLikeCompactStyle) {
      $bodyFontHalfPoints = if ($isCourseDesignReport -and $isInTable) { 18 } else { $styleSettings.BodyFontHalfPoints }
      Set-RunTypography -Paragraph $paragraph -FontName "宋体" -SizeHalfPoints $bodyFontHalfPoints
    }
    Set-ParagraphPagination -Paragraph $paragraph -KeepNext $false -KeepLines $false
    $styledBodyCount++
    if ($isInTable) { $styledTableParagraphCount++ }
  }

  if ($isCourseDesignReport) {
    Apply-CourseDesignCoverStyles -Body $body
  }

  $removedCourseDesignPlaceholderCount = 0
  if ($isCourseDesignReport) {
    $removedCourseDesignPlaceholderCount = Remove-CourseDesignDuplicatePlaceholderParagraphs -Body $body -NamespaceManager $script:namespaceManager
  }

  foreach ($row in @($documentXml.SelectNodes("//w:tbl[not(ancestor::w:tbl)]/w:tr", $script:namespaceManager))) {
    if (Normalize-TableRowLayout -Row $row -NormalizeCellMargins (-not $useTemplateLikeCompactStyle)) {
      $normalizedBodyRowCount++
    }
  }

  $removedTrailingEmptyParagraphCount = Remove-TrailingEmptyBodyParagraphs -Body $body -NamespaceManager $script:namespaceManager

  [System.IO.File]::WriteAllText($documentXmlPath, $documentXml.OuterXml, (New-Object System.Text.UTF8Encoding($false)))
  Write-OpenXmlPackage -SourceDirectory $tempRoot -DestinationPath $resolvedOutPath

  [pscustomobject]@{
    docxPath = $resolvedDocxPath
    outPath = $resolvedOutPath
    reportProfileName = [string]$reportProfile.name
    reportProfilePath = [string]$reportProfile.resolvedProfilePath
    profilePath = $(if ($null -ne $fileProfile) { [string]$fileProfile.path } else { $null })
    requestedProfile = $profileDecision.RequestedProfile
    resolvedProfile = $profileDecision.ResolvedProfile
    profileReason = $profileDecision.Reason
    styleProfile = $profileDecision.ResolvedProfile
    appliedSettings = $styleSettings
    styledTitleCount = $styledTitleCount
    styledHeadingCount = $styledHeadingCount
    styledBodyCount = $styledBodyCount
    styledCaptionCount = $styledCaptionCount
    styledImageCount = $styledImageCount
    styledMetadataCount = $styledMetadataCount
    styledListCount = $styledListCount
    styledCommandCount = $styledCommandCount
    styledCodeCount = $styledCodeCount
    styledTableParagraphCount = $styledTableParagraphCount
    normalizedBodyRowCount = $normalizedBodyRowCount
    removedCourseDesignPlaceholderCount = $removedCourseDesignPlaceholderCount
    removedTrailingEmptyParagraphCount = $removedTrailingEmptyParagraphCount
  }
} finally {
  if (Test-Path -LiteralPath $tempRoot) {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force
  }
}
