[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$DocxPath,

  [string]$MappingPath,

  [string]$ImagesJson,

  [string]$ReportProfileName = "experiment-report",

  [string]$ReportProfilePath,

  [string]$OutPath,

  [switch]$Overwrite
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot "report-profiles.ps1")

$script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
$relationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships"
$officeDocumentRelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
$drawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main"
$pictureNamespace = "http://schemas.openxmlformats.org/drawingml/2006/picture"
$wordprocessingDrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
$defaultImageWidthCm = 11.5
$sectionRules = @()
$sectionInputAliasLookup = @{}
$script:ImagePathProbeRoots = @()

function Get-WordIntAttribute {
  param(
    [AllowNull()]
    [System.Xml.XmlNode]$Node,

    [Parameter(Mandatory = $true)]
    [string]$Name,

    [int]$DefaultValue = 0
  )

  if ($null -eq $Node -or $null -eq $Node.Attributes) {
    return $DefaultValue
  }

  $attribute = $Node.Attributes.GetNamedItem($Name, $wordNamespace)
  if ($null -eq $attribute) {
    $attribute = $Node.Attributes.GetNamedItem($Name)
  }

  if ($null -eq $attribute -or [string]::IsNullOrWhiteSpace($attribute.Value) -or $attribute.Value -notmatch '^-?\d+$') {
    return $DefaultValue
  }

  return [int]$attribute.Value
}

function Get-DocumentBodyWidthCm {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$DocumentXml,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $defaultBodyWidthCm = 15.8
  $sectionProperties = @($DocumentXml.SelectNodes("//w:sectPr", $NamespaceManager))
  if ($sectionProperties.Count -eq 0) {
    return $defaultBodyWidthCm
  }

  $sectionProperty = $sectionProperties[$sectionProperties.Count - 1]
  $pageSize = $sectionProperty.SelectSingleNode("./w:pgSz", $NamespaceManager)
  $pageMargin = $sectionProperty.SelectSingleNode("./w:pgMar", $NamespaceManager)
  if ($null -eq $pageSize) {
    return $defaultBodyWidthCm
  }

  $pageWidthTwips = Get-WordIntAttribute -Node $pageSize -Name "w" -DefaultValue 0
  if ($pageWidthTwips -le 0) {
    return $defaultBodyWidthCm
  }

  $leftMarginTwips = Get-WordIntAttribute -Node $pageMargin -Name "left" -DefaultValue 1440
  $rightMarginTwips = Get-WordIntAttribute -Node $pageMargin -Name "right" -DefaultValue 1440
  $bodyWidthTwips = $pageWidthTwips - $leftMarginTwips - $rightMarginTwips
  if ($bodyWidthTwips -le 0) {
    return $defaultBodyWidthCm
  }

  return [Math]::Round(($bodyWidthTwips / 567.0), 2)
}

function Get-EffectiveRowImageWidthCm {
  param(
    [Parameter(Mandatory = $true)]
    [object]$ImageSpec,

    [Parameter(Mandatory = $true)]
    [int]$Columns,

    [Parameter(Mandatory = $true)]
    [double]$BodyWidthCm
  )

  if ($Columns -lt 2) {
    return [double]$ImageSpec.WidthCm
  }

  $cellMarginCm = 0.35
  $maxWidthCm = [Math]::Max(1.0, (($BodyWidthCm / $Columns) - $cellMarginCm))
  return [Math]::Round([Math]::Min([double]$ImageSpec.WidthCm, $maxWidthCm), 2)
}

function Add-UniqueProbeRoot {
  param(
    [AllowEmptyCollection()]
    [System.Collections.Generic.List[string]]$Roots,

    [AllowNull()]
    [string]$Path
  )

  if ([string]::IsNullOrWhiteSpace($Path)) {
    return
  }

  $candidate = $null
  try {
    $candidate = [System.IO.Path]::GetFullPath($Path)
  } catch {
    return
  }

  if ($Roots.Contains($candidate)) {
    return
  }

  $Roots.Add($candidate) | Out-Null
}

function Resolve-FileUriPath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  if ($Path -notmatch '^(?i)file://') {
    return $Path
  }

  try {
    return [System.Uri]$Path | ForEach-Object { $_.LocalPath }
  } catch {
    throw "Unsupported file URI path: $Path"
  }
}

function Get-ImagePathProbeRoots {
  param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath,

    [AllowNull()]
    [string]$MappingPath
  )

  $roots = New-Object 'System.Collections.Generic.List[string]'

  Add-UniqueProbeRoot -Roots $roots -Path (Get-Location).Path
  Add-UniqueProbeRoot -Roots $roots -Path (Split-Path -Parent $DocxPath)
  Add-UniqueProbeRoot -Roots $roots -Path $PSScriptRoot
  Add-UniqueProbeRoot -Roots $roots -Path $script:RepoRoot
  Add-UniqueProbeRoot -Roots $roots -Path (Split-Path -Parent $script:RepoRoot)

  if (-not [string]::IsNullOrWhiteSpace($MappingPath)) {
    Add-UniqueProbeRoot -Roots $roots -Path (Split-Path -Parent $MappingPath)
  }

  foreach ($envKey in @("OPENCLAW_WORKSPACE_DIR", "OPENCLAW_WORKSPACE", "OPENCLAW_SESSION_WORKSPACE")) {
    Add-UniqueProbeRoot -Roots $roots -Path ([System.Environment]::GetEnvironmentVariable($envKey))
  }

  return @($roots.ToArray())
}

function Resolve-ExistingImagePath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $rawPath = Resolve-FileUriPath -Path $Path
  if ([string]::IsNullOrWhiteSpace($rawPath)) {
    throw "Image path is empty."
  }

  $trimmedPath = $rawPath.Trim()
  $triedCandidates = New-Object 'System.Collections.Generic.List[string]'

  if ([System.IO.Path]::IsPathRooted($trimmedPath)) {
    $absoluteCandidate = [System.IO.Path]::GetFullPath($trimmedPath)
    $triedCandidates.Add($absoluteCandidate) | Out-Null
    if (Test-Path -LiteralPath $absoluteCandidate) {
      return (Resolve-Path -LiteralPath $absoluteCandidate).Path
    }
  } else {
    try {
      $resolvedDirect = (Resolve-Path -LiteralPath $trimmedPath).Path
      if (-not $triedCandidates.Contains($resolvedDirect)) {
        $triedCandidates.Add($resolvedDirect) | Out-Null
      }
      return $resolvedDirect
    } catch {
    }

    foreach ($root in $script:ImagePathProbeRoots) {
      $candidate = [System.IO.Path]::GetFullPath((Join-Path $root $trimmedPath))
      if (-not $triedCandidates.Contains($candidate)) {
        $triedCandidates.Add($candidate) | Out-Null
      }
      if (Test-Path -LiteralPath $candidate) {
        return (Resolve-Path -LiteralPath $candidate).Path
      }
    }
  }

  $triedText = if ($triedCandidates.Count -gt 0) { $triedCandidates -join "; " } else { "<none>" }
  throw "Image path was not found: $Path. Tried: $triedText"
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

function Normalize-TargetSelector {
  param(
    [AllowNull()]
    [string]$Selector
  )

  if ([string]::IsNullOrWhiteSpace($Selector)) {
    return $null
  }

  $trimmed = $Selector.Trim()
  if ($trimmed -match '^(?i)(anchor|location|target)\s*:\s*(?<value>P\d+|T\d+R\d+C\d+)\s*$') {
    return ("anchor:{0}" -f $matches["value"].ToUpperInvariant())
  }

  if ($trimmed -match '^(?i)(section|heading)\s*:\s*(?<value>.+)$') {
    return ("section:{0}" -f $matches["value"].Trim())
  }

  if ($trimmed -match '^(?i:P\d+|T\d+R\d+C\d+)$') {
    return $trimmed.ToUpperInvariant()
  }

  return $trimmed
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

function ConvertTo-ObjectArray {
  param(
    [AllowNull()]
    [object]$Value
  )

  if ($null -eq $Value) {
    return @()
  }

  if (($Value -is [System.Collections.IEnumerable]) -and ($Value -isnot [string])) {
    return @($Value)
  }

  return @($Value)
}

function Initialize-SectionRules {
  param(
    [Parameter(Mandatory = $true)]
    [psobject]$ReportProfile
  )

  $script:sectionRules = @(Get-ReportProfileSectionRules -Profile $ReportProfile)

  $script:sectionInputAliasLookup = @{}
  foreach ($rule in $script:sectionRules) {
    foreach ($alias in @($rule.inputAliases + $rule.id)) {
      $normalizedAlias = Normalize-FieldKey -Text $alias
      if (-not [string]::IsNullOrWhiteSpace($normalizedAlias)) {
        $script:sectionInputAliasLookup[$normalizedAlias] = $rule.id
      }
    }
  }
}

function Get-ImageMappingRootObject {
  param(
    [AllowNull()]
    [string]$PathToJson,

    [AllowNull()]
    [string]$InlineJson
  )

  if ([string]::IsNullOrWhiteSpace($PathToJson) -eq [string]::IsNullOrWhiteSpace($InlineJson)) {
    throw "Provide exactly one of -MappingPath or -ImagesJson."
  }

  if (-not [string]::IsNullOrWhiteSpace($PathToJson)) {
    $resolvedPath = Resolve-Path -LiteralPath $PathToJson
    $rootObject = (Get-Content -LiteralPath $resolvedPath.Path -Raw -Encoding UTF8) | ConvertFrom-Json
  } else {
    $resolvedPath = $null
    $rootObject = $InlineJson | ConvertFrom-Json
  }

  if ($null -eq $rootObject) {
    throw "Image mapping JSON is empty."
  }

  return [pscustomobject]@{
    RootObject = $rootObject
    ResolvedMappingPath = $(if ($null -ne $resolvedPath) { $resolvedPath.Path } else { $null })
  }
}

function Resolve-ImageMappingReportProfile {
  param(
    [AllowNull()]
    [object]$RootObject,

    [AllowNull()]
    [string]$ProfileName,

    [AllowNull()]
    [string]$ProfilePath
  )

  $summaryProfileName = $null
  $summaryProfilePath = $null
  if ($null -ne $RootObject -and ($RootObject -isnot [System.Collections.IEnumerable] -or $RootObject -is [string])) {
    $rootTable = ConvertTo-PlainHashtable -InputObject $RootObject
    if ($rootTable.ContainsKey("summary") -and $null -ne $rootTable["summary"]) {
      $summaryTable = ConvertTo-PlainHashtable -InputObject $rootTable["summary"]
      if ($summaryTable.ContainsKey("reportProfileName") -and -not [string]::IsNullOrWhiteSpace([string]$summaryTable["reportProfileName"])) {
        $summaryProfileName = [string]$summaryTable["reportProfileName"]
      }
      if ($summaryTable.ContainsKey("reportProfilePath") -and -not [string]::IsNullOrWhiteSpace([string]$summaryTable["reportProfilePath"])) {
        $summaryProfilePath = [string]$summaryTable["reportProfilePath"]
      }
    }
  }

  $effectiveProfilePath = if (-not [string]::IsNullOrWhiteSpace($ProfilePath)) {
    $ProfilePath
  } elseif (-not [string]::IsNullOrWhiteSpace($summaryProfilePath)) {
    $summaryProfilePath
  } else {
    $null
  }

  $effectiveProfileName = $ProfileName
  if ([string]::IsNullOrWhiteSpace($effectiveProfileName) -or ($effectiveProfileName -eq "experiment-report" -and -not [string]::IsNullOrWhiteSpace($summaryProfileName))) {
    $effectiveProfileName = $summaryProfileName
  }
  if ([string]::IsNullOrWhiteSpace($effectiveProfileName)) {
    $effectiveProfileName = "experiment-report"
  }

  return Get-ReportProfile -ProfileName $effectiveProfileName -ProfilePath $effectiveProfilePath -RepoRoot $script:RepoRoot
}

function Resolve-ImageLayoutSpec {
  param(
    [AllowNull()]
    [object]$LayoutValue,

    [Parameter(Mandatory = $true)]
    [hashtable]$ItemTable
  )

  $mode = $null
  $columns = $null
  $group = $null
  $groupAnchor = $null

  if ($null -ne $LayoutValue) {
    if ($LayoutValue -is [string]) {
      $mode = ([string]$LayoutValue).Trim()
    } else {
      $layoutTable = ConvertTo-PlainHashtable -InputObject $LayoutValue
      foreach ($key in @("mode", "type")) {
        if ($layoutTable.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$layoutTable[$key])) {
          $mode = ([string]$layoutTable[$key]).Trim()
          break
        }
      }
      foreach ($key in @("columns", "cols")) {
        if ($layoutTable.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$layoutTable[$key])) {
          $columns = [int]$layoutTable[$key]
          break
        }
      }
      foreach ($key in @("group", "groupId", "rowGroup")) {
        if ($layoutTable.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$layoutTable[$key])) {
          $group = ([string]$layoutTable[$key]).Trim()
          break
        }
      }
      foreach ($key in @("groupAnchor", "anchor", "target")) {
        if ($layoutTable.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$layoutTable[$key])) {
          $groupAnchor = Normalize-TargetSelector -Selector ([string]$layoutTable[$key])
          break
        }
      }
    }
  }

  foreach ($key in @("layoutMode", "mode")) {
    if ([string]::IsNullOrWhiteSpace($mode) -and $ItemTable.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$ItemTable[$key])) {
      $mode = ([string]$ItemTable[$key]).Trim()
      break
    }
  }

  foreach ($key in @("columns", "cols")) {
    if ($null -eq $columns -and $ItemTable.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$ItemTable[$key])) {
      $columns = [int]$ItemTable[$key]
      break
    }
  }

  foreach ($key in @("group", "groupId", "rowGroup")) {
    if ([string]::IsNullOrWhiteSpace($group) -and $ItemTable.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$ItemTable[$key])) {
      $group = ([string]$ItemTable[$key]).Trim()
      break
    }
  }

  foreach ($key in @("groupAnchor", "groupTarget", "layoutAnchor")) {
    if ([string]::IsNullOrWhiteSpace($groupAnchor) -and $ItemTable.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$ItemTable[$key])) {
      $groupAnchor = Normalize-TargetSelector -Selector ([string]$ItemTable[$key])
      break
    }
  }

  if ([string]::IsNullOrWhiteSpace($mode) -and $null -eq $columns -and [string]::IsNullOrWhiteSpace($group) -and [string]::IsNullOrWhiteSpace($groupAnchor)) {
    return $null
  }

  $normalizedMode = switch -Regex (([string]$mode).Trim().ToLowerInvariant()) {
    '^$' { "row" }
    '^(row|grid|gallery|sidebyside|side-by-side|twoup|two-up)$' { "row" }
    '^(stack|column)$' { "stack" }
    default { throw "Unsupported image layout mode: $mode" }
  }

  if ($normalizedMode -eq "row") {
    if ($null -eq $columns) {
      $columns = 2
    }
    if ($columns -lt 2) {
      throw "Row image layout must use at least 2 columns."
    }
  } else {
    $columns = $null
  }

  return [pscustomobject]@{
    Mode = $normalizedMode
    Columns = $columns
    Group = $(if ([string]::IsNullOrWhiteSpace($group)) { $null } else { $group })
    GroupAnchor = $(if ([string]::IsNullOrWhiteSpace($groupAnchor)) { $null } else { $groupAnchor })
  }
}

function Get-HeadingPattern {
  param(
    [Parameter(Mandatory = $true)]
    [string[]]$Aliases
  )

  $escapedAliases = foreach ($alias in $Aliases) {
    [regex]::Escape($alias)
  }

  return '^(?:#{1,6}\s*)?(?:第?[0-9一二三四五六七八九十]+(?:章|节)?[.\)\u3001]?\s*)?(?:' + ($escapedAliases -join '|') + ')\s*(?:[:\uFF1A])?$'
}

function Resolve-SectionId {
  param(
    [AllowNull()]
    [string]$SectionName
  )

  $normalized = Normalize-FieldKey -Text $SectionName
  if ([string]::IsNullOrWhiteSpace($normalized)) {
    return $null
  }

  if ($sectionInputAliasLookup.ContainsKey($normalized)) {
    return $sectionInputAliasLookup[$normalized]
  }

  return $null
}

function Resolve-SectionRuleFromHeading {
  param(
    [AllowNull()]
    [string]$HeadingText
  )

  if ([string]::IsNullOrWhiteSpace($HeadingText)) {
    return $null
  }

  foreach ($rule in $sectionRules) {
    if ($HeadingText -match (Get-HeadingPattern -Aliases $rule.headingAliases)) {
      return $rule
    }
  }

  return $null
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

function Get-ImageMappingItems {
  param(
    [Parameter(Mandatory = $true)]
    [object]$RootObject
  )

  if ($RootObject -is [System.Collections.IEnumerable] -and $RootObject -isnot [string]) {
    return @(ConvertTo-ObjectArray -Value $RootObject)
  }

  $rootTable = ConvertTo-PlainHashtable -InputObject $RootObject
  if ($rootTable.ContainsKey("images")) {
    return @(ConvertTo-ObjectArray -Value $rootTable["images"])
  }

  return @(ConvertTo-ObjectArray -Value $RootObject)
}

function Resolve-ImageSpecification {
  param(
    [Parameter(Mandatory = $true)]
    [object]$Item
  )

  $itemTable = ConvertTo-PlainHashtable -InputObject $Item
  $anchor = $null
  foreach ($key in @("anchor", "location", "target")) {
    if ($itemTable.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$itemTable[$key])) {
      $anchor = Normalize-TargetSelector -Selector ([string]$itemTable[$key])
      break
    }
  }

  $sectionName = $null
  foreach ($key in @("section", "heading", "sectionName")) {
    if ($itemTable.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$itemTable[$key])) {
      $sectionName = ([string]$itemTable[$key]).Trim()
      break
    }
  }

  if ([string]::IsNullOrWhiteSpace($anchor) -and [string]::IsNullOrWhiteSpace($sectionName)) {
    throw "Each image mapping item must include anchor/location/target or section/heading."
  }

  $imagePath = $null
  foreach ($key in @("path", "imagePath", "file")) {
    if ($itemTable.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$itemTable[$key])) {
      $imagePath = Resolve-ExistingImagePath -Path ([string]$itemTable[$key])
      break
    }
  }

  if ([string]::IsNullOrWhiteSpace($imagePath)) {
    throw "Each image mapping item must include path, imagePath, or file."
  }

  $caption = $null
  foreach ($key in @("caption", "title", "figureCaption")) {
    if ($itemTable.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$itemTable[$key])) {
      $caption = ([string]$itemTable[$key]).Trim()
      break
    }
  }

  $widthCm = $defaultImageWidthCm
  foreach ($key in @("widthCm", "width")) {
    if ($itemTable.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$itemTable[$key])) {
      $widthCm = [double]$itemTable[$key]
      break
    }
  }

  $layout = Resolve-ImageLayoutSpec -LayoutValue $(if ($itemTable.ContainsKey("layout")) { $itemTable["layout"] } else { $null }) -ItemTable $itemTable

  return [pscustomobject]@{
    Anchor = $anchor
    SectionName = $sectionName
    ImagePath = $imagePath
    Caption = $caption
    WidthCm = $widthCm
    Layout = $layout
  }
}

function Get-ImageContentType {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  switch ([System.IO.Path]::GetExtension($Path).ToLowerInvariant()) {
    ".png" { return "image/png" }
    ".jpg" { return "image/jpeg" }
    ".jpeg" { return "image/jpeg" }
    ".gif" { return "image/gif" }
    ".bmp" { return "image/bmp" }
    default { throw "Unsupported image format: $Path" }
  }
}

function Get-ImageSizeEmu {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [double]$WidthCm
  )

  if ($WidthCm -le 0) {
    throw "Image width must be positive: $WidthCm"
  }

  $image = [System.Drawing.Image]::FromFile($Path)
  try {
    if ($image.Width -le 0 -or $image.Height -le 0) {
      throw "Invalid image dimensions: $Path"
    }

    $widthEmu = [int64][Math]::Round($WidthCm * 360000.0)
    $heightEmu = [int64][Math]::Round($widthEmu * ($image.Height / [double]$image.Width))

    return [pscustomobject]@{
      WidthEmu = $widthEmu
      HeightEmu = $heightEmu
    }
  } finally {
    $image.Dispose()
  }
}

function Get-NextMediaIndex {
  param(
    [Parameter(Mandatory = $true)]
    [string]$MediaDirectory
  )

  $maxIndex = 0
  if (Test-Path -LiteralPath $MediaDirectory) {
    foreach ($file in Get-ChildItem -LiteralPath $MediaDirectory -File) {
      if ($file.BaseName -match '^image(?<index>\d+)$') {
        $maxIndex = [Math]::Max($maxIndex, [int]$matches["index"])
      }
    }
  }

  return ($maxIndex + 1)
}

function Get-NextRelationshipId {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$RelationshipsXml
  )

  $maxId = 0
  foreach ($relationshipNode in @($RelationshipsXml.SelectNodes("/*[local-name()='Relationships']/*[local-name()='Relationship']"))) {
    if ($null -eq $relationshipNode) {
      continue
    }

    $id = [string]$relationshipNode.Id
    if ($id -match '^rId(?<index>\d+)$') {
      $maxId = [Math]::Max($maxId, [int]$matches["index"])
    }
  }

  return ($maxId + 1)
}

function Get-NextDocPrId {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$DocumentXml,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $maxId = 0
  foreach ($node in @($DocumentXml.SelectNodes("//wp:docPr", $NamespaceManager))) {
    if ($null -eq $node) {
      continue
    }

    $idAttribute = $node.Attributes.GetNamedItem("id")
    if ($null -ne $idAttribute -and $idAttribute.Value -match '^\d+$') {
      $maxId = [Math]::Max($maxId, [int]$idAttribute.Value)
    }
  }

  return ($maxId + 1)
}

function Ensure-ContentTypeDefault {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$ContentTypesXml,

    [Parameter(Mandatory = $true)]
    [string]$Extension,

    [Parameter(Mandatory = $true)]
    [string]$ContentType
  )

  foreach ($defaultNode in @($ContentTypesXml.Types.Default)) {
    if ($null -ne $defaultNode -and [string]$defaultNode.Extension -eq $Extension) {
      return
    }
  }

  $defaultElement = $ContentTypesXml.CreateElement("Default", $ContentTypesXml.Types.NamespaceURI)
  $defaultElement.SetAttribute("Extension", $Extension)
  $defaultElement.SetAttribute("ContentType", $ContentType)
  $ContentTypesXml.Types.AppendChild($defaultElement) | Out-Null
}

function Add-ImageRelationship {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$RelationshipsXml,

    [Parameter(Mandatory = $true)]
    [string]$RelationshipId,

    [Parameter(Mandatory = $true)]
    [string]$Target
  )

  $relationshipElement = $RelationshipsXml.CreateElement("Relationship", $RelationshipsXml.Relationships.NamespaceURI)
  $relationshipElement.SetAttribute("Id", $RelationshipId)
  $relationshipElement.SetAttribute("Type", "$officeDocumentRelationshipNamespace/image")
  $relationshipElement.SetAttribute("Target", $Target)
  $RelationshipsXml.Relationships.AppendChild($relationshipElement) | Out-Null
}

function New-XmlNodeFromString {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$DocumentXml,

    [Parameter(Mandatory = $true)]
    [string]$XmlText
  )

  $fragment = $DocumentXml.CreateDocumentFragment()
  $fragment.InnerXml = $XmlText
  return ,$fragment.FirstChild
}

function Get-SingleXmlNode {
  param(
    [Parameter(Mandatory = $true)]
    [object]$NodeLike
  )

  if ($NodeLike -is [System.Xml.XmlNode]) {
    return $NodeLike
  }

  $candidates = @($NodeLike | Where-Object { $_ -is [System.Xml.XmlNode] })
  if ($candidates.Count -eq 0) {
    throw "Expected an XmlNode-compatible value."
  }

  return $candidates[0]
}

function New-ImageParagraph {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$DocumentXml,

    [Parameter(Mandatory = $true)]
    [string]$RelationshipId,

    [Parameter(Mandatory = $true)]
    [int]$DocPrId,

    [Parameter(Mandatory = $true)]
    [string]$Name,

    [Parameter(Mandatory = $true)]
    [int64]$WidthEmu,

    [Parameter(Mandatory = $true)]
    [int64]$HeightEmu
  )

  $escapedName = [System.Security.SecurityElement]::Escape($Name)
  $paragraphXml = @"
<w:p xmlns:w="$wordNamespace" xmlns:r="$officeDocumentRelationshipNamespace" xmlns:wp="$wordprocessingDrawingNamespace" xmlns:a="$drawingNamespace" xmlns:pic="$pictureNamespace">
  <w:pPr>
    <w:jc w:val="center"/>
    <w:spacing w:before="80" w:after="80"/>
  </w:pPr>
  <w:r>
    <w:drawing>
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="$WidthEmu" cy="$HeightEmu"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <wp:docPr id="$DocPrId" name="$escapedName"/>
        <wp:cNvGraphicFramePr/>
        <a:graphic>
          <a:graphicData uri="$pictureNamespace">
            <pic:pic>
              <pic:nvPicPr>
                <pic:cNvPr id="0" name="$escapedName"/>
                <pic:cNvPicPr/>
              </pic:nvPicPr>
              <pic:blipFill>
                <a:blip r:embed="$RelationshipId"/>
                <a:stretch>
                  <a:fillRect/>
                </a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="$WidthEmu" cy="$HeightEmu"/>
                </a:xfrm>
                <a:prstGeom prst="rect">
                  <a:avLst/>
                </a:prstGeom>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>
"@

  return ,(New-XmlNodeFromString -DocumentXml $DocumentXml -XmlText $paragraphXml)
}

function New-CaptionParagraph {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$DocumentXml,

    [Parameter(Mandatory = $true)]
    [string]$CaptionText
  )

  $escapedCaption = [System.Security.SecurityElement]::Escape($CaptionText)
  $paragraphXml = @"
<w:p xmlns:w="$wordNamespace">
  <w:pPr>
    <w:jc w:val="center"/>
    <w:spacing w:after="80"/>
  </w:pPr>
  <w:r>
    <w:t>$escapedCaption</w:t>
  </w:r>
</w:p>
"@

  return ,(New-XmlNodeFromString -DocumentXml $DocumentXml -XmlText $paragraphXml)
}

function New-EmptyParagraph {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$DocumentXml
  )

  $paragraphXml = @"
<w:p xmlns:w="$wordNamespace"/>
"@

  return ,(New-XmlNodeFromString -DocumentXml $DocumentXml -XmlText $paragraphXml)
}

function New-ImageTable {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$DocumentXml,

    [Parameter(Mandatory = $true)]
    [object[]]$CellEntries,

    [Parameter(Mandatory = $true)]
    [int]$Columns
  )

  if ($Columns -lt 2) {
    throw "Image row layout requires at least 2 columns."
  }

  $tableXml = @"
<w:tbl xmlns:w="$wordNamespace">
  <w:tblPr>
    <w:tblW w:w="5000" w:type="pct"/>
    <w:jc w:val="center"/>
    <w:tblLayout w:type="fixed"/>
    <w:tblBorders>
      <w:top w:val="nil"/>
      <w:left w:val="nil"/>
      <w:bottom w:val="nil"/>
      <w:right w:val="nil"/>
      <w:insideH w:val="nil"/>
      <w:insideV w:val="nil"/>
    </w:tblBorders>
    <w:tblCellMar>
      <w:left w:w="80" w:type="dxa"/>
      <w:right w:w="80" w:type="dxa"/>
    </w:tblCellMar>
  </w:tblPr>
  <w:tblGrid/>
</w:tbl>
"@

  $tableNode = New-XmlNodeFromString -DocumentXml $DocumentXml -XmlText $tableXml
  $gridNode = $tableNode.SelectSingleNode("./*[local-name()='tblGrid']")
  $cellWidthPct = [int][Math]::Floor(5000 / $Columns)

  for ($columnIndex = 0; $columnIndex -lt $Columns; $columnIndex++) {
    $gridCol = $DocumentXml.CreateElement("w", "gridCol", $wordNamespace)
    $gridCol.SetAttribute("w", $wordNamespace, [string]$cellWidthPct)
    $gridNode.AppendChild($gridCol) | Out-Null
  }

  for ($offset = 0; $offset -lt $CellEntries.Count; $offset += $Columns) {
    $rowNode = $DocumentXml.CreateElement("w", "tr", $wordNamespace)

    for ($columnIndex = 0; $columnIndex -lt $Columns; $columnIndex++) {
      $cellNode = New-XmlNodeFromString -DocumentXml $DocumentXml -XmlText @"
<w:tc xmlns:w="$wordNamespace">
  <w:tcPr>
    <w:tcW w:w="$cellWidthPct" w:type="pct"/>
    <w:vAlign w:val="top"/>
    <w:tcBorders>
      <w:top w:val="nil"/>
      <w:left w:val="nil"/>
      <w:bottom w:val="nil"/>
      <w:right w:val="nil"/>
    </w:tcBorders>
  </w:tcPr>
</w:tc>
"@

      $entryIndex = $offset + $columnIndex
      if ($entryIndex -lt $CellEntries.Count) {
        $entry = $CellEntries[$entryIndex]
        $cellNode.AppendChild((Get-SingleXmlNode -NodeLike $entry.ImageParagraph)) | Out-Null
        if ($null -ne $entry.CaptionParagraph) {
          $cellNode.AppendChild((Get-SingleXmlNode -NodeLike $entry.CaptionParagraph)) | Out-Null
        }
      } else {
        $cellNode.AppendChild((New-EmptyParagraph -DocumentXml $DocumentXml)) | Out-Null
      }

      $rowNode.AppendChild($cellNode) | Out-Null
    }

    $tableNode.AppendChild($rowNode) | Out-Null
  }

  return ,$tableNode
}

function Build-AnchorLookup {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Body,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $lookup = @{}
  $paragraphIndex = 0
  $tableIndex = 0

  foreach ($child in @($Body.ChildNodes)) {
    if ($child.LocalName -eq "p") {
      $paragraphIndex++
      $lookup[("P{0}" -f $paragraphIndex).ToUpperInvariant()] = $child
      continue
    }

    if ($child.LocalName -eq "tbl") {
      $tableIndex++
      $rowIndex = 0
      foreach ($row in @($child.SelectNodes("./w:tr", $NamespaceManager))) {
        $rowIndex++
        $cellIndex = 0
        foreach ($cell in @($row.SelectNodes("./w:tc", $NamespaceManager))) {
          $cellIndex++
          $lookup[("T{0}R{1}C{2}" -f $tableIndex, $rowIndex, $cellIndex).ToUpperInvariant()] = $cell
        }
      }
    }
  }

  return $lookup
}

function Add-SectionLookupEntry {
  param(
    [Parameter(Mandatory = $true)]
    [hashtable]$Lookup,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Paragraph,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $text = Get-NodeText -Node $Paragraph -NamespaceManager $NamespaceManager
  $rule = Resolve-SectionRuleFromHeading -HeadingText $text
  if ($null -eq $rule) {
    return
  }

  if (-not $Lookup.ContainsKey($rule.id)) {
    $Lookup[$rule.id] = [pscustomobject]@{
      Node = $Paragraph
      SectionId = $rule.id
      HeadingText = $text
    }
  }
}

function Build-SectionLookup {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Body,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $lookup = @{}

  foreach ($child in @($Body.ChildNodes)) {
    if ($child.LocalName -eq "p") {
      Add-SectionLookupEntry -Lookup $lookup -Paragraph $child -NamespaceManager $NamespaceManager
      continue
    }

    if ($child.LocalName -eq "tbl") {
      foreach ($row in @($child.SelectNodes("./w:tr", $NamespaceManager))) {
        foreach ($cell in @($row.SelectNodes("./w:tc", $NamespaceManager))) {
          foreach ($paragraph in @($cell.SelectNodes("./w:p", $NamespaceManager))) {
            Add-SectionLookupEntry -Lookup $lookup -Paragraph $paragraph -NamespaceManager $NamespaceManager
          }
        }
      }
    }
  }

  return $lookup
}

function Resolve-AnchorTarget {
  param(
    [Parameter(Mandatory = $true)]
    [object]$ImageSpec,

    [Parameter(Mandatory = $true)]
    [hashtable]$AnchorLookup,

    [Parameter(Mandatory = $true)]
    [hashtable]$SectionLookup
  )

  if (-not [string]::IsNullOrWhiteSpace($ImageSpec.Anchor)) {
    return Resolve-TargetReference -Selector $ImageSpec.Anchor -AnchorLookup $AnchorLookup -SectionLookup $SectionLookup -ContextLabel "Image target"
  }

  return Resolve-TargetReference -Selector $ImageSpec.SectionName -AnchorLookup $AnchorLookup -SectionLookup $SectionLookup -ContextLabel "Image section"
}

function Resolve-TargetReference {
  param(
    [AllowNull()]
    [string]$Selector,

    [Parameter(Mandatory = $true)]
    [hashtable]$AnchorLookup,

    [Parameter(Mandatory = $true)]
    [hashtable]$SectionLookup,

    [Parameter(Mandatory = $true)]
    [string]$ContextLabel
  )

  $normalizedSelector = Normalize-TargetSelector -Selector $Selector
  if ([string]::IsNullOrWhiteSpace($normalizedSelector)) {
    throw "$ContextLabel is empty."
  }

  $anchor = $null
  $sectionName = $null

  if ($normalizedSelector -match '^(?i)(anchor|location|target):(?<value>P\d+|T\d+R\d+C\d+)$') {
    $anchor = $matches["value"].ToUpperInvariant()
  } elseif ($normalizedSelector -match '^(?i)(section|heading):(?<value>.+)$') {
    $sectionName = $matches["value"].Trim()
  } elseif ($normalizedSelector -match '^(?i:P\d+|T\d+R\d+C\d+)$') {
    $anchor = $normalizedSelector.ToUpperInvariant()
  } else {
    $sectionName = $normalizedSelector
  }

  if (-not [string]::IsNullOrWhiteSpace($anchor)) {
    if (-not $AnchorLookup.ContainsKey($anchor)) {
      throw "$ContextLabel anchor was not found in the document: $anchor"
    }

    return [pscustomobject]@{
      Node = $AnchorLookup[$anchor]
      ResolutionKey = "anchor:$anchor"
      Description = $anchor
    }
  }

  $sectionId = Resolve-SectionId -SectionName $sectionName
  if ([string]::IsNullOrWhiteSpace($sectionId)) {
    throw "$ContextLabel was not recognized: $Selector"
  }
  if (-not $SectionLookup.ContainsKey($sectionId)) {
    $availableSections = @($SectionLookup.Keys | Sort-Object) -join ", "
    throw "$ContextLabel was not found in the document: $Selector. Available sections: $availableSections"
  }

  $sectionEntry = $SectionLookup[$sectionId]
  return [pscustomobject]@{
    Node = $sectionEntry.Node
    ResolutionKey = "section:$sectionId"
    Description = $sectionEntry.HeadingText
  }
}

function Get-SectionEndInsertionNode {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$TargetNode,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  if ($TargetNode.LocalName -ne "p" -or $TargetNode.ParentNode.LocalName -ne "body") {
    return $TargetNode
  }

  $insertionNode = $TargetNode
  $cursor = $TargetNode.NextSibling
  while ($null -ne $cursor) {
    if ($cursor.LocalName -eq "sectPr") {
      break
    }

    if ($cursor.LocalName -eq "p") {
      $text = Get-NodeText -Node $cursor -NamespaceManager $NamespaceManager
      if ($null -ne (Resolve-SectionRuleFromHeading -HeadingText $text)) {
        break
      }
    }

    $insertionNode = $cursor
    $cursor = $cursor.NextSibling
  }

  return $insertionNode
}

function Get-EffectiveInsertionNode {
  param(
    [Parameter(Mandatory = $true)]
    [object]$ResolvedTarget,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $targetNode = Get-SingleXmlNode -NodeLike $ResolvedTarget.Node
  if ([string]$ResolvedTarget.ResolutionKey -like "section:*") {
    return Get-SectionEndInsertionNode -TargetNode $targetNode -NamespaceManager $NamespaceManager
  }

  return $targetNode
}

function Get-EffectiveRowLayoutTarget {
  param(
    [Parameter(Mandatory = $true)]
    [object]$ResolvedEntry,

    [AllowNull()]
    [object]$FallbackGroupTarget
  )

  if ($null -ne $ResolvedEntry.GroupResolvedTarget) {
    return $ResolvedEntry.GroupResolvedTarget
  }

  if ($null -ne $FallbackGroupTarget) {
    return $FallbackGroupTarget
  }

  return $ResolvedEntry.ResolvedTarget
}

function New-PreparedImageBlock {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$DocumentXml,

    [Parameter(Mandatory = $true)]
    [xml]$RelationshipsXml,

    [Parameter(Mandatory = $true)]
    [xml]$ContentTypesXml,

    [Parameter(Mandatory = $true)]
    [string]$MediaDirectory,

    [Parameter(Mandatory = $true)]
    [object]$ImageSpec,

    [Parameter(Mandatory = $true)]
    [ref]$NextMediaIndex,

    [Parameter(Mandatory = $true)]
    [ref]$NextRelationshipId,

    [Parameter(Mandatory = $true)]
    [ref]$NextDocPrId,

    [double]$WidthCmOverride = 0
  )

  $contentType = Get-ImageContentType -Path $ImageSpec.ImagePath
  $extension = [System.IO.Path]::GetExtension($ImageSpec.ImagePath).TrimStart('.').ToLowerInvariant()
  $effectiveWidthCm = if ($WidthCmOverride -gt 0) { $WidthCmOverride } else { [double]$ImageSpec.WidthCm }
  $dimensions = Get-ImageSizeEmu -Path $ImageSpec.ImagePath -WidthCm $effectiveWidthCm

  $mediaFileName = "image{0}.{1}" -f $NextMediaIndex.Value, $extension
  Copy-Item -LiteralPath $ImageSpec.ImagePath -Destination (Join-Path $MediaDirectory $mediaFileName) -Force
  Ensure-ContentTypeDefault -ContentTypesXml $ContentTypesXml -Extension $extension -ContentType $contentType

  $relationshipId = "rId{0}" -f $NextRelationshipId.Value
  Add-ImageRelationship -RelationshipsXml $RelationshipsXml -RelationshipId $relationshipId -Target ("media/{0}" -f $mediaFileName)

  $imageParagraph = Get-SingleXmlNode -NodeLike (New-ImageParagraph -DocumentXml $DocumentXml -RelationshipId $relationshipId -DocPrId $NextDocPrId.Value -Name $mediaFileName -WidthEmu $dimensions.WidthEmu -HeightEmu $dimensions.HeightEmu)
  $captionParagraph = $null
  if (-not [string]::IsNullOrWhiteSpace($ImageSpec.Caption)) {
    $captionParagraph = Get-SingleXmlNode -NodeLike (New-CaptionParagraph -DocumentXml $DocumentXml -CaptionText $ImageSpec.Caption)
  }

  $NextMediaIndex.Value++
  $NextRelationshipId.Value++
  $NextDocPrId.Value++

  return [pscustomobject]@{
    ImageParagraph = $imageParagraph
    CaptionParagraph = $captionParagraph
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

$resolvedMappingPathForProbe = $null
if (-not [string]::IsNullOrWhiteSpace($MappingPath)) {
  $resolvedMappingPathForProbe = (Resolve-Path -LiteralPath $MappingPath).Path
}
$script:ImagePathProbeRoots = Get-ImagePathProbeRoots -DocxPath $resolvedDocxPath -MappingPath $resolvedMappingPathForProbe

$imageMappingDocument = Get-ImageMappingRootObject -PathToJson $MappingPath -InlineJson $ImagesJson
$reportProfile = Resolve-ImageMappingReportProfile -RootObject $imageMappingDocument.RootObject -ProfileName $ReportProfileName -ProfilePath $ReportProfilePath
Initialize-SectionRules -ReportProfile $reportProfile

$imageItems = Get-ImageMappingItems -RootObject $imageMappingDocument.RootObject
$imageSpecs = @($imageItems | ForEach-Object { Resolve-ImageSpecification -Item $_ })
if ($imageSpecs.Count -eq 0) {
  throw "No image mapping items were provided."
}

if ([string]::IsNullOrWhiteSpace($OutPath)) {
  $directory = Split-Path -Parent $resolvedDocxPath
  $fileName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedDocxPath)
  $OutPath = Join-Path $directory ($fileName + ".images.docx")
}

$resolvedOutPath = [System.IO.Path]::GetFullPath($OutPath)
if ((-not $Overwrite) -and (Test-Path -LiteralPath $resolvedOutPath)) {
  throw "Output file already exists: $resolvedOutPath. Re-run with -Overwrite to replace it."
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("openclaw-docx-images-" + [System.Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

try {
  [System.IO.Compression.ZipFile]::ExtractToDirectory($resolvedDocxPath, $tempRoot)

  $documentXmlPath = Join-Path $tempRoot "word\document.xml"
  $relationshipsPath = Join-Path $tempRoot "word\_rels\document.xml.rels"
  $contentTypesPath = Join-Path $tempRoot "[Content_Types].xml"
  $mediaDirectory = Join-Path $tempRoot "word\media"

  if (-not (Test-Path -LiteralPath $documentXmlPath)) {
    throw "word/document.xml was not found in $resolvedDocxPath"
  }
  if (-not (Test-Path -LiteralPath $relationshipsPath)) {
    throw "word/_rels/document.xml.rels was not found in $resolvedDocxPath"
  }
  if (-not (Test-Path -LiteralPath $contentTypesPath)) {
    throw "[Content_Types].xml was not found in $resolvedDocxPath"
  }

  New-Item -ItemType Directory -Path $mediaDirectory -Force | Out-Null

  [xml]$documentXml = [System.IO.File]::ReadAllText($documentXmlPath, (New-Object System.Text.UTF8Encoding($false)))
  [xml]$relationshipsXml = [System.IO.File]::ReadAllText($relationshipsPath, (New-Object System.Text.UTF8Encoding($false)))
  [xml]$contentTypesXml = [System.IO.File]::ReadAllText($contentTypesPath, (New-Object System.Text.UTF8Encoding($false)))

  $namespaceManager = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
  $namespaceManager.AddNamespace("w", $wordNamespace)
  $namespaceManager.AddNamespace("wp", $wordprocessingDrawingNamespace)

  $body = $documentXml.SelectSingleNode("/w:document/w:body", $namespaceManager)
  if ($null -eq $body) {
    throw "Could not locate /w:document/w:body in $resolvedDocxPath"
  }
  $bodyWidthCm = Get-DocumentBodyWidthCm -DocumentXml $documentXml -NamespaceManager $namespaceManager

  $anchorLookup = Build-AnchorLookup -Body $body -NamespaceManager $namespaceManager
  $sectionLookup = Build-SectionLookup -Body $body -NamespaceManager $namespaceManager
  $tailLookup = @{}
  $nextMediaIndex = Get-NextMediaIndex -MediaDirectory $mediaDirectory
  $nextRelationshipId = Get-NextRelationshipId -RelationshipsXml $relationshipsXml
  $nextDocPrId = Get-NextDocPrId -DocumentXml $documentXml -NamespaceManager $namespaceManager
  $insertedCaptionCount = 0
  $insertedImageCount = 0
  $resolvedTargets = New-Object System.Collections.Generic.List[string]
  $resolvedImageEntries = New-Object System.Collections.Generic.List[object]
  foreach ($imageSpec in $imageSpecs) {
    $resolvedTarget = Resolve-AnchorTarget -ImageSpec $imageSpec -AnchorLookup $anchorLookup -SectionLookup $sectionLookup
    [void]$resolvedTargets.Add($resolvedTarget.ResolutionKey)
    $groupResolvedTarget = $null
    if ($null -ne $imageSpec.Layout -and -not [string]::IsNullOrWhiteSpace([string]$imageSpec.Layout.GroupAnchor)) {
      $groupResolvedTarget = Resolve-TargetReference -Selector $imageSpec.Layout.GroupAnchor -AnchorLookup $anchorLookup -SectionLookup $sectionLookup -ContextLabel "Row layout group target"
    }
    $resolvedImageEntries.Add([pscustomobject]@{
        ImageSpec = $imageSpec
        ResolvedTarget = $resolvedTarget
        GroupResolvedTarget = $groupResolvedTarget
      }) | Out-Null
  }

  $index = 0
  while ($index -lt $resolvedImageEntries.Count) {
    $resolvedEntry = $resolvedImageEntries[$index]
    $imageSpec = $resolvedEntry.ImageSpec
    $resolvedTarget = $resolvedEntry.ResolvedTarget
    $rowLayoutEnabled = ($null -ne $imageSpec.Layout -and $imageSpec.Layout.Mode -eq "row" -and $imageSpec.Layout.Columns -ge 2)

    if ($rowLayoutEnabled) {
      $startIndex = $index
      $groupEntries = New-Object System.Collections.Generic.List[object]
      $groupColumns = $imageSpec.Layout.Columns
      $groupKey = if ([string]::IsNullOrWhiteSpace($imageSpec.Layout.Group)) { "__auto__" } else { $imageSpec.Layout.Group }
      $groupTarget = Get-EffectiveRowLayoutTarget -ResolvedEntry $resolvedEntry -FallbackGroupTarget $null
      $targetKey = $groupTarget.ResolutionKey

      while ($index -lt $resolvedImageEntries.Count) {
        $candidate = $resolvedImageEntries[$index]
        $candidateLayout = $candidate.ImageSpec.Layout
        if ($null -eq $candidateLayout -or $candidateLayout.Mode -ne "row" -or $candidateLayout.Columns -ne $groupColumns) {
          break
        }

        $candidateGroupKey = if ([string]::IsNullOrWhiteSpace($candidateLayout.Group)) { "__auto__" } else { $candidateLayout.Group }
        if ($candidateGroupKey -ne $groupKey) {
          break
        }

        $candidateGroupTarget = Get-EffectiveRowLayoutTarget -ResolvedEntry $candidate -FallbackGroupTarget $(if ($null -ne $resolvedEntry.GroupResolvedTarget) { $groupTarget } else { $null })
        if ($candidateGroupTarget.ResolutionKey -ne $targetKey) {
          break
        }

        $groupEntries.Add($candidate) | Out-Null
        $index++
      }

      if ($groupEntries.Count -ge 2) {
        $cellEntries = New-Object System.Collections.Generic.List[object]
        foreach ($groupEntry in $groupEntries) {
          $rowImageWidthCm = Get-EffectiveRowImageWidthCm -ImageSpec $groupEntry.ImageSpec -Columns $groupColumns -BodyWidthCm $bodyWidthCm
          $preparedImage = New-PreparedImageBlock -DocumentXml $documentXml -RelationshipsXml $relationshipsXml -ContentTypesXml $contentTypesXml -MediaDirectory $mediaDirectory -ImageSpec $groupEntry.ImageSpec -NextMediaIndex ([ref]$nextMediaIndex) -NextRelationshipId ([ref]$nextRelationshipId) -NextDocPrId ([ref]$nextDocPrId) -WidthCmOverride $rowImageWidthCm
          $cellEntries.Add($preparedImage) | Out-Null
          $insertedImageCount++
          if ($null -ne $preparedImage.CaptionParagraph) {
            $insertedCaptionCount++
          }
        }

        $layoutTable = Get-SingleXmlNode -NodeLike (New-ImageTable -DocumentXml $documentXml -CellEntries ($cellEntries.ToArray()) -Columns $groupColumns)
        $anchorNode = Get-EffectiveInsertionNode -ResolvedTarget $groupTarget -NamespaceManager $namespaceManager
        if ($anchorNode.LocalName -eq "p") {
          $insertAfter = if ($tailLookup.ContainsKey($targetKey)) { Get-SingleXmlNode -NodeLike $tailLookup[$targetKey] } else { $anchorNode }
          $parentNode = $insertAfter.ParentNode
          $parentNode.InsertAfter($layoutTable, $insertAfter) | Out-Null
          $tailLookup[$targetKey] = $layoutTable
        } else {
          $anchorNode.AppendChild($layoutTable) | Out-Null
        }

        continue
      }

      $index = $startIndex
      $resolvedEntry = $resolvedImageEntries[$index]
      $imageSpec = $resolvedEntry.ImageSpec
      $resolvedTarget = $resolvedEntry.ResolvedTarget
    }

    $preparedImage = New-PreparedImageBlock -DocumentXml $documentXml -RelationshipsXml $relationshipsXml -ContentTypesXml $contentTypesXml -MediaDirectory $mediaDirectory -ImageSpec $imageSpec -NextMediaIndex ([ref]$nextMediaIndex) -NextRelationshipId ([ref]$nextRelationshipId) -NextDocPrId ([ref]$nextDocPrId)

    $anchorNode = Get-EffectiveInsertionNode -ResolvedTarget $resolvedTarget -NamespaceManager $namespaceManager
    if ($anchorNode.LocalName -eq "p") {
      $tailKey = $resolvedTarget.ResolutionKey
      $insertAfter = if ($tailLookup.ContainsKey($tailKey)) { Get-SingleXmlNode -NodeLike $tailLookup[$tailKey] } else { $anchorNode }
      $parentNode = $insertAfter.ParentNode
      $parentNode.InsertAfter($preparedImage.ImageParagraph, $insertAfter) | Out-Null
      $insertAfter = $preparedImage.ImageParagraph
      if ($null -ne $preparedImage.CaptionParagraph) {
        $parentNode.InsertAfter($preparedImage.CaptionParagraph, $insertAfter) | Out-Null
        $insertAfter = $preparedImage.CaptionParagraph
        $insertedCaptionCount++
      }
      $tailLookup[$tailKey] = $insertAfter
    } else {
      $anchorNode.AppendChild($preparedImage.ImageParagraph) | Out-Null
      if ($null -ne $preparedImage.CaptionParagraph) {
        $anchorNode.AppendChild($preparedImage.CaptionParagraph) | Out-Null
        $insertedCaptionCount++
      }
    }

    $insertedImageCount++
    $index++
  }

  [System.IO.File]::WriteAllText($documentXmlPath, $documentXml.OuterXml, (New-Object System.Text.UTF8Encoding($false)))
  [System.IO.File]::WriteAllText($relationshipsPath, $relationshipsXml.OuterXml, (New-Object System.Text.UTF8Encoding($false)))
  [System.IO.File]::WriteAllText($contentTypesPath, $contentTypesXml.OuterXml, (New-Object System.Text.UTF8Encoding($false)))

  Write-OpenXmlPackage -SourceDirectory $tempRoot -DestinationPath $resolvedOutPath

  [pscustomobject]@{
    docxPath = $resolvedDocxPath
    outPath = $resolvedOutPath
    reportProfileName = [string]$reportProfile.name
    reportProfilePath = [string]$reportProfile.resolvedProfilePath
    insertedImageCount = $insertedImageCount
    insertedCaptionCount = $insertedCaptionCount
    anchorCount = @($resolvedTargets | Select-Object -Unique).Count
  }
} finally {
  if (Test-Path -LiteralPath $tempRoot) {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force
  }
}

