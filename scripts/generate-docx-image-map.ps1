[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$DocxPath,

  [string]$ImageSpecsPath,

  [string]$ImageSpecsJson,

  [string[]]$ImagePaths,

  [string]$ReportProfileName = "experiment-report",

  [string]$ReportProfilePath,

  [ValidateSet("json", "markdown")]
  [string]$Format = "json",

  [switch]$PlanOnly,

  [string]$OutFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

. (Join-Path $PSScriptRoot "report-profiles.ps1")

$script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$reportProfile = Get-ReportProfile -ProfileName $ReportProfileName -ProfilePath $ReportProfilePath -RepoRoot $script:RepoRoot
$wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
$defaultImageWidthCm = 10.5
$sectionRules = @()
$sectionInputAliasLookup = @{}
$sectionRuleLookup = @{}
$script:ImagePathProbeRoots = @()

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
    [string]$SpecsPath
  )

  $roots = New-Object 'System.Collections.Generic.List[string]'

  Add-UniqueProbeRoot -Roots $roots -Path (Get-Location).Path
  Add-UniqueProbeRoot -Roots $roots -Path (Split-Path -Parent $DocxPath)
  Add-UniqueProbeRoot -Roots $roots -Path $script:RepoRoot
  Add-UniqueProbeRoot -Roots $roots -Path (Split-Path -Parent $script:RepoRoot)

  if (-not [string]::IsNullOrWhiteSpace($SpecsPath)) {
    Add-UniqueProbeRoot -Roots $roots -Path (Split-Path -Parent $SpecsPath)
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

$sectionRules = @(Get-ReportProfileSectionRules -Profile $reportProfile)

foreach ($rule in $sectionRules) {
  $sectionRuleLookup[$rule.id] = $rule
  foreach ($alias in @($rule.inputAliases + $rule.id)) {
    $normalizedAlias = Normalize-FieldKey -Text $alias
    if (-not [string]::IsNullOrWhiteSpace($normalizedAlias)) {
      $sectionInputAliasLookup[$normalizedAlias] = $rule.id
    }
  }
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
    mode = $normalizedMode
    columns = $columns
    group = $(if ([string]::IsNullOrWhiteSpace($group)) { $null } else { $group })
    groupAnchor = $(if ([string]::IsNullOrWhiteSpace($groupAnchor)) { $null } else { $groupAnchor })
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

function Resolve-AvailableSectionId {
  param(
    [AllowNull()]
    [string]$RequestedSectionId,

    [AllowEmptyCollection()]
    [string[]]$AvailableSectionIds
  )

  $resolvedAvailableSectionIds = @($AvailableSectionIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
  if ([string]::IsNullOrWhiteSpace($RequestedSectionId) -or $resolvedAvailableSectionIds.Count -eq 0) {
    return $null
  }

  if ($resolvedAvailableSectionIds -contains $RequestedSectionId) {
    return $RequestedSectionId
  }

  $preferredOrder = New-Object System.Collections.Generic.List[string]
  foreach ($candidate in @($RequestedSectionId) + @(Get-ReportProfileImageFallbackSectionOrder -Profile $reportProfile)) {
    if (-not [string]::IsNullOrWhiteSpace($candidate) -and -not $preferredOrder.Contains($candidate)) {
      $preferredOrder.Add($candidate) | Out-Null
    }
  }

  foreach ($candidate in $preferredOrder) {
    if ($resolvedAvailableSectionIds -contains $candidate) {
      return $candidate
    }
  }

  return $resolvedAvailableSectionIds[0]
}

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

function Get-ImageInputItems {
  param(
    [AllowNull()]
    [string]$SpecsPath,

    [AllowNull()]
    [string]$SpecsJson,

    [AllowNull()]
    [string[]]$Paths
  )

  if ($null -ne $Paths -and $Paths.Count -eq 1 -and -not [string]::IsNullOrWhiteSpace([string]$Paths[0]) -and ([string]$Paths[0]).Contains(",")) {
    $splitPaths = @(([string]$Paths[0] -split '\s*,\s*') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($splitPaths.Count -gt 1) {
      $Paths = $splitPaths
    }
  }

  $providedCount = 0
  if (-not [string]::IsNullOrWhiteSpace($SpecsPath)) {
    $providedCount++
  }
  if (-not [string]::IsNullOrWhiteSpace($SpecsJson)) {
    $providedCount++
  }
  if ($null -ne $Paths -and $Paths.Count -gt 0) {
    $providedCount++
  }

  if ($providedCount -ne 1) {
    throw "Provide exactly one of -ImageSpecsPath, -ImageSpecsJson, or -ImagePaths."
  }

  if ($null -ne $Paths -and $Paths.Count -gt 0) {
    return @($Paths | ForEach-Object { [pscustomobject]@{ path = $_ } })
  }

  if (-not [string]::IsNullOrWhiteSpace($SpecsPath)) {
    $resolvedSpecsPath = (Resolve-Path -LiteralPath $SpecsPath).Path
    $rootObject = (Get-Content -LiteralPath $resolvedSpecsPath -Raw -Encoding UTF8) | ConvertFrom-Json
  } else {
    $rootObject = $SpecsJson | ConvertFrom-Json
  }

  if ($null -eq $rootObject) {
    throw "Image specs JSON is empty."
  }

  if ($rootObject -is [System.Collections.IEnumerable] -and $rootObject -isnot [string]) {
    return @(ConvertTo-ObjectArray -Value $rootObject)
  }

  $rootTable = ConvertTo-PlainHashtable -InputObject $rootObject
  if ($rootTable.ContainsKey("images")) {
    return @(ConvertTo-ObjectArray -Value $rootTable["images"])
  }

  return @(ConvertTo-ObjectArray -Value $rootObject)
}

function Resolve-ImageInputSpec {
  param(
    [Parameter(Mandatory = $true)]
    [object]$Item
  )

  if ($Item -is [string]) {
    $itemTable = @{ path = [string]$Item }
  } else {
    $itemTable = ConvertTo-PlainHashtable -InputObject $Item
  }

  $imagePath = $null
  foreach ($key in @("path", "imagePath", "file")) {
    if ($itemTable.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$itemTable[$key])) {
      $imagePath = Resolve-ExistingImagePath -Path ([string]$itemTable[$key])
      break
    }
  }
  if ([string]::IsNullOrWhiteSpace($imagePath)) {
    throw "Each image spec must include path, imagePath, or file."
  }

  $sectionName = $null
  foreach ($key in @("section", "sectionName", "heading")) {
    if ($itemTable.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$itemTable[$key])) {
      $sectionName = ([string]$itemTable[$key]).Trim()
      break
    }
  }

  $caption = $null
  foreach ($key in @("caption", "title", "figureCaption")) {
    if ($itemTable.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$itemTable[$key])) {
      $caption = ([string]$itemTable[$key]).Trim()
      break
    }
  }

  $widthCm = $null
  foreach ($key in @("widthCm", "width")) {
    if ($itemTable.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$itemTable[$key])) {
      $widthCm = [double]$itemTable[$key]
      break
    }
  }

  $anchor = $null
  foreach ($key in @("anchor", "location", "target")) {
    if ($itemTable.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$itemTable[$key])) {
      $anchor = Normalize-TargetSelector -Selector ([string]$itemTable[$key])
      break
    }
  }

  $layout = Resolve-ImageLayoutSpec -LayoutValue $(if ($itemTable.ContainsKey("layout")) { $itemTable["layout"] } else { $null }) -ItemTable $itemTable

  return [pscustomobject]@{
    ImagePath = $imagePath
    BaseName = [System.IO.Path]::GetFileNameWithoutExtension($imagePath)
    SectionName = $sectionName
    SectionProvided = (-not [string]::IsNullOrWhiteSpace($sectionName))
    Caption = $caption
    CaptionProvided = (-not [string]::IsNullOrWhiteSpace($caption))
    WidthCm = $widthCm
    Anchor = $anchor
    AnchorProvided = (-not [string]::IsNullOrWhiteSpace($anchor))
    Layout = $layout
  }
}

function New-ImageLayoutOutput {
  param(
    [AllowNull()]
    [object]$LayoutSpec
  )

  if ($null -eq $LayoutSpec) {
    return $null
  }

  $layoutOutput = [ordered]@{
    mode = [string]$LayoutSpec.mode
  }

  if ($null -ne $LayoutSpec.columns) {
    $layoutOutput["columns"] = [int]$LayoutSpec.columns
  }

  if (-not [string]::IsNullOrWhiteSpace([string]$LayoutSpec.group)) {
    $layoutOutput["group"] = [string]$LayoutSpec.group
  }

  if (-not [string]::IsNullOrWhiteSpace([string]$LayoutSpec.groupAnchor)) {
    $layoutOutput["groupAnchor"] = [string]$LayoutSpec.groupAnchor
  }

  return $layoutOutput
}

function Apply-RowGroupAnchors {
  param(
    [Parameter(Mandatory = $true)]
    [object[]]$Entries,

    [Parameter(Mandatory = $true)]
    [object]$Notes
  )

  $index = 0
  while ($index -lt $Entries.Count) {
    $entry = $Entries[$index]
    $layout = $entry.Layout
    if ($null -eq $layout -or [string]$layout["mode"] -ne "row" -or $null -eq $layout["columns"] -or [int]$layout["columns"] -lt 2 -or [string]::IsNullOrWhiteSpace([string]$layout["group"])) {
      $index++
      continue
    }

    $groupKey = [string]$layout["group"]
    $groupColumns = [int]$layout["columns"]
    $groupEntries = New-Object System.Collections.Generic.List[object]

    while ($index -lt $Entries.Count) {
      $candidate = $Entries[$index]
      $candidateLayout = $candidate.Layout
      if ($null -eq $candidateLayout -or [string]$candidateLayout["mode"] -ne "row" -or [int]$candidateLayout["columns"] -ne $groupColumns -or [string]$candidateLayout["group"] -ne $groupKey) {
        break
      }

      $groupEntries.Add($candidate) | Out-Null
      $index++
    }

    if ($groupEntries.Count -lt 2) {
      continue
    }

    $existingGroupAnchors = @(
      $groupEntries |
      ForEach-Object {
        if ($null -ne $_.Layout -and $_.Layout.Contains("groupAnchor") -and -not [string]::IsNullOrWhiteSpace([string]$_.Layout["groupAnchor"])) {
          [string]$_.Layout["groupAnchor"]
        }
      } |
      Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
      Select-Object -Unique
    )

    if ($existingGroupAnchors.Count -gt 1) {
      throw ("Row layout group '{0}' has conflicting groupAnchor values: {1}" -f $groupKey, ($existingGroupAnchors -join ", "))
    }

    $sharedGroupAnchor = if ($existingGroupAnchors.Count -eq 1) { [string]$existingGroupAnchors[0] } else { $null }
    $resolvedSectionIds = @($groupEntries | ForEach-Object { $_.ResolvedSectionId } | Select-Object -Unique)

    if ([string]::IsNullOrWhiteSpace($sharedGroupAnchor) -and $resolvedSectionIds.Count -gt 1) {
      $firstEntry = $groupEntries[0]
      if (-not [string]::IsNullOrWhiteSpace([string]$firstEntry.Anchor)) {
        $sharedGroupAnchor = [string]$firstEntry.Anchor
      } else {
        $sharedGroupAnchor = [string]$firstEntry.OutputEntry["section"]
      }

      $Notes.Add(("Row layout group {0} spans multiple sections and will use shared groupAnchor {1}." -f $groupKey, $sharedGroupAnchor)) | Out-Null
    }

    if (-not [string]::IsNullOrWhiteSpace($sharedGroupAnchor)) {
      foreach ($groupEntry in $groupEntries) {
        $groupEntry.Layout["groupAnchor"] = $sharedGroupAnchor
      }
    }
  }
}

function Apply-AutoRowLayouts {
  param(
    [Parameter(Mandatory = $true)]
    [object[]]$Entries,

    [Parameter(Mandatory = $true)]
    [object]$Notes,

    [int]$Columns = 2
  )

  if ($Columns -lt 2) {
    throw "Auto row layout must use at least 2 columns."
  }

  $index = 0
  $groupIndex = 0
  while ($index -lt $Entries.Count) {
    $entry = $Entries[$index]
    if ($null -ne $entry.Layout -or -not [string]::IsNullOrWhiteSpace([string]$entry.Anchor)) {
      $index++
      continue
    }

    $sectionId = [string]$entry.ResolvedSectionId
    $runEntries = New-Object System.Collections.Generic.List[object]
    while ($index -lt $Entries.Count) {
      $candidate = $Entries[$index]
      if (
        $null -ne $candidate.Layout -or
        -not [string]::IsNullOrWhiteSpace([string]$candidate.Anchor) -or
        [string]$candidate.ResolvedSectionId -ne $sectionId
      ) {
        break
      }

      $runEntries.Add($candidate) | Out-Null
      $index++
    }

    $layoutEntryCount = [int]([Math]::Floor($runEntries.Count / [double]$Columns) * $Columns)
    if ($layoutEntryCount -lt $Columns) {
      continue
    }

    $groupIndex++
    $groupName = "auto-{0}-{1:D2}" -f $sectionId, $groupIndex
    $groupAnchor = [string]$runEntries[0].OutputEntry["section"]
    for ($runIndex = 0; $runIndex -lt $layoutEntryCount; $runIndex++) {
      $layoutOutput = [ordered]@{
        mode = "row"
        columns = $Columns
        group = $groupName
        groupAnchor = $groupAnchor
      }

      $runEntry = $runEntries[$runIndex]
      $runEntry.OutputEntry["layout"] = $layoutOutput
      $runEntry.Layout = $layoutOutput
    }

    $Notes.Add(("Auto row layout grouped {0} image(s) for {1} with {2} columns." -f $layoutEntryCount, $groupAnchor, $Columns)) | Out-Null
  }
}

function Add-DiscoveredSection {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Paragraph,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $text = Get-NodeText -Node $Paragraph -NamespaceManager $NamespaceManager
  $rule = Resolve-SectionRuleFromHeading -HeadingText $text
  if ($null -eq $rule) {
    return $null
  }

  return [pscustomobject]@{
    id = $rule.id
    canonicalLabel = $rule.canonicalLabel
    headingText = $text
  }
}

function Get-DiscoveredSections {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
  try {
    $documentXmlText = Get-ZipEntryText -Archive $archive -EntryName "word/document.xml"
    if ([string]::IsNullOrWhiteSpace($documentXmlText)) {
      throw "word/document.xml was not found in $Path"
    }

    [xml]$documentXml = $documentXmlText
    $namespaceManager = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
    $namespaceManager.AddNamespace("w", $wordNamespace)

    $body = $documentXml.SelectSingleNode("/w:document/w:body", $namespaceManager)
    if ($null -eq $body) {
      throw "Could not locate /w:document/w:body in $Path"
    }

    $sections = @()
    foreach ($child in @($body.ChildNodes)) {
      if ($child.LocalName -eq "p") {
        $entry = Add-DiscoveredSection -Paragraph $child -NamespaceManager $namespaceManager
        if ($null -ne $entry) {
          $sections += $entry
        }
        continue
      }

      if ($child.LocalName -eq "tbl") {
        foreach ($row in @($child.SelectNodes("./w:tr", $namespaceManager))) {
          foreach ($cell in @($row.SelectNodes("./w:tc", $namespaceManager))) {
            foreach ($paragraph in @($cell.SelectNodes("./w:p", $namespaceManager))) {
              $entry = Add-DiscoveredSection -Paragraph $paragraph -NamespaceManager $namespaceManager
              if ($null -ne $entry) {
                $sections += $entry
              }
            }
          }
        }
      }
    }

    return @($sections)
  } finally {
    $archive.Dispose()
  }
}

function Infer-SectionIdFromBaseName {
  param(
    [AllowNull()]
    [string]$BaseName
  )

  $normalized = Normalize-FieldKey -Text $BaseName
  if ([string]::IsNullOrWhiteSpace($normalized)) {
    return $null
  }

  if ($normalized -match '拓扑|environment|vmware|server|机房|环境|network') {
    return "environment"
  }
  if ($normalized -match 'ping|result|reply|输出|结果|success') {
    return "result"
  }
  if ($normalized -match 'arp|analysis|error|problem|issue|异常') {
    return "analysis"
  }
  if ($normalized -match 'step|步骤|config|setup|install|command|cmd|命令|ipconfig|配置') {
    return "steps"
  }

  return $null
}

function Get-FallbackSectionId {
  param(
    [Parameter(Mandatory = $true)]
    [int]$Index,

    [Parameter(Mandatory = $true)]
    [int]$Total,

    [Parameter(Mandatory = $true)]
    [string[]]$AvailableSectionIds
  )

  $stepResultSplitIndex = [int][Math]::Ceiling($Total / 2.0)
  $fallbackSectionOrder = @(Get-ReportProfileImageFallbackSectionOrder -Profile $reportProfile)
  $preferredOrder = New-Object System.Collections.Generic.List[string]

  foreach ($preferred in @(
      $(if ($Total -eq 1) { "steps" } else { $null }),
      $(if ($Total -gt 1 -and $Index -le $stepResultSplitIndex) { "steps" } else { $null }),
      $(if ($Total -gt 1 -and $Index -gt $stepResultSplitIndex) { "result" } else { $null })
    )) {
    if (-not [string]::IsNullOrWhiteSpace($preferred) -and -not $preferredOrder.Contains($preferred)) {
      $preferredOrder.Add($preferred) | Out-Null
    }
  }

  foreach ($preferred in $fallbackSectionOrder) {
    if (-not [string]::IsNullOrWhiteSpace($preferred) -and -not $preferredOrder.Contains($preferred)) {
      $preferredOrder.Add($preferred) | Out-Null
    }
  }

  foreach ($preferred in $preferredOrder) {
    if (-not [string]::IsNullOrWhiteSpace($preferred) -and ($AvailableSectionIds -contains $preferred)) {
      return $preferred
    }
  }

  return $AvailableSectionIds[0]
}

function Get-DefaultCaption {
  param(
    [Parameter(Mandatory = $true)]
    [int]$Index,

    [Parameter(Mandatory = $true)]
    [string]$SectionId,

    [AllowNull()]
    [string]$BaseName
  )

  $body = Get-ReportProfileDefaultImageCaptionBody -Profile $reportProfile -SectionId $SectionId -BaseName $BaseName
  return ("图{0} {1}" -f $Index, $body)
}

$resolvedDocxPath = (Resolve-Path -LiteralPath $DocxPath).Path
if ([System.IO.Path]::GetExtension($resolvedDocxPath).ToLowerInvariant() -ne ".docx") {
  throw "Only .docx files are supported: $resolvedDocxPath"
}

$resolvedSpecsPathForProbe = $null
if (-not [string]::IsNullOrWhiteSpace($ImageSpecsPath)) {
  $resolvedSpecsPathForProbe = (Resolve-Path -LiteralPath $ImageSpecsPath).Path
}
$script:ImagePathProbeRoots = Get-ImagePathProbeRoots -DocxPath $resolvedDocxPath -SpecsPath $resolvedSpecsPathForProbe
$imageInputMode = if (-not [string]::IsNullOrWhiteSpace($ImageSpecsPath)) {
  "specs-path"
} elseif (-not [string]::IsNullOrWhiteSpace($ImageSpecsJson)) {
  "specs-json"
} else {
  "image-paths"
}

$rawItems = Get-ImageInputItems -SpecsPath $ImageSpecsPath -SpecsJson $ImageSpecsJson -Paths $ImagePaths
$imageSpecs = @($rawItems | ForEach-Object { Resolve-ImageInputSpec -Item $_ })
if ($imageSpecs.Count -eq 0) {
  throw "No image specs were provided."
}

$discoveredSections = @(Get-DiscoveredSections -Path $resolvedDocxPath)
$availableSectionIds = @($discoveredSections | ForEach-Object { $_.id } | Select-Object -Unique)
if ($availableSectionIds.Count -eq 0) {
  throw "No supported report sections were found in $resolvedDocxPath."
}

$resolvedImageEntries = New-Object System.Collections.Generic.List[object]
$notes = New-Object System.Collections.Generic.List[string]

for ($index = 0; $index -lt $imageSpecs.Count; $index++) {
  $spec = $imageSpecs[$index]
  $resolvedSectionId = $null
  $sectionSource = $null
  $sectionReason = $null
  $sectionConfidence = $null

  if ($spec.SectionProvided) {
    $resolvedSectionId = Resolve-SectionId -SectionName $spec.SectionName
    if ([string]::IsNullOrWhiteSpace($resolvedSectionId)) {
      throw "Image section was not recognized: $($spec.SectionName)"
    }
    if ($availableSectionIds -notcontains $resolvedSectionId) {
      $fallbackSectionId = Resolve-AvailableSectionId -RequestedSectionId $resolvedSectionId -AvailableSectionIds $availableSectionIds
      if ([string]::IsNullOrWhiteSpace($fallbackSectionId)) {
        throw "Image section is not available in the target docx: $($spec.SectionName)"
      }
      $notes.Add(("Image {0} requested section {1}, but the target docx does not contain it. Falling back to {2}." -f ($index + 1), $spec.SectionName, $fallbackSectionId)) | Out-Null
      $resolvedSectionId = $fallbackSectionId
      $sectionSource = "explicit-fallback"
      $sectionReason = "Requested section was not present in the target docx, so the nearest available section was used."
      $sectionConfidence = "medium"
    } else {
      $sectionSource = "explicit"
      $sectionReason = "Section was provided in the image spec."
      $sectionConfidence = "high"
    }
  } else {
    $resolvedSectionId = Infer-SectionIdFromBaseName -BaseName $spec.BaseName
    if (-not [string]::IsNullOrWhiteSpace($resolvedSectionId) -and ($availableSectionIds -contains $resolvedSectionId)) {
      $sectionSource = "filename"
      $sectionReason = "Section was inferred from the image file name."
      $sectionConfidence = "medium"
    } else {
      $resolvedSectionId = Get-FallbackSectionId -Index ($index + 1) -Total $imageSpecs.Count -AvailableSectionIds $availableSectionIds
      $sectionSource = "fallback-order"
      $sectionReason = "Section was assigned from image order because no file-name hint matched."
      $sectionConfidence = "low"
    }
    $notes.Add(("Image {0} section inferred as {1} from file name or fallback order." -f ($index + 1), $resolvedSectionId)) | Out-Null
  }

  $resolvedRule = $sectionRuleLookup[$resolvedSectionId]
  $caption = if ($spec.CaptionProvided) { $spec.Caption } else { Get-DefaultCaption -Index ($index + 1) -SectionId $resolvedSectionId -BaseName $spec.BaseName }
  $widthCm = if ($null -ne $spec.WidthCm) { $spec.WidthCm } else { $defaultImageWidthCm }

  $sectionHeading = ($discoveredSections | Where-Object { $_.id -eq $resolvedSectionId } | Select-Object -First 1 -ExpandProperty headingText)

  $imageEntry = [ordered]@{
    section = $resolvedRule.canonicalLabel
    path = $spec.ImagePath
    caption = $caption
    widthCm = $widthCm
    resolvedHeading = $sectionHeading
  }
  if ($spec.AnchorProvided) {
    $imageEntry["anchor"] = $spec.Anchor
  }
  $layoutOutput = New-ImageLayoutOutput -LayoutSpec $spec.Layout
  if ($null -ne $layoutOutput) {
    $imageEntry["layout"] = $layoutOutput
  }

  $resolvedImageEntries.Add([pscustomobject]@{
      OutputEntry = $imageEntry
      Layout = $layoutOutput
      Anchor = $spec.Anchor
      ResolvedSectionId = $resolvedSectionId
      SectionSource = $sectionSource
      SectionReason = $sectionReason
      SectionConfidence = $sectionConfidence
    }) | Out-Null
}

Apply-AutoRowLayouts -Entries ($resolvedImageEntries.ToArray()) -Notes $notes -Columns 2
Apply-RowGroupAnchors -Entries ($resolvedImageEntries.ToArray()) -Notes $notes

$resultImages = New-Object System.Collections.Generic.List[object]
foreach ($resolvedEntry in $resolvedImageEntries) {
  $resultImages.Add([pscustomobject]$resolvedEntry.OutputEntry) | Out-Null
}

$planEntries = New-Object System.Collections.Generic.List[object]
for ($planIndex = 0; $planIndex -lt $resolvedImageEntries.Count; $planIndex++) {
  $resolvedEntry = $resolvedImageEntries[$planIndex]
  $item = [pscustomobject]$resolvedEntry.OutputEntry
  $layoutText = "none"
  if ($item.PSObject.Properties.Name -contains 'layout' -and $null -ne $item.layout) {
    $layoutTable = ConvertTo-PlainHashtable -InputObject $item.layout
    $layoutParts = New-Object System.Collections.Generic.List[string]
    [void]$layoutParts.Add([string]$layoutTable["mode"])
    if ($layoutTable.ContainsKey("columns") -and $null -ne $layoutTable["columns"]) {
      [void]$layoutParts.Add(("columns={0}" -f $layoutTable["columns"]))
    }
    if ($layoutTable.ContainsKey("group") -and -not [string]::IsNullOrWhiteSpace([string]$layoutTable["group"])) {
      [void]$layoutParts.Add(("group={0}" -f $layoutTable["group"]))
    }
    if ($layoutTable.ContainsKey("groupAnchor") -and -not [string]::IsNullOrWhiteSpace([string]$layoutTable["groupAnchor"])) {
      [void]$layoutParts.Add(("groupAnchor={0}" -f $layoutTable["groupAnchor"]))
    }
    $layoutText = $layoutParts -join ", "
  }

  $planEntries.Add([pscustomobject]@{
      index = $planIndex + 1
      fileName = [System.IO.Path]::GetFileName([string]$item.path)
      path = [string]$item.path
      proposedSection = [string]$item.section
      resolvedHeading = [string]$item.resolvedHeading
      proposedCaption = [string]$item.caption
      widthCm = $item.widthCm
      layout = $layoutText
      sectionSource = [string]$resolvedEntry.SectionSource
      confidence = [string]$resolvedEntry.SectionConfidence
      reason = [string]$resolvedEntry.SectionReason
    }) | Out-Null
}

$result = [pscustomobject]@{
  summary = [pscustomobject]@{
    docxPath = $resolvedDocxPath
    reportProfileName = [string]$reportProfile.name
    reportProfilePath = [string]$reportProfile.resolvedProfilePath
    imageInputMode = $imageInputMode
    imageCount = $resultImages.Count
    availableSections = @($discoveredSections | ForEach-Object { $_.headingText } | Select-Object -Unique)
    planOnly = [bool]$PlanOnly
  }
  images = $resultImages
  plan = $planEntries
  notes = $notes
}

if ($Format -eq "json") {
  if ($PlanOnly) {
    $output = ([pscustomobject]@{
        summary = $result.summary
        plan = $planEntries
        notes = $notes
      }) | ConvertTo-Json -Depth 6
  } else {
    $output = $result | ConvertTo-Json -Depth 6
  }
} else {
  $lines = New-Object System.Collections.Generic.List[string]
  [void]$lines.Add($(if ($PlanOnly) { "# DOCX Image Plan" } else { "# DOCX Image Map" }))
  [void]$lines.Add("")
  [void]$lines.Add("- Docx: $resolvedDocxPath")
  [void]$lines.Add("- Image count: $($resultImages.Count)")
  [void]$lines.Add("- Available sections: $((@($result.summary.availableSections) -join ', '))")
  [void]$lines.Add("")
  if ($PlanOnly) {
    [void]$lines.Add("## Proposed Image Placement")
    [void]$lines.Add("")
    [void]$lines.Add("| # | File | Proposed section | Caption | Layout | Source | Confidence | Reason |")
    [void]$lines.Add("|---|---|---|---|---|---|---|---|")
    foreach ($item in $planEntries) {
      $rowValues = @(
        [string]$item.index,
        [string]$item.fileName,
        [string]$item.proposedSection,
        [string]$item.proposedCaption,
        [string]$item.layout,
        [string]$item.sectionSource,
        [string]$item.confidence,
        [string]$item.reason
      ) | ForEach-Object { ($_ -replace '\|', '\|') }
      [void]$lines.Add(("| {0} |" -f ($rowValues -join " | ")))
    }
    [void]$lines.Add("")
    [void]$lines.Add("> Review this plan before inserting images. Low-confidence rows came from fallback order and should be checked manually.")
  } else {
    [void]$lines.Add("## Images")
    foreach ($item in $resultImages) {
      $layoutSuffix = ""
      if ($item.PSObject.Properties.Name -contains 'layout' -and $null -ne $item.layout) {
        $layoutTable = ConvertTo-PlainHashtable -InputObject $item.layout
        $layoutSuffix = " [layout: {0}" -f $layoutTable["mode"]
        if ($layoutTable.ContainsKey("columns") -and $null -ne $layoutTable["columns"]) {
          $layoutSuffix += ", columns=$($layoutTable["columns"])"
        }
        if ($layoutTable.ContainsKey("group") -and -not [string]::IsNullOrWhiteSpace([string]$layoutTable["group"])) {
          $layoutSuffix += ", group=$($layoutTable["group"])"
        }
        if ($layoutTable.ContainsKey("groupAnchor") -and -not [string]::IsNullOrWhiteSpace([string]$layoutTable["groupAnchor"])) {
          $layoutSuffix += ", groupAnchor=$($layoutTable["groupAnchor"])"
        }
        $layoutSuffix += "]"
      }
      [void]$lines.Add("- $([System.IO.Path]::GetFileName($item.path)) -> $($item.section) -> $($item.caption)$layoutSuffix")
    }
  }
  if ($notes.Count -gt 0) {
    [void]$lines.Add("")
    [void]$lines.Add("## Notes")
    foreach ($note in $notes) {
      [void]$lines.Add("- $note")
    }
  }
  $output = $lines -join [Environment]::NewLine
}

if ([string]::IsNullOrWhiteSpace($OutFile)) {
  Write-Output $output
} else {
  [System.IO.File]::WriteAllText($OutFile, $output, (New-Object System.Text.UTF8Encoding($true)))
  if ($PlanOnly) {
    Write-Output "Wrote image plan to $OutFile"
  } else {
    Write-Output "Wrote image map to $OutFile"
  }
}





