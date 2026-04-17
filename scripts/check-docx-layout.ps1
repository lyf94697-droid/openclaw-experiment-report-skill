[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$DocxPath,

  [string]$ReportProfileName = "experiment-report",

  [string]$ReportProfilePath,

  [int]$ExpectedImageCount = -1,

  [int]$ExpectedCaptionCount = -1,

  [string]$OutFile,

  [ValidateSet("json", "markdown")]
  [string]$Format = "json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

. (Join-Path $PSScriptRoot "report-profiles.ps1")

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
  $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
  try {
    return $reader.ReadToEnd()
  } finally {
    $reader.Dispose()
    $stream.Dispose()
  }
}

function New-NamespaceManager {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$Document
  )

  $manager = New-Object System.Xml.XmlNamespaceManager($Document.NameTable)
  [void]$manager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
  [void]$manager.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
  [void]$manager.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
  [void]$manager.AddNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships")
  Write-Output -NoEnumerate $manager
}

function Get-ParagraphTexts {
  param(
    [Parameter(Mandatory = $true)]
    [xml]$Document,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $paragraphTexts = New-Object System.Collections.Generic.List[string]
  foreach ($paragraph in @($Document.SelectNodes("//w:p", $NamespaceManager))) {
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($textNode in @($paragraph.SelectNodes(".//w:t", $NamespaceManager))) {
      [void]$parts.Add([string]$textNode.InnerText)
    }
    $text = ($parts -join "").Trim()
    if (-not [string]::IsNullOrWhiteSpace($text)) {
      [void]$paragraphTexts.Add($text)
    }
  }

  return $paragraphTexts
}

function Add-Finding {
  param(
    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [System.Collections.Generic.List[object]]$Findings,

    [Parameter(Mandatory = $true)]
    [string]$Code,

    [Parameter(Mandatory = $true)]
    [string]$Message
  )

  [void]$Findings.Add([pscustomobject]@{
      code = $Code
      message = $Message
    })
}

function ConvertTo-SectionPattern {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Text
  )

  return [regex]::Escape($Text)
}

function ConvertTo-IntegerFromDigitString {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return $null
  }

  $value = 0
  foreach ($character in $Text.ToCharArray()) {
    $digit = [System.Globalization.CharUnicodeInfo]::GetDigitValue($character)
    if ($digit -lt 0) {
      return $null
    }
    $value = ($value * 10) + $digit
  }

  return $value
}

function New-LayoutCheckMessage {
  param(
    [Parameter(Mandatory = $true)]
    [bool]$Passed,

    [Parameter(Mandatory = $true)]
    [int]$ActualImageCount,

    [AllowNull()]
    [int]$ExpectedImageCount,

    [Parameter(Mandatory = $true)]
    [int]$ActualCaptionCount,

    [AllowNull()]
    [int]$ExpectedCaptionCount,

    [Parameter(Mandatory = $true)]
    [int]$PlaceholderCount,

    [Parameter(Mandatory = $true)]
    [int]$MissingSectionCount,

    [Parameter(Mandatory = $true)]
    [int]$ErrorCount,

    [Parameter(Mandatory = $true)]
    [int]$WarningCount
  )

  $imageSummary = if ($ExpectedImageCount -ge 0) {
    "images {0}/{1}" -f $ActualImageCount, $ExpectedImageCount
  } else {
    "images {0}" -f $ActualImageCount
  }

  $captionSummary = if ($ExpectedCaptionCount -ge 0) {
    "captions {0}/{1}" -f $ActualCaptionCount, $ExpectedCaptionCount
  } else {
    "captions {0}" -f $ActualCaptionCount
  }

  return "Layout check {0}: {1}, {2}, placeholders {3}, missing sections {4}, errors {5}, warnings {6}." -f $(if ($Passed) { "passed" } else { "failed" }), $imageSummary, $captionSummary, $PlaceholderCount, $MissingSectionCount, $ErrorCount, $WarningCount
}

$resolvedDocxPath = (Resolve-Path -LiteralPath $DocxPath).Path
if ([string]::IsNullOrWhiteSpace($ReportProfilePath)) {
  $resolvedReportProfilePath = $null
} else {
  $resolvedReportProfilePath = (Resolve-Path -LiteralPath $ReportProfilePath).Path
}
$reportProfile = Get-ReportProfile -ProfileName $ReportProfileName -ProfilePath $resolvedReportProfilePath -RepoRoot (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
if ([System.IO.Path]::GetExtension($resolvedDocxPath).ToLowerInvariant() -ne ".docx") {
  throw "Only .docx files are supported: $resolvedDocxPath"
}

$warnings = New-Object System.Collections.Generic.List[object]
$errors = New-Object System.Collections.Generic.List[object]
$placeholderMatches = New-Object System.Collections.Generic.List[object]

$archive = [System.IO.Compression.ZipFile]::OpenRead($resolvedDocxPath)
try {
  $documentText = Get-ZipEntryText -Archive $archive -EntryName "word/document.xml"
  if ([string]::IsNullOrWhiteSpace($documentText)) {
    throw "The docx package is missing word/document.xml: $resolvedDocxPath"
  }

  $relationshipsText = Get-ZipEntryText -Archive $archive -EntryName "word/_rels/document.xml.rels"

  [xml]$documentXml = $documentText
  $documentNamespaceManager = New-NamespaceManager -Document $documentXml
  $paragraphTexts = @(Get-ParagraphTexts -Document $documentXml -NamespaceManager $documentNamespaceManager)
  $allText = $paragraphTexts -join [Environment]::NewLine

  $imageDrawingCount = @($documentXml.SelectNodes("//a:blip[@r:embed or @r:link]", $documentNamespaceManager)).Count
  $imageRelationshipCount = 0
  if (-not [string]::IsNullOrWhiteSpace($relationshipsText)) {
    [xml]$relationshipsXml = $relationshipsText
    $relationshipNamespaceManager = New-NamespaceManager -Document $relationshipsXml
    $imageRelationshipCount = @($relationshipsXml.SelectNodes("//rel:Relationship[contains(@Type, '/image')]", $relationshipNamespaceManager)).Count
  }

  $captionTexts = @($paragraphTexts | Where-Object { $_ -match '^\s*\u56fe\s*\p{Nd}+' })
  $captionNumberMatches = New-Object System.Collections.Generic.List[object]
  foreach ($captionText in $captionTexts) {
    if ($captionText -match '^\s*\u56fe\s*(\p{Nd}+)') {
      $captionNumber = ConvertTo-IntegerFromDigitString -Text $matches[1]
      if ($null -ne $captionNumber) {
        [void]$captionNumberMatches.Add([pscustomobject]@{
            number = [int]$captionNumber
            text = $captionText
          })
      }
    }
  }

  if ($ExpectedImageCount -ge 0 -and $imageDrawingCount -ne $ExpectedImageCount) {
    Add-Finding -Findings $errors -Code "image-count-mismatch" -Message ("Expected {0} image drawing nodes, found {1}." -f $ExpectedImageCount, $imageDrawingCount)
  }

  if ($ExpectedCaptionCount -lt 0 -and $ExpectedImageCount -ge 0) {
    $ExpectedCaptionCount = $ExpectedImageCount
  }
  if ($ExpectedCaptionCount -ge 0 -and @($captionTexts).Count -ne $ExpectedCaptionCount) {
    Add-Finding -Findings $errors -Code "caption-count-mismatch" -Message ("Expected {0} figure captions, found {1}." -f $ExpectedCaptionCount, @($captionTexts).Count)
  }

  $captionNumbers = @($captionNumberMatches | ForEach-Object { [int]$_.number })
  $duplicateCaptionNumbers = @()
  $missingCaptionNumbers = @()
  $maxCaptionNumber = $null
  if ($captionNumbers.Count -gt 0) {
    $maxCaptionNumber = ($captionNumbers | Measure-Object -Maximum).Maximum
    $duplicateCaptionNumbers = @($captionNumbers | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { [int]$_.Name })
    $expectedCaptionNumbers = if ($ExpectedCaptionCount -ge 0) { 1..$ExpectedCaptionCount } else { 1..$maxCaptionNumber }
    $captionNumberSet = @{}
    foreach ($number in $captionNumbers) {
      $captionNumberSet[[int]$number] = $true
    }
    $missingCaptionNumbers = @($expectedCaptionNumbers | Where-Object { -not $captionNumberSet.ContainsKey([int]$_) })
  }
  if ($duplicateCaptionNumbers.Count -gt 0 -or $missingCaptionNumbers.Count -gt 0) {
    Add-Finding -Findings $errors -Code "caption-number-sequence" -Message ("Figure captions should be numbered continuously. Missing: {0}; duplicate: {1}." -f ($(if ($missingCaptionNumbers.Count -gt 0) { $missingCaptionNumbers -join ", " } else { "none" })), ($(if ($duplicateCaptionNumbers.Count -gt 0) { $duplicateCaptionNumbers -join ", " } else { "none" })))
  }

  if ($imageRelationshipCount -gt 0 -and $imageDrawingCount -ne $imageRelationshipCount) {
    Add-Finding -Findings $warnings -Code "image-relationship-drawing-mismatch" -Message ("Image relationship count is {0}, but document drawing count is {1}." -f $imageRelationshipCount, $imageDrawingCount)
  }

  $sectionGroups = foreach ($sectionField in (Get-ReportProfileSectionFields -Profile $reportProfile)) {
    [pscustomobject]@{
      label = [string]$sectionField.key
      heading = [string]$sectionField.heading
      patterns = @($sectionField.aliases | ForEach-Object { ConvertTo-SectionPattern -Text ([string]$_) })
    }
  }

  $sectionMatches = New-Object System.Collections.Generic.List[object]
  $missingSections = New-Object System.Collections.Generic.List[string]
  foreach ($group in $sectionGroups) {
    $matchedAlias = $null
    foreach ($pattern in $group.patterns) {
      if ($allText -match $pattern) {
        $matchedAlias = $pattern
        break
      }
    }

    [void]$sectionMatches.Add([pscustomobject]@{
        label = $group.label
        heading = $group.heading
        matched = (-not [string]::IsNullOrWhiteSpace($matchedAlias))
        matchedAlias = $matchedAlias
      })

    if ([string]::IsNullOrWhiteSpace($matchedAlias)) {
      [void]$missingSections.Add([string]$group.heading)
    }
  }

  if ($missingSections.Count -gt 0) {
    Add-Finding -Findings $warnings -Code "missing-common-sections" -Message ("Missing common report sections: {0}." -f ($missingSections -join ", "))
  }

  $placeholderPatterns = @(
    [pscustomobject]@{ name = "underscore-placeholder"; pattern = "_{3,}" },
    [pscustomobject]@{ name = "mustache-placeholder"; pattern = "\{\{[^}]+\}\}" },
    [pscustomobject]@{ name = "variable-placeholder"; pattern = "\$\{[^}]+\}" },
    [pscustomobject]@{ name = "todo-placeholder"; pattern = "(?i)\b(TODO|TBD)\b|\u5f85\u586b\u5199|\u5f85\u8865\u5145|\u8bf7\u586b\u5199|\u8bf7\u8f93\u5165" }
  )

  for ($index = 0; $index -lt $paragraphTexts.Count; $index++) {
    $paragraphText = [string]$paragraphTexts[$index]
    foreach ($placeholderPattern in $placeholderPatterns) {
      if ($paragraphText -match $placeholderPattern.pattern) {
        $excerpt = $paragraphText
        if ($excerpt.Length -gt 120) {
          $excerpt = $excerpt.Substring(0, 120) + "..."
        }
        [void]$placeholderMatches.Add([pscustomobject]@{
            paragraphIndex = $index + 1
            pattern = $placeholderPattern.name
            text = $excerpt
          })
      }
    }
  }

  if ($placeholderMatches.Count -gt 0) {
    Add-Finding -Findings $errors -Code "remaining-placeholders" -Message ("Found {0} possible unfilled placeholders." -f $placeholderMatches.Count)
  }

  $checkPassed = ($errors.Count -eq 0)
  $layoutMessage = New-LayoutCheckMessage `
    -Passed $checkPassed `
    -ActualImageCount $imageDrawingCount `
    -ExpectedImageCount $ExpectedImageCount `
    -ActualCaptionCount @($captionTexts).Count `
    -ExpectedCaptionCount $ExpectedCaptionCount `
    -PlaceholderCount $placeholderMatches.Count `
    -MissingSectionCount $missingSections.Count `
    -ErrorCount $errors.Count `
    -WarningCount $warnings.Count

  $result = [pscustomobject]@{
    docxPath = $resolvedDocxPath
    reportProfileName = [string]$reportProfile.name
    reportProfilePath = [string]$reportProfile.resolvedProfilePath
    passed = $checkPassed
    message = $layoutMessage
    summary = [pscustomobject]@{
      errorCount = $errors.Count
      warningCount = $warnings.Count
    }
    expected = [pscustomobject]@{
      imageCount = $(if ($ExpectedImageCount -ge 0) { $ExpectedImageCount } else { $null })
      captionCount = $(if ($ExpectedCaptionCount -ge 0) { $ExpectedCaptionCount } else { $null })
    }
    actual = [pscustomobject]@{
      imageDrawingCount = $imageDrawingCount
      imageRelationshipCount = $imageRelationshipCount
      captionCount = @($captionTexts).Count
      captions = $captionTexts
      captionNumbers = $captionNumbers
      paragraphCount = $paragraphTexts.Count
    }
    captionNumberCheck = [pscustomobject]@{
      maxNumber = $maxCaptionNumber
      missingNumbers = $missingCaptionNumbers
      duplicateNumbers = $duplicateCaptionNumbers
      continuous = ($duplicateCaptionNumbers.Count -eq 0 -and $missingCaptionNumbers.Count -eq 0)
    }
    sectionChecks = $sectionMatches
    missingSections = $missingSections
    placeholders = $placeholderMatches
    errors = $errors
    warnings = $warnings
  }

  if ($Format -eq "json") {
    $output = $result | ConvertTo-Json -Depth 8
  } else {
    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add("# DOCX Layout Check")
    [void]$lines.Add("")
    [void]$lines.Add("- Docx: $resolvedDocxPath")
    [void]$lines.Add("- Passed: $($result.passed)")
    [void]$lines.Add("- Message: $layoutMessage")
    [void]$lines.Add("- Images: $imageDrawingCount")
    [void]$lines.Add("- Captions: $(@($captionTexts).Count)")
    [void]$lines.Add("- Errors: $($errors.Count)")
    [void]$lines.Add("- Warnings: $($warnings.Count)")
    if ($errors.Count -gt 0) {
      [void]$lines.Add("")
      [void]$lines.Add("## Errors")
      foreach ($finding in $errors) {
        [void]$lines.Add("- $($finding.code): $($finding.message)")
      }
    }
    if ($warnings.Count -gt 0) {
      [void]$lines.Add("")
      [void]$lines.Add("## Warnings")
      foreach ($finding in $warnings) {
        [void]$lines.Add("- $($finding.code): $($finding.message)")
      }
    }
    $output = $lines -join [Environment]::NewLine
  }

  if ([string]::IsNullOrWhiteSpace($OutFile)) {
    Write-Output $output
  } else {
    $parent = Split-Path -Parent $OutFile
    if (-not [string]::IsNullOrWhiteSpace($parent)) {
      New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    [System.IO.File]::WriteAllText(([System.IO.Path]::GetFullPath($OutFile)), $output, (New-Object System.Text.UTF8Encoding($true)))
    Write-Output "Wrote layout check to $OutFile"
  }
} finally {
  $archive.Dispose()
}
