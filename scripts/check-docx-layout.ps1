[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$DocxPath,

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

$resolvedDocxPath = (Resolve-Path -LiteralPath $DocxPath).Path
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

  if ($ExpectedImageCount -ge 0 -and $imageDrawingCount -ne $ExpectedImageCount) {
    Add-Finding -Findings $errors -Code "image-count-mismatch" -Message ("Expected {0} image drawing nodes, found {1}." -f $ExpectedImageCount, $imageDrawingCount)
  }

  if ($ExpectedCaptionCount -lt 0 -and $ExpectedImageCount -ge 0) {
    $ExpectedCaptionCount = $ExpectedImageCount
  }
  if ($ExpectedCaptionCount -ge 0 -and @($captionTexts).Count -ne $ExpectedCaptionCount) {
    Add-Finding -Findings $errors -Code "caption-count-mismatch" -Message ("Expected {0} figure captions, found {1}." -f $ExpectedCaptionCount, @($captionTexts).Count)
  }

  if ($imageRelationshipCount -gt 0 -and $imageDrawingCount -ne $imageRelationshipCount) {
    Add-Finding -Findings $warnings -Code "image-relationship-drawing-mismatch" -Message ("Image relationship count is {0}, but document drawing count is {1}." -f $imageRelationshipCount, $imageDrawingCount)
  }

  $sectionGroups = @(
    [pscustomobject]@{ label = "purpose"; patterns = @("\u5b9e\u9a8c\u76ee\u7684") },
    [pscustomobject]@{ label = "environment"; patterns = @("\u5b9e\u9a8c\u73af\u5883", "\u5b9e\u9a8c\u8bbe\u5907", "\u5b9e\u9a8c\u5e73\u53f0") },
    [pscustomobject]@{ label = "theory-or-task"; patterns = @("\u5b9e\u9a8c\u539f\u7406", "\u4efb\u52a1\u8981\u6c42", "\u5b9e\u9a8c\u5185\u5bb9", "\u5b9e\u9a8c\u8981\u6c42") },
    [pscustomobject]@{ label = "steps"; patterns = @("\u5b9e\u9a8c\u6b65\u9aa4", "\u5b9e\u9a8c\u8fc7\u7a0b") },
    [pscustomobject]@{ label = "results"; patterns = @("\u5b9e\u9a8c\u7ed3\u679c", "\u5b9e\u9a8c\u73b0\u8c61", "\u7ed3\u679c\u8bb0\u5f55") },
    [pscustomobject]@{ label = "analysis"; patterns = @("\u95ee\u9898\u5206\u6790", "\u7ed3\u679c\u5206\u6790", "\u5b9e\u9a8c\u5206\u6790") },
    [pscustomobject]@{ label = "summary"; patterns = @("\u5b9e\u9a8c\u603b\u7ed3", "\u5b9e\u9a8c\u5c0f\u7ed3", "\u603b\u7ed3\u4e0e\u601d\u8003") }
  )

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
        matched = (-not [string]::IsNullOrWhiteSpace($matchedAlias))
        matchedAlias = $matchedAlias
      })

    if ([string]::IsNullOrWhiteSpace($matchedAlias)) {
      [void]$missingSections.Add([string]$group.label)
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

  $result = [pscustomobject]@{
    docxPath = $resolvedDocxPath
    passed = ($errors.Count -eq 0)
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
      paragraphCount = $paragraphTexts.Count
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
