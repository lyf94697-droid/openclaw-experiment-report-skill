[CmdletBinding()]
param(
  [string]$Path,

  [string]$Text,

  [string]$RequirementsPath,

  [string]$RequirementsJson,

  [ValidateSet("markdown", "json")]
  [string]$Format = "markdown",

  [string]$OutFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

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

function Get-StringList {
  param(
    [AllowNull()]
    [object]$Value
  )

  $items = New-Object System.Collections.Generic.List[string]

  if ($null -eq $Value) {
    return @()
  }

  if (($Value -is [System.Collections.IEnumerable]) -and ($Value -isnot [string])) {
    foreach ($item in $Value) {
      foreach ($text in (Get-StringList -Value $item)) {
        if (-not [string]::IsNullOrWhiteSpace($text)) {
          [void]$items.Add($text)
        }
      }
    }
    return $items.ToArray()
  }

  $textValue = [string]$Value
  if (-not [string]::IsNullOrWhiteSpace($textValue)) {
    [void]$items.Add($textValue.Trim())
  }

  return $items.ToArray()
}

function Get-ReportText {
  param(
    [AllowNull()]
    [string]$TextPath,

    [AllowNull()]
    [string]$InlineText
  )

  if ([string]::IsNullOrWhiteSpace($TextPath) -eq [string]::IsNullOrWhiteSpace($InlineText)) {
    throw "Provide exactly one of -Path or -Text."
  }

  if (-not [string]::IsNullOrWhiteSpace($TextPath)) {
    $resolvedPath = Resolve-Path -LiteralPath $TextPath
    return @{
      Source = $resolvedPath.Path
      Text = Get-Content -LiteralPath $resolvedPath.Path -Raw -Encoding UTF8
    }
  }

  return @{
    Source = "[inline text]"
    Text = $InlineText
  }
}

function Get-RequirementsRoot {
  param(
    [AllowNull()]
    [string]$ConfigPath,

    [AllowNull()]
    [string]$ConfigJson
  )

  if ([string]::IsNullOrWhiteSpace($ConfigPath) -and [string]::IsNullOrWhiteSpace($ConfigJson)) {
    $ConfigJson = @'
{
  "courseName": "",
  "experimentName": "",
  "minChars": 600,
  "sections": [
    { "name": "\u5B9E\u9A8C\u76EE\u7684", "aliases": ["\u5B9E\u9A8C\u76EE\u7684"], "minChars": 40 },
    { "name": "\u5B9E\u9A8C\u73AF\u5883", "aliases": ["\u5B9E\u9A8C\u73AF\u5883", "\u5B9E\u9A8C\u8BBE\u5907\u4E0E\u73AF\u5883"], "minChars": 40 },
    { "name": "\u5B9E\u9A8C\u539F\u7406\u6216\u4EFB\u52A1\u8981\u6C42", "aliases": ["\u5B9E\u9A8C\u539F\u7406\u6216\u4EFB\u52A1\u8981\u6C42", "\u5B9E\u9A8C\u539F\u7406", "\u4EFB\u52A1\u8981\u6C42"], "minChars": 40 },
    { "name": "\u5B9E\u9A8C\u6B65\u9AA4", "aliases": ["\u5B9E\u9A8C\u6B65\u9AA4", "\u5B9E\u9A8C\u8FC7\u7A0B", "\u5B9E\u9A8C\u6B65\u9AA4 / \u5173\u952E\u4EE3\u7801\u8BF4\u660E"], "minChars": 80 },
    { "name": "\u5B9E\u9A8C\u7ED3\u679C", "aliases": ["\u5B9E\u9A8C\u7ED3\u679C", "\u5B9E\u9A8C\u73B0\u8C61\u4E0E\u7ED3\u679C\u8BB0\u5F55"], "minChars": 60 },
    { "name": "\u95EE\u9898\u5206\u6790", "aliases": ["\u95EE\u9898\u5206\u6790", "\u7ED3\u679C\u5206\u6790"], "minChars": 40 },
    { "name": "\u5B9E\u9A8C\u603B\u7ED3", "aliases": ["\u5B9E\u9A8C\u603B\u7ED3", "\u603B\u7ED3\u4E0E\u601D\u8003"], "minChars": 40 }
  ],
  "forbiddenPatterns": [
    "TODO",
    "\u5F85\u8865\u5145",
    "\u81EA\u884C\u586B\u5199",
    "\u53EF\u6839\u636E\u5B9E\u9645\u60C5\u51B5\u4FEE\u6539",
    "\u793A\u4F8B",
    "\u6837\u4F8B",
    "ChatGPT",
    "Claude",
    "AI\u751F\u6210"
  ]
}
'@
  } elseif ([string]::IsNullOrWhiteSpace($ConfigPath) -eq [string]::IsNullOrWhiteSpace($ConfigJson)) {
    throw "Provide zero or one of -RequirementsPath and -RequirementsJson."
  }

  if (-not [string]::IsNullOrWhiteSpace($ConfigPath)) {
    $resolvedPath = Resolve-Path -LiteralPath $ConfigPath
    return ConvertTo-PlainHashtable -InputObject ((Get-Content -LiteralPath $resolvedPath.Path -Raw -Encoding UTF8) | ConvertFrom-Json)
  }

  return ConvertTo-PlainHashtable -InputObject ($ConfigJson | ConvertFrom-Json)
}

function Test-TextContains {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Haystack,

    [Parameter(Mandatory = $true)]
    [string]$Needle
  )

  return $Haystack.ToLowerInvariant().Contains($Needle.ToLowerInvariant())
}

function Get-HeadingPattern {
  param(
    [Parameter(Mandatory = $true)]
    [string[]]$Aliases
  )

  $escapedAliases = foreach ($alias in $Aliases) {
    [regex]::Escape($alias)
  }

  return '^(?:\u7B2C?[0-9\u4E00\u4E8C\u4E09\u56DB\u4E94\u516D\u4E03\u516B\u4E5D\u5341]+(?:\u7AE0|\u8282)?[.\)\u3001]?\s*)?(?:' + ($escapedAliases -join '|') + ')\s*(?:[:\uFF1A])?$'
}

function Add-Finding {
  param(
    [Parameter(Mandatory = $true)]
    [object]$Target,

    [Parameter(Mandatory = $true)]
    [string]$Severity,

    [Parameter(Mandatory = $true)]
    [string]$Code,

    [Parameter(Mandatory = $true)]
    [string]$Message
  )

  $Target.Add([pscustomobject]@{
      severity = $Severity
      code = $Code
      message = $Message
    }) | Out-Null
}

function New-SectionRule {
  param(
    [Parameter(Mandatory = $true)]
    [object]$Rule
  )

  $table = ConvertTo-PlainHashtable -InputObject $Rule
  if (-not $table.ContainsKey("name")) {
    throw "Section rules must include a name."
  }

  $name = [string]$table["name"]
  $aliases = @(Get-StringList -Value $table["aliases"])
  if ($aliases.Count -eq 0) {
    $aliases = @($name)
  }

  $minChars = 0
  if ($table.ContainsKey("minChars") -and -not [string]::IsNullOrWhiteSpace([string]$table["minChars"])) {
    $minChars = [int]$table["minChars"]
  }

  return [pscustomobject]@{
    name = $name
    aliases = $aliases
    minChars = $minChars
    headingPattern = Get-HeadingPattern -Aliases $aliases
  }
}

function Normalize-ContentChars {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Text
  )

  return (($Text -replace '\s+', '')).Length
}

$reportInfo = Get-ReportText -TextPath $Path -InlineText $Text
$reportText = [string]$reportInfo.Text
$requirementsRoot = Get-RequirementsRoot -ConfigPath $RequirementsPath -ConfigJson $RequirementsJson

$findings = New-Object System.Collections.Generic.List[object]
$lines = @($reportText -split "\r?\n")
$trimmedLines = @(foreach ($line in $lines) { $line.Trim() })
$normalizedReportChars = Normalize-ContentChars -Text $reportText

$sectionRules = New-Object System.Collections.Generic.List[object]
if ($requirementsRoot.ContainsKey("sections")) {
  foreach ($sectionRule in $requirementsRoot["sections"]) {
    $sectionRules.Add((New-SectionRule -Rule $sectionRule)) | Out-Null
  }
}

$sectionMatches = New-Object System.Collections.Generic.List[object]
foreach ($sectionRule in $sectionRules) {
  $headingIndex = $null
  $headingText = $null
  for ($lineIndex = 0; $lineIndex -lt $trimmedLines.Count; $lineIndex++) {
    $lineText = $trimmedLines[$lineIndex]
    if ([string]::IsNullOrWhiteSpace($lineText)) {
      continue
    }

    if ($lineText -match $sectionRule.headingPattern) {
      $headingIndex = $lineIndex
      $headingText = $lineText
      break
    }
  }

  $sectionMatches.Add([pscustomobject]@{
      name = $sectionRule.name
      aliases = $sectionRule.aliases
      minChars = $sectionRule.minChars
      headingIndex = $headingIndex
      headingText = $headingText
      contentChars = 0
      found = ($null -ne $headingIndex)
    }) | Out-Null
}

$orderedFoundSections = @($sectionMatches | Where-Object { $_.found } | Sort-Object headingIndex)
for ($matchIndex = 0; $matchIndex -lt $orderedFoundSections.Count; $matchIndex++) {
  $current = $orderedFoundSections[$matchIndex]
  $nextHeadingIndex = if (($matchIndex + 1) -lt $orderedFoundSections.Count) { $orderedFoundSections[$matchIndex + 1].headingIndex } else { $lines.Count }
  $contentLines = if (($current.headingIndex + 1) -lt $nextHeadingIndex) { $lines[($current.headingIndex + 1)..($nextHeadingIndex - 1)] } else { @() }
  $contentText = ($contentLines -join [Environment]::NewLine).Trim()
  $current.contentChars = Normalize-ContentChars -Text $contentText
}

foreach ($sectionMatch in $sectionMatches) {
  if (-not $sectionMatch.found) {
    Add-Finding -Target $findings -Severity "error" -Code "missing-section" -Message ("Missing required section: {0}" -f $sectionMatch.name)
    continue
  }

  if ($sectionMatch.minChars -gt 0 -and $sectionMatch.contentChars -lt $sectionMatch.minChars) {
    Add-Finding -Target $findings -Severity "error" -Code "short-section" -Message ("Section '{0}' is too short: {1} chars < {2}" -f $sectionMatch.name, $sectionMatch.contentChars, $sectionMatch.minChars)
  }
}

$courseName = if ($requirementsRoot.ContainsKey("courseName")) { [string]$requirementsRoot["courseName"] } else { "" }
if (-not [string]::IsNullOrWhiteSpace($courseName) -and -not (Test-TextContains -Haystack $reportText -Needle $courseName)) {
  Add-Finding -Target $findings -Severity "error" -Code "missing-course-name" -Message ("Report does not mention the course name: {0}" -f $courseName)
}

$experimentName = if ($requirementsRoot.ContainsKey("experimentName")) { [string]$requirementsRoot["experimentName"] } else { "" }
if (-not [string]::IsNullOrWhiteSpace($experimentName) -and -not (Test-TextContains -Haystack $reportText -Needle $experimentName)) {
  Add-Finding -Target $findings -Severity "error" -Code "missing-experiment-name" -Message ("Report does not mention the experiment name: {0}" -f $experimentName)
}

foreach ($keyword in (Get-StringList -Value $requirementsRoot["requiredKeywords"])) {
  if (-not (Test-TextContains -Haystack $reportText -Needle $keyword)) {
    Add-Finding -Target $findings -Severity "error" -Code "missing-keyword" -Message ("Report is missing required keyword: {0}" -f $keyword)
  }
}

foreach ($phrase in (Get-StringList -Value $requirementsRoot["requiredPhrases"])) {
  if (-not (Test-TextContains -Haystack $reportText -Needle $phrase)) {
    Add-Finding -Target $findings -Severity "error" -Code "missing-required-phrase" -Message ("Report is missing required phrase: {0}" -f $phrase)
  }
}

foreach ($pattern in (Get-StringList -Value $requirementsRoot["forbiddenPatterns"])) {
  if ([string]::IsNullOrWhiteSpace($pattern)) {
    continue
  }

  if ($reportText -match [regex]::Escape($pattern)) {
    Add-Finding -Target $findings -Severity "error" -Code "forbidden-pattern" -Message ("Report contains forbidden pattern: {0}" -f $pattern)
  }
}

$minChars = 0
if ($requirementsRoot.ContainsKey("minChars") -and -not [string]::IsNullOrWhiteSpace([string]$requirementsRoot["minChars"])) {
  $minChars = [int]$requirementsRoot["minChars"]
}
if ($minChars -gt 0 -and $normalizedReportChars -lt $minChars) {
  Add-Finding -Target $findings -Severity "error" -Code "short-report" -Message ("Report is too short: {0} chars < {1}" -f $normalizedReportChars, $minChars)
}

$minFigureRefs = 0
if ($requirementsRoot.ContainsKey("minFigureRefs") -and -not [string]::IsNullOrWhiteSpace([string]$requirementsRoot["minFigureRefs"])) {
  $minFigureRefs = [int]$requirementsRoot["minFigureRefs"]
}
if ($minFigureRefs -gt 0) {
  $figureRefCount = @([regex]::Matches($reportText, '\u56FE\s*\d+')).Count
  if ($figureRefCount -lt $minFigureRefs) {
    Add-Finding -Target $findings -Severity "error" -Code "missing-figure-refs" -Message ("Report references too few figures: {0} < {1}" -f $figureRefCount, $minFigureRefs)
  }
}

$errorCount = @($findings | Where-Object { $_.severity -eq "error" }).Count
$warningCount = @($findings | Where-Object { $_.severity -eq "warning" }).Count

$result = [pscustomobject]@{
  source = $reportInfo.Source
  passed = ($errorCount -eq 0)
  summary = [pscustomobject]@{
    charCount = $normalizedReportChars
    sectionCount = $sectionRules.Count
    foundSectionCount = @($sectionMatches | Where-Object { $_.found }).Count
    errorCount = $errorCount
    warningCount = $warningCount
  }
  sections = $sectionMatches
  findings = $findings
}

if ($Format -eq "json") {
  $output = $result | ConvertTo-Json -Depth 8
} else {
  $linesOut = New-Object System.Collections.Generic.List[string]
  [void]$linesOut.Add("# Report Validation")
  [void]$linesOut.Add("")
  [void]$linesOut.Add("- Source: $($result.source)")
  [void]$linesOut.Add("- Passed: $($result.passed)")
  [void]$linesOut.Add("- Char count: $($result.summary.charCount)")
  [void]$linesOut.Add("- Found sections: $($result.summary.foundSectionCount)/$($result.summary.sectionCount)")
  [void]$linesOut.Add("- Errors: $($result.summary.errorCount)")
  [void]$linesOut.Add("- Warnings: $($result.summary.warningCount)")
  [void]$linesOut.Add("")
  [void]$linesOut.Add("## Section Coverage")

  foreach ($sectionMatch in $sectionMatches) {
    $status = if ($sectionMatch.found) { "found" } else { "missing" }
    [void]$linesOut.Add("- $($sectionMatch.name): $status; content chars = $($sectionMatch.contentChars)")
  }

  [void]$linesOut.Add("")
  [void]$linesOut.Add("## Findings")
  if ($findings.Count -eq 0) {
    [void]$linesOut.Add("- No validation findings")
  } else {
    foreach ($finding in $findings) {
      [void]$linesOut.Add("- [$($finding.severity)] $($finding.code): $($finding.message)")
    }
  }

  $output = $linesOut -join [Environment]::NewLine
}

if ([string]::IsNullOrWhiteSpace($OutFile)) {
  Write-Output $output
} else {
  [System.IO.File]::WriteAllText($OutFile, $output, (New-Object System.Text.UTF8Encoding($true)))
  Write-Output "Wrote report validation to $OutFile"
}
