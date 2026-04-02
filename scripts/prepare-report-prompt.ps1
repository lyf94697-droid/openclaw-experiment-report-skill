[CmdletBinding()]
param(
  [string]$PromptPath,

  [string]$PromptText,

  [string[]]$ReferenceTextPaths,

  [string[]]$ReferenceUrls,

  [string]$OutFile,

  [string]$BrowserProfile = $env:OPENCLAW_BROWSER_PROFILE,

  [string]$OpenClawCmd = $env:OPENCLAW_CMD,

  [int]$ReferenceMaxChars = 30000
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Parse-ReferenceDocument {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Text,

    [Parameter(Mandatory = $true)]
    [string]$FallbackSource
  )

  $trimmed = $Text.Trim()
  if ([string]::IsNullOrWhiteSpace($trimmed)) {
    throw "Reference content is empty: $FallbackSource"
  }

  $lines = $trimmed -split "\r?\n"
  $currentIndex = 0
  $title = $null
  $url = $null

  if ($lines.Length -gt $currentIndex -and $lines[$currentIndex] -match '^TITLE:\s*(.*)$') {
    $title = $Matches[1].Trim()
    $currentIndex++
  }

  if ($lines.Length -gt $currentIndex -and $lines[$currentIndex] -match '^URL:\s*(.*)$') {
    $url = $Matches[1].Trim()
    $currentIndex++
  }

  if ($lines.Length -gt $currentIndex -and $lines[$currentIndex] -match '^TARGET:\s*(.*)$') {
    $currentIndex++
  }

  while ($lines.Length -gt $currentIndex -and [string]::IsNullOrWhiteSpace($lines[$currentIndex])) {
    $currentIndex++
  }

  $content = if ($currentIndex -lt $lines.Length) {
    (($lines[$currentIndex..($lines.Length - 1)] -join [Environment]::NewLine).Trim())
  } else {
    ""
  }

  if ([string]::IsNullOrWhiteSpace($content)) {
    throw "Reference content body is empty: $FallbackSource"
  }

  return [pscustomobject]@{
    source = $(if (-not [string]::IsNullOrWhiteSpace($url)) { $url } else { $FallbackSource })
    title = $title
    content = $content
  }
}

if ([string]::IsNullOrWhiteSpace($PromptPath) -eq [string]::IsNullOrWhiteSpace($PromptText)) {
  throw "Provide exactly one of -PromptPath or -PromptText."
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$basePromptText = if (-not [string]::IsNullOrWhiteSpace($PromptPath)) {
  Get-Content -LiteralPath (Resolve-Path -LiteralPath $PromptPath).Path -Raw -Encoding UTF8
} else {
  $PromptText
}

if ([string]::IsNullOrWhiteSpace($OutFile)) {
  $OutFile = Join-Path $repoRoot ("tests-output\prepared-prompt-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".txt")
}

$resolvedOutFile = [System.IO.Path]::GetFullPath($OutFile)
$outDir = Split-Path -Parent $resolvedOutFile
if (-not [string]::IsNullOrWhiteSpace($outDir)) {
  New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

$referenceDocuments = New-Object System.Collections.Generic.List[object]

foreach ($referenceTextPath in @($ReferenceTextPaths)) {
  if ([string]::IsNullOrWhiteSpace($referenceTextPath)) {
    continue
  }

  $resolvedReferencePath = (Resolve-Path -LiteralPath $referenceTextPath).Path
  $referenceText = Get-Content -LiteralPath $resolvedReferencePath -Raw -Encoding UTF8
  [void]$referenceDocuments.Add((Parse-ReferenceDocument -Text $referenceText -FallbackSource $resolvedReferencePath))
}

foreach ($referenceUrl in @($ReferenceUrls)) {
  if ([string]::IsNullOrWhiteSpace($referenceUrl)) {
    continue
  }

  $fetchedReference = & (Join-Path $PSScriptRoot "fetch-web-article.ps1") `
    -Url $referenceUrl `
    -BrowserProfile $BrowserProfile `
    -OpenClawCmd $OpenClawCmd `
    -MaxChars $ReferenceMaxChars
  $fetchedText = ($fetchedReference | Out-String)
  [void]$referenceDocuments.Add((Parse-ReferenceDocument -Text $fetchedText -FallbackSource $referenceUrl))
}

$assembledParts = New-Object System.Collections.Generic.List[string]
[void]$assembledParts.Add($basePromptText.Trim())

if ($referenceDocuments.Count -gt 0) {
  [void]$assembledParts.Add("")
  [void]$assembledParts.Add("The following reference materials may be used to complete theory, task requirements, procedure details, and background context. Do not copy them verbatim. If they conflict with user-provided screenshots, outputs, data, or explicit facts, trust the user facts. If real results are missing, do not fabricate them.")

  for ($index = 0; $index -lt $referenceDocuments.Count; $index++) {
    $reference = $referenceDocuments[$index]
    [void]$assembledParts.Add("")
    [void]$assembledParts.Add(("Reference Material {0}" -f ($index + 1)))
    [void]$assembledParts.Add(("Source: {0}" -f ([string]$reference.source)))
    if (-not [string]::IsNullOrWhiteSpace([string]$reference.title)) {
      [void]$assembledParts.Add(("Title: {0}" -f ([string]$reference.title)))
    }
    [void]$assembledParts.Add("Content:")
    [void]$assembledParts.Add([string]$reference.content)
  }
}

$preparedPrompt = ($assembledParts -join [Environment]::NewLine).Trim() + [Environment]::NewLine
[System.IO.File]::WriteAllText($resolvedOutFile, $preparedPrompt, (New-Object System.Text.UTF8Encoding($true)))

[pscustomobject]@{
  outPath = $resolvedOutFile
  referenceCount = $referenceDocuments.Count
  sources = @($referenceDocuments | ForEach-Object { $_.source })
}
