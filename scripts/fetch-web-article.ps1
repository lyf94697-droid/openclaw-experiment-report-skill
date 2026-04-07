param(
  [Parameter(Mandatory = $true)]
  [string]$Url,

  [string]$BrowserProfile = $env:OPENCLAW_BROWSER_PROFILE,

  [string]$OpenClawCmd = $env:OPENCLAW_CMD,

  [int]$MaxChars = 30000,

  [int]$TimeoutMs = 30000
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($BrowserProfile)) {
  $BrowserProfile = "openclaw"
}

function Resolve-OpenClawCommand {
  param(
    [string]$Candidate
  )

  if (-not [string]::IsNullOrWhiteSpace($Candidate)) {
    if (Test-Path $Candidate) {
      return (Resolve-Path $Candidate).Path
    }
    throw "OPENCLAW_CMD does not exist: $Candidate"
  }

  foreach ($name in @("openclaw.cmd", "openclaw")) {
    $cmd = Get-Command $name -ErrorAction SilentlyContinue
    if ($null -ne $cmd -and $cmd.Source) {
      return $cmd.Source
    }
  }

  throw "OpenClaw CLI not found. Set OPENCLAW_CMD or add openclaw.cmd to PATH."
}

function Parse-JsonFromOutput {
  param(
    [string]$Text
  )

  $trimmed = $Text.Trim()
  if ([string]::IsNullOrWhiteSpace($trimmed)) {
    throw "Expected JSON output, got empty text."
  }

  try {
    return $trimmed | ConvertFrom-Json
  } catch {
    $starts = @($trimmed.IndexOf("{"), $trimmed.IndexOf("["))
    $start = ($starts | Where-Object { $_ -ge 0 } | Measure-Object -Minimum).Minimum
    if ($null -eq $start) {
      throw "Failed to locate JSON in output: $trimmed"
    }
    return $trimmed.Substring([int]$start) | ConvertFrom-Json
  }
}

$cli = Resolve-OpenClawCommand -Candidate $OpenClawCmd
$uri = [Uri]$Url

$status = (& $cli browser status --browser-profile $BrowserProfile 2>&1 | Out-String).Trim()
if ($status -notmatch "running:\s*true") {
  & $cli browser start --browser-profile $BrowserProfile | Out-Null
}

$openRaw = (& $cli browser --browser-profile $BrowserProfile open $Url --json 2>&1 | Out-String).Trim()
$opened = Parse-JsonFromOutput -Text $openRaw
$targetId = $opened.targetId
if ([string]::IsNullOrWhiteSpace($targetId)) {
  throw "Failed to open URL: $Url"
}

& $cli browser --browser-profile $BrowserProfile focus $targetId | Out-Null
try {
  & $cli browser --browser-profile $BrowserProfile wait --target-id $targetId --load networkidle --timeout-ms $TimeoutMs | Out-Null
} catch {
  Start-Sleep -Seconds 3
}

# Keep using the target id returned by browser open/focus.
# Avoid reparsing the global tabs list here because page titles from unrelated
# tabs can contain unescaped quotes and break ConvertFrom-Json in legacy shells.

$titleFn = "() => document.title || ''"
$contentFn = "() => { const selectors = ['article','main article','[role=""main""] article','main','[role=""main""]','.article-content','.article-content-box','.blog-content-box','.markdown-body','.post-content','.entry-content','#content','.content','.article','.post']; let text = ''; for (const selector of selectors) { const node = document.querySelector(selector); if (node && node.innerText && node.innerText.trim().length > 200) { text = node.innerText; break; } } if (!text) { text = document.body ? document.body.innerText : ''; } return text.replace(/\u00a0/g, ' ').replace(/\r/g, '').replace(/\n{3,}/g, '\n\n').trim().slice(0, $MaxChars); }"

$titleRaw = (& $cli browser --browser-profile $BrowserProfile evaluate --target-id $targetId --fn $titleFn 2>&1 | Out-String).Trim()
$contentRaw = (& $cli browser --browser-profile $BrowserProfile evaluate --target-id $targetId --fn $contentFn 2>&1 | Out-String).Trim()

$title = $titleRaw.Trim('"')
$content = $contentRaw.Trim('"')
$content = $content -replace "\\n", "`n"
$content = $content -replace '\\"', '"'
$content = $content -replace "\\t", "`t"

Write-Output "TITLE: $title"
Write-Output "URL: $Url"
Write-Output "TARGET: $targetId"
Write-Output ""
Write-Output $content
