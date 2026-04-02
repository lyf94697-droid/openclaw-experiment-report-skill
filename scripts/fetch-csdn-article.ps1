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

& (Join-Path $PSScriptRoot "fetch-web-article.ps1") `
  -Url $Url `
  -BrowserProfile $BrowserProfile `
  -OpenClawCmd $OpenClawCmd `
  -MaxChars $MaxChars `
  -TimeoutMs $TimeoutMs
