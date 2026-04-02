param(
  [string]$BrowserProfile = $env:OPENCLAW_BROWSER_PROFILE,
  [string]$OpenClawCmd = $env:OPENCLAW_CMD
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

$cli = Resolve-OpenClawCommand -Candidate $OpenClawCmd

Write-Output "OpenClaw CLI: $cli"
Write-Output "Browser profile: $BrowserProfile"
Write-Output ""
Write-Output "browser status:"
& $cli browser status --browser-profile $BrowserProfile

