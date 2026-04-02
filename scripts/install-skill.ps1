[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [string]$AgentsHome = $env:AGENTS_HOME,

  [string]$OpenClawHome = $env:OPENCLAW_HOME,

  [string]$TargetDir,

  [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Copy-RepositoryContent {
  param(
    [Parameter(Mandatory = $true)]
    [string]$SourceRoot,

    [Parameter(Mandatory = $true)]
    [string]$DestinationRoot
  )

  $excludeNames = @(".git", "local-inputs", "outputs", "tests-output")
  $items = Get-ChildItem -LiteralPath $SourceRoot -Force | Where-Object { $_.Name -notin $excludeNames }
  foreach ($item in $items) {
    $destination = Join-Path $DestinationRoot $item.Name
    Copy-Item -LiteralPath $item.FullName -Destination $destination -Recurse -Force
  }
}

function Get-BackupDirectory {
  param(
    [Parameter(Mandatory = $true)]
    [string]$InstallDir,

    [Parameter(Mandatory = $true)]
    [string]$InstallParent
  )

  $timestamp = Get-Date -Format "yyyyMMddHHmmss"
  $installLeaf = Split-Path -Leaf $InstallDir
  $backupRoot = $InstallParent

  if ((Split-Path -Leaf $InstallParent).ToLowerInvariant() -eq "skills") {
    $backupRoot = Join-Path (Split-Path -Parent $InstallParent) "skill-backups"
  }

  if (-not (Test-Path -LiteralPath $backupRoot)) {
    New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
  }

  return (Join-Path $backupRoot ("{0}.bak-{1}" -f $installLeaf, $timestamp))
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
if ([string]::IsNullOrWhiteSpace($TargetDir)) {
  if (-not [string]::IsNullOrWhiteSpace($AgentsHome)) {
    $TargetDir = Join-Path (Join-Path $AgentsHome "skills") "experiment-report"
  } elseif (-not [string]::IsNullOrWhiteSpace($OpenClawHome)) {
    # Backward-compatible fallback for older local layouts.
    $TargetDir = Join-Path (Join-Path $OpenClawHome "skills") "experiment-report"
  } else {
    $TargetDir = Join-Path (Join-Path (Join-Path $HOME ".agents") "skills") "experiment-report"
  }
}

$targetParent = Split-Path -Parent $TargetDir
if ([string]::IsNullOrWhiteSpace($targetParent)) {
  throw "TargetDir must include a parent directory."
}

if (-not (Test-Path -LiteralPath $targetParent)) {
  if ($PSCmdlet.ShouldProcess($targetParent, "Create parent directory")) {
    New-Item -ItemType Directory -Path $targetParent -Force | Out-Null
  }
}

if (Test-Path -LiteralPath $TargetDir) {
  if (-not $Force) {
    throw "Target directory already exists: $TargetDir. Re-run with -Force to back it up and reinstall."
  }

  $backupDir = Get-BackupDirectory -InstallDir $TargetDir -InstallParent $targetParent
  if ($PSCmdlet.ShouldProcess($TargetDir, "Move existing install to $backupDir")) {
    Move-Item -LiteralPath $TargetDir -Destination $backupDir
  }
}

if ($PSCmdlet.ShouldProcess($TargetDir, "Install skill files")) {
  New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
  Copy-RepositoryContent -SourceRoot $repoRoot -DestinationRoot $TargetDir
}

Write-Output "Installed skill to $TargetDir"
