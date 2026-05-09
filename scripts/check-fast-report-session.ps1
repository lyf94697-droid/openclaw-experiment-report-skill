[CmdletBinding()]
param(
  [string[]]$TemplatePath = @(),

  [string]$CachePath,

  [int]$MaxAgeHours = 24,

  [switch]$RunSmokeWhenNeeded,

  [switch]$RecordPass,

  [switch]$Json
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path

if ([string]::IsNullOrWhiteSpace($CachePath)) {
  $CachePath = [System.IO.Path]::Combine($repoRoot, "tests-output", "fast-report-session-cache.json")
}

function Ensure-ParentDirectory {
  param([Parameter(Mandatory = $true)][string]$Path)

  $parent = Split-Path -Parent $Path
  if (-not [string]::IsNullOrWhiteSpace($parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
  }
}

function Get-StringHash {
  param([Parameter(Mandatory = $true)][string]$Text)

  $sha = [System.Security.Cryptography.SHA256]::Create()
  try {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    $hash = $sha.ComputeHash($bytes)
    return (($hash | ForEach-Object { $_.ToString("x2") }) -join "")
  } finally {
    $sha.Dispose()
  }
}

function Add-ExistingFile {
  param(
    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [System.Collections.Generic.List[string]]$Files,

    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  if (Test-Path -LiteralPath $Path -PathType Leaf) {
    [void]$Files.Add((Resolve-Path -LiteralPath $Path).Path)
  }
}

function Join-RepoRelativePath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$RelativePath
  )

  $parts = @($RelativePath -split '[\\/]+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
  return [System.IO.Path]::Combine([string[]](@($repoRoot) + $parts))
}

function Add-ExistingFilesFromDirectory {
  param(
    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [System.Collections.Generic.List[string]]$Files,

    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [string[]]$Include
  )

  if (Test-Path -LiteralPath $Path -PathType Container) {
    foreach ($file in Get-ChildItem -LiteralPath $Path -Recurse -File -Include $Include) {
      [void]$Files.Add($file.FullName)
    }
  }
}

function Get-RelativeOrFullPath {
  param([Parameter(Mandatory = $true)][string]$Path)

  if ($Path.StartsWith($repoRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    return $Path.Substring($repoRoot.Length).TrimStart("\", "/")
  }

  return $Path
}

function Get-FastSessionFingerprint {
  param([string[]]$ExtraTemplatePath)

  $files = New-Object System.Collections.Generic.List[string]

  foreach ($relativePath in @(
      "SKILL.md",
      "scripts/build-report.ps1",
      "scripts/build-report-from-feishu.ps1",
      "scripts/build-report-from-url.ps1",
      "scripts/check-docx-layout.ps1",
      "scripts/convert-docx-template-frame.ps1",
      "scripts/format-docx-report-style.ps1",
      "scripts/generate-docx-field-map.ps1",
      "scripts/generate-docx-image-map.ps1",
      "scripts/insert-course-design-tables.ps1",
      "scripts/insert-docx-images.ps1",
      "scripts/prepare-report-prompt.ps1",
      "scripts/report-defaults.ps1",
      "scripts/report-profiles.ps1",
      "scripts/render-vertical-lab-flowchart.py",
      "scripts/validate-report-draft.ps1"
    )) {
    Add-ExistingFile -Files $files -Path (Join-RepoRelativePath -RelativePath $relativePath)
  }

  Add-ExistingFilesFromDirectory -Files $files -Path (Join-Path $repoRoot "profiles") -Include @("*.json", "*.md")
  Add-ExistingFilesFromDirectory -Files $files -Path (Join-RepoRelativePath -RelativePath "examples/report-templates") -Include @("*.docx", "*.doc")

  foreach ($path in @($ExtraTemplatePath)) {
    if ([string]::IsNullOrWhiteSpace($path)) {
      continue
    }

    if (Test-Path -LiteralPath $path -PathType Leaf) {
      Add-ExistingFile -Files $files -Path $path
    } elseif (Test-Path -LiteralPath $path -PathType Container) {
      Add-ExistingFilesFromDirectory -Files $files -Path $path -Include @("*.docx", "*.doc", "*.dotx", "*.json")
    }
  }

  $entries = New-Object System.Collections.Generic.List[string]
  foreach ($file in @($files | Sort-Object -Unique)) {
    $hash = (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash.ToLowerInvariant()
    [void]$entries.Add(("{0}|{1}" -f (Get-RelativeOrFullPath -Path $file), $hash))
  }

  $manifest = $entries -join "`n"
  [pscustomobject]@{
    fingerprint = Get-StringHash -Text $manifest
    files = @($entries)
  }
}

function Write-Status {
  param([Parameter(Mandatory = $true)]$Status)

  if ($Json) {
    $Status | ConvertTo-Json -Depth 8
    return
  }

  if ($Status.canSkipFullSmoke) {
    Write-Output ("Fast report session cache hit. Full smoke can be skipped. Checked files: {0}. Cache: {1}" -f $Status.checkedFileCount, $Status.cachePath)
  } else {
    Write-Output ("Full smoke recommended before report generation: {0}. Cache: {1}" -f $Status.reason, $Status.cachePath)
    Write-Output "Run once: powershell -ExecutionPolicy Bypass -File .\scripts\check-fast-report-session.ps1 -RunSmokeWhenNeeded"
  }
}

$fingerprintInfo = Get-FastSessionFingerprint -ExtraTemplatePath $TemplatePath
$now = Get-Date
$cache = $null
$cacheExists = Test-Path -LiteralPath $CachePath -PathType Leaf
if ($cacheExists) {
  try {
    $cache = Get-Content -LiteralPath $CachePath -Raw -Encoding UTF8 | ConvertFrom-Json
  } catch {
    $cache = $null
  }
}

$reason = ""
$canSkip = $false
$ageHours = $null

if ($null -eq $cache) {
  $reason = "cache missing or unreadable"
} elseif (-not ($cache.PSObject.Properties.Name -contains "smokePassed") -or -not [bool]$cache.smokePassed) {
  $reason = "previous smoke status was not a pass"
} elseif (-not ($cache.PSObject.Properties.Name -contains "fingerprint") -or [string]$cache.fingerprint -ne [string]$fingerprintInfo.fingerprint) {
  $reason = "critical files changed"
} elseif (-not ($cache.PSObject.Properties.Name -contains "updatedAt")) {
  $reason = "cache timestamp missing"
} else {
  $updatedAt = [datetime]$cache.updatedAt
  $ageHours = ($now - $updatedAt).TotalHours
  if ($ageHours -gt $MaxAgeHours) {
    $reason = "cache is older than $MaxAgeHours hours"
  } else {
    $canSkip = $true
    $reason = "cache hit"
  }
}

if (($RecordPass -or ($RunSmokeWhenNeeded -and -not $canSkip)) -and -not $RecordPass) {
  & (Join-RepoRelativePath -RelativePath "scripts/run-smoke-tests.ps1") | Out-Host
}

if ($RecordPass -or ($RunSmokeWhenNeeded -and -not $canSkip)) {
  Ensure-ParentDirectory -Path $CachePath
  $cachePayload = [ordered]@{
    version = 1
    repoRoot = $repoRoot
    smokePassed = $true
    updatedAt = $now.ToString("o")
    maxAgeHours = $MaxAgeHours
    fingerprint = [string]$fingerprintInfo.fingerprint
    checkedFileCount = @($fingerprintInfo.files).Count
    files = @($fingerprintInfo.files)
  }
  $cachePayload | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $CachePath -Encoding UTF8
  $canSkip = $true
  $reason = if ($RecordPass) { "recorded pass" } else { "smoke passed and cache updated" }
  $ageHours = 0
}

$status = [pscustomobject]@{
  canSkipFullSmoke = $canSkip
  smokeRequired = (-not $canSkip)
  reason = $reason
  cachePath = [System.IO.Path]::GetFullPath($CachePath)
  fingerprint = [string]$fingerprintInfo.fingerprint
  checkedFileCount = @($fingerprintInfo.files).Count
  ageHours = $ageHours
  maxAgeHours = $MaxAgeHours
}

Write-Status -Status $status
