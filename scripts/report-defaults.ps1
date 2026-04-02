Set-StrictMode -Version Latest

function Resolve-ExperimentReportDefaultsPath {
  param(
    [string]$AgentsHome = $env:AGENTS_HOME
  )

  $agentsRoot = if (-not [string]::IsNullOrWhiteSpace($AgentsHome)) {
    [System.IO.Path]::GetFullPath($AgentsHome)
  } else {
    Join-Path $HOME ".agents"
  }

  return (Join-Path $agentsRoot "experiment-report.defaults.json")
}

function Read-ExperimentReportDefaults {
  param(
    [string]$DefaultsPath = (Resolve-ExperimentReportDefaultsPath)
  )

  if ([string]::IsNullOrWhiteSpace($DefaultsPath) -or -not (Test-Path -LiteralPath $DefaultsPath)) {
    return $null
  }

  $raw = Get-Content -LiteralPath $DefaultsPath -Raw -Encoding UTF8
  if ([string]::IsNullOrWhiteSpace($raw)) {
    return $null
  }

  return ($raw | ConvertFrom-Json)
}

function Resolve-ExperimentReportNames {
  param(
    [string]$CourseName,

    [string]$ExperimentName,

    [string]$DefaultsPath = (Resolve-ExperimentReportDefaultsPath)
  )

  $storedDefaults = Read-ExperimentReportDefaults -DefaultsPath $DefaultsPath
  $storedCourseName = if ($null -ne $storedDefaults) { [string]$storedDefaults.courseName } else { $null }
  $storedExperimentName = if ($null -ne $storedDefaults) { [string]$storedDefaults.experimentName } else { $null }

  $resolvedCourseName = if (-not [string]::IsNullOrWhiteSpace($CourseName)) { $CourseName } else { $storedCourseName }
  $resolvedExperimentName = if (-not [string]::IsNullOrWhiteSpace($ExperimentName)) { $ExperimentName } else { $storedExperimentName }

  return [pscustomobject]@{
    courseName = $resolvedCourseName
    experimentName = $resolvedExperimentName
    defaultsPath = $DefaultsPath
    loadedDefaults = ($null -ne $storedDefaults)
    usedStoredCourseName = ([string]::IsNullOrWhiteSpace($CourseName) -and -not [string]::IsNullOrWhiteSpace($storedCourseName))
    usedStoredExperimentName = ([string]::IsNullOrWhiteSpace($ExperimentName) -and -not [string]::IsNullOrWhiteSpace($storedExperimentName))
  }
}

function Save-ExperimentReportDefaults {
  param(
    [Parameter(Mandatory = $true)]
    [string]$CourseName,

    [Parameter(Mandatory = $true)]
    [string]$ExperimentName,

    [string]$DefaultsPath = (Resolve-ExperimentReportDefaultsPath)
  )

  if ([string]::IsNullOrWhiteSpace($CourseName) -or [string]::IsNullOrWhiteSpace($ExperimentName)) {
    throw "CourseName and ExperimentName are required to save experiment-report defaults."
  }

  $resolvedDefaultsPath = [System.IO.Path]::GetFullPath($DefaultsPath)
  $parent = Split-Path -Parent $resolvedDefaultsPath
  if (-not [string]::IsNullOrWhiteSpace($parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
  }

  $payload = [pscustomobject]@{
    courseName = $CourseName
    experimentName = $ExperimentName
    updatedAt = (Get-Date).ToString("s")
  }

  [System.IO.File]::WriteAllText($resolvedDefaultsPath, ($payload | ConvertTo-Json -Depth 4), (New-Object System.Text.UTF8Encoding($true)))
  return $resolvedDefaultsPath
}
