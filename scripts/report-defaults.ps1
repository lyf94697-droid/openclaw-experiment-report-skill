Set-StrictMode -Version Latest

function Resolve-ExperimentReportDefaultsPath {
  param(
    [string]$AgentsHome = $env:AGENTS_HOME,

    [AllowNull()]
    [string]$ReportProfileName = "experiment-report",

    [AllowNull()]
    [string]$ReportProfilePath
  )

  $agentsRoot = if (-not [string]::IsNullOrWhiteSpace($AgentsHome)) {
    [System.IO.Path]::GetFullPath($AgentsHome)
  } else {
    Join-Path $HOME ".agents"
  }

  $defaultsKey = if (-not [string]::IsNullOrWhiteSpace($ReportProfilePath)) {
    [System.IO.Path]::GetFileNameWithoutExtension($ReportProfilePath)
  } elseif (-not [string]::IsNullOrWhiteSpace($ReportProfileName)) {
    $ReportProfileName
  } else {
    "experiment-report"
  }

  return (Join-Path $agentsRoot ("{0}.defaults.json" -f $defaultsKey))
}

function Read-ExperimentReportDefaults {
  param(
    [AllowNull()]
    [string]$DefaultsPath,

    [AllowNull()]
    [string]$ReportProfileName = "experiment-report",

    [AllowNull()]
    [string]$ReportProfilePath
  )

  if ([string]::IsNullOrWhiteSpace($DefaultsPath)) {
    $DefaultsPath = Resolve-ExperimentReportDefaultsPath -ReportProfileName $ReportProfileName -ReportProfilePath $ReportProfilePath
  }

  if ([string]::IsNullOrWhiteSpace($DefaultsPath) -or -not (Test-Path -LiteralPath $DefaultsPath)) {
    return $null
  }

  $raw = Get-Content -LiteralPath $DefaultsPath -Raw -Encoding UTF8
  if ([string]::IsNullOrWhiteSpace($raw)) {
    return $null
  }

  return ($raw | ConvertFrom-Json)
}

function New-ExperimentReportTextFromCodePoints {
  param(
    [Parameter(Mandatory = $true)]
    [int[]]$CodePoints
  )

  $builder = New-Object System.Text.StringBuilder
  foreach ($codePoint in $CodePoints) {
    [void]$builder.Append([char]$codePoint)
  }

  return $builder.ToString()
}

function Get-ExperimentNameInferenceLabels {
  return @(
    (New-ExperimentReportTextFromCodePoints -CodePoints @(0x5B9E, 0x9A8C, 0x540D, 0x79F0)),
    (New-ExperimentReportTextFromCodePoints -CodePoints @(0x5B9E, 0x9A8C, 0x540D)),
    (New-ExperimentReportTextFromCodePoints -CodePoints @(0x5B9E, 0x9A8C, 0x9898, 0x76EE)),
    (New-ExperimentReportTextFromCodePoints -CodePoints @(0x5B9E, 0x9A8C, 0x9898, 0x76EE, 0x540D, 0x79F0)),
    (New-ExperimentReportTextFromCodePoints -CodePoints @(0x9898, 0x76EE))
  )
}

function Get-ExperimentNameRejectedCandidates {
  return @(
    (New-ExperimentReportTextFromCodePoints -CodePoints @(0x5B9E, 0x9A8C)),
    (New-ExperimentReportTextFromCodePoints -CodePoints @(0x5B9E, 0x9A8C, 0x62A5, 0x544A)),
    (New-ExperimentReportTextFromCodePoints -CodePoints @(0x62A5, 0x544A)),
    (New-ExperimentReportTextFromCodePoints -CodePoints @(0x4F60, 0x7684, 0x6587, 0x7AE0, 0x94FE, 0x63A5))
  )
}

function Get-ExperimentNameSiteSuffixes {
  return @(
    "CSDN",
    "Bilibili",
    (New-ExperimentReportTextFromCodePoints -CodePoints @(0x535A, 0x5BA2, 0x56ED)),
    (New-ExperimentReportTextFromCodePoints -CodePoints @(0x77E5, 0x4E4E)),
    (New-ExperimentReportTextFromCodePoints -CodePoints @(0x54D4, 0x54E9, 0x54D4, 0x54E9)),
    (New-ExperimentReportTextFromCodePoints -CodePoints @(0x7B80, 0x4E66))
  )
}

function Normalize-ExperimentNameCandidate {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return $null
  }

  $candidate = $Text.Trim()
  $labelPattern = ((Get-ExperimentNameInferenceLabels) | ForEach-Object { [regex]::Escape($_) }) -join '|'
  $siteSuffixPattern = ((Get-ExperimentNameSiteSuffixes) | ForEach-Object { [regex]::Escape($_) }) -join '|'
  $candidate = $candidate -replace '^(?i)\s*(?:title|experiment\s*name)\s*[:\uFF1A]\s*', ''
  $candidate = $candidate -replace ("^\s*(?:{0})\s*[:\uFF1A]\s*" -f $labelPattern), ''
  $candidate = $candidate -replace ("\s*(?:[-_\u2014\u2013\|\uFF5C]\s*)?(?:{0}).*$" -f $siteSuffixPattern),''
  $candidate = $candidate -replace '\.(?:html?|md|txt)$',''
  $candidate = $candidate -replace '[_+]+',' '
  $candidate = $candidate -replace '\s+',' '
  $trimChars = @(
    [char]0x20, [char]0x09, [char]0x0D, [char]0x0A, [char]0x22, [char]0x27,
    [char]0x201C, [char]0x201D, [char]0x2018, [char]0x2019,
    [char]0x005B, [char]0x005D, [char]0x3010, [char]0x3011,
    [char]0x0028, [char]0x0029, [char]0xFF08, [char]0xFF09,
    [char]0x002D, [char]0x005F, [char]0x2014, [char]0x2013,
    [char]0x007C, [char]0xFF5C, [char]0x003A, [char]0xFF1A
  )
  $candidate = $candidate.Trim($trimChars)

  if ([string]::IsNullOrWhiteSpace($candidate)) {
    return $null
  }

  if ($candidate -match '\s*(?:\u2014\u2014|--|\uFF5C|\|)\s*') {
    $parts = @($candidate -split '\s*(?:\u2014\u2014|--|\uFF5C|\|)\s*' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($parts.Count -gt 1) {
      $candidate = [string]$parts[$parts.Count - 1]
    }
  }

  $candidate = $candidate.Trim($trimChars)

  if ($candidate.Length -lt 3 -or $candidate.Length -gt 80) {
    return $null
  }

  if ($candidate -match '^(?i)(?:article|details|index|default|blog|post|url|html?)$') {
    return $null
  }
  if ((Get-ExperimentNameRejectedCandidates) -contains $candidate) {
    return $null
  }
  if ($candidate -match '^(?:https?|ftp)://' -or $candidate -match '^[0-9]+$') {
    return $null
  }

  return $candidate
}

function Get-ExperimentNameCandidateFromText {
  param(
    [AllowNull()]
    [string]$Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return $null
  }

  $labelPattern = ((Get-ExperimentNameInferenceLabels) | ForEach-Object { [regex]::Escape($_) }) -join '|'
  $patterns = @(
    ("(?im)^\s*(?:{0})\s*[:\uFF1A]\s*(?<name>.+?)\s*$" -f $labelPattern),
    '(?im)^\s*(?:Experiment\s*Name|Experiment\s*Title|Lab\s*Name|Lab\s*Title)\s*[:\uFF1A]\s*(?<name>.+?)\s*$',
    '(?im)^\s*TITLE\s*:\s*(?<name>.+?)\s*$'
  )

  foreach ($pattern in $patterns) {
    $match = [regex]::Match($Text, $pattern)
    if ($match.Success) {
      $candidate = Normalize-ExperimentNameCandidate -Text $match.Groups["name"].Value
      if (-not [string]::IsNullOrWhiteSpace($candidate)) {
        return $candidate
      }
    }
  }

  return $null
}

function Get-ExperimentNameCandidateFromUrl {
  param(
    [AllowNull()]
    [string]$Url
  )

  if ([string]::IsNullOrWhiteSpace($Url)) {
    return $null
  }

  try {
    $uri = [System.Uri]$Url
  } catch {
    return $null
  }

  $segments = @($uri.Segments | ForEach-Object { $_.Trim('/') } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
  for ($index = $segments.Count - 1; $index -ge 0; $index--) {
    $segment = [string]$segments[$index]
    if ($segment -match '^(?i)(?:article|details|blog|post|p|archives)$' -or $segment -match '^[0-9]+$') {
      continue
    }

    try {
      $decoded = [System.Uri]::UnescapeDataString($segment)
    } catch {
      $decoded = $segment
    }

    $candidate = Normalize-ExperimentNameCandidate -Text $decoded
    if (-not [string]::IsNullOrWhiteSpace($candidate)) {
      return $candidate
    }
  }

  return $null
}

function Resolve-InferredExperimentName {
  param(
    [AllowNull()]
    [string]$PromptText,

    [AllowNull()]
    [string]$PromptPath,

    [AllowNull()]
    [string[]]$ReferenceTextPaths,

    [AllowNull()]
    [string[]]$ReferenceUrls
  )

  if (-not [string]::IsNullOrWhiteSpace($PromptText)) {
    $candidate = Get-ExperimentNameCandidateFromText -Text $PromptText
    if (-not [string]::IsNullOrWhiteSpace($candidate)) {
      return $candidate
    }
  }

  if (-not [string]::IsNullOrWhiteSpace($PromptPath) -and (Test-Path -LiteralPath $PromptPath)) {
    $candidate = Get-ExperimentNameCandidateFromText -Text (Get-Content -LiteralPath $PromptPath -Raw -Encoding UTF8)
    if (-not [string]::IsNullOrWhiteSpace($candidate)) {
      return $candidate
    }
  }

  foreach ($referenceTextPath in @($ReferenceTextPaths)) {
    if ([string]::IsNullOrWhiteSpace($referenceTextPath) -or -not (Test-Path -LiteralPath $referenceTextPath)) {
      continue
    }

    $candidate = Get-ExperimentNameCandidateFromText -Text (Get-Content -LiteralPath $referenceTextPath -Raw -Encoding UTF8)
    if (-not [string]::IsNullOrWhiteSpace($candidate)) {
      return $candidate
    }
  }

  foreach ($referenceUrl in @($ReferenceUrls)) {
    $candidate = Get-ExperimentNameCandidateFromUrl -Url $referenceUrl
    if (-not [string]::IsNullOrWhiteSpace($candidate)) {
      return $candidate
    }
  }

  return $null
}

function Resolve-ExperimentReportNames {
  param(
    [string]$CourseName,

    [string]$ExperimentName,

    [string]$InferredExperimentName,

    [AllowNull()]
    [string]$DefaultsPath,

    [AllowNull()]
    [string]$ReportProfileName = "experiment-report",

    [AllowNull()]
    [string]$ReportProfilePath
  )

  if ([string]::IsNullOrWhiteSpace($DefaultsPath)) {
    $DefaultsPath = Resolve-ExperimentReportDefaultsPath -ReportProfileName $ReportProfileName -ReportProfilePath $ReportProfilePath
  }

  $storedDefaults = Read-ExperimentReportDefaults -DefaultsPath $DefaultsPath -ReportProfileName $ReportProfileName -ReportProfilePath $ReportProfilePath
  $storedCourseName = if ($null -ne $storedDefaults) { [string]$storedDefaults.courseName } else { $null }
  $storedExperimentName = if ($null -ne $storedDefaults) { [string]$storedDefaults.experimentName } else { $null }

  $resolvedCourseName = if (-not [string]::IsNullOrWhiteSpace($CourseName)) { $CourseName } else { $storedCourseName }
  $resolvedExperimentName = if (-not [string]::IsNullOrWhiteSpace($ExperimentName)) {
    $ExperimentName
  } elseif (-not [string]::IsNullOrWhiteSpace($InferredExperimentName)) {
    $InferredExperimentName
  } else {
    $storedExperimentName
  }

  return [pscustomobject]@{
    courseName = $resolvedCourseName
    experimentName = $resolvedExperimentName
    inferredExperimentName = $InferredExperimentName
    defaultsPath = $DefaultsPath
    loadedDefaults = ($null -ne $storedDefaults)
    usedStoredCourseName = ([string]::IsNullOrWhiteSpace($CourseName) -and -not [string]::IsNullOrWhiteSpace($storedCourseName))
    usedStoredExperimentName = ([string]::IsNullOrWhiteSpace($ExperimentName) -and [string]::IsNullOrWhiteSpace($InferredExperimentName) -and -not [string]::IsNullOrWhiteSpace($storedExperimentName))
    usedInferredExperimentName = ([string]::IsNullOrWhiteSpace($ExperimentName) -and -not [string]::IsNullOrWhiteSpace($InferredExperimentName))
  }
}

function Save-ExperimentReportDefaults {
  param(
    [Parameter(Mandatory = $true)]
    [string]$CourseName,

    [Parameter(Mandatory = $true)]
    [string]$ExperimentName,

    [AllowNull()]
    [string]$DefaultsPath,

    [AllowNull()]
    [string]$ReportProfileName = "experiment-report",

    [AllowNull()]
    [string]$ReportProfilePath
  )

  if ([string]::IsNullOrWhiteSpace($CourseName) -or [string]::IsNullOrWhiteSpace($ExperimentName)) {
    throw "CourseName and ExperimentName are required to save report defaults."
  }

  if ([string]::IsNullOrWhiteSpace($DefaultsPath)) {
    $DefaultsPath = Resolve-ExperimentReportDefaultsPath -ReportProfileName $ReportProfileName -ReportProfilePath $ReportProfilePath
  }

  $resolvedDefaultsPath = [System.IO.Path]::GetFullPath($DefaultsPath)
  $parent = Split-Path -Parent $resolvedDefaultsPath
  if (-not [string]::IsNullOrWhiteSpace($parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
  }

  $payload = [pscustomobject]@{
    courseName = $CourseName
    experimentName = $ExperimentName
    reportProfileName = $ReportProfileName
    reportProfilePath = $ReportProfilePath
    updatedAt = (Get-Date).ToString("s")
  }

  [System.IO.File]::WriteAllText($resolvedDefaultsPath, ($payload | ConvertTo-Json -Depth 4), (New-Object System.Text.UTF8Encoding($true)))
  return $resolvedDefaultsPath
}
