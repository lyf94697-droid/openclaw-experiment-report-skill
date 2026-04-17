[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$Name,

  [Parameter(Mandatory = $true)]
  [string]$DisplayName,

  [ValidateSet("auto", "default", "compact", "school")]
  [string]$DefaultStyleProfile = "school",

  [string]$DefaultExperimentProperty = "Custom report",

  [string]$CourseNameLabel = "Course Name",

  [string]$TitleNameLabel = "Report Title",

  [string]$StudentNameLabel = "Student Name",

  [string]$StudentIdLabel = "Student ID",

  [string]$ClassNameLabel = "Class",

  [string]$TeacherNameLabel = "Teacher",

  [string]$DateLabel = "Date",

  [string]$LocationLabel = "Environment",

  [string[]]$SectionHeadings,

  [string]$OutPath,

  [switch]$Force,

  [switch]$SkipValidation,

  [ValidateSet("text", "json")]
  [string]$Format = "text"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Assert-ProfileName {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ProfileName
  )

  if ($ProfileName -notmatch "^[a-z][a-z0-9]*(?:-[a-z0-9]+)*$") {
    throw "Profile name must use lower hyphen-case, for example 'weekly-report'."
  }
}

function Assert-NonEmpty {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Value,

    [Parameter(Mandatory = $true)]
    [string]$Name
  )

  if ([string]::IsNullOrWhiteSpace($Value)) {
    throw "$Name must not be empty."
  }
}

function New-LabeledField {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Key,

    [Parameter(Mandatory = $true)]
    [string]$Label,

    [string[]]$Aliases = @()
  )

  $field = [ordered]@{
    key = $Key
    label = $Label
  }

  $cleanAliases = @($Aliases | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
  if ($cleanAliases.Count -gt 0) {
    $field["aliases"] = @($cleanAliases)
  }

  return $field
}

function New-SectionField {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Key,

    [Parameter(Mandatory = $true)]
    [string]$Heading,

    [Parameter(Mandatory = $true)]
    [int]$StandardMinChars,

    [Parameter(Mandatory = $true)]
    [int]$FullMinChars
  )

  return [ordered]@{
    key = $Key
    heading = $Heading
    aliases = @($Heading)
    minChars = [ordered]@{
      standard = $StandardMinChars
      full = $FullMinChars
    }
  }
}

function Get-DefaultSectionHeadings {
  return @(
    "Purpose",
    "Environment",
    "Background And Requirements",
    "Process And Implementation",
    "Results",
    "Analysis And Improvement",
    "Summary"
  )
}

function Resolve-OutputPath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ProfileName,

    [AllowNull()]
    [string]$RequestedPath,

    [Parameter(Mandatory = $true)]
    [string]$RepositoryRoot
  )

  if ([string]::IsNullOrWhiteSpace($RequestedPath)) {
    return (Join-Path (Join-Path $RepositoryRoot "profiles") ("{0}.json" -f $ProfileName))
  }

  return $RequestedPath
}

function Invoke-ProfileValidation {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ProfilePath,

    [Parameter(Mandatory = $true)]
    [string]$ValidatorPath
  )

  $powerShellExe = if ($PSVersionTable.PSEdition -eq "Core") { "pwsh" } else { "powershell" }
  $validationOutput = & $powerShellExe -NoProfile -ExecutionPolicy Bypass -File $ValidatorPath -ProfilePath $ProfilePath -Format json | Out-String
  $validationExitCode = $LASTEXITCODE
  if ($validationExitCode -ne 0) {
    throw ("Profile validation failed for '{0}'.`n{1}" -f $ProfilePath, $validationOutput.Trim())
  }

  return ($validationOutput | ConvertFrom-Json)
}

function Write-Utf8NoBom {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [string]$Text
  )

  $parent = Split-Path -Parent $Path
  if (-not [string]::IsNullOrWhiteSpace($parent) -and -not (Test-Path -LiteralPath $parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
  }

  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($Path, $Text, $utf8NoBom)
}

Assert-ProfileName -ProfileName $Name
Assert-NonEmpty -Value $DisplayName -Name "DisplayName"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$resolvedOutPath = Resolve-OutputPath -ProfileName $Name -RequestedPath $OutPath -RepositoryRoot $repoRoot
$resolvedOutPath = [System.IO.Path]::GetFullPath($resolvedOutPath)

$outPathBaseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedOutPath)
if ($outPathBaseName -ne $Name) {
  throw "OutPath filename must match the profile name: $Name.json"
}

if ((Test-Path -LiteralPath $resolvedOutPath) -and -not $Force) {
  throw "Profile file already exists: $resolvedOutPath. Re-run with -Force to overwrite it."
}

$headings = if ($null -eq $SectionHeadings -or $SectionHeadings.Count -eq 0) {
  @(Get-DefaultSectionHeadings)
} else {
  @($SectionHeadings)
}

if ($headings.Count -ne 7) {
  throw "SectionHeadings must contain exactly 7 headings: purpose, environment, theory, steps, results, analysis, summary."
}

foreach ($index in 0..6) {
  Assert-NonEmpty -Value ([string]$headings[$index]) -Name ("SectionHeadings[{0}]" -f $index)
}

$profile = [ordered]@{
  name = $Name
  displayName = $DisplayName
  defaultStyleProfile = $DefaultStyleProfile
  defaultExperimentProperty = $DefaultExperimentProperty
  metadataFields = @(
    (New-LabeledField -Key "Name" -Label $StudentNameLabel -Aliases @("Name")),
    (New-LabeledField -Key "StudentId" -Label $StudentIdLabel -Aliases @("Student ID")),
    (New-LabeledField -Key "ClassName" -Label $ClassNameLabel),
    (New-LabeledField -Key "TeacherName" -Label $TeacherNameLabel -Aliases @("Instructor")),
    (New-LabeledField -Key "CourseName" -Label $CourseNameLabel -Aliases @("Course")),
    (New-LabeledField -Key "ExperimentName" -Label $TitleNameLabel -Aliases @("Title", "Project")),
    (New-LabeledField -Key "ExperimentProperty" -Label "Report Type" -Aliases @("Type")),
    (New-LabeledField -Key "ExperimentDate" -Label $DateLabel -Aliases @("Date")),
    (New-LabeledField -Key "ExperimentLocation" -Label $LocationLabel -Aliases @("Location", "Environment"))
  )
  extraLabels = @(
    (New-LabeledField -Key "FinalEdition" -Label "Final Edition")
  )
  sectionFields = @(
    (New-SectionField -Key "Purpose" -Heading ([string]$headings[0]) -StandardMinChars 30 -FullMinChars 70),
    (New-SectionField -Key "Environment" -Heading ([string]$headings[1]) -StandardMinChars 40 -FullMinChars 90),
    (New-SectionField -Key "Theory" -Heading ([string]$headings[2]) -StandardMinChars 40 -FullMinChars 100),
    (New-SectionField -Key "Steps" -Heading ([string]$headings[3]) -StandardMinChars 80 -FullMinChars 180),
    (New-SectionField -Key "Results" -Heading ([string]$headings[4]) -StandardMinChars 60 -FullMinChars 140),
    (New-SectionField -Key "Analysis" -Heading ([string]$headings[5]) -StandardMinChars 40 -FullMinChars 100),
    (New-SectionField -Key "Summary" -Heading ([string]$headings[6]) -StandardMinChars 40 -FullMinChars 100)
  )
  extraSectionHeadings = @(
    "Appendix",
    "Evidence"
  )
  imagePlacementDefaults = [ordered]@{
    fallbackSectionOrder = @("steps", "result", "analysis", "environment", "summary", "purpose")
    defaultCaptions = [ordered]@{
      environment = ("{0} screenshot" -f [string]$headings[1])
      steps = ("{0} screenshot" -f [string]$headings[3])
      result = ("{0} screenshot" -f [string]$headings[4])
      analysis = ("{0} screenshot" -f [string]$headings[5])
      summary = ("{0} screenshot" -f [string]$headings[6])
      default = ("{0} screenshot" -f $DisplayName)
    }
    filenameCaptionRules = @(
      [ordered]@{ pattern = "setup|config|process|step"; caption = ("{0} screenshot" -f [string]$headings[3]) },
      [ordered]@{ pattern = "result|output|pass|verify"; caption = ("{0} screenshot" -f [string]$headings[4]) },
      [ordered]@{ pattern = "error|issue|problem|analysis"; caption = ("{0} screenshot" -f [string]$headings[5]) },
      [ordered]@{ pattern = "env|environment|tool"; caption = ("{0} screenshot" -f [string]$headings[1]) }
    )
  }
  paginationRiskThresholds = [ordered]@{
    longSectionChars = 900
    denseSectionChars = 550
    denseSectionParagraphs = 2
    figureClusterRefs = 3
  }
  fieldMapCompositeRules = @(
    [ordered]@{
      matchAll = @([string]$headings[0], [string]$headings[2])
      blocks = @(
        [ordered]@{ heading = [string]$headings[0]; sectionIds = @("purpose") },
        [ordered]@{ heading = [string]$headings[2]; sectionIds = @("theory") }
      )
    },
    [ordered]@{
      matchAll = @([string]$headings[3], [string]$headings[6])
      blocks = @(
        [ordered]@{ heading = [string]$headings[1]; sectionIds = @("environment") },
        [ordered]@{ heading = [string]$headings[3]; sectionIds = @("steps") },
        [ordered]@{ heading = [string]$headings[4]; sectionIds = @("result") },
        [ordered]@{ heading = [string]$headings[5]; sectionIds = @("analysis") },
        [ordered]@{ heading = [string]$headings[6]; sectionIds = @("summary") }
      )
    }
  )
  detailProfiles = [ordered]@{
    standard = [ordered]@{
      minChars = 1000
      promptGuidance = @(
        "Prefer a submit-ready school report instead of a short outline.",
        "Keep each required section concrete and grounded in the provided materials."
      )
    }
    full = [ordered]@{
      minChars = 1600
      promptGuidance = @(
        "Prefer a submit-ready school report with substantial detail instead of a terse outline.",
        "Unless source material is very sparse, aim for roughly 1600 to 2400 Chinese characters.",
        "Explain the process, evidence, results, risks, and follow-up improvements in concrete terms."
      )
    }
  }
  forbiddenPatterns = @(
    "TODO",
    "TBD",
    "fill in later",
    "teacher comments"
  )
}

$profileJson = ($profile | ConvertTo-Json -Depth 12)
Write-Utf8NoBom -Path $resolvedOutPath -Text ($profileJson + [Environment]::NewLine)

$validationResult = $null
if (-not $SkipValidation) {
  $validatorPath = Join-Path $PSScriptRoot "validate-report-profiles.ps1"
  $validationResult = Invoke-ProfileValidation -ProfilePath $resolvedOutPath -ValidatorPath $validatorPath
}

$result = [pscustomobject]@{
  path = $resolvedOutPath
  name = $Name
  displayName = $DisplayName
  validationPassed = if ($null -eq $validationResult) { $null } else { [bool]$validationResult.passed }
  validationErrorCount = if ($null -eq $validationResult) { $null } else { [int]$validationResult.summary.errorCount }
}

if ($Format -eq "json") {
  Write-Output ($result | ConvertTo-Json -Depth 6)
} else {
  Write-Output ("Created report profile scaffold: {0}" -f $resolvedOutPath)
  if ($null -ne $validationResult) {
    Write-Output ("Validation passed: {0}" -f [bool]$validationResult.passed)
  }
}
