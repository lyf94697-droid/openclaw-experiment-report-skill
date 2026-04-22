[CmdletBinding()]
param(
  [string[]]$ReportProfileName,

  [string[]]$ReportProfilePath,

  [string]$OutputDir,

  [switch]$Overwrite
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "report-profiles.ps1")

function Get-TargetProfilePaths {
  param(
    [AllowNull()]
    [string[]]$RequestedProfileNames,

    [AllowNull()]
    [string[]]$RequestedProfilePaths,

    [Parameter(Mandatory = $true)]
    [string]$RepoRoot
  )

  $resolved = New-Object System.Collections.Generic.List[string]

  foreach ($profilePath in @($RequestedProfilePaths)) {
    if (-not [string]::IsNullOrWhiteSpace($profilePath)) {
      [void]$resolved.Add((Resolve-Path -LiteralPath $profilePath).Path)
    }
  }

  foreach ($profileName in @($RequestedProfileNames)) {
    if (-not [string]::IsNullOrWhiteSpace($profileName)) {
      [void]$resolved.Add((Resolve-ReportProfilePath -ProfileName $profileName -RepoRoot $RepoRoot))
    }
  }

  if ($resolved.Count -gt 0) {
    return @($resolved | Select-Object -Unique)
  }

  return @(
    Get-ChildItem -LiteralPath (Join-Path $RepoRoot "profiles") -Filter "*.json" -File |
      Where-Object { $_.Name -ne "report-profile.schema.json" } |
      Sort-Object -Property Name |
      ForEach-Object { $_.FullName }
  )
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
  $OutputDir = Join-Path $repoRoot "examples\report-templates"
}

$resolvedOutputDir = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null

$targetProfilePaths = @(Get-TargetProfilePaths -RequestedProfileNames $ReportProfileName -RequestedProfilePaths $ReportProfilePath -RepoRoot $repoRoot)
if ($targetProfilePaths.Count -eq 0) {
  throw "No report profiles were selected."
}

$profileSpecs = New-Object System.Collections.Generic.List[object]
$generated = New-Object System.Collections.Generic.List[object]

foreach ($profilePath in $targetProfilePaths) {
  $profile = Get-ReportProfile -ProfilePath $profilePath -RepoRoot $repoRoot
  $outputPath = Join-Path $resolvedOutputDir ("{0}-template.docx" -f [string]$profile.name)

  if ((Test-Path -LiteralPath $outputPath) -and (-not $Overwrite)) {
    throw "Output template already exists: $outputPath. Pass -Overwrite to replace it."
  }

  $metadataLabels = @(
    @($profile.metadataFields) |
      ForEach-Object { [string]$_.label } |
      Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
  )

  $sectionHeadings = @(
    @($profile.sectionFields) |
      ForEach-Object { [string]$_.heading } |
      Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
  )

  [void]$profileSpecs.Add([ordered]@{
      name = [string]$profile.name
      displayName = [string](Get-ReportProfileDisplayName -Profile $profile -Fallback ([string]$profile.name))
      metadataLabels = $metadataLabels
      sectionHeadings = $sectionHeadings
      outputPath = $outputPath
      profilePath = [string]$profile.resolvedProfilePath
    })

  [void]$generated.Add([pscustomobject]@{
      reportProfileName = [string]$profile.name
      reportProfileDisplayName = [string](Get-ReportProfileDisplayName -Profile $profile -Fallback ([string]$profile.name))
      profilePath = [string]$profile.resolvedProfilePath
      outputPath = $outputPath
    })
}

$tempFileBase = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ("report-template-examples-" + [Guid]::NewGuid().ToString("N")))
$specPath = "$tempFileBase.spec.json"
[System.IO.File]::WriteAllText(
  $specPath,
  (($profileSpecs | ConvertTo-Json -Depth 6) + [Environment]::NewLine),
  (New-Object System.Text.UTF8Encoding($true))
)

$pythonScript = @'
import json
import math
import os
import sys

try:
    from docx import Document
    from docx.enum.section import WD_SECTION_START
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt
except ImportError as exc:
    raise SystemExit("python-docx is required: %s" % exc)


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_run_font(run, size_pt, bold=False):
    run.bold = bold
    run.font.size = Pt(size_pt)
    run.font.name = "SimSun"
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    r_fonts.set(qn("w:eastAsia"), "SimSun")
    r_fonts.set(qn("w:ascii"), "Times New Roman")
    r_fonts.set(qn("w:hAnsi"), "Times New Roman")


def set_paragraph_spacing(paragraph, before=0, after=0, line=1.5):
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.line_spacing = line


def add_placeholder_paragraph(doc):
    paragraph = doc.add_paragraph()
    set_paragraph_spacing(paragraph, before=0, after=6, line=1.5)
    run = paragraph.add_run("____________________________________________________________")
    set_run_font(run, 12, bold=False)
    return paragraph


def add_heading(doc, index, heading_text):
    numerals = [
        "\u4e00", "\u4e8c", "\u4e09", "\u56db", "\u4e94", "\u516d",
        "\u4e03", "\u516b", "\u4e5d", "\u5341", "\u5341\u4e00", "\u5341\u4e8c"
    ]
    prefix = numerals[index] if index < len(numerals) else str(index + 1)
    paragraph = doc.add_paragraph()
    set_paragraph_spacing(paragraph, before=10, after=4, line=1.2)
    run = paragraph.add_run(f"{prefix}\u3001{heading_text}")
    set_run_font(run, 14, bold=True)
    return paragraph


def build_document(spec):
    document = Document()
    section = document.sections[0]
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.0)
    section.right_margin = Cm(2.5)
    section.header_distance = Cm(1.5)
    section.footer_distance = Cm(1.5)
    section.start_type = WD_SECTION_START.NEW_PAGE

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(title, before=0, after=12, line=1.0)
    title_run = title.add_run(spec["displayName"])
    set_run_font(title_run, 18, bold=True)

    labels = list(spec["metadataLabels"])
    if not labels:
        labels = ["Name", "ID", "Group", "Date"]

    rows = int(math.ceil(len(labels) / 2.0))
    table = document.add_table(rows=rows, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    table.autofit = True

    for row_index in range(rows):
        first = labels[row_index * 2] if row_index * 2 < len(labels) else ""
        second = labels[row_index * 2 + 1] if row_index * 2 + 1 < len(labels) else ""
        pairs = [
            (table.cell(row_index, 0), table.cell(row_index, 1), first),
            (table.cell(row_index, 2), table.cell(row_index, 3), second),
        ]
        for label_cell, value_cell, label_text in pairs:
            label_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            value_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_shading(label_cell, "EDEDED")
            label_paragraph = label_cell.paragraphs[0]
            label_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_paragraph_spacing(label_paragraph, before=0, after=0, line=1.0)
            label_run = label_paragraph.add_run(label_text if label_text else "")
            set_run_font(label_run, 11, bold=True)

            value_paragraph = value_cell.paragraphs[0]
            set_paragraph_spacing(value_paragraph, before=0, after=0, line=1.0)
            value_run = value_paragraph.add_run("__________________" if label_text else "")
            set_run_font(value_run, 11, bold=False)

    document.add_paragraph()

    for index, heading in enumerate(spec["sectionHeadings"]):
        add_heading(document, index, heading)
        add_placeholder_paragraph(document)

    document.save(spec["outputPath"])


def main():
    spec_path = sys.argv[1]
    with open(spec_path, "r", encoding="utf-8-sig") as handle:
        specs = json.load(handle)
    for spec in specs:
        os.makedirs(os.path.dirname(spec["outputPath"]), exist_ok=True)
        build_document(spec)


if __name__ == "__main__":
    main()
'@

$pythonScriptPath = "$tempFileBase.generator.py"
[System.IO.File]::WriteAllText(
  $pythonScriptPath,
  $pythonScript,
  (New-Object System.Text.UTF8Encoding($false))
)

try {
  $null = & python $pythonScriptPath $specPath
  if ($LASTEXITCODE -ne 0) {
    throw "Failed to generate report template examples."
  }
} finally {
  foreach ($tempPath in @($specPath, $pythonScriptPath)) {
    if (Test-Path -LiteralPath $tempPath) {
      Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
    }
  }
}

$summaryPath = Join-Path $resolvedOutputDir "report-template-examples-summary.json"
$summary = [pscustomobject]@{
  outputDir = $resolvedOutputDir
  generatedCount = $generated.Count
  generated = $generated.ToArray()
}
[System.IO.File]::WriteAllText(
  $summaryPath,
  (($summary | ConvertTo-Json -Depth 6) + [Environment]::NewLine),
  (New-Object System.Text.UTF8Encoding($true))
)

$markdownPath = Join-Path $resolvedOutputDir "README.md"
$markdownLines = New-Object System.Collections.Generic.List[string]
[void]$markdownLines.Add("# Report Template Examples")
[void]$markdownLines.Add("")
[void]$markdownLines.Add("Generated by `scripts/export-report-template-examples.ps1`.")
[void]$markdownLines.Add("")
[void]$markdownLines.Add("| Profile | Display name | Template |")
[void]$markdownLines.Add("| --- | --- | --- |")
foreach ($item in $generated) {
  [void]$markdownLines.Add(("| {0} | {1} | {2} |" -f [string]$item.reportProfileName, [string]$item.reportProfileDisplayName, [System.IO.Path]::GetFileName([string]$item.outputPath)))
}
[System.IO.File]::WriteAllText(
  $markdownPath,
  (($markdownLines -join [Environment]::NewLine) + [Environment]::NewLine),
  (New-Object System.Text.UTF8Encoding($true))
)

$summary
