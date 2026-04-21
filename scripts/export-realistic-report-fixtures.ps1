[CmdletBinding()]
param(
  [string]$OutputDir,

  [switch]$Overwrite
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Drawing

function New-CourseSelectionFlowchartPng {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $width = 1400
  $height = 1980
  $bitmap = New-Object System.Drawing.Bitmap($width, $height)
  $graphics = [System.Drawing.Graphics]::FromImage($bitmap)

  $titleFont = New-Object System.Drawing.Font("Microsoft YaHei", 30, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
  $font = New-Object System.Drawing.Font("Microsoft YaHei", 30, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
  $smallFont = New-Object System.Drawing.Font("Microsoft YaHei", 24, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
  $badgeFont = New-Object System.Drawing.Font("Microsoft YaHei", 22, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)

  $inkBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black)
  $headerBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
  $noteBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black)
  $linePen = New-Object System.Drawing.Pen([System.Drawing.Color]::Black, 3)
  $thinLinePen = New-Object System.Drawing.Pen([System.Drawing.Color]::Black, 2)
  $shadowBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
  $shadowPen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 1)
  $boxBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
  $decisionBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
  $endBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
  $warnBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
  $whiteBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)

  function New-Rect {
    param([int]$X, [int]$Y, [int]$W, [int]$H)
    return (New-Object System.Drawing.Rectangle($X, $Y, $W, $H))
  }

  function Draw-CenteredText {
    param(
      [System.Drawing.Graphics]$G,
      [System.Drawing.Rectangle]$Rect,
      [string]$Text,
      [System.Drawing.Font]$TextFont,
      [System.Drawing.Brush]$Brush
    )

    $format = New-Object System.Drawing.StringFormat
    $format.Alignment = [System.Drawing.StringAlignment]::Center
    $format.LineAlignment = [System.Drawing.StringAlignment]::Center
    $rectF = New-Object System.Drawing.RectangleF([single]$Rect.X, [single]$Rect.Y, [single]$Rect.Width, [single]$Rect.Height)
    $G.DrawString($Text, $TextFont, $Brush, $rectF, $format)
    $format.Dispose()
  }

  function Draw-RoundedRectangle {
    param(
      [System.Drawing.Graphics]$G,
      [System.Drawing.Rectangle]$Rect,
      [int]$Radius,
      [System.Drawing.Brush]$Fill,
      [System.Drawing.Pen]$Pen
    )

    $diameter = $Radius * 2
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc($Rect.X, $Rect.Y, $diameter, $diameter, 180, 90)
    $path.AddArc($Rect.Right - $diameter, $Rect.Y, $diameter, $diameter, 270, 90)
    $path.AddArc($Rect.Right - $diameter, $Rect.Bottom - $diameter, $diameter, $diameter, 0, 90)
    $path.AddArc($Rect.X, $Rect.Bottom - $diameter, $diameter, $diameter, 90, 90)
    $path.CloseFigure()
    $G.FillPath($Fill, $path)
    $G.DrawPath($Pen, $path)
    $path.Dispose()
  }

  function Draw-Diamond {
    param(
      [System.Drawing.Graphics]$G,
      [System.Drawing.Rectangle]$Rect,
      [System.Drawing.Brush]$Fill,
      [System.Drawing.Pen]$Pen
    )

    $points = @(
      (New-Object System.Drawing.Point(($Rect.X + [int]($Rect.Width / 2)), $Rect.Y)),
      (New-Object System.Drawing.Point($Rect.Right, ($Rect.Y + [int]($Rect.Height / 2)))),
      (New-Object System.Drawing.Point(($Rect.X + [int]($Rect.Width / 2)), $Rect.Bottom)),
      (New-Object System.Drawing.Point($Rect.X, ($Rect.Y + [int]($Rect.Height / 2))))
    )
    $G.FillPolygon($Fill, $points)
    $G.DrawPolygon($Pen, $points)
  }

  function Draw-Node {
    param(
      [System.Drawing.Graphics]$G,
      [string]$Kind,
      [System.Drawing.Rectangle]$Rect,
      [string]$Text,
      [System.Drawing.Brush]$Fill,
      [string]$Badge = ""
    )

    if ($Kind -eq "round") {
      Draw-RoundedRectangle -G $G -Rect $Rect -Radius 32 -Fill $Fill -Pen $thinLinePen
    } elseif ($Kind -eq "diamond") {
      Draw-Diamond -G $G -Rect $Rect -Fill $Fill -Pen $thinLinePen
    } else {
      $G.FillRectangle($Fill, $Rect)
      $G.DrawRectangle($thinLinePen, $Rect)
    }

    Draw-CenteredText -G $G -Rect $Rect -Text $Text -TextFont $font -Brush $inkBrush
  }

  function Draw-Arrow {
    param(
      [System.Drawing.Graphics]$G,
      [int]$X1,
      [int]$Y1,
      [int]$X2,
      [int]$Y2,
      [string]$Label = "",
      [int]$LabelDx = 0,
      [int]$LabelDy = 0
    )

    $G.DrawLine($linePen, $X1, $Y1, $X2, $Y2)
    $angle = [Math]::Atan2(($Y2 - $Y1), ($X2 - $X1))
    $size = 18
    $a1 = $angle + [Math]::PI * 0.82
    $a2 = $angle - [Math]::PI * 0.82
    $points = @(
      (New-Object System.Drawing.Point($X2, $Y2)),
      (New-Object System.Drawing.Point(([int]($X2 + $size * [Math]::Cos($a1))), ([int]($Y2 + $size * [Math]::Sin($a1))))),
      (New-Object System.Drawing.Point(([int]($X2 + $size * [Math]::Cos($a2))), ([int]($Y2 + $size * [Math]::Sin($a2)))))
    )
    $G.FillPolygon((New-Object System.Drawing.SolidBrush($linePen.Color)), $points)

    if (-not [string]::IsNullOrWhiteSpace($Label)) {
      $labelX = [int](($X1 + $X2) / 2 + $LabelDx)
      $labelY = [int](($Y1 + $Y2) / 2 + $LabelDy)
      $labelSize = $G.MeasureString($Label, $smallFont)
      $labelRect = New-Object System.Drawing.Rectangle(
        ([int]($labelX - 8)),
        ([int]($labelY - 4)),
        ([int]($labelSize.Width + 16)),
        ([int]($labelSize.Height + 8))
      )
      $G.FillRectangle($whiteBrush, $labelRect)
      $G.DrawString($Label, $smallFont, $inkBrush, $labelX, $labelY)
    }
  }

  try {
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $graphics.Clear([System.Drawing.Color]::White)

    $titleRect = New-Object System.Drawing.Rectangle(0, 24, $width, 58)
    Draw-CenteredText -G $graphics -Rect $titleRect -Text "学生选课系统流程图" -TextFont $titleFont -Brush $inkBrush

    Draw-Node -G $graphics -Kind "round" -Rect (New-Rect 550 120 300 90) -Text "开始" -Fill $endBrush -Badge "1"
    Draw-Node -G $graphics -Kind "box" -Rect (New-Rect 470 290 460 100) -Text "输入用户名和密码" -Fill $boxBrush -Badge "2"
    Draw-Node -G $graphics -Kind "diamond" -Rect (New-Rect 500 480 400 190) -Text "用户名和密码`n是否正确" -Fill $decisionBrush
    Draw-Node -G $graphics -Kind "box" -Rect (New-Rect 100 520 320 120) -Text "提示登录失败`n重新输入" -Fill $warnBrush -Badge "3"
    Draw-Node -G $graphics -Kind "box" -Rect (New-Rect 470 770 460 110) -Text "进入对应角色页面" -Fill $boxBrush -Badge "4"
    Draw-Node -G $graphics -Kind "box" -Rect (New-Rect 470 980 460 110) -Text "选择选课 / 退课 / 查询" -Fill $boxBrush -Badge "5"
    Draw-Node -G $graphics -Kind "box" -Rect (New-Rect 470 1180 460 110) -Text "提交操作请求" -Fill $boxBrush -Badge "6"
    Draw-Node -G $graphics -Kind "diamond" -Rect (New-Rect 500 1380 400 190) -Text "课程容量和时间`n是否满足" -Fill $decisionBrush
    Draw-Node -G $graphics -Kind "box" -Rect (New-Rect 970 1420 330 120) -Text "提示容量不足`n或时间冲突" -Fill $warnBrush -Badge "7"
    Draw-Node -G $graphics -Kind "box" -Rect (New-Rect 470 1660 460 110) -Text "写入选课记录`n刷新课表" -Fill $boxBrush -Badge "8"
    Draw-Node -G $graphics -Kind "round" -Rect (New-Rect 550 1830 300 80) -Text "结束" -Fill $endBrush -Badge "9"

    Draw-Arrow -G $graphics -X1 700 -Y1 210 -X2 700 -Y2 290
    Draw-Arrow -G $graphics -X1 700 -Y1 390 -X2 700 -Y2 480
    Draw-Arrow -G $graphics -X1 700 -Y1 670 -X2 700 -Y2 770 -Label "是" -LabelDx 22 -LabelDy -8
    Draw-Arrow -G $graphics -X1 700 -Y1 880 -X2 700 -Y2 980
    Draw-Arrow -G $graphics -X1 700 -Y1 1090 -X2 700 -Y2 1180
    Draw-Arrow -G $graphics -X1 700 -Y1 1290 -X2 700 -Y2 1380
    Draw-Arrow -G $graphics -X1 700 -Y1 1570 -X2 700 -Y2 1660 -Label "是" -LabelDx 22 -LabelDy -8
    Draw-Arrow -G $graphics -X1 700 -Y1 1770 -X2 700 -Y2 1830
    Draw-Arrow -G $graphics -X1 500 -Y1 575 -X2 420 -Y2 575 -Label "否" -LabelDx 0 -LabelDy -36
    $graphics.DrawLine($linePen, 100, 580, 60, 580)
    $graphics.DrawLine($linePen, 60, 580, 60, 340)
    $graphics.DrawLine($linePen, 60, 340, 430, 340)
    Draw-Arrow -G $graphics -X1 430 -Y1 340 -X2 470 -Y2 340
    Draw-Arrow -G $graphics -X1 900 -Y1 1475 -X2 970 -Y2 1475 -Label "否" -LabelDx 10 -LabelDy -36
    $graphics.DrawLine($linePen, 1135, 1420, 1135, 1040)
    $graphics.DrawLine($linePen, 1135, 1040, 970, 1040)
    Draw-Arrow -G $graphics -X1 970 -Y1 1040 -X2 930 -Y2 1040

    $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
  } finally {
    foreach ($resource in @($graphics, $bitmap, $titleFont, $font, $smallFont, $badgeFont, $inkBrush, $headerBrush, $noteBrush, $linePen, $thinLinePen, $shadowBrush, $shadowPen, $boxBrush, $decisionBrush, $endBrush, $warnBrush, $whiteBrush)) {
      if ($null -ne $resource) {
        $resource.Dispose()
      }
    }
  }
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
  $OutputDir = Join-Path $repoRoot "examples\realistic-report-fixtures"
}

$resolvedOutputDir = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $resolvedOutputDir -Force | Out-Null

$fixtureFileNames = @(
  "single-table-experiment-filled.docx",
  "integrated-experiment-multi-table.docx",
  "course-design-full-example.docx",
  "student-course-selection-flowchart.png",
  "README.md",
  "realistic-report-fixtures-summary.json"
)

foreach ($fileName in $fixtureFileNames) {
  $targetPath = Join-Path $resolvedOutputDir $fileName
  if ((Test-Path -LiteralPath $targetPath) -and (-not $Overwrite)) {
    throw "Output fixture already exists: $targetPath. Pass -Overwrite to replace it."
  }
}

$specPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ("realistic-report-fixtures-" + [Guid]::NewGuid().ToString("N") + ".json"))
$scriptPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ("realistic-report-fixtures-" + [Guid]::NewGuid().ToString("N") + ".py"))
$flowchartPath = Join-Path $resolvedOutputDir "student-course-selection-flowchart.png"

New-CourseSelectionFlowchartPng -Path $flowchartPath

[System.IO.File]::WriteAllText(
  $specPath,
  (([pscustomobject]@{
      outputDir = $resolvedOutputDir
      flowchartPath = $flowchartPath
    } | ConvertTo-Json -Depth 4) + [Environment]::NewLine),
  (New-Object System.Text.UTF8Encoding($true))
)

$pythonScript = @'
import json
import os
import struct
import sys
import tempfile
import zlib

try:
    from docx import Document
    from docx.enum.section import WD_SECTION_START
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
except ImportError as exc:
    raise SystemExit("python-docx is required: %s" % exc)


def write_png(path, width, height, color):
    def chunk(kind, data):
        payload = kind + data
        return struct.pack(">I", len(data)) + payload + struct.pack(">I", zlib.crc32(payload) & 0xFFFFFFFF)

    row = b"\x00" + bytes(color) * width
    raw = row * height
    png = (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw, 9))
        + chunk(b"IEND", b"")
    )
    with open(path, "wb") as handle:
        handle.write(png)


def ensure_font(run, size=12, bold=False, color=None):
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = "SimSun"
    if color:
        run.font.color.rgb = RGBColor(*color)
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    r_fonts.set(qn("w:eastAsia"), "SimSun")
    r_fonts.set(qn("w:ascii"), "Times New Roman")
    r_fonts.set(qn("w:hAnsi"), "Times New Roman")


def set_paragraph_spacing(paragraph, before=0, after=4, line=1.35):
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = line


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_width(cell, width_twips):
    tc_pr = cell._tc.get_or_add_tcPr()
    width = OxmlElement("w:tcW")
    width.set(qn("w:w"), str(width_twips))
    width.set(qn("w:type"), "dxa")
    tc_pr.append(width)


def add_text(cell, text, size=11, bold=False):
    cell.text = ""
    paragraph = cell.paragraphs[0]
    set_paragraph_spacing(paragraph, after=0, line=1.2)
    run = paragraph.add_run(text)
    ensure_font(run, size=size, bold=bold)


def append_cell_paragraph(cell, text="", size=11, bold=False, align=None):
    paragraph = cell.add_paragraph()
    if align is not None:
        paragraph.alignment = align
    set_paragraph_spacing(paragraph, after=3, line=1.35)
    if text:
        run = paragraph.add_run(text)
        ensure_font(run, size=size, bold=bold)
    return paragraph


def add_doc_title(document, text, size=18):
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(paragraph, after=14, line=1.0)
    run = paragraph.add_run(text)
    ensure_font(run, size=size, bold=True)
    return paragraph


def add_heading(document, text, level=1):
    paragraph = document.add_paragraph()
    set_paragraph_spacing(paragraph, before=10 if level == 1 else 6, after=4, line=1.2)
    run = paragraph.add_run(text)
    ensure_font(run, size=14 if level == 1 else 12, bold=True)
    return paragraph


def add_body_paragraph(document, text):
    paragraph = document.add_paragraph()
    set_paragraph_spacing(paragraph, after=4, line=1.45)
    paragraph.paragraph_format.first_line_indent = Cm(0.74)
    run = paragraph.add_run(text)
    ensure_font(run, size=11)
    return paragraph


def add_caption(document, text):
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(paragraph, after=6, line=1.1)
    run = paragraph.add_run(text)
    ensure_font(run, size=10)
    return paragraph


def add_picture_block(document, image_path, caption):
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(paragraph, before=4, after=2, line=1.0)
    paragraph.add_run().add_picture(image_path, width=Cm(11.5))
    add_caption(document, caption)


def setup_document(document):
    section = document.sections[0]
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.8)
    section.right_margin = Cm(2.4)
    section.header_distance = Cm(1.5)
    section.footer_distance = Cm(1.5)
    section.start_type = WD_SECTION_START.NEW_PAGE
    style = document.styles["Normal"]
    style.font.name = "SimSun"
    style.font.size = Pt(11)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "SimSun")


def build_single_table_experiment(path, assets):
    document = Document()
    setup_document(document)
    add_doc_title(document, "信息学院实验报告", size=18)

    table = document.add_table(rows=6, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    table.autofit = True

    rows = [
        ["学号：20260001", "姓名：张同学", "班级：计科2301"],
        ["课程名称：操作系统", "实验内容：进程调度算法模拟", "上机实验性质：综合性实验"],
        ["实验时间：2026年4月18日", "实验地点：实验楼211", "实验设备：Windows 11 / Python 3.12"],
    ]
    for row_index, row_values in enumerate(rows):
        for col_index, value in enumerate(row_values):
            cell = table.cell(row_index, col_index)
            set_cell_width(cell, 3000)
            add_text(cell, value, size=10)

    body_cell = table.cell(3, 0).merge(table.cell(3, 2))
    add_text(body_cell, "实验报告：（包括：目的、方法、原理、结果或实验小结等）。", size=10, bold=True)
    sections = [
        ("一、实验目的", [
            "理解先来先服务、短作业优先和时间片轮转调度算法的基本思想。",
            "通过程序模拟观察不同调度策略下平均周转时间和平均带权周转时间的差异。"
        ]),
        ("二、实验方法", [
            "构造五个进程的到达时间、服务时间和优先级数据，分别调用三种调度函数输出执行序列。",
            "记录每个进程的开始时间、完成时间、周转时间和带权周转时间，并对结果进行横向比较。"
        ]),
        ("三、实验原理", [
            "FCFS按照到达顺序分配处理机，算法简单但可能造成短作业等待时间过长。",
            "SJF优先执行服务时间较短的进程，通常能降低平均等待时间，但需要提前估计运行时长。",
            "RR按照固定时间片轮转运行各进程，响应更均衡，适合分时系统。"
        ]),
        ("四、实验结果", [
            "运行结果表明，SJF在本组数据中平均周转时间最低，RR的响应更均匀但上下文切换次数更多。",
            "图1展示了三种算法的平均周转时间对比，图2展示了时间片轮转调度的执行序列。"
        ]),
    ]
    for heading, paragraphs in sections:
        append_cell_paragraph(body_cell, heading, size=12, bold=True)
        for text in paragraphs:
            append_cell_paragraph(body_cell, text, size=11)
    picture_paragraph = body_cell.add_paragraph()
    picture_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    picture_paragraph.add_run().add_picture(assets["chart"], width=Cm(10.2))
    append_cell_paragraph(body_cell, "图1 三种调度算法平均周转时间对比", size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
    picture_paragraph = body_cell.add_paragraph()
    picture_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    picture_paragraph.add_run().add_picture(assets["terminal"], width=Cm(10.2))
    append_cell_paragraph(body_cell, "图2 时间片轮转调度输出片段", size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
    append_cell_paragraph(body_cell, "五、实验小结", size=12, bold=True)
    append_cell_paragraph(body_cell, "本次实验将调度算法的理论指标转化为可运行的模拟结果，加深了对调度公平性和效率之间权衡关系的理解。", size=11)

    comment_cell = table.cell(4, 0).merge(table.cell(4, 2))
    add_text(comment_cell, "任课教师评语：\n\n教师签字：                年    月    日", size=10)
    note_cell = table.cell(5, 0).merge(table.cell(5, 2))
    add_text(note_cell, "注：本合成样例仅用于测试模板结构，不包含真实学生信息。", size=9)
    document.save(path)


def build_integrated_experiment(path, assets):
    document = Document()
    setup_document(document)
    for _ in range(3):
        document.add_paragraph()
    add_doc_title(document, "本科学生综合性实验报告", size=20)
    for text in [
        "学号 20260001  姓名 张同学（组长）",
        "学号 20260002  姓名 李同学（组员）",
        "学院 信息学院  专业、班级 网络工程2301",
        "实验课程名称 计算机网络综合实验",
        "教师及职称 王老师 / 讲师",
        "开课学期 2025 至 2026 学年 第二学期",
        "填报时间 2026 年 4 月 18 日",
    ]:
        paragraph = document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(text)
        ensure_font(run, size=12)
    document.add_paragraph()
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ensure_font(paragraph.add_run("教务处编印"), size=11)
    document.add_page_break()

    add_heading(document, "一．实验设计方案", level=1)
    table1 = document.add_table(rows=6, cols=4)
    table1.style = "Table Grid"
    table1.alignment = WD_TABLE_ALIGNMENT.CENTER
    basic = [
        ("实验序号", "10", "实验名称", "动态路由与连通性验证"),
        ("实验时间", "2026年4月18日", "实验室", "网络工程实验室"),
    ]
    for row_index, values in enumerate(basic):
        for col_index, value in enumerate(values):
            cell = table1.cell(row_index, col_index)
            add_text(cell, value, size=10, bold=(col_index % 2 == 0))
            if col_index % 2 == 0:
                set_cell_shading(cell, "EDEDED")
    purpose = table1.cell(2, 0).merge(table1.cell(2, 3))
    add_text(purpose, "1．实验目的", size=11, bold=True)
    append_cell_paragraph(purpose, "掌握RIP和OSPF动态路由协议的基础配置方法，理解路由表生成和链路状态更新过程。", size=10)
    append_cell_paragraph(purpose, "通过拓扑搭建、地址规划、协议配置和连通性验证，形成完整的网络实验记录。", size=10)
    principle = table1.cell(3, 0).merge(table1.cell(3, 3))
    add_text(principle, "2．实验原理、实验流程或装置示意图", size=11, bold=True)
    append_cell_paragraph(principle, "动态路由协议通过周期性通告或链路状态同步维护路由表，路由器根据目的网段选择下一跳转发路径。", size=10)
    pic = principle.add_paragraph()
    pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic.add_run().add_picture(assets["topology"], width=Cm(10.5))
    append_cell_paragraph(principle, "图1 实验网络拓扑示意图", size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
    equipment = table1.cell(4, 0).merge(table1.cell(4, 3))
    add_text(equipment, "3．实验设备及材料", size=11, bold=True)
    append_cell_paragraph(equipment, "Windows 11、Cisco Packet Tracer、三台路由器、三台PC、串口链路和以太网链路。", size=10)
    blank = table1.cell(5, 0).merge(table1.cell(5, 3))
    add_text(blank, "预留：教师检查记录", size=10)

    table2 = document.add_table(rows=3, cols=1)
    table2.style = "Table Grid"
    method = table2.cell(0, 0)
    add_text(method, "4．实验方法步骤及注意事项", size=11, bold=True)
    for text in [
        "按照地址规划表完成PC和路由器接口IP配置。",
        "启用RIP v2并关闭自动汇总，检查路由表中是否出现远端网段。",
        "切换到OSPF单区域配置，观察邻接关系建立过程。",
        "使用ping和tracert验证跨网段通信路径。"
    ]:
        append_cell_paragraph(method, text, size=10)
    data = table2.cell(1, 0)
    add_text(data, "5．实验数据处理方法", size=11, bold=True)
    append_cell_paragraph(data, "记录路由表条目、下一跳、管理距离、度量值和连通性测试结果，并用表格对比不同协议的路径选择。", size=10)
    refs = table2.cell(2, 0)
    add_text(refs, "6．参考文献", size=11, bold=True)
    append_cell_paragraph(refs, "[1] 计算机网络实验指导书，校内讲义，2026。", size=10)

    add_heading(document, "二．实验报告", level=1)
    table3 = document.add_table(rows=2, cols=1)
    table3.style = "Table Grid"
    result = table3.cell(0, 0)
    add_text(result, "1．实验现象与结果", size=11, bold=True)
    append_cell_paragraph(result, "RIP配置完成后，R1能够学习到远端192.168.2.0/24网段，PC1到PC2的ping测试成功。", size=10)
    pic = result.add_paragraph()
    pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic.add_run().add_picture(assets["terminal"], width=Cm(10.5))
    append_cell_paragraph(result, "图2 ping与tracert验证结果", size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
    analysis = table3.cell(1, 0)
    add_text(analysis, "2．对实验现象、实验结果的分析及其结论", size=11, bold=True)
    append_cell_paragraph(analysis, "路由表中远端网段的下一跳与拓扑设计一致，说明协议通告和接口状态均正常。OSPF收敛速度更快，适合更复杂的多区域网络。", size=10)

    table4 = document.add_table(rows=3, cols=1)
    table4.style = "Table Grid"
    reserve = table4.cell(0, 0)
    add_text(reserve, "补充记录：配置命令、路由表截图和测试截图均已按实验步骤归档。", size=10)
    summary = table4.cell(1, 0)
    add_text(summary, "3．实验总结", size=11, bold=True)
    append_cell_paragraph(summary, "本次综合实验把地址规划、协议配置、结果验证和故障排查串成完整流程，训练了从拓扑到结论的报告写作方式。", size=10)
    teacher = table4.cell(2, 0)
    add_text(teacher, "教师评语及评分：\n\n签名：                年    月    日", size=10)
    document.save(path)


def add_metadata_table(document):
    table = document.add_table(rows=6, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    values = [
        ("题 目", "学生选课系统设计与实现"),
        ("专 业", "软件工程"),
        ("班 级", "软工2302"),
        ("学 号", "20260001"),
        ("姓 名", "张同学"),
        ("时 间", "2026年4月"),
    ]
    for row_index, (label, value) in enumerate(values):
        add_text(table.cell(row_index, 0), label, size=11, bold=True)
        set_cell_shading(table.cell(row_index, 0), "EDEDED")
        add_text(table.cell(row_index, 1), value, size=11)


def add_grading_table(document):
    document.add_paragraph()
    table = document.add_table(rows=4, cols=4)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    rows = [
        ["评价项目", "设计内容", "文档质量", "总评"],
        ["权重", "40%", "40%", "20%"],
        ["得分", "", "", ""],
        ["教师签名", "", "日期", ""],
    ]
    for row_index, row in enumerate(rows):
        for col_index, value in enumerate(row):
            cell = table.cell(row_index, col_index)
            add_text(cell, value, size=10, bold=(row_index == 0 or col_index == 0))
            if row_index == 0:
                set_cell_shading(cell, "EDEDED")


def build_course_design(path, assets):
    document = Document()
    setup_document(document)
    for _ in range(2):
        document.add_paragraph()
    add_doc_title(document, "《数据结构》课程设计报告", size=20)
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ensure_font(paragraph.add_run("（2025-2026学年第二学期）"), size=12)
    document.add_paragraph()
    add_metadata_table(document)
    add_grading_table(document)
    document.add_page_break()

    add_heading(document, "摘要：", level=1)
    add_body_paragraph(document, "本文围绕学生选课系统展开课程设计，完成了需求分析、总体设计、数据库设计、核心流程设计和运行测试。系统采用分层结构组织登录、选课、退课、成绩录入和基础信息维护等模块。")
    add_heading(document, "关键词：", level=1)
    add_body_paragraph(document, "课程设计；学生选课系统；数据库设计；模块化；流程图")

    add_heading(document, "一、课程设计的目的与要求", level=1)
    add_body_paragraph(document, "课程设计要求综合运用数据结构、数据库和程序设计知识，完成一个具有明确业务流程和可验证结果的小型管理系统。")
    add_body_paragraph(document, "系统需要支持学生选课退课、教师成绩提交、管理员维护课程与用户信息，并输出可复核的运行结果。")

    add_heading(document, "二、设计正文", level=1)
    add_heading(document, "2.1 开发环境", level=2)
    add_body_paragraph(document, "课程设计在 Windows 11、Python 3、SQLite 和 VS Code 环境下完成。SQLite 用于保存学生、教师、课程和选课记录，便于在本地快速演示和验证核心业务流程。")

    add_heading(document, "2.2 需求分析", level=2)
    add_body_paragraph(document, "系统面向学生、教师和管理员三类用户。学生用户完成选课、退课和查看成绩；教师用户提交课程成绩；管理员维护学生、教师、课程和院系基础信息。")
    add_picture_block(document, assets["topology"], "图1 系统功能结构图")

    add_heading(document, "2.3 总体设计", level=2)
    add_body_paragraph(document, "系统采用界面层、业务逻辑层和数据访问层三层结构。界面层负责交互，业务逻辑层封装选课规则，数据访问层完成数据库读写。")
    module_table = document.add_table(rows=5, cols=3)
    module_table.style = "Table Grid"
    module_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    module_rows = [
        ["模块", "主要功能", "输出结果"],
        ["学生模块", "选课、退课、查询成绩", "选课记录和成绩列表"],
        ["教师模块", "提交和修改成绩", "课程成绩表"],
        ["管理员模块", "维护用户、课程和院系", "基础数据表"],
        ["公共模块", "登录、权限判断、安全退出", "会话状态"],
    ]
    for row_index, row in enumerate(module_rows):
        for col_index, value in enumerate(row):
            add_text(module_table.cell(row_index, col_index), value, size=10, bold=(row_index == 0))
            if row_index == 0:
                set_cell_shading(module_table.cell(row_index, col_index), "EDEDED")
    add_caption(document, "表3-1 学生选课系统功能模块表")

    add_heading(document, "2.4 数据库设计", level=2)
    add_body_paragraph(document, "数据库至少包含学生表、教师表、课程表、选课表和管理员表。选课表通过学生编号和课程编号建立业务关联，并记录选课状态。")
    db_table = document.add_table(rows=6, cols=4)
    db_table.style = "Table Grid"
    db_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    db_rows = [
        ["字段名", "类型", "约束", "说明"],
        ["StudentId", "varchar(20)", "PK", "学生编号"],
        ["StudentName", "nvarchar(50)", "NOT NULL", "学生姓名"],
        ["CourseId", "varchar(20)", "FK", "课程编号"],
        ["SelectedAt", "datetime", "NOT NULL", "选课时间"],
        ["Status", "int", "NOT NULL", "选课状态"],
    ]
    for row_index, row in enumerate(db_rows):
        for col_index, value in enumerate(row):
            add_text(db_table.cell(row_index, col_index), value, size=10, bold=(row_index == 0))
            if row_index == 0:
                set_cell_shading(db_table.cell(row_index, col_index), "EDEDED")
    add_caption(document, "表3-2 Elect选课信息表")

    add_heading(document, "2.5 详细设计与实现", level=2)
    add_body_paragraph(document, "用户登录后，系统根据角色加载不同功能菜单。学生提交选课请求时，业务层先检查课程容量和时间冲突，再写入选课表。")
    add_picture_block(document, assets["flow"], "图2 学生选课流程图")

    add_heading(document, "2.6 测试结果", level=2)
    test_table = document.add_table(rows=4, cols=4)
    test_table.style = "Table Grid"
    test_rows = [
        ["测试项", "输入", "预期结果", "实际结果"],
        ["学生登录", "正确账号密码", "进入学生主页", "通过"],
        ["重复选课", "已选课程再次提交", "提示不能重复选课", "通过"],
        ["教师提交成绩", "课程成绩明细", "保存并可查询", "通过"],
    ]
    for row_index, row in enumerate(test_rows):
        for col_index, value in enumerate(row):
            add_text(test_table.cell(row_index, col_index), value, size=10, bold=(row_index == 0))
            if row_index == 0:
                set_cell_shading(test_table.cell(row_index, col_index), "EDEDED")
    add_caption(document, "表5-1 系统功能测试表")
    add_picture_block(document, assets["terminal"], "图3 运行结果与测试输出")

    add_heading(document, "2.7 问题与改进", level=2)
    add_body_paragraph(document, "当前样例系统重点验证主流程，异常提示、日志记录和并发访问仍较简单。后续可以补充分页查询、课程推荐、权限细分和数据库事务处理，使系统更接近真实部署场景。")

    add_heading(document, "三、总结", level=1)
    add_body_paragraph(document, "本次课程设计完成了从需求到实现、从数据库到流程图、从测试到总结的完整文档链路。后续可以继续补充异常处理、日志记录和更细粒度的权限控制。")
    add_heading(document, "参考文献", level=1)
    add_body_paragraph(document, "[1] 数据结构课程设计指导书，校内讲义，2026。")
    add_body_paragraph(document, "[2] 数据库系统概论，第五版，高等教育出版社。")
    document.save(path)


def write_readme(path, fixtures):
    lines = [
        "# Realistic Report Fixtures",
        "",
        "Generated by `scripts/export-realistic-report-fixtures.ps1`.",
        "",
        "These files are synthetic, anonymized fixtures derived from common university report structures. They are safe to commit because they do not contain real student identities, school-owned source templates, or copied report bodies.",
        "",
        "| Fixture | Pattern | Purpose |",
        "| --- | --- | --- |",
    ]
    for fixture in fixtures:
        lines.append(f"| `{fixture['fileName']}` | {fixture['pattern']} | {fixture['purpose']} |")
    lines.extend([
        "",
        "Regenerate the pack with:",
        "",
        "```powershell",
        "powershell -ExecutionPolicy Bypass -File .\\scripts\\export-realistic-report-fixtures.ps1 -Overwrite",
        "```",
        "",
        "The standalone `student-course-selection-flowchart.png` asset is also generated and embedded as `图4-1` in `course-design-full-example.docx`.",
        "",
        "Use `scripts\\extract-docx-template.ps1` to inspect outlines before changing field-map or layout logic.",
        "",
    ])
    with open(path, "w", encoding="utf-8-sig", newline="\n") as handle:
        handle.write("\n".join(lines))


def main():
    with open(sys.argv[1], "r", encoding="utf-8-sig") as handle:
        spec = json.load(handle)
    output_dir = spec["outputDir"]
    os.makedirs(output_dir, exist_ok=True)

    fixtures = [
        {
            "fileName": "single-table-experiment-filled.docx",
            "pattern": "single-table framed experiment report",
            "purpose": "Filled one-table lab report with metadata, body sections, figures, and teacher-comment row.",
        },
        {
            "fileName": "integrated-experiment-multi-table.docx",
            "pattern": "multi-table integrated experiment report",
            "purpose": "Cover plus four report tables for design plan, method/data/reference, results, analysis, summary, and score.",
        },
        {
            "fileName": "course-design-full-example.docx",
            "pattern": "full course-design report example",
            "purpose": "Cover, grading table, abstract, keywords, module tables, database tables, diagrams, tests, and references.",
        },
    ]

    with tempfile.TemporaryDirectory() as asset_dir:
        assets = {
            "chart": os.path.join(asset_dir, "chart.png"),
            "terminal": os.path.join(asset_dir, "terminal.png"),
            "topology": os.path.join(asset_dir, "topology.png"),
            "flow": spec["flowchartPath"],
        }
        write_png(assets["chart"], 900, 360, (214, 231, 246))
        write_png(assets["terminal"], 900, 360, (232, 238, 224))
        write_png(assets["topology"], 900, 360, (240, 230, 216))

        build_single_table_experiment(os.path.join(output_dir, fixtures[0]["fileName"]), assets)
        build_integrated_experiment(os.path.join(output_dir, fixtures[1]["fileName"]), assets)
        build_course_design(os.path.join(output_dir, fixtures[2]["fileName"]), assets)

    for fixture in fixtures:
        fixture["outputPath"] = os.path.join(output_dir, fixture["fileName"])

    write_readme(os.path.join(output_dir, "README.md"), fixtures)
    summary = {
        "generatedCount": len(fixtures),
        "outputDir": output_dir,
        "assets": {
            "studentCourseSelectionFlowchart": spec["flowchartPath"],
        },
        "fixtures": fixtures,
    }
    with open(os.path.join(output_dir, "realistic-report-fixtures-summary.json"), "w", encoding="utf-8-sig", newline="\n") as handle:
        json.dump(summary, handle, ensure_ascii=False, indent=2)
        handle.write("\n")
    print(json.dumps(summary, ensure_ascii=False))


if __name__ == "__main__":
    main()
'@

[System.IO.File]::WriteAllText($scriptPath, $pythonScript, (New-Object System.Text.UTF8Encoding($true)))

try {
  $python = Get-Command python -ErrorAction SilentlyContinue
  if ($null -eq $python) {
    throw "python is required to generate realistic report fixtures."
  }

  $output = & $python.Source $scriptPath $specPath
  if ($LASTEXITCODE -ne 0) {
    throw "Fixture generation failed with exit code $LASTEXITCODE."
  }

  $summaryPath = Join-Path $resolvedOutputDir "realistic-report-fixtures-summary.json"
  if (-not (Test-Path -LiteralPath $summaryPath -PathType Leaf)) {
    throw "Fixture generation did not write summary: $summaryPath"
  }

  Get-Content -LiteralPath $summaryPath -Raw -Encoding UTF8 | ConvertFrom-Json
} finally {
  foreach ($tempPath in @($specPath, $scriptPath)) {
    if (Test-Path -LiteralPath $tempPath) {
      Remove-Item -LiteralPath $tempPath -Force
    }
  }
}
