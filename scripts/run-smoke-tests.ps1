[CmdletBinding()]
param(
  [string]$OpenClawCmd = $env:OPENCLAW_CMD
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Drawing

function Assert-True {
  param(
    [Parameter(Mandatory = $true)]
    [bool]$Condition,

    [Parameter(Mandatory = $true)]
    [string]$Message
  )

  if (-not $Condition) {
    throw $Message
  }
}

function Assert-ValidationPaginationRiskSummary {
  param(
    [Parameter(Mandatory = $true)]
    [object]$Summary,

    [Parameter(Mandatory = $true)]
    [string]$MessagePrefix,

    [AllowNull()]
    [string]$ExpectedLongRemediationPattern,

    [AllowNull()]
    [string]$ExpectedDenseRemediationPattern,

    [AllowNull()]
    [string]$ExpectedFigureRemediationPattern
  )

  $warningCodes = @($Summary.validationWarningCodes | ForEach-Object { [string]$_ })
  $warningSummaryCodes = @($Summary.validationWarningSummary | ForEach-Object { [string]$_.code })

  Assert-True -Condition ([bool]$Summary.validationPassed) -Message "$MessagePrefix should keep validationPassed=true when only pagination warnings are present."
  Assert-True -Condition ([int]$Summary.validationErrorCount -eq 0) -Message "$MessagePrefix should not report validation errors for the pagination-warning fixture."
  Assert-True -Condition ([int]$Summary.validationStructuralIssueCount -eq 0) -Message "$MessagePrefix should not report structural issues for the pagination-warning fixture."
  Assert-True -Condition ([int]$Summary.validationWarningCount -ge 3) -Message "$MessagePrefix should report pagination warning findings."
  Assert-True -Condition ([int]$Summary.validationPaginationRiskCount -ge 3) -Message "$MessagePrefix should report pagination risk findings."
  Assert-True -Condition ($warningCodes -contains 'pagination-risk-long-section') -Message "$MessagePrefix should expose pagination-risk-long-section."
  Assert-True -Condition ($warningCodes -contains 'pagination-risk-dense-section-block') -Message "$MessagePrefix should expose pagination-risk-dense-section-block."
  Assert-True -Condition ($warningCodes -contains 'pagination-risk-figure-cluster') -Message "$MessagePrefix should expose pagination-risk-figure-cluster."
  Assert-True -Condition ($warningSummaryCodes -contains 'pagination-risk-long-section') -Message "$MessagePrefix should include pagination-risk-long-section in validationWarningSummary."
  Assert-True -Condition ($warningSummaryCodes -contains 'pagination-risk-dense-section-block') -Message "$MessagePrefix should include pagination-risk-dense-section-block in validationWarningSummary."
  Assert-True -Condition ($warningSummaryCodes -contains 'pagination-risk-figure-cluster') -Message "$MessagePrefix should include pagination-risk-figure-cluster in validationWarningSummary."
  $longWarningSummary = @($Summary.validationWarningSummary | Where-Object { [string]$_.code -eq 'pagination-risk-long-section' })[0]
  $denseWarningSummary = @($Summary.validationWarningSummary | Where-Object { [string]$_.code -eq 'pagination-risk-dense-section-block' })[0]
  $figureWarningSummary = @($Summary.validationWarningSummary | Where-Object { [string]$_.code -eq 'pagination-risk-figure-cluster' })[0]
  Assert-True -Condition (-not [string]::IsNullOrWhiteSpace([string]$longWarningSummary.remediation)) -Message "$MessagePrefix should expose long-section remediation guidance."
  Assert-True -Condition (-not [string]::IsNullOrWhiteSpace([string]$denseWarningSummary.remediation)) -Message "$MessagePrefix should expose dense-section remediation guidance."
  Assert-True -Condition (-not [string]::IsNullOrWhiteSpace([string]$figureWarningSummary.remediation)) -Message "$MessagePrefix should expose figure-cluster remediation guidance."
  if ((-not [string]::IsNullOrWhiteSpace($ExpectedLongRemediationPattern)) -or (-not [string]::IsNullOrWhiteSpace($ExpectedDenseRemediationPattern)) -or (-not [string]::IsNullOrWhiteSpace($ExpectedFigureRemediationPattern))) {
    Assert-True -Condition ($Summary.PSObject.Properties.Name -contains 'validationPaginationRiskRemediations') -Message "$MessagePrefix should expose validationPaginationRiskRemediations."
  }
  if (-not [string]::IsNullOrWhiteSpace($ExpectedLongRemediationPattern)) {
    Assert-True -Condition ([string]$Summary.validationPaginationRiskRemediations.'pagination-risk-long-section' -match $ExpectedLongRemediationPattern) -Message "$MessagePrefix should expose the custom long-section remediation override."
    Assert-True -Condition ([string]$longWarningSummary.remediation -match $ExpectedLongRemediationPattern) -Message "$MessagePrefix should carry the custom long-section remediation into validationWarningSummary."
  }
  if (-not [string]::IsNullOrWhiteSpace($ExpectedDenseRemediationPattern)) {
    Assert-True -Condition ([string]$Summary.validationPaginationRiskRemediations.'pagination-risk-dense-section-block' -match $ExpectedDenseRemediationPattern) -Message "$MessagePrefix should expose the custom dense-section remediation override."
    Assert-True -Condition ([string]$denseWarningSummary.remediation -match $ExpectedDenseRemediationPattern) -Message "$MessagePrefix should carry the custom dense-section remediation into validationWarningSummary."
  }
  if (-not [string]::IsNullOrWhiteSpace($ExpectedFigureRemediationPattern)) {
    Assert-True -Condition ([string]$Summary.validationPaginationRiskRemediations.'pagination-risk-figure-cluster' -match $ExpectedFigureRemediationPattern) -Message "$MessagePrefix should expose the custom figure-cluster remediation override."
    Assert-True -Condition ([string]$figureWarningSummary.remediation -match $ExpectedFigureRemediationPattern) -Message "$MessagePrefix should carry the custom figure-cluster remediation into validationWarningSummary."
  }
}

function Normalize-OutlineForComparison {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Text
  )

  $normalizedLines = foreach ($line in ($Text -split "\r?\n")) {
    if ($line -notmatch '^\s*-\s*Source:\s+') {
      $line
    }
  }

  return (($normalizedLines -join [Environment]::NewLine).Trim())
}

function New-SampleTemplateDocx {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"@

  $relationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"@

  $documentRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>
"@

  $document = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>计算机网络实验报告</w:t></w:r></w:p>
    <w:p><w:r><w:t>课程名称：</w:t></w:r></w:p>
    <w:p><w:r><w:t>班级：__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>实验名称：保留原值</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>姓名</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>学号</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>指导教师：__________</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>实验目的</w:t></w:r></w:p>
    <w:p><w:r><w:t></w:t></w:r></w:p>
    <w:p><w:r><w:t>实验步骤</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>实验结果</w:t></w:r></w:p>
    <w:p><w:r><w:t></w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>
"@

  $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
  try {
    foreach ($entrySpec in @(
        @{ Name = "[Content_Types].xml"; Text = $contentTypes },
        @{ Name = "_rels/.rels"; Text = $relationships },
        @{ Name = "word/_rels/document.xml.rels"; Text = $documentRelationships },
        @{ Name = "word/document.xml"; Text = $document }
      )) {
      $entry = $zip.CreateEntry($entrySpec.Name)
      $writer = New-Object System.IO.StreamWriter($entry.Open(), (New-Object System.Text.UTF8Encoding($false)))
      try {
        $writer.Write($entrySpec.Text)
      } finally {
        $writer.Dispose()
      }
    }
  } finally {
    $zip.Dispose()
  }
}

function New-CourseDesignTemplateDocx {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"@

  $relationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"@

  $documentRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>
"@

  $document = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>课程设计报告</w:t></w:r></w:p>
    <w:p><w:r><w:t>课程名称：__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>课题名称：__________</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>学生姓名</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>学号</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>指导老师：__________</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>完成时间：__________</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>设计地点：__________</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>设计目标</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>开发环境</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>需求分析</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>方案设计与实现</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>运行结果</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>问题与改进</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>设计总结</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>
"@

  $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
  try {
    foreach ($entrySpec in @(
        @{ Name = "[Content_Types].xml"; Text = $contentTypes },
        @{ Name = "_rels/.rels"; Text = $relationships },
        @{ Name = "word/_rels/document.xml.rels"; Text = $documentRelationships },
        @{ Name = "word/document.xml"; Text = $document }
      )) {
      $entry = $zip.CreateEntry($entrySpec.Name)
      $writer = New-Object System.IO.StreamWriter($entry.Open(), (New-Object System.Text.UTF8Encoding($false)))
      try {
        $writer.Write($entrySpec.Text)
      } finally {
        $writer.Dispose()
      }
    }
  } finally {
    $zip.Dispose()
  }
}

function New-MonthlyReportTemplateDocx {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"@

  $relationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"@

  $documentRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>
"@

  $document = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>项目月报</w:t></w:r></w:p>
    <w:p><w:r><w:t>项目名称：__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>月报主题：__________</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>提交人</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>工号/学号</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>小组/班级</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>审核人</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>月份周期：__________</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>工作环境：__________</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>报告类型：__________</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>一、本月目标</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>协作环境</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>二、本月任务与输入</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>三、本月完成事项</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>四、阶段成果与数据</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>五、问题与改进</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>六、下月计划</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>
"@

  $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
  try {
    foreach ($entrySpec in @(
        @{ Name = "[Content_Types].xml"; Text = $contentTypes },
        @{ Name = "_rels/.rels"; Text = $relationships },
        @{ Name = "word/_rels/document.xml.rels"; Text = $documentRelationships },
        @{ Name = "word/document.xml"; Text = $document }
      )) {
      $entry = $zip.CreateEntry($entrySpec.Name)
      $writer = New-Object System.IO.StreamWriter($entry.Open(), (New-Object System.Text.UTF8Encoding($false)))
      try {
        $writer.Write($entrySpec.Text)
      } finally {
        $writer.Dispose()
      }
    }
  } finally {
    $zip.Dispose()
  }
}

function New-WeeklyReportTemplateDocx {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"@

  $relationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"@

  $documentRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>
"@

  $document = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>项目周报</w:t></w:r></w:p>
    <w:p><w:r><w:t>项目名称：__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>周报主题：__________</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>提交人</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>工号/学号</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>小组/班级</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>审核人</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>周次：__________</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>工作环境：__________</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>报告类型：__________</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>一、工作目标</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>协作环境</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>二、本周任务与输入</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>三、本周完成事项</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>四、阶段成果</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>五、风险与改进</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>六、下周计划</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>
"@

  $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
  try {
    foreach ($entrySpec in @(
        @{ Name = "[Content_Types].xml"; Text = $contentTypes },
        @{ Name = "_rels/.rels"; Text = $relationships },
        @{ Name = "word/_rels/document.xml.rels"; Text = $documentRelationships },
        @{ Name = "word/document.xml"; Text = $document }
      )) {
      $entry = $zip.CreateEntry($entrySpec.Name)
      $writer = New-Object System.IO.StreamWriter($entry.Open(), (New-Object System.Text.UTF8Encoding($false)))
      try {
        $writer.Write($entrySpec.Text)
      } finally {
        $writer.Dispose()
      }
    }
  } finally {
    $zip.Dispose()
  }
}

function New-SoftwareTestTemplateDocx {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"@

  $relationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"@

  $documentRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>
"@

  $document = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>软件测试报告</w:t></w:r></w:p>
    <w:p><w:r><w:t>课程名称：__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>测试项目：__________</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>学生姓名</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>学号</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>班级</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>指导教师</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>测试时间：__________</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>测试类型：__________</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>测试环境：__________</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>一、测试目标</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>二、测试环境</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>三、测试范围与依据</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>四、测试用例设计与执行</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>五、测试结果</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>六、缺陷分析与改进</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>七、测试总结</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>
"@

  $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
  try {
    foreach ($entrySpec in @(
        @{ Name = "[Content_Types].xml"; Text = $contentTypes },
        @{ Name = "_rels/.rels"; Text = $relationships },
        @{ Name = "word/_rels/document.xml.rels"; Text = $documentRelationships },
        @{ Name = "word/document.xml"; Text = $document }
      )) {
      $entry = $zip.CreateEntry($entrySpec.Name)
      $writer = New-Object System.IO.StreamWriter($entry.Open(), (New-Object System.Text.UTF8Encoding($false)))
      try {
        $writer.Write($entrySpec.Text)
      } finally {
        $writer.Dispose()
      }
    }
  } finally {
    $zip.Dispose()
  }
}

function New-DeploymentTemplateDocx {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"@

  $relationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"@

  $documentRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>
"@

  $document = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>部署运维报告</w:t></w:r></w:p>
    <w:p><w:r><w:t>课程名称：__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>部署项目：__________</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>学生姓名</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>学号</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>班级</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>指导教师</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>部署时间：__________</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>部署类型：__________</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>部署环境：__________</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>一、部署目标</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>二、部署环境</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>三、部署方案与架构</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>四、部署步骤与配置</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>五、验证结果</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>六、问题处理与回滚预案</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>七、部署总结</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>
"@

  $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
  try {
    foreach ($entrySpec in @(
        @{ Name = "[Content_Types].xml"; Text = $contentTypes },
        @{ Name = "_rels/.rels"; Text = $relationships },
        @{ Name = "word/_rels/document.xml.rels"; Text = $documentRelationships },
        @{ Name = "word/document.xml"; Text = $document }
      )) {
      $entry = $zip.CreateEntry($entrySpec.Name)
      $writer = New-Object System.IO.StreamWriter($entry.Open(), (New-Object System.Text.UTF8Encoding($false)))
      try {
        $writer.Write($entrySpec.Text)
      } finally {
        $writer.Dispose()
      }
    }
  } finally {
    $zip.Dispose()
  }
}

function New-MeetingMinutesTemplateDocx {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"@

  $relationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"@

  $documentRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>
"@

  $document = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>会议纪要</w:t></w:r></w:p>
    <w:p><w:r><w:t>项目名称：__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>会议主题：__________</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>记录人</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>工号/学号</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>参会小组</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>主持人</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>会议日期：__________</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>纪要类型：__________</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>会议地点：__________</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>一、会议目标</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>二、参会信息与背景</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>三、议题与输入</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>四、讨论过程与决议</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>五、当前结论</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>六、风险与争议</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>七、后续安排</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>
"@

  $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
  try {
    foreach ($entrySpec in @(
        @{ Name = "[Content_Types].xml"; Text = $contentTypes },
        @{ Name = "_rels/.rels"; Text = $relationships },
        @{ Name = "word/_rels/document.xml.rels"; Text = $documentRelationships },
        @{ Name = "word/document.xml"; Text = $document }
      )) {
      $entry = $zip.CreateEntry($entrySpec.Name)
      $writer = New-Object System.IO.StreamWriter($entry.Open(), (New-Object System.Text.UTF8Encoding($false)))
      try {
        $writer.Write($entrySpec.Text)
      } finally {
        $writer.Dispose()
      }
    }
  } finally {
    $zip.Dispose()
  }
}

function New-InternshipTemplateDocx {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"@

  $relationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"@

  $documentRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>
"@

  $document = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>专业实习报告</w:t></w:r></w:p>
    <w:p><w:r><w:t>专业名称：__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>实习项目：__________</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>学生姓名</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>学号</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>校内指导教师：__________</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>实习时间：__________</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>班级：__________</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>实习单位：__________</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>实习目的</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>实习单位与环境</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>实习任务与要求</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>实习过程与内容</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>实习成果</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>问题分析与改进</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>实习总结</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>
"@

  $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
  try {
    foreach ($entrySpec in @(
        @{ Name = "[Content_Types].xml"; Text = $contentTypes },
        @{ Name = "_rels/.rels"; Text = $relationships },
        @{ Name = "word/_rels/document.xml.rels"; Text = $documentRelationships },
        @{ Name = "word/document.xml"; Text = $document }
      )) {
      $entry = $zip.CreateEntry($entrySpec.Name)
      $writer = New-Object System.IO.StreamWriter($entry.Open(), (New-Object System.Text.UTF8Encoding($false)))
      try {
        $writer.Write($entrySpec.Text)
      } finally {
        $writer.Dispose()
      }
    }
  } finally {
    $zip.Dispose()
  }
}

function New-CoverBodyTemplateDocx {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"@

  $relationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"@

  $documentRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>
"@

  $document = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>实 验 报 告</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>学号：</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>姓名：</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>班级：</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>课程名称：</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>实验名称：</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>实验性质： ①综合性实验 ②设计性实验 ③验证性实验：</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>实验时间：</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>实验地点：</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>一. 实验目的二. 实验内容</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>三. 实验步骤六.实验小结</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t></w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>
"@

  $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
  try {
    foreach ($entrySpec in @(
        @{ Name = "[Content_Types].xml"; Text = $contentTypes },
        @{ Name = "_rels/.rels"; Text = $relationships },
        @{ Name = "word/_rels/document.xml.rels"; Text = $documentRelationships },
        @{ Name = "word/document.xml"; Text = $document }
      )) {
      $entry = $zip.CreateEntry($entrySpec.Name)
      $writer = New-Object System.IO.StreamWriter($entry.Open(), (New-Object System.Text.UTF8Encoding($false)))
      try {
        $writer.Write($entrySpec.Text)
      } finally {
        $writer.Dispose()
      }
    }
  } finally {
    $zip.Dispose()
  }
}

function New-ParagraphCoverTemplateDocx {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"@

  $relationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"@

  $documentRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>
"@

  $document = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>计算机网络实验报告</w:t></w:r></w:p>
    <w:p><w:r><w:t>课程名称：计算机网络</w:t></w:r></w:p>
    <w:p><w:r><w:t>实验名称：局域网搭建与常用 DOS 命令使用</w:t></w:r></w:p>
    <w:p><w:r><w:t>姓名：张三</w:t></w:r></w:p>
    <w:p><w:r><w:t>学号：20260001</w:t></w:r></w:p>
    <w:p><w:r><w:t>班级：计科 2201</w:t></w:r></w:p>
    <w:p><w:r><w:t>实验目的</w:t></w:r></w:p>
    <w:p><w:r><w:t>通过本次实验掌握局域网的基本搭建方法，并理解 DOS 命令在网络排查中的作用。</w:t></w:r></w:p>
    <w:p><w:r><w:t>实验总结</w:t></w:r></w:p>
    <w:p><w:r><w:t>本次实验完成了局域网搭建与常用 DOS 命令使用。</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>
"@

  $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
  try {
    foreach ($entrySpec in @(
        @{ Name = "[Content_Types].xml"; Text = $contentTypes },
        @{ Name = "_rels/.rels"; Text = $relationships },
        @{ Name = "word/_rels/document.xml.rels"; Text = $documentRelationships },
        @{ Name = "word/document.xml"; Text = $document }
      )) {
      $entry = $zip.CreateEntry($entrySpec.Name)
      $writer = New-Object System.IO.StreamWriter($entry.Open(), (New-Object System.Text.UTF8Encoding($false)))
      try {
        $writer.Write($entrySpec.Text)
      } finally {
        $writer.Dispose()
      }
    }
  } finally {
    $zip.Dispose()
  }
}

function New-FieldMapDiagnosticTemplateDocx {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"@

  $relationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"@

  $documentRelationships = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>
"@

  $document = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>模板适配诊断用例</w:t></w:r></w:p>
    <w:p><w:r><w:t>实验地点：__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>实验台号：__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>问题分析</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:p><w:r><w:t>实验器材与拓扑</w:t></w:r></w:p>
    <w:p><w:r><w:t>__________</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>实验目的 / 实验结果</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr/>
  </w:body>
</w:document>
"@

  $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
  try {
    foreach ($entrySpec in @(
        @{ Name = "[Content_Types].xml"; Text = $contentTypes },
        @{ Name = "_rels/.rels"; Text = $relationships },
        @{ Name = "word/_rels/document.xml.rels"; Text = $documentRelationships },
        @{ Name = "word/document.xml"; Text = $document }
      )) {
      $entry = $zip.CreateEntry($entrySpec.Name)
      $writer = New-Object System.IO.StreamWriter($entry.Open(), (New-Object System.Text.UTF8Encoding($false)))
      try {
        $writer.Write($entrySpec.Text)
      } finally {
        $writer.Dispose()
      }
    }
  } finally {
    $zip.Dispose()
  }
}

function New-SamplePngImage {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [string]$Text,

    [string]$BackgroundHex = "#E8F1FB",

    [int]$Width = 360,

    [int]$Height = 200
  )

  $bitmap = New-Object System.Drawing.Bitmap $Width, $Height
  $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
  try {
    $graphics.Clear([System.Drawing.ColorTranslator]::FromHtml($BackgroundHex))
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $font = New-Object System.Drawing.Font("Microsoft YaHei", 18, [System.Drawing.FontStyle]::Bold)
    $brush = [System.Drawing.Brushes]::Black
    $graphics.DrawString($Text, $font, $brush, 18, 74)
    $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
  } finally {
    if ($null -ne $font) {
      $font.Dispose()
    }
    $graphics.Dispose()
    $bitmap.Dispose()
  }
}

function Test-PowerShellSyntax {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $tokens = $null
  $errors = $null
  [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors) | Out-Null
  Assert-True -Condition ($errors.Count -eq 0) -Message ("Syntax errors in {0}: {1}" -f $Path, (($errors | ForEach-Object { $_.Message }) -join "; "))
}

function Test-HasUtf8Bom {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $bytes = [System.IO.File]::ReadAllBytes($Path)
  return $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
}

function Test-HasNonAsciiText {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $text = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
  foreach ($character in $text.ToCharArray()) {
    if ([int][char]$character -gt 127) {
      return $true
    }
  }

  return $false
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("openclaw-exp-report-smoke-" + [System.Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

$results = New-Object System.Collections.Generic.List[string]

try {
  foreach ($requiredPath in @(
      (Join-Path $repoRoot 'SKILL.md'),
      (Join-Path $repoRoot '.gitattributes'),
      (Join-Path $repoRoot 'CHANGELOG.md'),
      (Join-Path $repoRoot 'CODE_OF_CONDUCT.md'),
      (Join-Path $repoRoot 'CONTRIBUTING.md'),
      (Join-Path $repoRoot 'ROADMAP.md'),
      (Join-Path $repoRoot 'SECURITY.md'),
      (Join-Path $repoRoot 'SUPPORT.md'),
      (Join-Path $repoRoot 'agents\openai.yaml'),
      (Join-Path $repoRoot '.github\ISSUE_TEMPLATE\bug_report.md'),
      (Join-Path $repoRoot '.github\ISSUE_TEMPLATE\feature_request.md'),
      (Join-Path $repoRoot '.github\ISSUE_TEMPLATE\document_profile_request.md'),
      (Join-Path $repoRoot '.github\ISSUE_TEMPLATE\config.yml'),
      (Join-Path $repoRoot '.github\pull_request_template.md'),
      (Join-Path $repoRoot '.github\workflows\quality.yml'),
      (Join-Path $repoRoot '.github\workflows\roadmap-triage.yml'),
      (Join-Path $repoRoot '.github\workflows\smoke-tests.yml'),
      (Join-Path $repoRoot 'demo\README.md'),
      (Join-Path $repoRoot 'demo\assets\step-network-config.png'),
      (Join-Path $repoRoot 'demo\assets\step-ipconfig.png'),
      (Join-Path $repoRoot 'demo\assets\result-ping.png'),
      (Join-Path $repoRoot 'demo\assets\result-arp.png'),
      (Join-Path $repoRoot 'examples\docx-field-map.json'),
      (Join-Path $repoRoot 'examples\docx-image-map.json'),
      (Join-Path $repoRoot 'examples\docx-image-map-row.json'),
      (Join-Path $repoRoot 'examples\docx-image-specs.json'),
      (Join-Path $repoRoot 'examples\docx-image-specs-row.json'),
      (Join-Path $repoRoot 'examples\docx-report-metadata.json'),
      (Join-Path $repoRoot 'examples\profile-presets\README.md'),
      (Join-Path $repoRoot 'examples\profile-presets\weekly-report.json'),
      (Join-Path $repoRoot 'examples\profile-presets\meeting-minutes.json'),
      (Join-Path $repoRoot 'examples\profile-presets\monthly-report.json'),
      (Join-Path $repoRoot 'examples\feishu-uploaded-images-docx-prompt.md'),
      (Join-Path $repoRoot 'examples\local-uploaded-images-docx-prompt.md'),
      (Join-Path $repoRoot 'examples\sample-report.txt'),
      (Join-Path $repoRoot 'examples\e2e-sample-requirements.json'),
      (Join-Path $repoRoot 'profiles\experiment-report.json'),
      (Join-Path $repoRoot 'profiles\course-design-report.json'),
      (Join-Path $repoRoot 'profiles\internship-report.json'),
      (Join-Path $repoRoot 'profiles\software-test-report.json'),
      (Join-Path $repoRoot 'profiles\deployment-report.json'),
      (Join-Path $repoRoot 'profiles\weekly-report.json'),
      (Join-Path $repoRoot 'profiles\meeting-minutes.json'),
      (Join-Path $repoRoot 'profiles\monthly-report.json'),
      (Join-Path $repoRoot 'profiles\report-profile.schema.json'),
      (Join-Path $repoRoot 'references\template-fit.md'),
      (Join-Path $repoRoot 'scripts\analyze-roadmap-next-step.ps1'),
      (Join-Path $repoRoot 'scripts\apply-docx-field-map.ps1'),
      (Join-Path $repoRoot 'scripts\build-report.ps1'),
      (Join-Path $repoRoot 'scripts\build-report-from-feishu.ps1'),
      (Join-Path $repoRoot 'scripts\build-report-from-url.ps1'),
      (Join-Path $repoRoot 'scripts\check-report-profile-template-fit.ps1'),
      (Join-Path $repoRoot 'scripts\check-docx-layout.ps1'),
      (Join-Path $repoRoot 'scripts\convert-docx-template-frame.ps1'),
      (Join-Path $repoRoot 'scripts\extract-docx-template.ps1'),
      (Join-Path $repoRoot 'scripts\fetch-csdn-article.ps1'),
      (Join-Path $repoRoot 'scripts\fetch-web-article.ps1'),
      (Join-Path $repoRoot 'scripts\format-docx-report-style.ps1'),
      (Join-Path $repoRoot 'scripts\generate-docx-field-map.ps1'),
      (Join-Path $repoRoot 'scripts\generate-docx-image-map.ps1'),
      (Join-Path $repoRoot 'scripts\generate-report-inputs.ps1'),
      (Join-Path $repoRoot 'scripts\generate-report-chat.ps1'),
      (Join-Path $repoRoot 'scripts\install-skill.ps1'),
      (Join-Path $repoRoot 'scripts\insert-docx-images.ps1'),
      (Join-Path $repoRoot 'scripts\new-report-profile.ps1'),
      (Join-Path $repoRoot 'scripts\prepare-report-prompt.ps1'),
      (Join-Path $repoRoot 'scripts\report-defaults.ps1'),
      (Join-Path $repoRoot 'scripts\report-profiles.ps1'),
      (Join-Path $repoRoot 'scripts\reset-openclaw-session.ps1'),
      (Join-Path $repoRoot 'scripts\run-e2e-sample.ps1'),
      (Join-Path $repoRoot 'scripts\run-profile-preset-samples.ps1'),
      (Join-Path $repoRoot 'scripts\self-check.ps1'),
      (Join-Path $repoRoot 'scripts\validate-report-draft.ps1'),
      (Join-Path $repoRoot 'scripts\validate-report-profiles.ps1')
    )) {
    Assert-True -Condition (Test-Path -LiteralPath $requiredPath) -Message "Missing required path: $requiredPath"
  }
  $results.Add('repository structure OK') | Out-Null

  Assert-True -Condition (-not (Test-HasUtf8Bom -Path (Join-Path $repoRoot 'SKILL.md'))) -Message 'SKILL.md must not start with a UTF-8 BOM because OpenClaw frontmatter parsing will fail.'
  $results.Add('skill frontmatter encoding OK') | Out-Null

  $nonAsciiPowerShellScriptsWithoutBom = @(
    Get-ChildItem -LiteralPath (Join-Path $repoRoot 'scripts') -Filter *.ps1 |
      Where-Object { (Test-HasNonAsciiText -Path $_.FullName) -and -not (Test-HasUtf8Bom -Path $_.FullName) } |
      ForEach-Object { $_.FullName }
  )
  Assert-True -Condition ($nonAsciiPowerShellScriptsWithoutBom.Count -eq 0) -Message ("PowerShell scripts with non-ASCII text must include a UTF-8 BOM for Windows PowerShell 5.1: {0}" -f ($nonAsciiPowerShellScriptsWithoutBom -join ", "))
  $results.Add('PowerShell script encoding OK') | Out-Null

  $exampleFieldMap = (Get-Content -LiteralPath (Join-Path $repoRoot 'examples\docx-field-map.json') -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ($null -ne $exampleFieldMap) -Message 'Example field-map JSON did not parse.'
  Assert-True -Condition ($exampleFieldMap.PSObject.Properties.Name -contains '课程名称') -Message 'Example field-map JSON is missing the course-name key.'
  Assert-True -Condition ($exampleFieldMap.PSObject.Properties.Name -contains 'P4') -Message 'Example field-map JSON is missing the example location key.'
  Assert-True -Condition ($exampleFieldMap.PSObject.Properties.Name -contains '实验目的') -Message 'Example field-map JSON is missing the example block key.'
  Assert-True -Condition ($exampleFieldMap.实验目的.paragraphs.Count -ge 2) -Message 'Example field-map JSON is missing example block paragraphs.'
  Assert-True -Condition ($exampleFieldMap.PSObject.Properties.Name -contains 'P10') -Message 'Example field-map JSON is missing the example block location key.'
  $results.Add('example field map JSON OK') | Out-Null

  $exampleImageMap = (Get-Content -LiteralPath (Join-Path $repoRoot 'examples\docx-image-map.json') -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ($null -ne $exampleImageMap) -Message 'Example image-map JSON did not parse.'
  Assert-True -Condition (@($exampleImageMap.images).Count -ge 2) -Message 'Example image-map JSON is missing example images.'
  Assert-True -Condition ([string]$exampleImageMap.images[0].anchor -eq 'P8') -Message 'Example image-map JSON is missing the expected paragraph anchor.'
  Assert-True -Condition ([string]$exampleImageMap.images[1].anchor -eq 'T1R6C1') -Message 'Example image-map JSON is missing the expected cell anchor.'
  $results.Add('example image map JSON OK') | Out-Null

  $exampleRowImageMap = (Get-Content -LiteralPath (Join-Path $repoRoot 'examples\docx-image-map-row.json') -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ($null -ne $exampleRowImageMap) -Message 'Example row image-map JSON did not parse.'
  Assert-True -Condition (@($exampleRowImageMap.images).Count -eq 4) -Message 'Example row image-map JSON should include four images.'
  Assert-True -Condition ([string]$exampleRowImageMap.images[0].layout.mode -eq 'row') -Message 'Example row image-map JSON is missing the row layout mode.'
  Assert-True -Condition ([int]$exampleRowImageMap.images[0].layout.columns -eq 2) -Message 'Example row image-map JSON is missing the expected layout columns.'
  $results.Add('example row image map JSON OK') | Out-Null

  $exampleImageSpecs = (Get-Content -LiteralPath (Join-Path $repoRoot 'examples\docx-image-specs.json') -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ($null -ne $exampleImageSpecs) -Message 'Example image-specs JSON did not parse.'
  Assert-True -Condition (@($exampleImageSpecs.images).Count -ge 2) -Message 'Example image-specs JSON is missing example images.'
  Assert-True -Condition ([string]$exampleImageSpecs.images[0].section -eq '实验步骤') -Message 'Example image-specs JSON is missing the expected first section.'
  Assert-True -Condition ([string]$exampleImageSpecs.images[1].section -eq '实验结果') -Message 'Example image-specs JSON is missing the expected second section.'
  $results.Add('example image specs JSON OK') | Out-Null

  $exampleRowImageSpecs = (Get-Content -LiteralPath (Join-Path $repoRoot 'examples\docx-image-specs-row.json') -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ($null -ne $exampleRowImageSpecs) -Message 'Example row image-specs JSON did not parse.'
  Assert-True -Condition (@($exampleRowImageSpecs.images).Count -eq 4) -Message 'Example row image-specs JSON should include four images.'
  Assert-True -Condition ([string]$exampleRowImageSpecs.images[0].layout.mode -eq 'row') -Message 'Example row image-specs JSON is missing the row layout mode.'
  Assert-True -Condition ([int]$exampleRowImageSpecs.images[0].layout.columns -eq 2) -Message 'Example row image-specs JSON is missing the expected layout columns.'
  $results.Add('example row image specs JSON OK') | Out-Null

  $exampleMetadata = (Get-Content -LiteralPath (Join-Path $repoRoot 'examples\docx-report-metadata.json') -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ($null -ne $exampleMetadata) -Message 'Example docx metadata JSON did not parse.'
  Assert-True -Condition ([string]$exampleMetadata.姓名 -eq '张三') -Message 'Example docx metadata JSON is missing the expected student name.'
  Assert-True -Condition ([string]$exampleMetadata.学号 -eq '20260001') -Message 'Example docx metadata JSON is missing the expected student id.'
  Assert-True -Condition ([string]$exampleMetadata.课程名称 -eq '计算机网络') -Message 'Example docx metadata JSON is missing the expected course name.'
  Assert-True -Condition ([string]$exampleMetadata.实验性质 -eq '③验证性实验') -Message 'Example docx metadata JSON is missing the expected experiment property.'
  $results.Add('example docx metadata JSON OK') | Out-Null

  $feishuUploadedPromptExample = Get-Content -LiteralPath (Join-Path $repoRoot 'examples\feishu-uploaded-images-docx-prompt.md') -Raw -Encoding UTF8
  Assert-True -Condition ($feishuUploadedPromptExample -match '\[media attached') -Message 'Feishu uploaded-images prompt example is missing the media-attached extraction guidance.'
  Assert-True -Condition ($feishuUploadedPromptExample -match '最终 docx 必须真正插入图片文件') -Message 'Feishu uploaded-images prompt example is missing the final docx insertion requirement.'
  Assert-True -Condition ($feishuUploadedPromptExample -match '可以省略') -Message 'Feishu uploaded-images prompt example is missing the remembered-name guidance.'
  $results.Add('example Feishu uploaded-images prompt OK') | Out-Null

  $localUploadedPromptExample = Get-Content -LiteralPath (Join-Path $repoRoot 'examples\local-uploaded-images-docx-prompt.md') -Raw -Encoding UTF8
  Assert-True -Condition ($localUploadedPromptExample -match '\[media attached') -Message 'Local uploaded-images prompt example is missing the media-attached extraction guidance.'
  Assert-True -Condition ($localUploadedPromptExample -match '本地上传图片直接插入 docx') -Message 'Local uploaded-images prompt example is missing the local-upload insertion guard.'
  Assert-True -Condition ($localUploadedPromptExample -match '最终 docx 必须真正插入图片文件') -Message 'Local uploaded-images prompt example is missing the final docx insertion requirement.'
  $results.Add('example local uploaded-images prompt OK') | Out-Null

  $oneShotUploadedPromptExample = Get-Content -LiteralPath (Join-Path $repoRoot 'examples\one-shot-uploaded-images-docx-prompt.md') -Raw -Encoding UTF8
  Assert-True -Condition ($oneShotUploadedPromptExample -match '不要中途让我确认') -Message 'One-shot uploaded-images prompt example is missing the no-confirmation guidance.'
  Assert-True -Condition ($oneShotUploadedPromptExample -match '-PlanOnly') -Message 'One-shot uploaded-images prompt example is missing the image placement planning command.'
  Assert-True -Condition ($oneShotUploadedPromptExample -match '低置信度') -Message 'One-shot uploaded-images prompt example is missing the low-confidence handling guidance.'
  $results.Add('example one-shot uploaded-images prompt OK') | Out-Null

  $exampleRequirements = (Get-Content -LiteralPath (Join-Path $repoRoot 'examples\e2e-sample-requirements.json') -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ($null -ne $exampleRequirements) -Message 'Example e2e requirements JSON did not parse.'
  Assert-True -Condition ($exampleRequirements.sections.Count -ge 7) -Message 'Example e2e requirements JSON is missing required sections.'
  Assert-True -Condition ($exampleRequirements.requiredKeywords.Count -ge 5) -Message 'Example e2e requirements JSON is missing required keywords.'
  $results.Add('example e2e requirements JSON OK') | Out-Null

  $demoReadme = Get-Content -LiteralPath (Join-Path $repoRoot 'demo\README.md') -Raw -Encoding UTF8
  Assert-True -Condition ($demoReadme -match '2x2 Layout Preview') -Message 'Demo README is missing the expected layout preview section.'
  Assert-True -Condition ($demoReadme -match 'demo-grid') -Message 'Demo README is missing the shared row-layout example.'
  $results.Add('demo documentation OK') | Out-Null

  $repoReadme = Get-Content -LiteralPath (Join-Path $repoRoot 'README.md') -Raw -Encoding UTF8
  Assert-True -Condition ($repoReadme -match 'build-report-from-feishu\.ps1') -Message 'README is missing the Feishu wrapper script documentation.'
  Assert-True -Condition ($repoReadme -match 'DetailLevel full') -Message 'README is missing the richer detail-level usage guidance.'
  Assert-True -Condition ($repoReadme -match 'uploaded images and you also provide local image paths') -Message 'README is missing the hybrid uploaded-image plus local-path guidance.'
  Assert-True -Condition ($repoReadme -match 'media/inbound/example\.png') -Message 'README is missing the uploaded-attachment path guidance.'
  Assert-True -Condition ($repoReadme -match 'can omit `-ExperimentName`') -Message 'README is missing the remembered experiment-name guidance.'
  Assert-True -Condition ($repoReadme -match 'generate-report-inputs\.ps1') -Message 'README is missing the report-input generation script documentation.'
  Assert-True -Condition ($repoReadme -match '按 report profile 隔离保存') -Message 'README is missing the per-profile defaults guidance.'
  Assert-True -Condition ($repoReadme -match 'ROADMAP\.md') -Message 'README is missing the roadmap link.'
  Assert-True -Condition ($repoReadme -match 'Repository Health') -Message 'README is missing the repository health section.'
  $results.Add('README wrapper documentation OK') | Out-Null

  $roadmapText = Get-Content -LiteralPath (Join-Path $repoRoot 'ROADMAP.md') -Raw -Encoding UTF8
  Assert-True -Condition ($roadmapText -match 'document profiles') -Message 'ROADMAP.md is missing the document-profile direction.'
  $results.Add('roadmap documentation OK') | Out-Null

  . (Join-Path $repoRoot 'scripts\report-defaults.ps1')
  $defaultsTempPath = Join-Path $tempRoot 'experiment-report.defaults.json'
  $savedDefaultsPath = Save-ExperimentReportDefaults -CourseName '计算机网络' -ExperimentName '局域网搭建与常用 DOS 命令使用' -DefaultsPath $defaultsTempPath
  Assert-True -Condition ($savedDefaultsPath -eq $defaultsTempPath) -Message 'Report-defaults helper returned an unexpected defaults path.'
  Assert-True -Condition (Test-Path -LiteralPath $defaultsTempPath) -Message 'Report-defaults helper did not create the defaults file.'
  $resolvedSavedNames = Resolve-ExperimentReportNames -CourseName '' -ExperimentName '' -DefaultsPath $defaultsTempPath
  Assert-True -Condition ([string]$resolvedSavedNames.courseName -eq '计算机网络') -Message 'Report-defaults helper did not restore the saved course name.'
  Assert-True -Condition ([string]$resolvedSavedNames.experimentName -eq '局域网搭建与常用 DOS 命令使用') -Message 'Report-defaults helper did not restore the saved experiment name.'
  Assert-True -Condition ([bool]$resolvedSavedNames.usedStoredExperimentName) -Message 'Report-defaults helper should report that it reused the stored experiment name.'
  $promptInferredExperimentName = Resolve-InferredExperimentName -PromptText '实验名称：交换机 VLAN 配置实验' -ReferenceUrls @() -ReferenceTextPaths @()
  Assert-True -Condition ([string]$promptInferredExperimentName -eq '交换机 VLAN 配置实验') -Message 'Report-defaults helper did not infer the experiment name from prompt text.'
  $referenceTitlePath = Join-Path $tempRoot 'reference-title.txt'
  @'
TITLE: 路由器静态路由配置实验 - CSDN博客
URL: https://example.com/network-lab

正文内容
'@ | Set-Content -LiteralPath $referenceTitlePath -Encoding UTF8
  $referenceInferredExperimentName = Resolve-InferredExperimentName -ReferenceTextPaths @($referenceTitlePath) -ReferenceUrls @()
  Assert-True -Condition ([string]$referenceInferredExperimentName -eq '路由器静态路由配置实验') -Message 'Report-defaults helper did not infer the experiment name from reference text title.'
  $urlInferredExperimentName = Resolve-InferredExperimentName -ReferenceUrls @('https://example.com/labs/%E5%B1%80%E5%9F%9F%E7%BD%91%E6%90%AD%E5%BB%BA%E4%B8%8E%E5%B8%B8%E7%94%A8DOS%E5%91%BD%E4%BB%A4%E4%BD%BF%E7%94%A8.html')
  Assert-True -Condition ([string]$urlInferredExperimentName -eq '局域网搭建与常用DOS命令使用') -Message 'Report-defaults helper did not infer the experiment name from the URL slug.'
  $resolvedInferredNames = Resolve-ExperimentReportNames -CourseName '计算机网络' -ExperimentName '' -InferredExperimentName '交换机 VLAN 配置实验' -DefaultsPath $defaultsTempPath
  Assert-True -Condition ([string]$resolvedInferredNames.experimentName -eq '交换机 VLAN 配置实验') -Message 'Report-defaults helper should prefer inferred experiment names over stored defaults.'
  Assert-True -Condition ([bool]$resolvedInferredNames.usedInferredExperimentName) -Message 'Report-defaults helper should report that it used an inferred experiment name.'
  Assert-True -Condition (-not [bool]$resolvedInferredNames.usedStoredExperimentName) -Message 'Report-defaults helper should not report stored experiment-name reuse when inference wins.'
  $originalAgentsHome = $env:AGENTS_HOME
  $profileDefaultsHome = Join-Path $tempRoot 'profile-defaults-home'
  try {
    $env:AGENTS_HOME = $profileDefaultsHome
    $savedExperimentDefaultsPath = Save-ExperimentReportDefaults -CourseName '计算机网络' -ExperimentName '局域网搭建与常用 DOS 命令使用'
    $savedCourseDesignDefaultsPath = Save-ExperimentReportDefaults -CourseName '软件工程综合实践' -ExperimentName '校园导览小程序设计' -ReportProfileName 'course-design-report'
    Assert-True -Condition ((Split-Path -Leaf $savedExperimentDefaultsPath) -eq 'experiment-report.defaults.json') -Message 'Report-defaults helper should keep the experiment-report defaults file name.'
    Assert-True -Condition ((Split-Path -Leaf $savedCourseDesignDefaultsPath) -eq 'course-design-report.defaults.json') -Message 'Report-defaults helper should isolate course-design defaults by profile name.'
    $resolvedExperimentDefaults = Resolve-ExperimentReportNames -CourseName '' -ExperimentName ''
    $resolvedCourseDesignDefaults = Resolve-ExperimentReportNames -CourseName '' -ExperimentName '' -ReportProfileName 'course-design-report'
    Assert-True -Condition ([string]$resolvedExperimentDefaults.experimentName -eq '局域网搭建与常用 DOS 命令使用') -Message 'Report-defaults helper lost the experiment-report stored title.'
    Assert-True -Condition ([string]$resolvedCourseDesignDefaults.experimentName -eq '校园导览小程序设计') -Message 'Report-defaults helper should restore the course-design stored title from its own defaults file.'
    Assert-True -Condition ([string]$resolvedExperimentDefaults.defaultsPath -ne [string]$resolvedCourseDesignDefaults.defaultsPath) -Message 'Report-defaults helper should not share defaults paths across report profiles.'
  } finally {
    $env:AGENTS_HOME = $originalAgentsHome
  }
  $results.Add('report defaults helper OK') | Out-Null

  . (Join-Path $repoRoot 'scripts\report-profiles.ps1')
  $reportProfile = Get-ReportProfile -RepoRoot $repoRoot
  Assert-True -Condition ([string]$reportProfile.name -eq 'experiment-report') -Message 'Report profile loader returned an unexpected profile name.'
  Assert-True -Condition ([string]$reportProfile.displayName -eq '实验报告') -Message 'Report profile loader returned an unexpected display name.'
  Assert-True -Condition ([string]$reportProfile.defaultExperimentProperty -eq '③验证性实验') -Message 'Report profile is missing the default experiment property.'
  $reportProfileLabels = Get-ReportProfileLabels -Profile $reportProfile
  Assert-True -Condition ([string]$reportProfileLabels['CourseName'] -eq '课程名称') -Message 'Report profile labels are missing the course-name field.'
  Assert-True -Condition ([string]$reportProfileLabels['Results'] -eq '实验结果') -Message 'Report profile labels are missing the results heading.'
  $reportProfileSections = @(Get-ReportProfileSectionFields -Profile $reportProfile)
  Assert-True -Condition ($reportProfileSections.Count -ge 7) -Message 'Report profile is missing required section definitions.'
  Assert-True -Condition ((Get-ReportProfileRequiredHeadings -Profile $reportProfile) -contains '问题分析') -Message 'Report profile required headings are missing 问题分析.'
  $fullDetailProfile = Get-ReportProfileDetailProfile -Profile $reportProfile -DetailLevel 'full'
  Assert-True -Condition ([int]$fullDetailProfile.minChars -eq 1100) -Message 'Report profile full detail level is missing the expected minChars.'
  Assert-True -Condition ([string](Get-ReportProfileDefaultStyleProfile -Profile $reportProfile) -eq 'auto') -Message 'Report profile is missing the expected defaultStyleProfile.'
  Assert-True -Condition ((Get-ReportProfileMetadataPrefixes -Profile $reportProfile) -contains '课程名称') -Message 'Report profile metadata prefixes are missing 课程名称.'
  Assert-True -Condition ((Get-ReportProfileExtraSectionHeadings -Profile $reportProfile) -contains '实验内容') -Message 'Report profile extra section headings are missing 实验内容.'
  Assert-True -Condition ([string](Get-ReportProfileDefaultImageCaptionBody -Profile $reportProfile -SectionId 'steps' -BaseName 'setup-step') -eq '实验步骤截图') -Message 'Report profile image caption defaults are missing the steps caption.'
  Assert-True -Condition (@(Get-ReportProfileFieldMapCompositeRules -Profile $reportProfile).Count -ge 2) -Message 'Report profile field-map composite rules are missing.'
  $reportProfilePaginationRiskThresholds = Get-ReportProfilePaginationRiskThresholds -Profile $reportProfile
  Assert-True -Condition ([int]$reportProfilePaginationRiskThresholds.longSectionChars -eq 900) -Message 'Report profile pagination-risk long-section threshold is missing.'
  Assert-True -Condition ([int]$reportProfilePaginationRiskThresholds.denseSectionChars -eq 550) -Message 'Report profile pagination-risk dense-section threshold is missing.'
  Assert-True -Condition ([int]$reportProfilePaginationRiskThresholds.denseSectionParagraphs -eq 2) -Message 'Report profile pagination-risk dense paragraph threshold is missing.'
  Assert-True -Condition ([int]$reportProfilePaginationRiskThresholds.figureClusterRefs -eq 3) -Message 'Report profile pagination-risk figure-cluster threshold is missing.'
  $courseDesignProfile = Get-ReportProfile -ProfileName 'course-design-report' -RepoRoot $repoRoot
  Assert-True -Condition ([string]$courseDesignProfile.name -eq 'course-design-report') -Message 'Course-design profile loader returned an unexpected profile name.'
  Assert-True -Condition ([string]$courseDesignProfile.displayName -eq '课程设计报告') -Message 'Course-design profile loader returned an unexpected display name.'
  Assert-True -Condition ([string](Get-ReportProfileDefaultStyleProfile -Profile $courseDesignProfile) -eq 'school') -Message 'Course-design profile is missing the expected defaultStyleProfile.'
  Assert-True -Condition ((Get-ReportProfileMetadataPrefixes -Profile $courseDesignProfile) -contains '指导老师') -Message 'Course-design profile metadata prefixes are missing 指导老师.'
  Assert-True -Condition ((Get-ReportProfileRequiredHeadings -Profile $courseDesignProfile) -contains '方案设计与实现') -Message 'Course-design profile required headings are missing 方案设计与实现.'
  Assert-True -Condition ([string](Get-ReportProfileDefaultImageCaptionBody -Profile $courseDesignProfile -SectionId 'result' -BaseName 'ui-home') -eq '运行结果截图') -Message 'Course-design profile image caption defaults are missing the result caption.'
  $internshipProfile = Get-ReportProfile -ProfileName 'internship-report' -RepoRoot $repoRoot
  Assert-True -Condition ([string]$internshipProfile.name -eq 'internship-report') -Message 'Internship profile loader returned an unexpected profile name.'
  Assert-True -Condition ([string]$internshipProfile.displayName -eq '专业实习报告') -Message 'Internship profile loader returned an unexpected display name.'
  Assert-True -Condition ([string](Get-ReportProfileDefaultStyleProfile -Profile $internshipProfile) -eq 'school') -Message 'Internship profile is missing the expected defaultStyleProfile.'
  Assert-True -Condition ((Get-ReportProfileMetadataPrefixes -Profile $internshipProfile) -contains '校内指导教师') -Message 'Internship profile metadata prefixes are missing 校内指导教师.'
  Assert-True -Condition ((Get-ReportProfileRequiredHeadings -Profile $internshipProfile) -contains '实习过程与内容') -Message 'Internship profile required headings are missing 实习过程与内容.'
  Assert-True -Condition ([string](Get-ReportProfileDefaultImageCaptionBody -Profile $internshipProfile -SectionId 'result' -BaseName 'intern-home') -eq '实习成果截图') -Message 'Internship profile image caption defaults are missing the result caption.'
  $softwareTestProfile = Get-ReportProfile -ProfileName 'software-test-report' -RepoRoot $repoRoot
  Assert-True -Condition ([string]$softwareTestProfile.name -eq 'software-test-report') -Message 'Software-test profile loader returned an unexpected profile name.'
  Assert-True -Condition ([string]$softwareTestProfile.displayName -eq '软件测试报告') -Message 'Software-test profile loader returned an unexpected display name.'
  Assert-True -Condition ([string](Get-ReportProfileDefaultStyleProfile -Profile $softwareTestProfile) -eq 'school') -Message 'Software-test profile is missing the expected defaultStyleProfile.'
  Assert-True -Condition ((Get-ReportProfileMetadataPrefixes -Profile $softwareTestProfile) -contains '测试项目') -Message 'Software-test profile metadata prefixes are missing 测试项目.'
  Assert-True -Condition ((Get-ReportProfileRequiredHeadings -Profile $softwareTestProfile) -contains '测试用例设计与执行') -Message 'Software-test profile required headings are missing 测试用例设计与执行.'
  Assert-True -Condition ([string](Get-ReportProfileDefaultImageCaptionBody -Profile $softwareTestProfile -SectionId 'result' -BaseName 'result-pass') -eq '测试结果截图') -Message 'Software-test profile image caption defaults are missing the result caption.'
  $softwareTestPaginationRemediations = Get-ReportProfilePaginationRiskRemediations -Profile $softwareTestProfile
  Assert-True -Condition ([string]$softwareTestPaginationRemediations.'pagination-risk-long-section' -match 'case group') -Message 'Software-test profile pagination-risk remediations are missing the long-section override.'
  Assert-True -Condition ([string]$softwareTestPaginationRemediations.'pagination-risk-dense-section-block' -match 'expected results, actual results') -Message 'Software-test profile pagination-risk remediations are missing the dense-section override.'
  Assert-True -Condition ([string]$softwareTestPaginationRemediations.'pagination-risk-figure-cluster' -match '测试结果') -Message 'Software-test profile pagination-risk remediations are missing the figure-cluster override.'
  $deploymentProfile = Get-ReportProfile -ProfileName 'deployment-report' -RepoRoot $repoRoot
  Assert-True -Condition ([string]$deploymentProfile.name -eq 'deployment-report') -Message 'Deployment profile loader returned an unexpected profile name.'
  Assert-True -Condition ([string]$deploymentProfile.displayName -eq '部署运维报告') -Message 'Deployment profile loader returned an unexpected display name.'
  Assert-True -Condition ([string](Get-ReportProfileDefaultStyleProfile -Profile $deploymentProfile) -eq 'school') -Message 'Deployment profile is missing the expected defaultStyleProfile.'
  Assert-True -Condition ((Get-ReportProfileMetadataPrefixes -Profile $deploymentProfile) -contains '部署项目') -Message 'Deployment profile metadata prefixes are missing 部署项目.'
  Assert-True -Condition ((Get-ReportProfileRequiredHeadings -Profile $deploymentProfile) -contains '部署步骤与配置') -Message 'Deployment profile required headings are missing 部署步骤与配置.'
  Assert-True -Condition ([string](Get-ReportProfileDefaultImageCaptionBody -Profile $deploymentProfile -SectionId 'result' -BaseName 'health-result') -eq '验证结果截图') -Message 'Deployment profile image caption defaults are missing the result caption.'
  $deploymentPaginationRemediations = Get-ReportProfilePaginationRiskRemediations -Profile $deploymentProfile
  Assert-True -Condition ([string]$deploymentPaginationRemediations.'pagination-risk-long-section' -match 'rollback subsections') -Message 'Deployment profile pagination-risk remediations are missing the long-section override.'
  Assert-True -Condition ([string]$deploymentPaginationRemediations.'pagination-risk-dense-section-block' -match 'commands, configuration snippets') -Message 'Deployment profile pagination-risk remediations are missing the dense-section override.'
  Assert-True -Condition ([string]$deploymentPaginationRemediations.'pagination-risk-figure-cluster' -match '验证结果') -Message 'Deployment profile pagination-risk remediations are missing the figure-cluster override.'
  $weeklyReportProfile = Get-ReportProfile -ProfileName 'weekly-report' -RepoRoot $repoRoot
  Assert-True -Condition ([string]$weeklyReportProfile.name -eq 'weekly-report') -Message 'Weekly-report profile loader returned an unexpected profile name.'
  Assert-True -Condition ([string]$weeklyReportProfile.displayName -eq '周报') -Message 'Weekly-report profile loader returned an unexpected display name.'
  Assert-True -Condition ([string](Get-ReportProfileDefaultStyleProfile -Profile $weeklyReportProfile) -eq 'compact') -Message 'Weekly-report profile is missing the expected defaultStyleProfile.'
  Assert-True -Condition ((Get-ReportProfileMetadataPrefixes -Profile $weeklyReportProfile) -contains '周报主题') -Message 'Weekly-report profile metadata prefixes are missing 周报主题.'
  Assert-True -Condition ((Get-ReportProfileRequiredHeadings -Profile $weeklyReportProfile) -contains '本周完成事项') -Message 'Weekly-report profile required headings are missing 本周完成事项.'
  Assert-True -Condition ([string](Get-ReportProfileDefaultImageCaptionBody -Profile $weeklyReportProfile -SectionId 'result' -BaseName 'demo-result') -eq '阶段成果截图') -Message 'Weekly-report profile image caption defaults are missing the result caption.'
  $weeklyPaginationThresholds = Get-ReportProfilePaginationRiskThresholds -Profile $weeklyReportProfile
  Assert-True -Condition ([int]$weeklyPaginationThresholds.longSectionChars -eq 1200) -Message 'Weekly-report profile pagination-risk thresholds are missing.'
  $meetingMinutesProfile = Get-ReportProfile -ProfileName 'meeting-minutes' -RepoRoot $repoRoot
  Assert-True -Condition ([string]$meetingMinutesProfile.name -eq 'meeting-minutes') -Message 'Meeting-minutes profile loader returned an unexpected profile name.'
  Assert-True -Condition ([string]$meetingMinutesProfile.displayName -eq '会议纪要') -Message 'Meeting-minutes profile loader returned an unexpected display name.'
  Assert-True -Condition ([string](Get-ReportProfileDefaultStyleProfile -Profile $meetingMinutesProfile) -eq 'default') -Message 'Meeting-minutes profile is missing the expected defaultStyleProfile.'
  Assert-True -Condition ((Get-ReportProfileMetadataPrefixes -Profile $meetingMinutesProfile) -contains '会议主题') -Message 'Meeting-minutes profile metadata prefixes are missing 会议主题.'
  Assert-True -Condition ((Get-ReportProfileRequiredHeadings -Profile $meetingMinutesProfile) -contains '讨论过程与决议') -Message 'Meeting-minutes profile required headings are missing 讨论过程与决议.'
  Assert-True -Condition ([string](Get-ReportProfileDefaultImageCaptionBody -Profile $meetingMinutesProfile -SectionId 'result' -BaseName 'decision-confirmed') -eq '结论确认截图') -Message 'Meeting-minutes profile image caption defaults are missing the result caption.'
  $meetingMinutesPaginationThresholds = Get-ReportProfilePaginationRiskThresholds -Profile $meetingMinutesProfile
  Assert-True -Condition ([int]$meetingMinutesPaginationThresholds.longSectionChars -eq 1000) -Message 'Meeting-minutes profile pagination-risk thresholds are missing.'
  $monthlyReportProfile = Get-ReportProfile -ProfileName 'monthly-report' -RepoRoot $repoRoot
  Assert-True -Condition ([string]$monthlyReportProfile.name -eq 'monthly-report') -Message 'Monthly-report profile loader returned an unexpected profile name.'
  Assert-True -Condition ([string]$monthlyReportProfile.displayName -eq '月报') -Message 'Monthly-report profile loader returned an unexpected display name.'
  Assert-True -Condition ([string](Get-ReportProfileDefaultStyleProfile -Profile $monthlyReportProfile) -eq 'compact') -Message 'Monthly-report profile is missing the expected defaultStyleProfile.'
  Assert-True -Condition ((Get-ReportProfileMetadataPrefixes -Profile $monthlyReportProfile) -contains '月报主题') -Message 'Monthly-report profile metadata prefixes are missing 月报主题.'
  Assert-True -Condition ((Get-ReportProfileRequiredHeadings -Profile $monthlyReportProfile) -contains '本月完成事项') -Message 'Monthly-report profile required headings are missing 本月完成事项.'
  Assert-True -Condition ([string](Get-ReportProfileDefaultImageCaptionBody -Profile $monthlyReportProfile -SectionId 'result' -BaseName 'monthly-metrics') -eq '阶段成果与数据截图') -Message 'Monthly-report profile image caption defaults are missing the result caption.'
  $monthlyPaginationThresholds = Get-ReportProfilePaginationRiskThresholds -Profile $monthlyReportProfile
  Assert-True -Condition ([int]$monthlyPaginationThresholds.longSectionChars -eq 1600) -Message 'Monthly-report profile pagination-risk thresholds are missing.'
  $monthlyPaginationRemediations = Get-ReportProfilePaginationRiskRemediations -Profile $monthlyReportProfile
  Assert-True -Condition ([string]$monthlyPaginationRemediations.'pagination-risk-long-section' -match 'week-by-week') -Message 'Monthly-report profile pagination-risk remediations are missing the long-section override.'
  Assert-True -Condition ([string]$monthlyPaginationRemediations.'pagination-risk-dense-section-block' -match 'objective, action, evidence, and impact') -Message 'Monthly-report profile pagination-risk remediations are missing the dense-section override.'
  Assert-True -Condition ([string]$monthlyPaginationRemediations.'pagination-risk-figure-cluster' -match '阶段成果与数据') -Message 'Monthly-report profile pagination-risk remediations are missing the figure-cluster override.'
  $experimentPromptText = New-ReportProfileAutoPromptText -ResolvedCourseName '计算机网络' -ResolvedExperimentName '交换机 VLAN 配置实验' -Profile $reportProfile -DetailLevel 'standard'
  Assert-True -Condition ($experimentPromptText -match '实验报告 body') -Message 'Auto prompt helper did not use the experiment-report display name.'
  Assert-True -Condition ($experimentPromptText -match '课程名称: 计算机网络') -Message 'Auto prompt helper did not emit the experiment-report course-name label.'
  Assert-True -Condition ($experimentPromptText -match '实验名称: 交换机 VLAN 配置实验') -Message 'Auto prompt helper did not emit the experiment-report title label.'
  $experimentRequirements = (New-ReportProfileAutoRequirementsJson -ResolvedCourseName '计算机网络' -ResolvedExperimentName '交换机 VLAN 配置实验' -Profile $reportProfile -ExtraKeywords @('VLAN', '交换机 VLAN 配置实验') -DetailLevel 'standard') | ConvertFrom-Json
  Assert-True -Condition ([string]$experimentRequirements.courseName -eq '计算机网络') -Message 'Auto requirements helper did not preserve the experiment-report course name.'
  Assert-True -Condition ([int]$experimentRequirements.minChars -eq 700) -Message 'Auto requirements helper did not use the experiment-report standard minChars.'
  Assert-True -Condition (@($experimentRequirements.requiredKeywords).Count -eq 3) -Message 'Auto requirements helper should keep unique course/title/extra keywords.'
  Assert-True -Condition ([int]$experimentRequirements.paginationRiskThresholds.longSectionChars -eq 900) -Message 'Auto requirements helper did not include profile pagination-risk thresholds.'
  $experimentMetadata = (New-ReportProfileAutoMetadataJson -ResolvedCourseName '计算机网络' -ResolvedExperimentName '交换机 VLAN 配置实验' -Profile $reportProfile -ResolvedStudentName '张三' -ResolvedStudentId '20260001' -ResolvedClassName '计科 2201' -ResolvedTeacherName '李老师' -ResolvedExperimentProperty '③验证性实验' -ResolvedExperimentDate '2026-04-09' -ResolvedExperimentLocation '实验楼 A201') | ConvertFrom-Json
  Assert-True -Condition ([string]$experimentMetadata.姓名 -eq '张三') -Message 'Auto metadata helper did not emit the experiment-report student label.'
  Assert-True -Condition ([string]$experimentMetadata.日期 -eq '2026-04-09') -Message 'Auto metadata helper did not emit the experiment-report extra date label.'
  $courseDesignPromptText = New-ReportProfileAutoPromptText -ResolvedCourseName '软件工程综合实践' -ResolvedExperimentName '校园导览小程序设计' -Profile $courseDesignProfile -DetailLevel 'full'
  Assert-True -Condition ($courseDesignPromptText -match '课程设计报告 body') -Message 'Auto prompt helper did not use the course-design display name.'
  Assert-True -Condition ($courseDesignPromptText -match '课程名称: 软件工程综合实践') -Message 'Auto prompt helper did not emit the course-design course-name label.'
  Assert-True -Condition ($courseDesignPromptText -match '课题名称: 校园导览小程序设计') -Message 'Auto prompt helper did not emit the course-design title label.'
  Assert-True -Condition ($courseDesignPromptText -match '方案设计与实现') -Message 'Auto prompt helper did not include course-design required headings.'
  $courseDesignRequirements = (New-ReportProfileAutoRequirementsJson -ResolvedCourseName '软件工程综合实践' -ResolvedExperimentName '校园导览小程序设计' -Profile $courseDesignProfile -ExtraKeywords @('小程序', '校园导览小程序设计') -DetailLevel 'full') | ConvertFrom-Json
  Assert-True -Condition ([int]$courseDesignRequirements.minChars -eq 1400) -Message 'Auto requirements helper did not use the course-design full minChars.'
  Assert-True -Condition (@($courseDesignRequirements.sections | Where-Object { $_.name -eq '方案设计与实现' }).Count -eq 1) -Message 'Auto requirements helper did not preserve the course-design implementation section heading.'
  $courseDesignMetadata = (New-ReportProfileAutoMetadataJson -ResolvedCourseName '软件工程综合实践' -ResolvedExperimentName '校园导览小程序设计' -Profile $courseDesignProfile -ResolvedStudentName '李四' -ResolvedStudentId '20261234' -ResolvedClassName '软工 2302' -ResolvedTeacherName '王老师' -ResolvedExperimentProperty '课程设计' -ResolvedExperimentDate '2026-04-08' -ResolvedExperimentLocation '实验楼 A201') | ConvertFrom-Json
  Assert-True -Condition ([string]$courseDesignMetadata.学生姓名 -eq '李四') -Message 'Auto metadata helper did not emit the course-design student label.'
  Assert-True -Condition ([string]$courseDesignMetadata.课题名称 -eq '校园导览小程序设计') -Message 'Auto metadata helper did not emit the course-design title label.'
  $internshipPromptText = New-ReportProfileAutoPromptText -ResolvedCourseName '软件工程' -ResolvedExperimentName '企业门户管理后台开发' -Profile $internshipProfile -DetailLevel 'full'
  Assert-True -Condition ($internshipPromptText -match '专业实习报告 body') -Message 'Auto prompt helper did not use the internship display name.'
  Assert-True -Condition ($internshipPromptText -match '专业名称: 软件工程') -Message 'Auto prompt helper did not emit the internship course-name label.'
  Assert-True -Condition ($internshipPromptText -match '实习项目: 企业门户管理后台开发') -Message 'Auto prompt helper did not emit the internship title label.'
  Assert-True -Condition ($internshipPromptText -match '实习过程与内容') -Message 'Auto prompt helper did not include internship required headings.'
  $internshipRequirements = (New-ReportProfileAutoRequirementsJson -ResolvedCourseName '软件工程' -ResolvedExperimentName '企业门户管理后台开发' -Profile $internshipProfile -ExtraKeywords @('后台开发', '企业门户管理后台开发') -DetailLevel 'full') | ConvertFrom-Json
  Assert-True -Condition ([int]$internshipRequirements.minChars -eq 1600) -Message 'Auto requirements helper did not use the internship full minChars.'
  Assert-True -Condition (@($internshipRequirements.sections | Where-Object { $_.name -eq '实习过程与内容' }).Count -eq 1) -Message 'Auto requirements helper did not preserve the internship process section heading.'
  $internshipMetadata = (New-ReportProfileAutoMetadataJson -ResolvedCourseName '软件工程' -ResolvedExperimentName '企业门户管理后台开发' -Profile $internshipProfile -ResolvedStudentName '王敏' -ResolvedStudentId '20262345' -ResolvedClassName '软工 2303' -ResolvedTeacherName '周老师' -ResolvedExperimentProperty '专业实习' -ResolvedExperimentDate '2026-03-01 至 2026-03-28' -ResolvedExperimentLocation '杭州云帆科技有限公司（滨江区）') | ConvertFrom-Json
  Assert-True -Condition ([string]$internshipMetadata.学生姓名 -eq '王敏') -Message 'Auto metadata helper did not emit the internship student label.'
  Assert-True -Condition ([string]$internshipMetadata.实习项目 -eq '企业门户管理后台开发') -Message 'Auto metadata helper did not emit the internship title label.'
  $softwareTestPromptText = New-ReportProfileAutoPromptText -ResolvedCourseName '软件测试技术' -ResolvedExperimentName '图书管理系统功能测试' -Profile $softwareTestProfile -DetailLevel 'full'
  Assert-True -Condition ($softwareTestPromptText -match '软件测试报告 body') -Message 'Auto prompt helper did not use the software-test display name.'
  Assert-True -Condition ($softwareTestPromptText -match '课程名称: 软件测试技术') -Message 'Auto prompt helper did not emit the software-test course-name label.'
  Assert-True -Condition ($softwareTestPromptText -match '测试项目: 图书管理系统功能测试') -Message 'Auto prompt helper did not emit the software-test project label.'
  Assert-True -Condition ($softwareTestPromptText -match '测试用例设计与执行') -Message 'Auto prompt helper did not include software-test required headings.'
  $softwareTestRequirements = (New-ReportProfileAutoRequirementsJson -ResolvedCourseName '软件测试技术' -ResolvedExperimentName '图书管理系统功能测试' -Profile $softwareTestProfile -ExtraKeywords @('登录测试', '图书管理系统功能测试') -DetailLevel 'full') | ConvertFrom-Json
  Assert-True -Condition ([int]$softwareTestRequirements.minChars -eq 1600) -Message 'Auto requirements helper did not use the software-test full minChars.'
  Assert-True -Condition (@($softwareTestRequirements.sections | Where-Object { $_.name -eq '测试用例设计与执行' }).Count -eq 1) -Message 'Auto requirements helper did not preserve the software-test case section heading.'
  Assert-True -Condition ([string]$softwareTestRequirements.paginationRiskRemediations.'pagination-risk-long-section' -match 'case group') -Message 'Auto requirements helper did not carry the software-test long-section remediation.'
  Assert-True -Condition ([string]$softwareTestRequirements.paginationRiskRemediations.'pagination-risk-dense-section-block' -match 'expected results, actual results') -Message 'Auto requirements helper did not carry the software-test dense-section remediation.'
  Assert-True -Condition ([string]$softwareTestRequirements.paginationRiskRemediations.'pagination-risk-figure-cluster' -match '缺陷分析与改进') -Message 'Auto requirements helper did not carry the software-test figure-cluster remediation.'
  $softwareTestMetadata = (New-ReportProfileAutoMetadataJson -ResolvedCourseName '软件测试技术' -ResolvedExperimentName '图书管理系统功能测试' -Profile $softwareTestProfile -ResolvedStudentName '赵强' -ResolvedStudentId '20263456' -ResolvedClassName '软工 2304' -ResolvedTeacherName '陈老师' -ResolvedExperimentProperty '功能测试' -ResolvedExperimentDate '2026-04-10' -ResolvedExperimentLocation 'Chrome 122 / Windows 11 / MySQL 8.0') | ConvertFrom-Json
  Assert-True -Condition ([string]$softwareTestMetadata.学生姓名 -eq '赵强') -Message 'Auto metadata helper did not emit the software-test student label.'
  Assert-True -Condition ([string]$softwareTestMetadata.测试项目 -eq '图书管理系统功能测试') -Message 'Auto metadata helper did not emit the software-test project label.'
  $deploymentPromptText = New-ReportProfileAutoPromptText -ResolvedCourseName '云平台运维实践' -ResolvedExperimentName '校园门户系统容器化部署' -Profile $deploymentProfile -DetailLevel 'full'
  Assert-True -Condition ($deploymentPromptText -match '部署运维报告 body') -Message 'Auto prompt helper did not use the deployment display name.'
  Assert-True -Condition ($deploymentPromptText -match '课程名称: 云平台运维实践') -Message 'Auto prompt helper did not emit the deployment course-name label.'
  Assert-True -Condition ($deploymentPromptText -match '部署项目: 校园门户系统容器化部署') -Message 'Auto prompt helper did not emit the deployment project label.'
  Assert-True -Condition ($deploymentPromptText -match '部署步骤与配置') -Message 'Auto prompt helper did not include deployment required headings.'
  $deploymentRequirements = (New-ReportProfileAutoRequirementsJson -ResolvedCourseName '云平台运维实践' -ResolvedExperimentName '校园门户系统容器化部署' -Profile $deploymentProfile -ExtraKeywords @('Docker', 'Nginx', '校园门户系统容器化部署') -DetailLevel 'full') | ConvertFrom-Json
  Assert-True -Condition ([int]$deploymentRequirements.minChars -eq 1600) -Message 'Auto requirements helper did not use the deployment full minChars.'
  Assert-True -Condition (@($deploymentRequirements.sections | Where-Object { $_.name -eq '部署步骤与配置' }).Count -eq 1) -Message 'Auto requirements helper did not preserve the deployment steps section heading.'
  Assert-True -Condition ([string]$deploymentRequirements.paginationRiskRemediations.'pagination-risk-long-section' -match 'rollback subsections') -Message 'Auto requirements helper did not carry the deployment long-section remediation.'
  Assert-True -Condition ([string]$deploymentRequirements.paginationRiskRemediations.'pagination-risk-dense-section-block' -match 'commands, configuration snippets') -Message 'Auto requirements helper did not carry the deployment dense-section remediation.'
  Assert-True -Condition ([string]$deploymentRequirements.paginationRiskRemediations.'pagination-risk-figure-cluster' -match '问题处理与回滚预案') -Message 'Auto requirements helper did not carry the deployment figure-cluster remediation.'
  $deploymentMetadata = (New-ReportProfileAutoMetadataJson -ResolvedCourseName '云平台运维实践' -ResolvedExperimentName '校园门户系统容器化部署' -Profile $deploymentProfile -ResolvedStudentName '刘洋' -ResolvedStudentId '20264567' -ResolvedClassName '网工 2301' -ResolvedTeacherName '孙老师' -ResolvedExperimentProperty '系统部署' -ResolvedExperimentDate '2026-04-12' -ResolvedExperimentLocation 'Ubuntu 22.04 / Docker 26 / Nginx 1.24') | ConvertFrom-Json
  Assert-True -Condition ([string]$deploymentMetadata.学生姓名 -eq '刘洋') -Message 'Auto metadata helper did not emit the deployment student label.'
  Assert-True -Condition ([string]$deploymentMetadata.部署项目 -eq '校园门户系统容器化部署') -Message 'Auto metadata helper did not emit the deployment project label.'
  $weeklyPromptText = New-ReportProfileAutoPromptText -ResolvedCourseName '校园导览小程序' -ResolvedExperimentName '第 6 周迭代周报' -Profile $weeklyReportProfile -DetailLevel 'full'
  Assert-True -Condition ($weeklyPromptText -match '周报 body') -Message 'Auto prompt helper did not use the weekly-report display name.'
  Assert-True -Condition ($weeklyPromptText -match '项目名称: 校园导览小程序') -Message 'Auto prompt helper did not emit the weekly-report project label.'
  Assert-True -Condition ($weeklyPromptText -match '周报主题: 第 6 周迭代周报') -Message 'Auto prompt helper did not emit the weekly-report title label.'
  Assert-True -Condition ($weeklyPromptText -match '本周完成事项') -Message 'Auto prompt helper did not include weekly-report required headings.'
  $weeklyRequirements = (New-ReportProfileAutoRequirementsJson -ResolvedCourseName '校园导览小程序' -ResolvedExperimentName '第 6 周迭代周报' -Profile $weeklyReportProfile -ExtraKeywords @('迭代', '校园导览小程序') -DetailLevel 'full') | ConvertFrom-Json
  Assert-True -Condition ([int]$weeklyRequirements.minChars -eq 1400) -Message 'Auto requirements helper did not use the weekly-report full minChars.'
  Assert-True -Condition (@($weeklyRequirements.sections | Where-Object { $_.name -eq '本周完成事项' }).Count -eq 1) -Message 'Auto requirements helper did not preserve the weekly-report completion section heading.'
  $weeklyMetadata = (New-ReportProfileAutoMetadataJson -ResolvedCourseName '校园导览小程序' -ResolvedExperimentName '第 6 周迭代周报' -Profile $weeklyReportProfile -ResolvedStudentName '李四' -ResolvedStudentId '20261234' -ResolvedClassName '软工 2302' -ResolvedTeacherName '王老师' -ResolvedExperimentProperty '项目周报' -ResolvedExperimentDate '第 6 周' -ResolvedExperimentLocation 'GitHub + 飞书 + 本地开发环境') | ConvertFrom-Json
  Assert-True -Condition ([string]$weeklyMetadata.提交人 -eq '李四') -Message 'Auto metadata helper did not emit the weekly-report owner label.'
  Assert-True -Condition ([string]$weeklyMetadata.周报主题 -eq '第 6 周迭代周报') -Message 'Auto metadata helper did not emit the weekly-report title label.'
  $meetingPromptText = New-ReportProfileAutoPromptText -ResolvedCourseName '校园导览小程序' -ResolvedExperimentName '第 6 周迭代评审会' -Profile $meetingMinutesProfile -DetailLevel 'full'
  Assert-True -Condition ($meetingPromptText -match '会议纪要 body') -Message 'Auto prompt helper did not use the meeting-minutes display name.'
  Assert-True -Condition ($meetingPromptText -match '项目名称: 校园导览小程序') -Message 'Auto prompt helper did not emit the meeting-minutes project label.'
  Assert-True -Condition ($meetingPromptText -match '会议主题: 第 6 周迭代评审会') -Message 'Auto prompt helper did not emit the meeting-minutes title label.'
  Assert-True -Condition ($meetingPromptText -match '讨论过程与决议') -Message 'Auto prompt helper did not include meeting-minutes required headings.'
  $meetingRequirements = (New-ReportProfileAutoRequirementsJson -ResolvedCourseName '校园导览小程序' -ResolvedExperimentName '第 6 周迭代评审会' -Profile $meetingMinutesProfile -ExtraKeywords @('迭代评审', '校园导览小程序') -DetailLevel 'full') | ConvertFrom-Json
  Assert-True -Condition ([int]$meetingRequirements.minChars -eq 1300) -Message 'Auto requirements helper did not use the meeting-minutes full minChars.'
  Assert-True -Condition (@($meetingRequirements.sections | Where-Object { $_.name -eq '讨论过程与决议' }).Count -eq 1) -Message 'Auto requirements helper did not preserve the meeting-minutes discussion section heading.'
  $meetingMetadata = (New-ReportProfileAutoMetadataJson -ResolvedCourseName '校园导览小程序' -ResolvedExperimentName '第 6 周迭代评审会' -Profile $meetingMinutesProfile -ResolvedStudentName '李四' -ResolvedStudentId '20261234' -ResolvedClassName '软工 2302' -ResolvedTeacherName '王老师' -ResolvedExperimentProperty '评审会议纪要' -ResolvedExperimentDate '2026-04-13' -ResolvedExperimentLocation 'GitHub + 飞书会议') | ConvertFrom-Json
  Assert-True -Condition ([string]$meetingMetadata.记录人 -eq '李四') -Message 'Auto metadata helper did not emit the meeting-minutes recorder label.'
  Assert-True -Condition ([string]$meetingMetadata.会议主题 -eq '第 6 周迭代评审会') -Message 'Auto metadata helper did not emit the meeting-minutes title label.'
  $monthlyPromptText = New-ReportProfileAutoPromptText -ResolvedCourseName '校园导览小程序' -ResolvedExperimentName '2026 年 4 月项目月报' -Profile $monthlyReportProfile -DetailLevel 'full'
  Assert-True -Condition ($monthlyPromptText -match '月报 body') -Message 'Auto prompt helper did not use the monthly-report display name.'
  Assert-True -Condition ($monthlyPromptText -match '项目名称: 校园导览小程序') -Message 'Auto prompt helper did not emit the monthly-report project label.'
  Assert-True -Condition ($monthlyPromptText -match '月报主题: 2026 年 4 月项目月报') -Message 'Auto prompt helper did not emit the monthly-report title label.'
  Assert-True -Condition ($monthlyPromptText -match '本月完成事项') -Message 'Auto prompt helper did not include monthly-report required headings.'
  $monthlyRequirements = (New-ReportProfileAutoRequirementsJson -ResolvedCourseName '校园导览小程序' -ResolvedExperimentName '2026 年 4 月项目月报' -Profile $monthlyReportProfile -ExtraKeywords @('月度进展', '校园导览小程序') -DetailLevel 'full') | ConvertFrom-Json
  Assert-True -Condition ([int]$monthlyRequirements.minChars -eq 1800) -Message 'Auto requirements helper did not use the monthly-report full minChars.'
  Assert-True -Condition (@($monthlyRequirements.sections | Where-Object { $_.name -eq '本月完成事项' }).Count -eq 1) -Message 'Auto requirements helper did not preserve the monthly-report completion section heading.'
  Assert-True -Condition ([string]$monthlyRequirements.paginationRiskRemediations.'pagination-risk-long-section' -match 'week-by-week') -Message 'Auto requirements helper did not carry the monthly-report long-section remediation.'
  Assert-True -Condition ([string]$monthlyRequirements.paginationRiskRemediations.'pagination-risk-figure-cluster' -match '阶段成果与数据') -Message 'Auto requirements helper did not carry the monthly-report figure-cluster remediation.'
  $monthlyMetadata = (New-ReportProfileAutoMetadataJson -ResolvedCourseName '校园导览小程序' -ResolvedExperimentName '2026 年 4 月项目月报' -Profile $monthlyReportProfile -ResolvedStudentName '李四' -ResolvedStudentId '20261234' -ResolvedClassName '软工 2302' -ResolvedTeacherName '王老师' -ResolvedExperimentProperty '项目月报' -ResolvedExperimentDate '2026-04' -ResolvedExperimentLocation 'GitHub + 飞书 + 本地开发环境') | ConvertFrom-Json
  Assert-True -Condition ([string]$monthlyMetadata.提交人 -eq '李四') -Message 'Auto metadata helper did not emit the monthly-report owner label.'
  Assert-True -Condition ([string]$monthlyMetadata.月报主题 -eq '2026 年 4 月项目月报') -Message 'Auto metadata helper did not emit the monthly-report title label.'
  $results.Add('report profile loader OK') | Out-Null

  $reportProfileSchema = (Get-Content -LiteralPath (Join-Path $repoRoot 'profiles\report-profile.schema.json') -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$reportProfileSchema.title -eq 'OpenClaw report profile') -Message 'Report profile schema did not parse.'
  $profileValidation = (& (Join-Path $repoRoot 'scripts\validate-report-profiles.ps1') -Format json | Out-String) | ConvertFrom-Json
  Assert-True -Condition ([bool]$profileValidation.passed) -Message 'Report profile validation failed.'
  Assert-True -Condition ([int]$profileValidation.summary.profileCount -ge 8) -Message 'Report profile validation did not cover built-in profiles.'
  Assert-True -Condition ([int]$profileValidation.summary.errorCount -eq 0) -Message 'Report profile validation reported unexpected errors.'
  $results.Add('report profile schema validation OK') | Out-Null

  $exampleProfilePresetValidation = (& (Join-Path $repoRoot 'scripts\validate-report-profiles.ps1') -ProfileDir (Join-Path $repoRoot 'examples\profile-presets') -Format json | Out-String) | ConvertFrom-Json
  Assert-True -Condition ([bool]$exampleProfilePresetValidation.passed) -Message 'Example profile presets failed validation.'
  Assert-True -Condition ([int]$exampleProfilePresetValidation.summary.profileCount -eq 3) -Message 'Example profile preset validation should cover the curated preset trio.'
  Assert-True -Condition ([int]$exampleProfilePresetValidation.summary.errorCount -eq 0) -Message 'Example profile preset validation reported unexpected errors.'
  $exampleWeeklyPreset = (Get-Content -LiteralPath (Join-Path $repoRoot 'examples\profile-presets\weekly-report.json') -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$exampleWeeklyPreset.defaultStyleProfile -eq 'compact') -Message 'Weekly preset should demonstrate compact as the default style profile.'
  Assert-True -Condition ([string]$exampleWeeklyPreset.sectionFields[3].heading -eq '本周完成事项') -Message 'Weekly preset is missing the expected steps heading.'
  Assert-True -Condition ([int]$exampleWeeklyPreset.paginationRiskThresholds.longSectionChars -eq 1200) -Message 'Weekly preset should demonstrate custom pagination-risk thresholds.'
  $exampleMeetingPreset = (Get-Content -LiteralPath (Join-Path $repoRoot 'examples\profile-presets\meeting-minutes.json') -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$exampleMeetingPreset.defaultStyleProfile -eq 'default') -Message 'Meeting-minutes preset should demonstrate default as the default style profile.'
  Assert-True -Condition ([string]$exampleMeetingPreset.sectionFields[3].heading -eq '讨论过程与决议') -Message 'Meeting-minutes preset is missing the expected steps heading.'
  Assert-True -Condition ([int]$exampleMeetingPreset.paginationRiskThresholds.longSectionChars -eq 1000) -Message 'Meeting-minutes preset should demonstrate custom pagination-risk thresholds.'
  $exampleMonthlyPreset = (Get-Content -LiteralPath (Join-Path $repoRoot 'examples\profile-presets\monthly-report.json') -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$exampleMonthlyPreset.defaultStyleProfile -eq 'compact') -Message 'Monthly preset should demonstrate compact as the default style profile.'
  Assert-True -Condition ([string]$exampleMonthlyPreset.sectionFields[3].heading -eq '本月完成事项') -Message 'Monthly preset is missing the expected steps heading.'
  Assert-True -Condition ([int]$exampleMonthlyPreset.paginationRiskThresholds.longSectionChars -eq 1600) -Message 'Monthly preset should demonstrate custom pagination-risk thresholds.'
  Assert-True -Condition ([string]$exampleMonthlyPreset.paginationRiskRemediations.'pagination-risk-long-section' -match 'week-by-week') -Message 'Monthly preset should demonstrate custom pagination-risk remediations.'
  $results.Add('example profile presets OK') | Out-Null

  $profilePresetSamplesDir = Join-Path $tempRoot 'profile-preset-samples'
  $profilePresetSamples = (& (Join-Path $repoRoot 'scripts\run-profile-preset-samples.ps1') -OutputDir $profilePresetSamplesDir -Format json | Out-String) | ConvertFrom-Json
  Assert-True -Condition ([int]$profilePresetSamples.generatedCount -eq 3) -Message 'Profile preset sample runner should generate all curated preset samples.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$profilePresetSamples.summaryPath)) -Message 'Profile preset sample runner did not write its summary JSON.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$profilePresetSamples.markdownPath)) -Message 'Profile preset sample runner did not write its markdown index.'
  $profilePresetSamplesMarkdown = Get-Content -LiteralPath ([string]$profilePresetSamples.markdownPath) -Raw -Encoding UTF8
  Assert-True -Condition ($profilePresetSamplesMarkdown -match 'Profile Preset Samples') -Message 'Profile preset sample markdown is missing the expected title.'
  Assert-True -Condition ($profilePresetSamplesMarkdown -match 'weekly-report') -Message 'Profile preset sample markdown is missing weekly-report.'
  Assert-True -Condition ($profilePresetSamplesMarkdown -match 'meeting-minutes') -Message 'Profile preset sample markdown is missing meeting-minutes.'
  Assert-True -Condition ($profilePresetSamplesMarkdown -match 'monthly-report') -Message 'Profile preset sample markdown is missing monthly-report.'
  $weeklyPresetSample = @($profilePresetSamples.generated | Where-Object { [string]$_.reportProfileName -eq 'weekly-report' })[0]
  Assert-True -Condition (Test-Path -LiteralPath ([string]$weeklyPresetSample.promptPath)) -Message 'Weekly preset sample runner did not create prompt.txt.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$weeklyPresetSample.metadataPath)) -Message 'Weekly preset sample runner did not create metadata.auto.json.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$weeklyPresetSample.requirementsPath)) -Message 'Weekly preset sample runner did not create requirements.auto.json.'
  $weeklyPresetRequirements = (Get-Content -LiteralPath ([string]$weeklyPresetSample.requirementsPath) -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition (@($weeklyPresetRequirements.sections | Where-Object { [string]$_.name -eq '本周完成事项' }).Count -eq 1) -Message 'Weekly preset sample requirements are missing the expected completion section.'
  Assert-True -Condition ([int]$weeklyPresetRequirements.paginationRiskThresholds.longSectionChars -eq 1200) -Message 'Weekly preset sample requirements did not preserve custom pagination-risk thresholds.'
  $meetingPresetSample = @($profilePresetSamples.generated | Where-Object { [string]$_.reportProfileName -eq 'meeting-minutes' })[0]
  Assert-True -Condition (Test-Path -LiteralPath ([string]$meetingPresetSample.promptPath)) -Message 'Meeting-minutes preset sample runner did not create prompt.txt.'
  $meetingPresetPrompt = Get-Content -LiteralPath ([string]$meetingPresetSample.promptPath) -Raw -Encoding UTF8
  Assert-True -Condition ($meetingPresetPrompt -match '会议纪要 body') -Message 'Meeting-minutes preset sample prompt is missing the profile display name.'
  $monthlyPresetSample = @($profilePresetSamples.generated | Where-Object { [string]$_.reportProfileName -eq 'monthly-report' })[0]
  Assert-True -Condition (Test-Path -LiteralPath ([string]$monthlyPresetSample.promptPath)) -Message 'Monthly preset sample runner did not create prompt.txt.'
  $monthlyPresetRequirements = (Get-Content -LiteralPath ([string]$monthlyPresetSample.requirementsPath) -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition (@($monthlyPresetRequirements.sections | Where-Object { [string]$_.name -eq '本月完成事项' }).Count -eq 1) -Message 'Monthly preset sample requirements are missing the expected completion section.'
  Assert-True -Condition ([int]$monthlyPresetRequirements.paginationRiskThresholds.longSectionChars -eq 1600) -Message 'Monthly preset sample requirements did not preserve custom pagination-risk thresholds.'
  Assert-True -Condition ([string]$monthlyPresetRequirements.paginationRiskRemediations.'pagination-risk-long-section' -match 'week-by-week') -Message 'Monthly preset sample requirements did not preserve the custom long-section remediation.'
  Assert-True -Condition ([string]$monthlyPresetRequirements.paginationRiskRemediations.'pagination-risk-figure-cluster' -match '阶段成果与数据') -Message 'Monthly preset sample requirements did not preserve the custom figure-cluster remediation.'
  $results.Add('profile preset sample runner OK') | Out-Null

  $roadmapTriageOutputDir = Join-Path $tempRoot 'roadmap-triage'
  $roadmapTriage = (& (Join-Path $repoRoot 'scripts\analyze-roadmap-next-step.ps1') -OutputDir $roadmapTriageOutputDir -Format json | Out-String) | ConvertFrom-Json
  Assert-True -Condition (Test-Path -LiteralPath ([string]$roadmapTriage.jsonPath)) -Message 'Roadmap triage did not write JSON output.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$roadmapTriage.markdownPath)) -Message 'Roadmap triage did not write markdown output.'
  Assert-True -Condition ([int]$roadmapTriage.summary.candidateCount -ge 1) -Message 'Roadmap triage should find at least one roadmap candidate.'
  Assert-True -Condition ([int]$roadmapTriage.summary.topCandidateCount -ge 1) -Message 'Roadmap triage should expose at least one top candidate.'
  $roadmapTriageMarkdown = Get-Content -LiteralPath ([string]$roadmapTriage.markdownPath) -Raw -Encoding UTF8
  Assert-True -Condition ($roadmapTriageMarkdown -match 'Roadmap Daily Triage') -Message 'Roadmap triage markdown is missing the expected title.'
  Assert-True -Condition ($roadmapTriageMarkdown -match 'Smoke-Coverable') -Message 'Roadmap triage markdown is missing the smoke-coverable section.'
  Assert-True -Condition (@($roadmapTriage.candidates | Where-Object { [string]$_.roadmapItem -eq 'weekly-report is listed in the recommended expansion order but is not a built-in profile.' }).Count -eq 0) -Message 'Roadmap triage should not keep recommending weekly-report after promotion.'
  Assert-True -Condition (@($roadmapTriage.candidates | Where-Object { [string]$_.roadmapItem -eq 'meeting-minutes is listed in the recommended expansion order but is not a built-in profile.' }).Count -eq 0) -Message 'Roadmap triage should not keep recommending meeting-minutes after promotion.'
  Assert-True -Condition (@($roadmapTriage.candidates | Where-Object { [string]$_.roadmapItem -eq 'monthly-report is listed in the recommended expansion order but is not a built-in profile.' }).Count -eq 0) -Message 'Roadmap triage should not keep recommending monthly-report after promotion.'
  $roadmapTriageWorkflow = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\roadmap-triage.yml') -Raw -Encoding UTF8
  Assert-True -Condition ($roadmapTriageWorkflow -match 'schedule:') -Message 'Roadmap triage workflow is missing a schedule trigger.'
  Assert-True -Condition ($roadmapTriageWorkflow -match 'workflow_dispatch:') -Message 'Roadmap triage workflow is missing a manual trigger.'
  Assert-True -Condition ($roadmapTriageWorkflow -match 'analyze-roadmap-next-step\.ps1') -Message 'Roadmap triage workflow does not run the analyzer script.'
  $results.Add('roadmap daily triage automation OK') | Out-Null

  $newReportProfilePath = Join-Path $tempRoot 'weekly-report.json'
  $newReportProfileResult = (& (Join-Path $repoRoot 'scripts\new-report-profile.ps1') `
      -Name 'weekly-report' `
      -DisplayName '周报' `
      -DefaultExperimentProperty '周报' `
      -CourseNameLabel '项目名称' `
      -TitleNameLabel '周报主题' `
      -DateLabel '周次' `
      -LocationLabel '工作环境' `
      -SectionHeadings @('工作目标', '工作环境', '工作范围与依据', '完成事项', '工作结果', '问题与改进', '下周计划') `
      -OutPath $newReportProfilePath `
      -Format json | Out-String) | ConvertFrom-Json
  Assert-True -Condition ([bool]$newReportProfileResult.validationPassed) -Message 'Profile scaffold generator did not validate the generated profile.'
  Assert-True -Condition (Test-Path -LiteralPath $newReportProfilePath) -Message 'Profile scaffold generator did not create the profile JSON.'
  $generatedWeeklyProfile = (Get-Content -LiteralPath $newReportProfilePath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$generatedWeeklyProfile.name -eq 'weekly-report') -Message 'Profile scaffold generator wrote an unexpected profile name.'
  Assert-True -Condition ([string]$generatedWeeklyProfile.displayName -eq '周报') -Message 'Profile scaffold generator wrote an unexpected display name.'
  Assert-True -Condition ([string]$generatedWeeklyProfile.sectionFields[3].heading -eq '完成事项') -Message 'Profile scaffold generator did not preserve the steps heading.'
  Assert-True -Condition ([int]$generatedWeeklyProfile.paginationRiskThresholds.longSectionChars -eq 900) -Message 'Profile scaffold generator did not include pagination-risk thresholds.'
  $generatedWeeklyProfileValidation = (& (Join-Path $repoRoot 'scripts\validate-report-profiles.ps1') -ProfilePath $newReportProfilePath -Format json | Out-String) | ConvertFrom-Json
  Assert-True -Condition ([bool]$generatedWeeklyProfileValidation.passed) -Message 'Generated weekly profile failed standalone validation.'
  $results.Add('report profile scaffold generator OK') | Out-Null

  $originalInputsAgentsHome = $env:AGENTS_HOME
  try {
    $env:AGENTS_HOME = (Join-Path $tempRoot 'report-inputs-agents-home')
    $reportInputsOutputDir = Join-Path $tempRoot 'report-inputs-output'
    & (Join-Path $repoRoot 'scripts\generate-report-inputs.ps1') `
      -CourseName '软件工程综合实践' `
      -ExperimentName '校园导览小程序设计' `
      -StudentName '李四' `
      -StudentId '20261234' `
      -ClassName '软工 2302' `
      -TeacherName '王老师' `
      -ExperimentProperty '课程设计' `
      -ExperimentDate '2026-04-08' `
      -ExperimentLocation '实验楼 A201' `
      -ReportProfileName 'course-design-report' `
      -RequiredKeywords @('小程序', '校园导览') `
      -OutputDir $reportInputsOutputDir `
      -DetailLevel full | Out-Null
    $reportInputsSummaryPath = Join-Path $reportInputsOutputDir 'report-inputs-summary.json'
    Assert-True -Condition (Test-Path -LiteralPath $reportInputsSummaryPath) -Message 'Report-input generation did not create the summary JSON.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $reportInputsOutputDir 'prompt.txt')) -Message 'Report-input generation did not create prompt.txt.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $reportInputsOutputDir 'metadata.auto.json')) -Message 'Report-input generation did not create metadata.auto.json.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $reportInputsOutputDir 'requirements.auto.json')) -Message 'Report-input generation did not create requirements.auto.json.'
    $reportInputsSummary = (Get-Content -LiteralPath $reportInputsSummaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
    Assert-True -Condition ([string]$reportInputsSummary.reportProfileName -eq 'course-design-report') -Message 'Report-input generation summary is missing the expected report profile name.'
    Assert-True -Condition ([string]$reportInputsSummary.reportProfileDisplayName -eq '课程设计报告') -Message 'Report-input generation summary is missing the expected display name.'
    Assert-True -Condition ((Split-Path -Leaf ([string]$reportInputsSummary.defaultsPath)) -eq 'course-design-report.defaults.json') -Message 'Report-input generation should persist defaults under the profile-specific defaults file.'
    $preparedInputsContext = Resolve-PreparedInputsSummaryContext -PreparedInputsSummaryPath $reportInputsSummaryPath
    Assert-True -Condition ([string]$preparedInputsContext.reportProfileName -eq 'course-design-report') -Message 'Prepared-input summary context should inherit the report profile name from the summary.'
    Assert-True -Condition ([string]$preparedInputsContext.reportProfilePath -eq [string]$reportInputsSummary.reportProfilePath) -Message 'Prepared-input summary context should inherit the report profile path from the summary.'
    Assert-True -Condition ([string]$preparedInputsContext.detailLevel -eq 'full') -Message 'Prepared-input summary context should inherit the detail level from the summary.'
    $generatedPromptText = Get-Content -LiteralPath (Join-Path $reportInputsOutputDir 'prompt.txt') -Raw -Encoding UTF8
    Assert-True -Condition ($generatedPromptText -match '课程设计报告 body') -Message 'Report-input generation did not emit the expected course-design prompt body.'
    Assert-True -Condition ($generatedPromptText -match '课题名称: 校园导览小程序设计') -Message 'Report-input generation did not emit the expected course-design title label.'
    $generatedMetadata = (Get-Content -LiteralPath (Join-Path $reportInputsOutputDir 'metadata.auto.json') -Raw -Encoding UTF8) | ConvertFrom-Json
    Assert-True -Condition ([string]$generatedMetadata.学生姓名 -eq '李四') -Message 'Report-input generation metadata is missing the course-design student label.'
    Assert-True -Condition ([string]$generatedMetadata.课题名称 -eq '校园导览小程序设计') -Message 'Report-input generation metadata is missing the course-design title label.'
    $generatedRequirements = (Get-Content -LiteralPath (Join-Path $reportInputsOutputDir 'requirements.auto.json') -Raw -Encoding UTF8) | ConvertFrom-Json
    Assert-True -Condition ([int]$generatedRequirements.minChars -eq 1400) -Message 'Report-input generation requirements are missing the expected course-design minChars.'
    Assert-True -Condition (@($generatedRequirements.sections | Where-Object { $_.name -eq '方案设计与实现' }).Count -eq 1) -Message 'Report-input generation requirements are missing the course-design implementation section.'
    Assert-True -Condition ([int]$generatedRequirements.paginationRiskThresholds.longSectionChars -eq 1000) -Message 'Report-input generation requirements are missing profile pagination-risk thresholds.'

    $weeklyInputsOutputDir = Join-Path $tempRoot 'weekly-report-inputs-output'
    & (Join-Path $repoRoot 'scripts\generate-report-inputs.ps1') `
      -CourseName '校园导览小程序' `
      -ExperimentName '第 6 周迭代周报' `
      -StudentName '李四' `
      -StudentId '20261234' `
      -ClassName '软工 2302' `
      -TeacherName '王老师' `
      -ExperimentProperty '项目周报' `
      -ExperimentDate '第 6 周' `
      -ExperimentLocation 'GitHub + 飞书 + 本地开发环境' `
      -ReportProfileName 'weekly-report' `
      -RequiredKeywords @('迭代', '阶段成果') `
      -OutputDir $weeklyInputsOutputDir `
      -DetailLevel full | Out-Null
    $weeklyInputsSummaryPath = Join-Path $weeklyInputsOutputDir 'report-inputs-summary.json'
    Assert-True -Condition (Test-Path -LiteralPath $weeklyInputsSummaryPath) -Message 'Weekly report-input generation did not create the summary JSON.'
    $weeklyInputsSummary = (Get-Content -LiteralPath $weeklyInputsSummaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
    Assert-True -Condition ([string]$weeklyInputsSummary.reportProfileName -eq 'weekly-report') -Message 'Weekly report-input generation summary is missing the expected profile name.'
    Assert-True -Condition ([string]$weeklyInputsSummary.reportProfileDisplayName -eq '周报') -Message 'Weekly report-input generation summary is missing the expected display name.'
    Assert-True -Condition ((Split-Path -Leaf ([string]$weeklyInputsSummary.defaultsPath)) -eq 'weekly-report.defaults.json') -Message 'Weekly report-input generation should persist defaults under the profile-specific defaults file.'
    $weeklyGeneratedPrompt = Get-Content -LiteralPath (Join-Path $weeklyInputsOutputDir 'prompt.txt') -Raw -Encoding UTF8
    Assert-True -Condition ($weeklyGeneratedPrompt -match '周报 body') -Message 'Weekly report-input generation did not emit the expected prompt body.'
    Assert-True -Condition ($weeklyGeneratedPrompt -match '周报主题: 第 6 周迭代周报') -Message 'Weekly report-input generation did not emit the expected title label.'
    $weeklyGeneratedMetadata = (Get-Content -LiteralPath (Join-Path $weeklyInputsOutputDir 'metadata.auto.json') -Raw -Encoding UTF8) | ConvertFrom-Json
    Assert-True -Condition ([string]$weeklyGeneratedMetadata.提交人 -eq '李四') -Message 'Weekly report-input generation metadata is missing the owner label.'
    Assert-True -Condition ([string]$weeklyGeneratedMetadata.周报主题 -eq '第 6 周迭代周报') -Message 'Weekly report-input generation metadata is missing the title label.'
    $weeklyGeneratedRequirements = (Get-Content -LiteralPath (Join-Path $weeklyInputsOutputDir 'requirements.auto.json') -Raw -Encoding UTF8) | ConvertFrom-Json
    Assert-True -Condition ([int]$weeklyGeneratedRequirements.minChars -eq 1400) -Message 'Weekly report-input generation requirements are missing the expected minChars.'
    Assert-True -Condition (@($weeklyGeneratedRequirements.sections | Where-Object { $_.name -eq '本周完成事项' }).Count -eq 1) -Message 'Weekly report-input generation requirements are missing the completion section.'
    Assert-True -Condition ([int]$weeklyGeneratedRequirements.paginationRiskThresholds.longSectionChars -eq 1200) -Message 'Weekly report-input generation requirements are missing pagination-risk thresholds.'

    $meetingInputsOutputDir = Join-Path $tempRoot 'meeting-minutes-inputs-output'
    & (Join-Path $repoRoot 'scripts\generate-report-inputs.ps1') `
      -CourseName '校园导览小程序' `
      -ExperimentName '第 6 周迭代评审会' `
      -StudentName '李四' `
      -StudentId '20261234' `
      -ClassName '软工 2302' `
      -TeacherName '王老师' `
      -ExperimentProperty '评审会议纪要' `
      -ExperimentDate '2026-04-13' `
      -ExperimentLocation 'GitHub + 飞书会议' `
      -ReportProfileName 'meeting-minutes' `
      -RequiredKeywords @('迭代评审', '行动项') `
      -OutputDir $meetingInputsOutputDir `
      -DetailLevel full | Out-Null
    $meetingInputsSummaryPath = Join-Path $meetingInputsOutputDir 'report-inputs-summary.json'
    Assert-True -Condition (Test-Path -LiteralPath $meetingInputsSummaryPath) -Message 'Meeting-minutes report-input generation did not create the summary JSON.'
    $meetingInputsSummary = (Get-Content -LiteralPath $meetingInputsSummaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
    Assert-True -Condition ([string]$meetingInputsSummary.reportProfileName -eq 'meeting-minutes') -Message 'Meeting-minutes report-input generation summary is missing the expected profile name.'
    Assert-True -Condition ([string]$meetingInputsSummary.reportProfileDisplayName -eq '会议纪要') -Message 'Meeting-minutes report-input generation summary is missing the expected display name.'
    Assert-True -Condition ((Split-Path -Leaf ([string]$meetingInputsSummary.defaultsPath)) -eq 'meeting-minutes.defaults.json') -Message 'Meeting-minutes report-input generation should persist defaults under the profile-specific defaults file.'
    $meetingGeneratedPrompt = Get-Content -LiteralPath (Join-Path $meetingInputsOutputDir 'prompt.txt') -Raw -Encoding UTF8
    Assert-True -Condition ($meetingGeneratedPrompt -match '会议纪要 body') -Message 'Meeting-minutes report-input generation did not emit the expected prompt body.'
    Assert-True -Condition ($meetingGeneratedPrompt -match '会议主题: 第 6 周迭代评审会') -Message 'Meeting-minutes report-input generation did not emit the expected title label.'
    $meetingGeneratedMetadata = (Get-Content -LiteralPath (Join-Path $meetingInputsOutputDir 'metadata.auto.json') -Raw -Encoding UTF8) | ConvertFrom-Json
    Assert-True -Condition ([string]$meetingGeneratedMetadata.记录人 -eq '李四') -Message 'Meeting-minutes report-input generation metadata is missing the recorder label.'
    Assert-True -Condition ([string]$meetingGeneratedMetadata.会议主题 -eq '第 6 周迭代评审会') -Message 'Meeting-minutes report-input generation metadata is missing the title label.'
    $meetingGeneratedRequirements = (Get-Content -LiteralPath (Join-Path $meetingInputsOutputDir 'requirements.auto.json') -Raw -Encoding UTF8) | ConvertFrom-Json
    Assert-True -Condition ([int]$meetingGeneratedRequirements.minChars -eq 1300) -Message 'Meeting-minutes report-input generation requirements are missing the expected minChars.'
    Assert-True -Condition (@($meetingGeneratedRequirements.sections | Where-Object { $_.name -eq '讨论过程与决议' }).Count -eq 1) -Message 'Meeting-minutes report-input generation requirements are missing the discussion section.'
    Assert-True -Condition ([int]$meetingGeneratedRequirements.paginationRiskThresholds.longSectionChars -eq 1000) -Message 'Meeting-minutes report-input generation requirements are missing pagination-risk thresholds.'

    $monthlyInputsOutputDir = Join-Path $tempRoot 'monthly-report-inputs-output'
    & (Join-Path $repoRoot 'scripts\generate-report-inputs.ps1') `
      -CourseName '校园导览小程序' `
      -ExperimentName '2026 年 4 月项目月报' `
      -StudentName '李四' `
      -StudentId '20261234' `
      -ClassName '软工 2302' `
      -TeacherName '王老师' `
      -ExperimentProperty '项目月报' `
      -ExperimentDate '2026-04' `
      -ExperimentLocation 'GitHub + 飞书 + 本地开发环境' `
      -ReportProfileName 'monthly-report' `
      -RequiredKeywords @('月度进展', '阶段成果') `
      -OutputDir $monthlyInputsOutputDir `
      -DetailLevel full | Out-Null
    $monthlyInputsSummaryPath = Join-Path $monthlyInputsOutputDir 'report-inputs-summary.json'
    Assert-True -Condition (Test-Path -LiteralPath $monthlyInputsSummaryPath) -Message 'Monthly report-input generation did not create the summary JSON.'
    $monthlyInputsSummary = (Get-Content -LiteralPath $monthlyInputsSummaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
    Assert-True -Condition ([string]$monthlyInputsSummary.reportProfileName -eq 'monthly-report') -Message 'Monthly report-input generation summary is missing the expected profile name.'
    Assert-True -Condition ([string]$monthlyInputsSummary.reportProfileDisplayName -eq '月报') -Message 'Monthly report-input generation summary is missing the expected display name.'
    Assert-True -Condition ((Split-Path -Leaf ([string]$monthlyInputsSummary.defaultsPath)) -eq 'monthly-report.defaults.json') -Message 'Monthly report-input generation should persist defaults under the profile-specific defaults file.'
    $monthlyGeneratedPrompt = Get-Content -LiteralPath (Join-Path $monthlyInputsOutputDir 'prompt.txt') -Raw -Encoding UTF8
    Assert-True -Condition ($monthlyGeneratedPrompt -match '月报 body') -Message 'Monthly report-input generation did not emit the expected prompt body.'
    Assert-True -Condition ($monthlyGeneratedPrompt -match '月报主题: 2026 年 4 月项目月报') -Message 'Monthly report-input generation did not emit the expected title label.'
    $monthlyGeneratedMetadata = (Get-Content -LiteralPath (Join-Path $monthlyInputsOutputDir 'metadata.auto.json') -Raw -Encoding UTF8) | ConvertFrom-Json
    Assert-True -Condition ([string]$monthlyGeneratedMetadata.提交人 -eq '李四') -Message 'Monthly report-input generation metadata is missing the owner label.'
    Assert-True -Condition ([string]$monthlyGeneratedMetadata.月报主题 -eq '2026 年 4 月项目月报') -Message 'Monthly report-input generation metadata is missing the title label.'
    $monthlyGeneratedRequirements = (Get-Content -LiteralPath (Join-Path $monthlyInputsOutputDir 'requirements.auto.json') -Raw -Encoding UTF8) | ConvertFrom-Json
    Assert-True -Condition ([int]$monthlyGeneratedRequirements.minChars -eq 1800) -Message 'Monthly report-input generation requirements are missing the expected minChars.'
    Assert-True -Condition (@($monthlyGeneratedRequirements.sections | Where-Object { $_.name -eq '本月完成事项' }).Count -eq 1) -Message 'Monthly report-input generation requirements are missing the completion section.'
    Assert-True -Condition ([int]$monthlyGeneratedRequirements.paginationRiskThresholds.longSectionChars -eq 1600) -Message 'Monthly report-input generation requirements are missing pagination-risk thresholds.'
    Assert-True -Condition ([string]$monthlyGeneratedRequirements.paginationRiskRemediations.'pagination-risk-long-section' -match 'week-by-week') -Message 'Monthly report-input generation requirements are missing the custom long-section remediation.'
    Assert-True -Condition ([string]$monthlyGeneratedRequirements.paginationRiskRemediations.'pagination-risk-dense-section-block' -match 'objective, action, evidence, and impact') -Message 'Monthly report-input generation requirements are missing the custom dense-section remediation.'
  } finally {
    $env:AGENTS_HOME = $originalInputsAgentsHome
  }
  $results.Add('report inputs generation OK') | Out-Null

  foreach ($scriptPath in Get-ChildItem -LiteralPath (Join-Path $repoRoot 'scripts') -Filter *.ps1 | Select-Object -ExpandProperty FullName) {
    Test-PowerShellSyntax -Path $scriptPath
  }
  $results.Add('PowerShell syntax OK') | Out-Null

  $sampleDocx = Join-Path $tempRoot 'sample-template.docx'
  New-SampleTemplateDocx -Path $sampleDocx
  Assert-True -Condition (Test-Path -LiteralPath $sampleDocx) -Message 'Failed to create sample docx fixture.'
  $results.Add('sample docx fixture OK') | Out-Null

  $markdownOutput = & (Join-Path $repoRoot 'scripts\extract-docx-template.ps1') -Path $sampleDocx -Format markdown | Out-String
  Assert-True -Condition ($markdownOutput -match 'DOCX Template Outline') -Message 'Markdown extractor output missing header.'
  Assert-True -Condition ($markdownOutput -match '课程名称') -Message 'Markdown extractor output missing expected paragraph text.'
  Assert-True -Condition ($markdownOutput -match 'T1R1C2') -Message 'Markdown extractor output missing expected table cell location.'
  Assert-True -Condition ($markdownOutput -match '实验目的') -Message 'Markdown extractor output missing expected section heading.'
  $results.Add('docx extractor markdown OK') | Out-Null

  $jsonOutput = & (Join-Path $repoRoot 'scripts\extract-docx-template.ps1') -Path $sampleDocx -Format json | Out-String
  $jsonResult = $jsonOutput | ConvertFrom-Json
  Assert-True -Condition ($jsonResult.summary.tableCount -eq 1) -Message 'JSON extractor reported unexpected table count.'
  Assert-True -Condition ($jsonResult.summary.paragraphCount -ge 9) -Message 'JSON extractor reported unexpected paragraph count.'
  Assert-True -Condition ($jsonResult.likelyFields.Count -ge 5) -Message 'JSON extractor reported too few likely fields.'
  Assert-True -Condition (($jsonResult.likelyFields | Where-Object { $_.reason -eq 'common-report-section-heading' }).Count -ge 2) -Message 'JSON extractor did not detect common section headings.'
  $results.Add('docx extractor json OK') | Out-Null

  $repeatMarkdownOutput = & (Join-Path $repoRoot 'scripts\extract-docx-template.ps1') -Path $sampleDocx -Format markdown | Out-String
  Assert-True -Condition ((Normalize-OutlineForComparison -Text $markdownOutput) -eq (Normalize-OutlineForComparison -Text $repeatMarkdownOutput)) -Message 'Extractor output changed between repeated runs.'
  $results.Add('docx extractor repeatability OK') | Out-Null

  $invalidFile = Join-Path $tempRoot 'not-a-docx.txt'
  [System.IO.File]::WriteAllText($invalidFile, 'placeholder', (New-Object System.Text.UTF8Encoding($true)))
  $invalidRejected = $false
  try {
    & (Join-Path $repoRoot 'scripts\extract-docx-template.ps1') -Path $invalidFile -Format markdown | Out-Null
  } catch {
    $invalidRejected = $true
  }
  Assert-True -Condition $invalidRejected -Message 'Extractor should reject non-docx input.'
  $results.Add('docx extractor invalid-input guard OK') | Out-Null

  $labelMappingFile = Join-Path $tempRoot 'label-field-map.json'
  @'
{
  "课程名称": "计算机网络",
  "班级": "计科 2201",
  "姓名": "张三",
  "学号": "20260001",
  "指导教师": "李老师",
  "实验名称": "不应覆盖",
  "实验目的": {
    "mode": "after",
    "paragraphs": [
      "掌握网络拓扑搭建流程。",
      "理解常用 DOS 命令的作用。"
    ]
  },
  "实验步骤": [
    "配置虚拟机网络参数。",
    "执行 ipconfig 与 ping 验证连通性。"
  ]
}
'@ | Set-Content -LiteralPath $labelMappingFile -Encoding UTF8

  $labelFilledDocx = Join-Path $tempRoot 'sample-template.label-filled.docx'
  $labelFillResult = & (Join-Path $repoRoot 'scripts\apply-docx-field-map.ps1') -TemplatePath $sampleDocx -MappingPath $labelMappingFile -OutPath $labelFilledDocx
  Assert-True -Condition (Test-Path -LiteralPath $labelFilledDocx) -Message 'Label-based fill did not create the filled docx.'
  Assert-True -Condition ($labelFillResult.labelFillCount -ge 4) -Message 'Label-based fill applied too few fields.'
  Assert-True -Condition ($labelFillResult.blockFillCount -ge 2) -Message 'Label-based fill did not report expected block fills.'
  Assert-True -Condition ($labelFillResult.insertedParagraphCount -ge 2) -Message 'Label-based fill did not insert expected continuation paragraphs.'
  $labelFilledOutline = & (Join-Path $repoRoot 'scripts\extract-docx-template.ps1') -Path $labelFilledDocx -Format markdown | Out-String
  Assert-True -Condition ($labelFilledOutline -match '课程名称：计算机网络') -Message 'Label-based fill did not update the course name paragraph.'
  Assert-True -Condition ($labelFilledOutline -match '班级：计科 2201') -Message 'Label-based fill did not update the class paragraph.'
  Assert-True -Condition ($labelFilledOutline -match 'T1R1C2: 张三') -Message 'Label-based fill did not update the name cell.'
  Assert-True -Condition ($labelFilledOutline -match 'T1R2C2: 20260001') -Message 'Label-based fill did not update the student id cell.'
  Assert-True -Condition ($labelFilledOutline -match '指导教师：李老师') -Message 'Label-based fill did not update the teacher field.'
  Assert-True -Condition ($labelFilledOutline -match '实验名称：保留原值') -Message 'Label-based fill should not overwrite existing non-placeholder text.'
  Assert-True -Condition ($labelFilledOutline -match '实验目的') -Message 'Label-based fill should keep the section heading paragraph.'
  Assert-True -Condition ($labelFilledOutline -match '掌握网络拓扑搭建流程。') -Message 'Label-based block fill did not write the first purpose paragraph.'
  Assert-True -Condition ($labelFilledOutline -match '理解常用 DOS 命令的作用。') -Message 'Label-based block fill did not write the second purpose paragraph.'
  Assert-True -Condition ($labelFilledOutline -match '配置虚拟机网络参数。') -Message 'Label-based block fill did not write the first procedure paragraph.'
  Assert-True -Condition ($labelFilledOutline -match '执行 ipconfig 与 ping 验证连通性。') -Message 'Label-based block fill did not write the second procedure paragraph.'
  $results.Add('docx fill label mapping OK') | Out-Null

  $locationMappingFile = Join-Path $tempRoot 'location-field-map.json'
  @'
{
  "P4": "实验名称：已按位置覆盖",
  "T1R2C2": "20261234",
  "P10": [
    "实验结果第一段：成功获取 IP 地址。",
    "实验结果第二段：连通性测试通过。"
  ]
}
'@ | Set-Content -LiteralPath $locationMappingFile -Encoding UTF8

  $locationFilledDocx = Join-Path $tempRoot 'sample-template.location-filled.docx'
  $locationFillResult = & (Join-Path $repoRoot 'scripts\apply-docx-field-map.ps1') -TemplatePath $sampleDocx -MappingPath $locationMappingFile -OutPath $locationFilledDocx
  Assert-True -Condition (Test-Path -LiteralPath $locationFilledDocx) -Message 'Location-based fill did not create the filled docx.'
  Assert-True -Condition ($locationFillResult.directFillCount -ge 3) -Message 'Location-based fill applied too few direct fields.'
  Assert-True -Condition ($locationFillResult.blockFillCount -ge 1) -Message 'Location-based fill did not report the expected block fill.'
  $locationFilledOutline = & (Join-Path $repoRoot 'scripts\extract-docx-template.ps1') -Path $locationFilledDocx -Format markdown | Out-String
  Assert-True -Condition ($locationFilledOutline -match '实验名称：已按位置覆盖') -Message 'Location-based fill did not override the experiment name paragraph.'
  Assert-True -Condition ($locationFilledOutline -match 'T1R2C2: 20261234') -Message 'Location-based fill did not override the target table cell.'
  Assert-True -Condition ($locationFilledOutline -match '实验结果第一段：成功获取 IP 地址。') -Message 'Location-based block fill did not write the first result paragraph.'
  Assert-True -Condition ($locationFilledOutline -match '实验结果第二段：连通性测试通过。') -Message 'Location-based block fill did not write the second result paragraph.'
  $results.Add('docx fill location mapping OK') | Out-Null

  $repeatFilledDocx = Join-Path $tempRoot 'sample-template.label-filled-repeat.docx'
  & (Join-Path $repoRoot 'scripts\apply-docx-field-map.ps1') -TemplatePath $sampleDocx -MappingPath $labelMappingFile -OutPath $repeatFilledDocx | Out-Null
  $repeatFilledOutline = & (Join-Path $repoRoot 'scripts\extract-docx-template.ps1') -Path $repeatFilledDocx -Format markdown | Out-String
  Assert-True -Condition ((Normalize-OutlineForComparison -Text $labelFilledOutline) -eq (Normalize-OutlineForComparison -Text $repeatFilledOutline)) -Message 'Label-based fill output changed between repeated runs.'
  $results.Add('docx fill repeatability OK') | Out-Null

  $sampleReportPath = Join-Path $tempRoot 'sample-report.md'
  @'
计算机网络实验报告

课程名称：计算机网络
实验名称：局域网搭建与常用 DOS 命令使用

一、实验目的
通过本次实验掌握局域网的基本搭建方法，并理解 DOS 命令在网络排查中的作用。
通过独立完成地址配置、连通性验证和缓存查看，熟悉常见网络实验的基本检查流程。

二、实验环境
实验环境为 Windows 11 主机、VMware Workstation 以及两台 Windows Server 2019 虚拟机。
两台虚拟机均使用仅主机网络模式，便于在受控环境中完成局域网连通测试。

三、实验原理或任务要求
本实验要求将两台主机配置到同一网段，通过 ipconfig、ping 和 arp 等命令检查地址信息与互通情况。
当主机位于同一网段且地址配置正确时，可以通过 ICMP 回显测试和 ARP 缓存记录判断网络通信是否正常建立。

四、实验步骤
首先将两台虚拟机配置在 192.168.10.0/24 网段，其中主机 A 为 192.168.10.11，主机 B 为 192.168.10.12。
1. 在主机 A 上查看地址并测试到主机 B 的连通性。

ipconfig

ping 192.168.10.12

arp -a

2. 在主机 B 上查看地址并测试到主机 A 的连通性。

ipconfig

ping 192.168.10.11

arp -a

最后对两台主机的命令输出进行核对，确认子网掩码均为 255.255.255.0。
在实验过程中还需要反复核对网卡是否启用以及地址是否写入到正确的网络接口，避免由于配置位置错误导致测试结果失真。

五、实验结果
实验结果表明，两台主机能够正常互通，ping 测试延迟稳定且无丢包。
通过 arp -a 可以看到对端主机的缓存记录，说明局域网通信建立正常。
从命令输出可以确认两台主机均获取到了预期的静态地址，网络参数与实验要求保持一致。

六、问题分析
如果网段配置错误或子网掩码不一致，局域网主机之间将无法正常通信，因此在 DOS 命令检查中必须优先确认基础地址参数。
如果只关注 ping 结果而忽略 ipconfig 与 arp 信息，容易遗漏地址冲突、接口选错或缓存未更新等隐蔽问题。

七、实验总结
本次实验完成了局域网搭建与常用 DOS 命令使用，进一步掌握了网络参数查看和互通测试方法。
通过将 DOS 命令结果与网络配置过程对应分析，可以更加系统地理解局域网实验中地址规划、连通验证和故障定位之间的关系。
'@ | Set-Content -LiteralPath $sampleReportPath -Encoding UTF8

  $validationOutput = & (Join-Path $repoRoot 'scripts\validate-report-draft.ps1') -Path $sampleReportPath -RequirementsPath (Join-Path $repoRoot 'examples\e2e-sample-requirements.json') -Format json | Out-String
  $validationResult = $validationOutput | ConvertFrom-Json
  Assert-True -Condition ($validationResult.passed) -Message 'Report validation should pass for the sample report.'
  Assert-True -Condition ($validationResult.summary.errorCount -eq 0) -Message 'Report validation returned unexpected errors for the sample report.'
  $defaultValidationResult = (& (Join-Path $repoRoot 'scripts\validate-report-draft.ps1') -Path $sampleReportPath -Format json | Out-String) | ConvertFrom-Json
  Assert-True -Condition ([bool]$defaultValidationResult.passed) -Message 'Profile-backed default report validation should pass for the sample report.'
  Assert-True -Condition ([string]$defaultValidationResult.reportProfileName -eq 'experiment-report') -Message 'Default report validation is missing the expected profile name.'
  Assert-True -Condition ([int]$defaultValidationResult.summary.sectionCount -ge 7) -Message 'Default report validation did not load the profile section rules.'
  $results.Add('report validation OK') | Out-Null

  $experimentDenseResults = ((@(
        '实验结果表明网络参数检查、ping 连通验证和 arp 缓存比对保持一致，交换机链路状态与主机地址配置保持稳定，因此可以持续确认局域网搭建过程满足实验目标并具备重复验证条件，相关截图见图1、图2和图3。'
      ) * 18) -join '')
  $experimentStructureRiskReportPath = Join-Path $tempRoot 'experiment-structure-risk-report.md'
  @(
    '计算机网络实验报告',
    '',
    '课程名称：计算机网络',
    '实验名称：局域网搭建与常用 DOS 命令使用',
    '',
    '一、实验目的',
    '本节说明局域网搭建实验的验证目标，并明确需要通过命令观察网络参数、地址规划和主机连通状态之间的对应关系。',
    '为了保证后续分析有依据，报告还需要把网络配置目的与命令验证思路联系起来，而不是只给出结论。',
    '',
    '二、实验环境',
    '实验环境包括两台处于同一网段的虚拟机、一台二层交换设备以及用于查看网络参数的 DOS 命令窗口，所有节点都保持固定地址配置。',
    '在操作前先确认虚拟机网卡模式、交换连接关系和主机名设置一致，以便后续步骤能够稳定复现。',
    '',
    '三、实验结果',
    $experimentDenseResults,
    '',
    '四、实验步骤',
    '先配置两台主机的静态地址和子网掩码，再分别执行 ipconfig、ping 与 arp -a，对照输出结果检查网络参数、邻居缓存和主机互通状态是否符合预期。',
    '',
    '五、实验总结',
    '__________',
    '',
    '六、实验目的',
    '重复标题段落仅用于触发结构校验。'
  ) | Set-Content -LiteralPath $experimentStructureRiskReportPath -Encoding UTF8
  $experimentStructureRiskValidation = (& (Join-Path $repoRoot 'scripts\validate-report-draft.ps1') -Path $experimentStructureRiskReportPath -Format json | Out-String) | ConvertFrom-Json
  $experimentStructureRiskCodes = @($experimentStructureRiskValidation.findings | ForEach-Object { [string]$_.code })
  $experimentStructureRiskWarningCodes = @($experimentStructureRiskValidation.summary.warningCodes)
  Assert-True -Condition (-not [bool]$experimentStructureRiskValidation.passed) -Message 'Experiment structural-risk fixture should not pass validation.'
  Assert-True -Condition ([string]$experimentStructureRiskValidation.reportProfileName -eq 'experiment-report') -Message 'Experiment structural-risk fixture should use the default experiment profile.'
  Assert-True -Condition ($experimentStructureRiskCodes -contains 'missing-profile-required-heading') -Message 'Experiment structural-risk fixture should report missing-profile-required-heading.'
  Assert-True -Condition ($experimentStructureRiskCodes -contains 'duplicate-section-heading') -Message 'Experiment structural-risk fixture should report duplicate-section-heading.'
  Assert-True -Condition ($experimentStructureRiskCodes -contains 'section-order-anomaly') -Message 'Experiment structural-risk fixture should report section-order-anomaly.'
  Assert-True -Condition ($experimentStructureRiskCodes -contains 'placeholder-only-section') -Message 'Experiment structural-risk fixture should report placeholder-only-section.'
  Assert-True -Condition (([int]$experimentStructureRiskValidation.summary.paginationRiskCount) -ge 3) -Message 'Experiment structural-risk fixture should surface pagination risk warnings.'
  Assert-True -Condition ($experimentStructureRiskWarningCodes -contains 'pagination-risk-long-section') -Message 'Experiment structural-risk fixture should report pagination-risk-long-section.'
  Assert-True -Condition ($experimentStructureRiskWarningCodes -contains 'pagination-risk-dense-section-block') -Message 'Experiment structural-risk fixture should report pagination-risk-dense-section-block.'
  Assert-True -Condition ($experimentStructureRiskWarningCodes -contains 'pagination-risk-figure-cluster') -Message 'Experiment structural-risk fixture should report pagination-risk-figure-cluster.'
  $duplicateHeadingFinding = @($experimentStructureRiskValidation.findings | Where-Object { [string]$_.code -eq 'duplicate-section-heading' })[0]
  $placeholderFinding = @($experimentStructureRiskValidation.findings | Where-Object { [string]$_.code -eq 'placeholder-only-section' })[0]
  $longPaginationFinding = @($experimentStructureRiskValidation.findings | Where-Object { [string]$_.code -eq 'pagination-risk-long-section' })[0]
  Assert-True -Condition ([string]$duplicateHeadingFinding.remediation -match 'canonical heading') -Message 'Experiment structural-risk fixture should include duplicate-heading remediation guidance.'
  Assert-True -Condition ([string]$placeholderFinding.remediation -match 'Replace placeholders') -Message 'Experiment structural-risk fixture should include placeholder remediation guidance.'
  Assert-True -Condition ([string]$longPaginationFinding.remediation -match 'paginationRiskThresholds\.longSectionChars') -Message 'Experiment structural-risk fixture should include pagination threshold remediation guidance.'
  $results.Add('experiment structural validation OK') | Out-Null

  $experimentPaginationRiskReportPath = Join-Path $tempRoot 'experiment-pagination-risk-report.md'
  @(
    '计算机网络实验报告',
    '',
    '课程名称：计算机网络',
    '实验名称：局域网搭建与常用 DOS 命令使用',
    '',
    '一、实验目的',
    '本节说明局域网搭建实验的验证目标，并明确需要通过命令观察网络参数、地址规划和主机连通状态之间的对应关系。',
    '为了保证后续分析有依据，报告还需要把网络配置目的与命令验证思路联系起来，而不是只给出结论。',
    '',
    '二、实验环境',
    '实验环境包括两台处于同一网段的虚拟机、一台二层交换设备以及用于查看网络参数的 DOS 命令窗口，所有节点都保持固定地址配置。',
    '在操作前先确认虚拟机网卡模式、交换连接关系和主机名设置一致，以便后续步骤能够稳定复现。',
    '',
    '三、实验原理或任务要求',
    '同一局域网内的主机需要具备一致的网络号和正确的子网掩码，通信过程中可以通过 ICMP 回显和 ARP 地址解析观察链路是否正常。',
    '任务要求依次完成地址配置、ipconfig 参数检查、ping 连通验证和 arp 缓存查看，并结合输出解释局域网通信是否建立。',
    '',
    '四、实验步骤',
    '先配置两台主机的静态地址和子网掩码，再分别执行 ipconfig、ping 与 arp -a，对照输出结果检查网络参数、邻居缓存和主机互通状态是否符合预期。',
    '',
    '五、实验结果',
    $experimentDenseResults,
    '',
    '六、问题分析',
    '如果 ping 不通，应优先检查 IP 地址、子网掩码、虚拟网卡模式和防火墙策略，再结合 arp 输出判断是否已经完成地址解析。',
    '如果只观察单次 ping 结果而忽略 ipconfig 和 arp 信息，可能遗漏网卡选错、地址冲突或缓存未更新等问题。',
    '',
    '七、实验总结',
    '本次实验完成了局域网搭建和常用 DOS 命令验证，能够从地址配置、连通测试和缓存记录三个角度说明实验结果。',
    '通过把命令输出与配置步骤逐项对应，进一步理解了局域网通信中地址规划、协议验证和故障定位之间的关系。'
  ) | Set-Content -LiteralPath $experimentPaginationRiskReportPath -Encoding UTF8
  $experimentPaginationRiskValidation = (& (Join-Path $repoRoot 'scripts\validate-report-draft.ps1') -Path $experimentPaginationRiskReportPath -Format json | Out-String) | ConvertFrom-Json
  Assert-True -Condition ([bool]$experimentPaginationRiskValidation.passed) -Message 'Experiment pagination-risk fixture should pass with warnings only.'
  Assert-True -Condition ([int]$experimentPaginationRiskValidation.summary.paginationRiskCount -ge 3) -Message 'Experiment pagination-risk fixture should surface default pagination warnings.'
  Assert-True -Condition ([int]$experimentPaginationRiskValidation.summary.paginationRiskThresholds.longSectionChars -eq 900) -Message 'Validation summary should expose default pagination-risk thresholds.'

  $quietPaginationProfilePath = Join-Path $tempRoot 'experiment-report-quiet-pagination.json'
  $quietPaginationProfile = (Get-Content -LiteralPath (Join-Path $repoRoot 'profiles\experiment-report.json') -Raw -Encoding UTF8) | ConvertFrom-Json
  $quietPaginationProfile.name = 'experiment-report-quiet-pagination'
  $quietPaginationProfile.paginationRiskThresholds = [pscustomobject]@{
    longSectionChars = 99999
    denseSectionChars = 99999
    denseSectionParagraphs = 1
    figureClusterRefs = 999
  }
  $quietPaginationProfile | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $quietPaginationProfilePath -Encoding UTF8
  $quietPaginationValidation = (& (Join-Path $repoRoot 'scripts\validate-report-draft.ps1') -Path $experimentPaginationRiskReportPath -ReportProfilePath $quietPaginationProfilePath -Format json | Out-String) | ConvertFrom-Json
  Assert-True -Condition ([bool]$quietPaginationValidation.passed) -Message 'Custom pagination-threshold profile should keep validation passing.'
  Assert-True -Condition ([int]$quietPaginationValidation.summary.paginationRiskCount -eq 0) -Message 'Custom pagination-threshold profile should suppress pagination warnings for the same report.'
  Assert-True -Condition ([int]$quietPaginationValidation.summary.paginationRiskThresholds.longSectionChars -eq 99999) -Message 'Validation summary should expose custom pagination-risk thresholds.'
  $results.Add('profile pagination-risk thresholds OK') | Out-Null

  $referenceTextPath = Join-Path $tempRoot 'tutorial-reference.txt'
  @'
TITLE: 局域网实验参考流程
URL: https://example.com/network-lab

首先为两台虚拟机配置同一网段地址，并确认子网掩码一致。
随后使用 ipconfig 查看地址配置，再通过 ping 验证两台主机之间是否可以正常通信。
最后执行 arp -a 查看邻居缓存，结合命令结果分析局域网通信是否已经建立。
'@ | Set-Content -LiteralPath $referenceTextPath -Encoding UTF8
  $preparedPromptPath = Join-Path $tempRoot 'prepared-prompt.txt'
  $preparedPromptResult = & (Join-Path $repoRoot 'scripts\prepare-report-prompt.ps1') `
    -PromptText @'
/experiment-report
写一份完整的实验报告正文。
'@ `
    -ReferenceTextPaths $referenceTextPath `
    -OutFile $preparedPromptPath
  Assert-True -Condition (Test-Path -LiteralPath $preparedPromptPath) -Message 'prepare-report-prompt did not create the prepared prompt file.'
  Assert-True -Condition ([int]$preparedPromptResult.referenceCount -eq 1) -Message 'prepare-report-prompt reported an unexpected reference count.'
  Assert-True -Condition ([string]$preparedPromptResult.sources[0] -eq 'https://example.com/network-lab') -Message 'prepare-report-prompt did not preserve the reference source URL.'
  $preparedPromptText = Get-Content -LiteralPath $preparedPromptPath -Raw -Encoding UTF8
  Assert-True -Condition ($preparedPromptText -match '/experiment-report') -Message 'prepare-report-prompt did not preserve the base prompt.'
  Assert-True -Condition ($preparedPromptText -match 'Reference Material 1') -Message 'prepare-report-prompt is missing the appended reference section.'
  Assert-True -Condition ($preparedPromptText -match 'Source:\s+https://example.com/network-lab') -Message 'prepare-report-prompt is missing the appended reference source.'
  Assert-True -Condition ($preparedPromptText -match '不要逐字照抄|Do not copy them verbatim') -Message 'prepare-report-prompt is missing the anti-copying guidance.'
  Assert-True -Condition ($preparedPromptText -match 'ipconfig') -Message 'prepare-report-prompt is missing the reference content body.'
  $results.Add('report prompt preparation OK') | Out-Null

  $generatedFieldMapPath = Join-Path $tempRoot 'generated-field-map.json'
  & (Join-Path $repoRoot 'scripts\generate-docx-field-map.ps1') `
    -TemplatePath $sampleDocx `
    -ReportPath $sampleReportPath `
    -MetadataPath (Join-Path $repoRoot 'examples\docx-report-metadata.json') `
    -Format json `
    -OutFile $generatedFieldMapPath | Out-Null
  Assert-True -Condition (Test-Path -LiteralPath $generatedFieldMapPath) -Message 'Field-map generator did not write the output file.'
  $generatedFieldMapRoot = (Get-Content -LiteralPath $generatedFieldMapPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$generatedFieldMapRoot.reportProfileName -eq 'experiment-report') -Message 'Field-map generator is missing the expected report profile name.'
  Assert-True -Condition ([string]$generatedFieldMapRoot.reportInputMode -eq 'path') -Message 'Field-map generator should record reportInputMode=path for file-backed reports.'
  Assert-True -Condition ([string]$generatedFieldMapRoot.metadataInputMode -eq 'path') -Message 'Field-map generator should record metadataInputMode=path for metadata files.'
  Assert-True -Condition ($generatedFieldMapRoot.summary.fieldCount -ge 7) -Message 'Field-map generator produced too few fields.'
  Assert-True -Condition ($generatedFieldMapRoot.fieldMap.PSObject.Properties.Name -contains '课程名称') -Message 'Field-map generator did not map the course name.'
  Assert-True -Condition ($generatedFieldMapRoot.fieldMap.PSObject.Properties.Name -contains '班级') -Message 'Field-map generator did not map the class field.'
  Assert-True -Condition ($generatedFieldMapRoot.fieldMap.PSObject.Properties.Name -contains '姓名') -Message 'Field-map generator did not map the student name field.'
  Assert-True -Condition ($generatedFieldMapRoot.fieldMap.PSObject.Properties.Name -contains '实验目的') -Message 'Field-map generator did not map the purpose section.'
  Assert-True -Condition ($generatedFieldMapRoot.fieldMap.实验目的.mode -eq 'after') -Message 'Field-map generator should preserve the purpose heading and fill after it.'
  Assert-True -Condition (@($generatedFieldMapRoot.fieldMap.实验目的.paragraphs).Count -ge 1) -Message 'Field-map generator did not include purpose section paragraphs.'
  Assert-True -Condition ($generatedFieldMapRoot.fieldMap.实验步骤.mode -eq 'after') -Message 'Field-map generator should preserve the procedure heading and fill after it.'
  $results.Add('docx field-map generation OK') | Out-Null

  $inlineReportText = Get-Content -LiteralPath $sampleReportPath -Raw -Encoding UTF8
  $inlineMetadataJson = Get-Content -LiteralPath (Join-Path $repoRoot 'examples\docx-report-metadata.json') -Raw -Encoding UTF8
  $inlineFieldMapRoot = ((& (Join-Path $repoRoot 'scripts\generate-docx-field-map.ps1') `
      -TemplatePath $sampleDocx `
      -ReportText $inlineReportText `
      -MetadataJson $inlineMetadataJson `
      -Format json) | Out-String) | ConvertFrom-Json
  Assert-True -Condition ([string]$inlineFieldMapRoot.reportInputMode -eq 'inline') -Message 'Field-map generator should record reportInputMode=inline for inline report text.'
  Assert-True -Condition ([string]$inlineFieldMapRoot.metadataInputMode -eq 'inline') -Message 'Field-map generator should record metadataInputMode=inline for inline metadata JSON.'
  Assert-True -Condition ([string]$inlineFieldMapRoot.reportSource -eq '[inline text]') -Message 'Field-map generator should expose [inline text] as the reportSource for inline report input.'
  Assert-True -Condition ([string]$inlineFieldMapRoot.fieldMap.课程名称 -eq '计算机网络') -Message 'Field-map generator should still map metadata values when using inline inputs.'
  Assert-True -Condition ($inlineFieldMapRoot.fieldMap.实验步骤.mode -eq 'after') -Message 'Field-map generator should still preserve section headings when using inline report text.'
  $results.Add('docx field-map inline inputs OK') | Out-Null

  $generatedFieldMapMarkdown = & (Join-Path $repoRoot 'scripts\generate-docx-field-map.ps1') `
    -TemplatePath $sampleDocx `
    -ReportPath $sampleReportPath `
    -MetadataPath (Join-Path $repoRoot 'examples\docx-report-metadata.json') `
    -Format markdown | Out-String
  Assert-True -Condition ($generatedFieldMapMarkdown -match 'DOCX Field Map') -Message 'Field-map generator markdown output missing header.'
  Assert-True -Condition ($generatedFieldMapMarkdown -match '课程名称') -Message 'Field-map generator markdown output missing expected content.'
  $results.Add('docx field-map markdown OK') | Out-Null

  $generatedFilledDocx = Join-Path $tempRoot 'sample-template.generated-filled.docx'
  $generatedFillResult = & (Join-Path $repoRoot 'scripts\apply-docx-field-map.ps1') -TemplatePath $sampleDocx -MappingPath $generatedFieldMapPath -OutPath $generatedFilledDocx
  Assert-True -Condition (Test-Path -LiteralPath $generatedFilledDocx) -Message 'Generated field-map fill did not create the filled docx.'
  Assert-True -Condition ($generatedFillResult.labelFillCount -ge 5) -Message 'Generated field-map fill applied too few label fields.'
  Assert-True -Condition ($generatedFillResult.blockFillCount -ge 3) -Message 'Generated field-map fill applied too few block fills.'
  $generatedFilledOutline = & (Join-Path $repoRoot 'scripts\extract-docx-template.ps1') -Path $generatedFilledDocx -Format markdown | Out-String
  Assert-True -Condition ($generatedFilledOutline -match '课程名称：计算机网络') -Message 'Generated field-map fill did not update the course name paragraph.'
  Assert-True -Condition ($generatedFilledOutline -match '班级：计科 2201') -Message 'Generated field-map fill did not update the class paragraph.'
  Assert-True -Condition ($generatedFilledOutline -match 'T1R1C2: 张三') -Message 'Generated field-map fill did not update the student name cell.'
  Assert-True -Condition ($generatedFilledOutline -match 'T1R2C2: 20260001') -Message 'Generated field-map fill did not update the student id cell.'
  Assert-True -Condition ($generatedFilledOutline -match '指导教师：李老师') -Message 'Generated field-map fill did not update the teacher field.'
  Assert-True -Condition ($generatedFilledOutline -match '实验名称：保留原值') -Message 'Generated field-map fill should not overwrite locked experiment text.'
  Assert-True -Condition ($generatedFilledOutline -match '通过本次实验掌握局域网的基本搭建方法') -Message 'Generated field-map fill did not write the purpose section text.'
  Assert-True -Condition ($generatedFilledOutline -match '首先将两台虚拟机配置在 192.168.10.0/24 网段') -Message 'Generated field-map fill did not write the procedure section text.'
  Assert-True -Condition ($generatedFilledOutline -match '实验结果表明，两台主机能够正常互通') -Message 'Generated field-map fill did not write the result section text.'
  $results.Add('docx fill from generated field map OK') | Out-Null

  $coverBodyDocx = Join-Path $tempRoot 'cover-body-template.docx'
  New-CoverBodyTemplateDocx -Path $coverBodyDocx
  Assert-True -Condition (Test-Path -LiteralPath $coverBodyDocx) -Message 'Failed to create the cover-body template fixture.'

  $coverBodyFieldMapPath = Join-Path $tempRoot 'cover-body-field-map.json'
  & (Join-Path $repoRoot 'scripts\generate-docx-field-map.ps1') `
    -TemplatePath $coverBodyDocx `
    -ReportPath $sampleReportPath `
    -MetadataPath (Join-Path $repoRoot 'examples\docx-report-metadata.json') `
    -Format json `
    -OutFile $coverBodyFieldMapPath | Out-Null
  $coverBodyFieldMapRoot = (Get-Content -LiteralPath $coverBodyFieldMapPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ($coverBodyFieldMapRoot.fieldMap.PSObject.Properties.Name -contains 'T1R3C1') -Message 'Cover-body template mapping did not include the experiment-property option row.'
  Assert-True -Condition ([string]$coverBodyFieldMapRoot.fieldMap.T1R3C1 -match '√③验证性实验') -Message 'Cover-body template mapping did not mark the selected experiment property option.'
  Assert-True -Condition ($coverBodyFieldMapRoot.fieldMap.PSObject.Properties.Name -contains 'T1R5C1') -Message 'Cover-body template mapping is missing the composite body start cell.'
  Assert-True -Condition ([string]$coverBodyFieldMapRoot.fieldMap.T1R5C1.mode -eq 'after-table') -Message 'Cover-body template mapping should use after-table mode.'
  Assert-True -Condition ([string]$coverBodyFieldMapRoot.fieldMap.T1R5C1.through -eq 'T1R6C1') -Message 'Cover-body template mapping should span the full composite body row range.'
  Assert-True -Condition (@($coverBodyFieldMapRoot.fieldMap.T1R5C1.paragraphs).Count -ge 6) -Message 'Cover-body template mapping is missing expected body paragraphs.'
  Assert-True -Condition ((@($coverBodyFieldMapRoot.diagnostics | Where-Object { $_.code -eq 'composite_body_after_table' }).Count) -ge 1) -Message 'Cover-body template mapping should emit the structured composite-body diagnostic.'
  $results.Add('docx cover-body field-map generation OK') | Out-Null

  $coverBodyFilledDocx = Join-Path $tempRoot 'cover-body-template.filled.docx'
  $coverBodyFillResult = & (Join-Path $repoRoot 'scripts\apply-docx-field-map.ps1') -TemplatePath $coverBodyDocx -MappingPath $coverBodyFieldMapPath -OutPath $coverBodyFilledDocx
  Assert-True -Condition (Test-Path -LiteralPath $coverBodyFilledDocx) -Message 'Cover-body fill did not create the filled docx.'
  Assert-True -Condition ($coverBodyFillResult.removedTableRowCount -eq 2) -Message 'Cover-body fill should remove the two composite body rows from the table.'
  Assert-True -Condition ($coverBodyFillResult.blockFillCount -ge 1) -Message 'Cover-body fill did not report the moved body block.'
  $coverBodyFilledOutlineJson = (& (Join-Path $repoRoot 'scripts\extract-docx-template.ps1') -Path $coverBodyFilledDocx -Format json | Out-String) | ConvertFrom-Json
  $coverBodyTableBlock = @($coverBodyFilledOutlineJson.blocks | Where-Object { $_.type -eq 'table' })[0]
  Assert-True -Condition ($coverBodyTableBlock.rows.Count -eq 4) -Message 'Cover-body fill should leave only the cover rows in the table.'
  Assert-True -Condition ([string]$coverBodyTableBlock.rows[0].cells[0].text -eq '学号：20260001') -Message 'Cover-body fill did not update the student id in the cover table.'
  Assert-True -Condition ([string]$coverBodyTableBlock.rows[0].cells[1].text -eq '姓名：张三') -Message 'Cover-body fill did not update the student name in the cover table.'
  Assert-True -Condition ([string]$coverBodyTableBlock.rows[1].cells[0].text -eq '课程名称：计算机网络') -Message 'Cover-body fill did not update the course name in the cover table.'
  Assert-True -Condition ([string]$coverBodyTableBlock.rows[2].cells[0].text -match '实验性质： ①综合性实验 ②设计性实验 √③验证性实验') -Message 'Cover-body fill did not mark the selected experiment property option in the cover table.'
  Assert-True -Condition ([string]$coverBodyFilledOutlineJson.blocks[2].text -eq '一. 实验目的') -Message 'Cover-body fill should insert the purpose heading immediately after the table.'
  Assert-True -Condition ((@($coverBodyFilledOutlineJson.blocks | Where-Object { $_.type -eq 'paragraph' -and $_.text -match '通过本次实验掌握局域网的基本搭建方法' }).Count) -ge 1) -Message 'Cover-body fill did not move the purpose content after the table.'
  Assert-True -Condition ((@($coverBodyFilledOutlineJson.blocks | Where-Object { $_.type -eq 'paragraph' -and $_.text -eq '三. 实验步骤' }).Count) -ge 1) -Message 'Cover-body fill did not insert the procedure heading after the table.'
  Assert-True -Condition ((@($coverBodyFilledOutlineJson.blocks | Where-Object { $_.type -eq 'paragraph' -and $_.text -match '实验结果表明，两台主机能够正常互通' }).Count) -ge 1) -Message 'Cover-body fill did not move the result content after the table.'
  $results.Add('docx cover-body fill OK') | Out-Null

  $customCompositeProfilePath = Join-Path $tempRoot 'custom-field-map-composite-profile.json'
  @'
{
  "name": "field-map-composite-custom-profile",
  "displayName": "自定义模板复合规则实验报告",
  "defaultStyleProfile": "auto",
  "defaultExperimentProperty": "③验证性实验",
  "metadataFields": [
    { "key": "Name", "label": "姓名" }
  ],
  "extraLabels": [],
  "sectionFields": [
    { "key": "Purpose", "heading": "实验目的", "aliases": ["实验目的"], "minChars": { "standard": 30, "full": 60 } },
    { "key": "Environment", "heading": "实验环境", "aliases": ["实验环境", "实验设备与环境"], "minChars": { "standard": 30, "full": 60 } },
    { "key": "Theory", "heading": "实验原理或任务要求", "aliases": ["实验原理或任务要求", "实验原理", "任务要求"], "minChars": { "standard": 30, "full": 80 } },
    { "key": "Steps", "heading": "实验步骤", "aliases": ["实验步骤", "实验过程"], "minChars": { "standard": 60, "full": 140 } },
    { "key": "Results", "heading": "实验结果", "aliases": ["实验结果", "实验现象与结果记录"], "minChars": { "standard": 50, "full": 120 } },
    { "key": "Analysis", "heading": "问题分析", "aliases": ["问题分析", "结果分析"], "minChars": { "standard": 30, "full": 80 } },
    { "key": "Summary", "heading": "实验总结", "aliases": ["实验总结", "总结与思考", "实验小结"], "minChars": { "standard": 30, "full": 80 } }
  ],
  "fieldMapCompositeRules": [
    {
      "matchAll": ["实验目的", "实验内容"],
      "blocks": [
        { "heading": "甲. 实验目标", "sectionIds": ["purpose"] },
        { "heading": "乙. 环境与原理", "sectionIds": ["environment", "theory"] }
      ]
    },
    {
      "matchAll": ["实验步骤", "实验小结"],
      "blocks": [
        { "heading": "丙. 操作步骤", "sectionIds": ["steps"] },
        { "heading": "丁. 实验结果", "sectionIds": ["result"] },
        { "heading": "戊. 问题分析", "sectionIds": ["analysis"] },
        { "heading": "己. 实验小结", "sectionIds": ["summary"] }
      ]
    }
  ],
  "detailProfiles": {
    "standard": { "minChars": 700, "promptGuidance": [] },
    "full": { "minChars": 1100, "promptGuidance": [] }
  }
}
'@ | Set-Content -LiteralPath $customCompositeProfilePath -Encoding UTF8
  $resolvedCustomCompositeProfilePath = (Resolve-Path -LiteralPath $customCompositeProfilePath).Path

  $customCompositeFieldMapPath = Join-Path $tempRoot 'cover-body-field-map-custom-profile.json'
  & (Join-Path $repoRoot 'scripts\generate-docx-field-map.ps1') `
    -TemplatePath $coverBodyDocx `
    -ReportPath $sampleReportPath `
    -MetadataPath (Join-Path $repoRoot 'examples\docx-report-metadata.json') `
    -ReportProfilePath $customCompositeProfilePath `
    -Format json `
    -OutFile $customCompositeFieldMapPath | Out-Null
  $customCompositeFieldMapRoot = (Get-Content -LiteralPath $customCompositeFieldMapPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$customCompositeFieldMapRoot.reportProfileName -eq 'field-map-composite-custom-profile') -Message 'Custom composite field-map generator is missing the expected report profile name.'
  Assert-True -Condition ([string]$customCompositeFieldMapRoot.reportProfilePath -eq $resolvedCustomCompositeProfilePath) -Message 'Custom composite field-map generator is missing the expected report profile path.'
  Assert-True -Condition ([string]$customCompositeFieldMapRoot.fieldMap.T1R5C1.paragraphs[0] -eq '甲. 实验目标') -Message 'Custom composite field-map generator did not use the profile-defined first composite heading.'
  Assert-True -Condition (@($customCompositeFieldMapRoot.fieldMap.T1R5C1.paragraphs | Where-Object { $_ -eq '乙. 环境与原理' }).Count -ge 1) -Message 'Custom composite field-map generator did not use the profile-defined merged content heading.'
  Assert-True -Condition (@($customCompositeFieldMapRoot.fieldMap.T1R5C1.paragraphs | Where-Object { $_ -eq '丙. 操作步骤' }).Count -ge 1) -Message 'Custom composite field-map generator did not use the profile-defined procedure heading.'

  $customCompositeFilledDocx = Join-Path $tempRoot 'cover-body-template.custom-composite.filled.docx'
  & (Join-Path $repoRoot 'scripts\apply-docx-field-map.ps1') -TemplatePath $coverBodyDocx -MappingPath $customCompositeFieldMapPath -OutPath $customCompositeFilledDocx | Out-Null
  $customCompositeOutlineJson = (& (Join-Path $repoRoot 'scripts\extract-docx-template.ps1') -Path $customCompositeFilledDocx -Format json | Out-String) | ConvertFrom-Json
  Assert-True -Condition ((@($customCompositeOutlineJson.blocks | Where-Object { $_.type -eq 'paragraph' -and $_.text -eq '甲. 实验目标' }).Count) -ge 1) -Message 'Custom composite field-map fill did not insert the custom first heading after the table.'
  Assert-True -Condition ((@($customCompositeOutlineJson.blocks | Where-Object { $_.type -eq 'paragraph' -and $_.text -eq '乙. 环境与原理' }).Count) -ge 1) -Message 'Custom composite field-map fill did not insert the custom merged heading after the table.'
  Assert-True -Condition ((@($customCompositeOutlineJson.blocks | Where-Object { $_.type -eq 'paragraph' -and $_.text -eq '丙. 操作步骤' }).Count) -ge 1) -Message 'Custom composite field-map fill did not insert the custom procedure heading after the table.'
  $results.Add('docx cover-body field-map custom profile OK') | Out-Null

  $diagnosticReportPath = Join-Path $tempRoot 'diagnostic-report.md'
  @'
计算机网络实验报告

课程名称：计算机网络
实验名称：局域网搭建与常用 DOS 命令使用

一、实验目的
通过本次实验掌握局域网的基本搭建方法，并理解 DOS 命令在网络排查中的作用。

四、实验步骤
首先将两台虚拟机配置在同一网段，并通过 ipconfig 与 ping 命令确认网络参数和连通性。

五、实验结果
实验结果表明，两台主机能够正常互通，网络参数符合实验要求。
'@ | Set-Content -LiteralPath $diagnosticReportPath -Encoding UTF8

  $diagnosticTemplateDocx = Join-Path $tempRoot 'field-map-diagnostic-template.docx'
  New-FieldMapDiagnosticTemplateDocx -Path $diagnosticTemplateDocx
  Assert-True -Condition (Test-Path -LiteralPath $diagnosticTemplateDocx) -Message 'Failed to create the field-map diagnostic template fixture.'

  $diagnosticFieldMapPath = Join-Path $tempRoot 'field-map-diagnostic-output.json'
  & (Join-Path $repoRoot 'scripts\generate-docx-field-map.ps1') `
    -TemplatePath $diagnosticTemplateDocx `
    -ReportPath $diagnosticReportPath `
    -MetadataPath (Join-Path $repoRoot 'examples\docx-report-metadata.json') `
    -Format json `
    -OutFile $diagnosticFieldMapPath | Out-Null
  $diagnosticFieldMapRoot = (Get-Content -LiteralPath $diagnosticFieldMapPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ($diagnosticFieldMapRoot.summary.diagnosticCount -ge 4) -Message 'Field-map diagnostics fixture should produce multiple diagnostics.'
  Assert-True -Condition (($diagnosticFieldMapRoot.summary.diagnosticCountsByCode.missing_metadata_value -as [int]) -ge 1) -Message 'Field-map diagnostics should count missing metadata values.'
  Assert-True -Condition (($diagnosticFieldMapRoot.summary.diagnosticCountsByCode.unrecognized_template_metadata_label -as [int]) -ge 1) -Message 'Field-map diagnostics should count unrecognized metadata labels.'
  Assert-True -Condition (($diagnosticFieldMapRoot.summary.diagnosticCountsByCode.missing_report_section -as [int]) -ge 1) -Message 'Field-map diagnostics should count missing report sections.'
  Assert-True -Condition (($diagnosticFieldMapRoot.summary.diagnosticCountsByCode.unrecognized_template_section_heading -as [int]) -ge 1) -Message 'Field-map diagnostics should count unrecognized section headings.'
  Assert-True -Condition (($diagnosticFieldMapRoot.summary.diagnosticCountsByCode.unmatched_composite_template_cell -as [int]) -ge 1) -Message 'Field-map diagnostics should count unmatched composite template cells.'
  Assert-True -Condition ((@($diagnosticFieldMapRoot.diagnostics | Where-Object { $_.code -eq 'missing_metadata_value' -and $_.context.label -eq '实验地点' }).Count) -ge 1) -Message 'Field-map diagnostics should identify the missing 实验地点 metadata value.'
  Assert-True -Condition ((@($diagnosticFieldMapRoot.diagnostics | Where-Object { $_.code -eq 'unrecognized_template_metadata_label' -and $_.context.label -eq '实验台号' }).Count) -ge 1) -Message 'Field-map diagnostics should identify the unrecognized 实验台号 label.'
  Assert-True -Condition ((@($diagnosticFieldMapRoot.diagnostics | Where-Object { $_.code -eq 'missing_report_section' -and $_.context.heading -eq '问题分析' }).Count) -ge 1) -Message 'Field-map diagnostics should identify the missing 问题分析 section.'
  Assert-True -Condition ((@($diagnosticFieldMapRoot.diagnostics | Where-Object { $_.code -eq 'unrecognized_template_section_heading' -and $_.context.heading -eq '实验器材与拓扑' }).Count) -ge 1) -Message 'Field-map diagnostics should identify the unrecognized 实验器材与拓扑 heading.'
  Assert-True -Condition ((@($diagnosticFieldMapRoot.diagnostics | Where-Object { $_.code -eq 'unmatched_composite_template_cell' -and $_.context.location -eq 'T1R1C1' }).Count) -ge 1) -Message 'Field-map diagnostics should identify unmatched composite template cells.'
  Assert-True -Condition ((@($diagnosticFieldMapRoot.notes | Where-Object { $_ -match '实验台号' }).Count) -ge 1) -Message 'Field-map diagnostics should continue mirroring messages into notes.'
  $results.Add('docx field-map diagnostics OK') | Out-Null

  $templateFitCheckPath = Join-Path $tempRoot 'template-fit-check.json'
  & (Join-Path $repoRoot 'scripts\check-report-profile-template-fit.ps1') `
    -TemplatePath $diagnosticTemplateDocx `
    -ReportPath $diagnosticReportPath `
    -MetadataPath (Join-Path $repoRoot 'examples\docx-report-metadata.json') `
    -Format json `
    -OutFile $templateFitCheckPath | Out-Null
  $templateFitCheckRoot = (Get-Content -LiteralPath $templateFitCheckPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$templateFitCheckRoot.reportProfileName -eq 'experiment-report') -Message 'Template-fit checker is missing the expected report profile name.'
  Assert-True -Condition ([int]$templateFitCheckRoot.summary.profileChangeSuggestionCount -ge 3) -Message 'Template-fit checker should suggest profile changes for metadata, sections, and composite rules.'
  Assert-True -Condition ([int]$templateFitCheckRoot.summary.inputGapCount -ge 2) -Message 'Template-fit checker should report both metadata and report-content input gaps.'
  Assert-True -Condition ((@($templateFitCheckRoot.suggestions.metadataFieldsToAdd | Where-Object { $_.label -eq '实验台号' }).Count) -ge 1) -Message 'Template-fit checker should suggest adding the 实验台号 metadata field.'
  Assert-True -Condition ((@($templateFitCheckRoot.inputGaps.missingMetadataValues | Where-Object { $_.label -eq '实验地点' }).Count) -ge 1) -Message 'Template-fit checker should surface the missing 实验地点 metadata value.'
  Assert-True -Condition ((@($templateFitCheckRoot.suggestions.sectionAliasesToAdd | Where-Object { $_.heading -eq '实验器材与拓扑' -and $_.suggestedSectionId -eq 'environment' }).Count) -ge 1) -Message 'Template-fit checker should suggest mapping 实验器材与拓扑 to the environment section.'
  Assert-True -Condition ((@($templateFitCheckRoot.inputGaps.missingReportSections | Where-Object { $_.heading -eq '问题分析' }).Count) -ge 1) -Message 'Template-fit checker should surface the missing 问题分析 report section.'
  Assert-True -Condition ((@($templateFitCheckRoot.suggestions.compositeRulesToAdd | Where-Object { $_.cellText -eq '实验目的 / 实验结果' }).Count) -ge 1) -Message 'Template-fit checker should suggest a composite rule for the unmatched cover/body cell.'
  Assert-True -Condition ((@($templateFitCheckRoot.suggestions.compositeRulesToAdd | Where-Object { $_.cellText -eq '实验目的 / 实验结果' -and @($_.suggestedProfilePatch.matchAll).Count -ge 2 }).Count) -ge 1) -Message 'Template-fit checker should emit a scaffolded matchAll array for composite-rule suggestions.'
  Assert-True -Condition ([string]$templateFitCheckRoot.reportInputMode -eq 'path') -Message 'Template-fit checker should record reportInputMode=path for file-backed reports.'
  Assert-True -Condition ([string]$templateFitCheckRoot.metadataInputMode -eq 'path') -Message 'Template-fit checker should record metadataInputMode=path for metadata files.'

  $inlineDiagnosticReportText = Get-Content -LiteralPath $diagnosticReportPath -Raw -Encoding UTF8
  $inlineDiagnosticMetadataJson = Get-Content -LiteralPath (Join-Path $repoRoot 'examples\docx-report-metadata.json') -Raw -Encoding UTF8
  $inlineTemplateFitRoot = ((& (Join-Path $repoRoot 'scripts\check-report-profile-template-fit.ps1') `
      -TemplatePath $diagnosticTemplateDocx `
      -ReportText $inlineDiagnosticReportText `
      -MetadataJson $inlineDiagnosticMetadataJson `
      -Format json) | Out-String) | ConvertFrom-Json
  Assert-True -Condition ([string]$inlineTemplateFitRoot.reportInputMode -eq 'inline') -Message 'Template-fit checker should record reportInputMode=inline for inline report text.'
  Assert-True -Condition ([string]$inlineTemplateFitRoot.metadataInputMode -eq 'inline') -Message 'Template-fit checker should record metadataInputMode=inline for inline metadata JSON.'
  Assert-True -Condition ([string]$inlineTemplateFitRoot.reportSource -eq '[inline text]') -Message 'Template-fit checker should expose [inline text] as the reportSource for inline report input.'
  Assert-True -Condition ([int]$inlineTemplateFitRoot.summary.profileChangeSuggestionCount -ge 3) -Message 'Template-fit checker should keep profile-change suggestions when using inline inputs.'
  Assert-True -Condition ([int]$inlineTemplateFitRoot.summary.inputGapCount -ge 2) -Message 'Template-fit checker should keep input-gap diagnostics when using inline inputs.'
  $results.Add('docx template-fit inline inputs OK') | Out-Null

  $templateFitMarkdown = & (Join-Path $repoRoot 'scripts\check-report-profile-template-fit.ps1') `
    -TemplatePath $diagnosticTemplateDocx `
    -ReportPath $diagnosticReportPath `
    -MetadataPath (Join-Path $repoRoot 'examples\docx-report-metadata.json') `
    -Format markdown | Out-String
  Assert-True -Condition ($templateFitMarkdown -match 'Report Profile Template Fit') -Message 'Template-fit checker markdown output is missing the title.'
  Assert-True -Condition ($templateFitMarkdown -match 'Profile Changes') -Message 'Template-fit checker markdown output is missing the profile-change section.'
  Assert-True -Condition ($templateFitMarkdown -match '实验台号') -Message 'Template-fit checker markdown output is missing the unrecognized metadata label.'
  Assert-True -Condition ($templateFitMarkdown -match '实验器材与拓扑') -Message 'Template-fit checker markdown output is missing the unrecognized section heading.'
  $results.Add('docx template-fit checker OK') | Out-Null

  $sampleImageOne = Join-Path $tempRoot 'sample-image-1.png'
  $sampleImageTwo = Join-Path $tempRoot 'sample-image-2.png'
  $sampleImageThree = Join-Path $tempRoot 'sample-image-3.png'
  $sampleImageFour = Join-Path $tempRoot 'sample-image-4.png'
  New-SamplePngImage -Path $sampleImageOne -Text 'Step 1'
  New-SamplePngImage -Path $sampleImageTwo -Text 'Result 1' -BackgroundHex '#FCEFD8'
  New-SamplePngImage -Path $sampleImageThree -Text 'Result 2' -BackgroundHex '#FDE7D6'
  New-SamplePngImage -Path $sampleImageFour -Text 'ARP A' -BackgroundHex '#F2E7FE'
  Assert-True -Condition (Test-Path -LiteralPath $sampleImageOne) -Message 'Failed to create the first sample image fixture.'
  Assert-True -Condition (Test-Path -LiteralPath $sampleImageTwo) -Message 'Failed to create the second sample image fixture.'
  Assert-True -Condition (Test-Path -LiteralPath $sampleImageThree) -Message 'Failed to create the third sample image fixture.'
  Assert-True -Condition (Test-Path -LiteralPath $sampleImageFour) -Message 'Failed to create the fourth sample image fixture.'

  $softwareTestReportPath = Join-Path $tempRoot 'software-test-report.md'
  @'
软件测试报告

课程名称：软件测试技术
测试项目：图书管理系统功能测试
学生姓名：赵强
学号：20263456
指导教师：陈老师
测试类型：功能测试
测试时间：2026-04-10
测试环境：Chrome 122 / Windows 11 / MySQL 8.0

一、测试目标
本次软件测试的目标是围绕图书管理系统的登录、图书查询、借阅登记和归还处理等核心功能进行验证，确认系统在常见用户路径下能够稳定返回正确结果。
测试过程中重点关注输入校验、状态流转和异常提示，确保学生用户、管理员用户和无效输入场景都能得到明确反馈。

二、测试环境
测试环境使用 Windows 11、Chrome 122、Node.js 本地服务和 MySQL 8.0 测试库，浏览器缓存会在每轮关键用例执行前清理。
测试账号分为学生账号、管理员账号和锁定账号三类，数据库预置了可借图书、已借出图书和不存在编号等多组数据，便于覆盖正常流程和异常流程。

三、测试范围与依据
测试范围包括账号登录、图书检索、借阅登记、归还确认、库存数量更新和错误提示展示，不覆盖后台统计报表和权限配置等后续迭代功能。
测试依据来自课程给定的需求说明、页面原型和接口字段约定，判断标准是页面展示、接口返回、数据库状态三者保持一致。

四、测试用例设计与执行
用例设计采用等价类和边界值思路，先验证正确账号登录、错误密码提示和空用户名拦截，再验证图书名称关键字搜索、编号精确查询和无结果提示。
借阅流程中，使用学生账号选择一本库存大于零的图书，点击借阅后检查页面提示、借阅记录和库存扣减是否同步更新；归还流程则检查记录状态是否从借阅中变为已归还，并确认库存数量恢复。
针对异常场景，额外执行重复借阅、无库存借阅、已归还记录再次归还和接口断开等用例，观察系统是否给出清晰错误信息，而不是出现白屏或状态混乱。

五、测试结果
本轮测试共执行十八条功能用例，其中十六条通过，两条记录为问题项。登录、查询、正常借阅和正常归还流程均能按照预期完成，页面提示和数据库状态一致。
发现的问题主要集中在无库存借阅时按钮未及时禁用，以及接口超时后页面仍保留上一次成功提示。两个问题不会阻断主流程，但会影响用户对当前操作结果的判断。

六、缺陷分析与改进
无库存借阅问题的原因是前端只在页面首次加载时读取库存状态，借阅按钮没有在库存字段变化后重新计算禁用条件，导致边界数据下仍然可以触发提交动作。
接口超时提示问题则来自请求异常分支没有清理旧状态，建议在进入新请求前统一重置提示区域，并在失败分支加入明确的重试说明。后续还应补充并发借阅和弱网环境测试，降低状态不同步风险。

七、测试总结
通过本次测试，可以确认图书管理系统的核心业务路径已经基本可用，登录、查询、借阅和归还流程具备课程演示所需的完整性。
测试也暴露出前端状态刷新和异常提示处理仍有细节不足，后续改进应优先围绕边界库存、接口失败和重复提交三个方向补充自动化回归用例。
'@ | Set-Content -LiteralPath $softwareTestReportPath -Encoding UTF8

  $softwareTestValidation = (& (Join-Path $repoRoot 'scripts\validate-report-draft.ps1') -Path $softwareTestReportPath -ReportProfileName 'software-test-report' -Format json | Out-String) | ConvertFrom-Json
  Assert-True -Condition ([bool]$softwareTestValidation.passed) -Message 'Software-test report validation should pass for the software-test profile.'
  Assert-True -Condition ([string]$softwareTestValidation.reportProfileName -eq 'software-test-report') -Message 'Software-test report validation is missing the expected profile name.'
  Assert-True -Condition ([int]$softwareTestValidation.summary.sectionCount -ge 7) -Message 'Software-test report validation did not load the expected section rules.'
  Assert-True -Condition ([string]$softwareTestValidation.summary.paginationRiskRemediations.'pagination-risk-dense-section-block' -match 'expected results, actual results') -Message 'Software-test validation summary should expose the custom dense-section remediation.'

  $softwareTestRiskDenseBlock = ((@('本轮测试把登录、查询、借阅、归还、库存边界、接口异常、重复提交、弱网提示、数据库状态核对和页面反馈串成同一组回归证据。') * 32) -join '') + '见图1、图2、图3、图4、图5。'
  $softwareTestRiskReportPath = Join-Path $tempRoot 'software-test-report-pagination-risk.md'
  @(
    '软件测试报告',
    '',
    '课程名称：软件测试技术',
    '测试项目：图书管理系统功能测试',
    '学生姓名：赵强',
    '学号：20263456',
    '指导教师：陈老师',
    '测试类型：功能测试',
    '测试时间：2026-04-10',
    '测试环境：Chrome 122 / Windows 11 / MySQL 8.0',
    '',
    '一、测试目标',
    '测试目标是验证图书管理系统的登录、查询、借阅和归还主流程是否满足课程验收要求，并观察异常输入下的提示是否清晰。',
    '',
    '二、测试环境',
    '测试环境使用 Windows 11、Chrome 浏览器、本地 Node.js 服务和 MySQL 测试库，测试前会清理浏览器缓存并重置样例数据。',
    '',
    '三、测试范围与依据',
    '测试范围覆盖账号登录、图书检索、借阅登记、归还确认、库存更新和错误提示展示，依据来自需求说明、页面原型和接口字段约定。',
    '',
    '四、测试用例设计与执行',
    $softwareTestRiskDenseBlock,
    '',
    '五、测试结果',
    '本轮测试执行后，登录、查询、正常借阅和正常归还流程均可通过，异常输入场景能够返回提示，测试结果具备继续修复、复测、回归统计和课堂验收说明的依据。',
    '',
    '六、缺陷分析与改进',
    '当前缺陷主要集中在库存边界和接口超时提示上，建议把前置条件、输入数据、预期结果、实际结果和缺陷备注分开记录，避免测试证据挤在同一段中。',
    '',
    '七、测试总结',
    '本次测试说明核心业务流程已经可用，后续应继续补充弱网、并发借阅和重复提交等回归用例，并把截图证据按用例组分散到对应章节。'
  ) | Set-Content -LiteralPath $softwareTestRiskReportPath -Encoding UTF8
  $softwareTestRiskValidation = (& (Join-Path $repoRoot 'scripts\validate-report-draft.ps1') -Path $softwareTestRiskReportPath -ReportProfileName 'software-test-report' -Format json | Out-String) | ConvertFrom-Json
  $softwareTestRiskWarningCodes = @($softwareTestRiskValidation.summary.warningCodes | ForEach-Object { [string]$_ })
  $softwareTestRiskLongFinding = @($softwareTestRiskValidation.findings | Where-Object { [string]$_.code -eq 'pagination-risk-long-section' })[0]
  Assert-True -Condition ([bool]$softwareTestRiskValidation.passed) -Message 'Software-test pagination-risk fixture should pass with warnings only.'
  Assert-True -Condition ([int]$softwareTestRiskValidation.summary.paginationRiskCount -ge 3) -Message 'Software-test pagination-risk fixture should surface profile warnings.'
  Assert-True -Condition ($softwareTestRiskWarningCodes -contains 'pagination-risk-long-section') -Message 'Software-test pagination-risk fixture should report pagination-risk-long-section.'
  Assert-True -Condition ($softwareTestRiskWarningCodes -contains 'pagination-risk-dense-section-block') -Message 'Software-test pagination-risk fixture should report pagination-risk-dense-section-block.'
  Assert-True -Condition ($softwareTestRiskWarningCodes -contains 'pagination-risk-figure-cluster') -Message 'Software-test pagination-risk fixture should report pagination-risk-figure-cluster.'
  Assert-True -Condition ([string]$softwareTestRiskValidation.summary.paginationRiskRemediations.'pagination-risk-long-section' -match 'case group') -Message 'Software-test validation summary should expose the custom long-section remediation.'
  Assert-True -Condition ([string]$softwareTestRiskValidation.summary.paginationRiskRemediations.'pagination-risk-figure-cluster' -match '缺陷分析与改进') -Message 'Software-test validation summary should expose the custom figure-cluster remediation.'
  Assert-True -Condition ([string]$softwareTestRiskLongFinding.remediation -match 'case group') -Message 'Software-test validation finding should use the custom long-section remediation.'
  $results.Add('software-test profile validation OK') | Out-Null

  $deploymentReportPath = Join-Path $tempRoot 'deployment-report.md'
  @'
部署运维报告

课程名称：云平台运维实践
部署项目：校园门户系统容器化部署
学生姓名：刘洋
学号：20264567
指导教师：孙老师
部署类型：系统部署
部署时间：2026-04-12
部署环境：Ubuntu 22.04 / Docker 26 / Nginx 1.24

一、部署目标
本次部署任务的目标是将校园门户系统从本地开发环境迁移到 Linux 服务器容器环境中运行，使前端静态页面、后端接口服务和数据库连接能够形成完整访问链路。
部署完成后需要能够通过浏览器访问门户首页、登录接口和健康检查地址，并保证服务重启后仍然可以按预期恢复运行。

二、部署环境
部署环境使用 Ubuntu 22.04 服务器，基础组件包括 Docker 26、Docker Compose、Nginx 1.24、MySQL 8.0 和 Node.js 运行镜像。
服务器开放 80、443 和后端内部服务端口，部署前已经完成防火墙规则检查、镜像仓库登录、项目配置文件整理和数据库初始化脚本准备。

三、部署方案与架构
整体方案采用 Nginx 反向代理加多容器编排结构，前端构建产物由 Nginx 容器提供静态访问，后端服务以应用容器形式运行，数据库使用独立 MySQL 容器保存业务数据。
配置层面将端口、数据库连接、日志目录和上传目录拆分到环境变量与挂载卷中，减少后续重新发布时对镜像内容的直接修改。

四、部署步骤与配置
部署前先拉取项目代码并确认分支版本，然后执行前端构建命令生成静态资源，再构建后端应用镜像并检查镜像标签是否与发布版本一致。
随后编写 docker-compose 配置文件，分别声明前端、后端和数据库服务的镜像、端口、环境变量、依赖关系和数据卷挂载路径，确保应用启动顺序能够满足数据库先就绪、接口再连接的要求。
启动服务后，通过 docker ps 查看容器状态，通过 docker logs 检查后端启动日志，并在 Nginx 配置中补充静态资源缓存、接口代理和超时参数，最后重新加载 Nginx 使配置生效。

五、验证结果
部署完成后，浏览器能够正常访问校园门户首页，登录接口返回成功状态，健康检查地址显示服务正常，数据库中也可以看到初始化后的用户和菜单数据。
通过 curl 请求后端健康检查接口返回 200 状态码，查看容器日志没有持续报错，重启后端容器后服务可以在短时间内恢复访问，说明本次部署满足基本运行和恢复要求。

六、问题处理与回滚预案
部署过程中出现过后端容器首次启动无法连接数据库的问题，排查后发现数据库容器虽然已启动但初始化尚未完成，因此在 compose 配置中补充健康检查和重试等待逻辑。
如果后续发布版本出现严重异常，可以先切回上一版镜像标签并重新执行 docker compose up，同时保留当前数据库卷不变；若数据库结构变更导致问题，则优先使用发布前导出的备份文件恢复。

七、部署总结
通过本次部署任务，完整梳理了从构建镜像、编写 compose 配置、配置 Nginx 代理到验证服务状态的流程，理解了环境变量、数据卷和日志检查在运维工作中的作用。
本次系统已经能够稳定完成基础访问和健康检查，但后续仍需补充自动化发布脚本、HTTPS 证书续期提醒和更细粒度的监控告警，以提高实际运维场景下的可靠性。
'@ | Set-Content -LiteralPath $deploymentReportPath -Encoding UTF8

  $deploymentValidation = (& (Join-Path $repoRoot 'scripts\validate-report-draft.ps1') -Path $deploymentReportPath -ReportProfileName 'deployment-report' -Format json | Out-String) | ConvertFrom-Json
  Assert-True -Condition ([bool]$deploymentValidation.passed) -Message 'Deployment report validation should pass for the deployment profile.'
  Assert-True -Condition ([string]$deploymentValidation.reportProfileName -eq 'deployment-report') -Message 'Deployment report validation is missing the expected profile name.'
  Assert-True -Condition ([int]$deploymentValidation.summary.sectionCount -ge 7) -Message 'Deployment report validation did not load the expected section rules.'
  Assert-True -Condition ([string]$deploymentValidation.summary.paginationRiskRemediations.'pagination-risk-dense-section-block' -match 'commands, configuration snippets') -Message 'Deployment validation summary should expose the custom dense-section remediation.'

  $deploymentRiskDenseBlock = ((@('部署过程把镜像构建、环境变量、端口映射、数据卷挂载、Nginx 代理、服务启动、日志检查、健康检查、回滚标签和恢复验证串成同一条发布证据链。') * 30) -join '') + '见图1、图2、图3、图4、图5。'
  $deploymentRiskReportPath = Join-Path $tempRoot 'deployment-report-pagination-risk.md'
  @(
    '部署运维报告',
    '',
    '课程名称：云平台运维实践',
    '部署项目：校园门户系统容器化部署',
    '学生姓名：刘洋',
    '学号：20264567',
    '指导教师：孙老师',
    '部署类型：系统部署',
    '部署时间：2026-04-12',
    '部署环境：Ubuntu 22.04 / Docker 26 / Nginx 1.24',
    '',
    '一、部署目标',
    '部署目标是把校园门户系统迁移到容器化运行环境中，使前端、后端和数据库能够形成可访问、可验证、可回滚的完整链路。',
    '',
    '二、部署环境',
    '部署环境包括 Ubuntu 服务器、Docker Compose、Nginx、MySQL 和后端应用镜像，部署前已检查端口、防火墙、环境变量和初始化脚本。',
    '',
    '三、部署方案与架构',
    '部署方案采用 Nginx 反向代理加多容器编排，前端、后端和数据库通过服务名互相访问，并通过环境变量和挂载卷降低重发版成本。',
    '',
    '四、部署步骤与配置',
    $deploymentRiskDenseBlock,
    '',
    '五、验证结果',
    '验证结果显示门户首页、登录接口和健康检查地址均可访问，容器日志没有持续报错，后端服务重启后能够在短时间内恢复，并能继续支撑课堂演示和发布复盘。',
    '',
    '六、问题处理与回滚预案',
    '当前风险集中在数据库初始化等待和发布回滚路径上，建议把命令、配置片段、日志观察和操作解释分开记录，并把截图分散到验证和回滚章节。',
    '',
    '七、部署总结',
    '本次部署梳理了镜像构建、compose 编排、Nginx 代理、健康检查和回滚预案，后续应继续补充监控告警、证书续期和自动化发布脚本。'
  ) | Set-Content -LiteralPath $deploymentRiskReportPath -Encoding UTF8
  $deploymentRiskValidation = (& (Join-Path $repoRoot 'scripts\validate-report-draft.ps1') -Path $deploymentRiskReportPath -ReportProfileName 'deployment-report' -Format json | Out-String) | ConvertFrom-Json
  $deploymentRiskWarningCodes = @($deploymentRiskValidation.summary.warningCodes | ForEach-Object { [string]$_ })
  $deploymentRiskLongFinding = @($deploymentRiskValidation.findings | Where-Object { [string]$_.code -eq 'pagination-risk-long-section' })[0]
  Assert-True -Condition ([bool]$deploymentRiskValidation.passed) -Message 'Deployment pagination-risk fixture should pass with warnings only.'
  Assert-True -Condition ([int]$deploymentRiskValidation.summary.paginationRiskCount -ge 3) -Message 'Deployment pagination-risk fixture should surface profile warnings.'
  Assert-True -Condition ($deploymentRiskWarningCodes -contains 'pagination-risk-long-section') -Message 'Deployment pagination-risk fixture should report pagination-risk-long-section.'
  Assert-True -Condition ($deploymentRiskWarningCodes -contains 'pagination-risk-dense-section-block') -Message 'Deployment pagination-risk fixture should report pagination-risk-dense-section-block.'
  Assert-True -Condition ($deploymentRiskWarningCodes -contains 'pagination-risk-figure-cluster') -Message 'Deployment pagination-risk fixture should report pagination-risk-figure-cluster.'
  Assert-True -Condition ([string]$deploymentRiskValidation.summary.paginationRiskRemediations.'pagination-risk-long-section' -match 'rollback subsections') -Message 'Deployment validation summary should expose the custom long-section remediation.'
  Assert-True -Condition ([string]$deploymentRiskValidation.summary.paginationRiskRemediations.'pagination-risk-figure-cluster' -match '问题处理与回滚预案') -Message 'Deployment validation summary should expose the custom figure-cluster remediation.'
  Assert-True -Condition ([string]$deploymentRiskLongFinding.remediation -match 'rollback subsections') -Message 'Deployment validation finding should use the custom long-section remediation.'
  $results.Add('deployment profile validation OK') | Out-Null

  $weeklyReportPath = Join-Path $tempRoot 'weekly-report.md'
  @'
项目周报

项目名称：校园导览小程序
周报主题：第 6 周迭代周报
提交人：李四
工号/学号：20261234
审核人：王老师
报告类型：项目周报
周次：第 6 周
工作环境：GitHub + 飞书 + 本地开发环境

一、工作目标
本周工作目标是把校园导览小程序从页面原型推进到可演示版本，重点完成首页分类、地点检索、详情展示和基础路线提示四个关键路径。
为了让周报能够服务后续复盘，团队还要求每项进展都对应明确输入、执行动作和可检查产出，而不是只记录笼统的完成状态。

二、协作环境
本周协作环境以 GitHub 仓库、飞书任务看板和本地微信开发者工具为主，需求拆分、接口字段、页面截图和问题记录都集中到同一迭代空间中维护。
开发过程中前端使用本地 mock 数据先行验证交互，后端接口按地点分类、关键词检索和详情查询三个方向分批补齐，保证每天都能形成可运行的小版本。

三、本周任务与输入
本周输入包括课程设计需求说明、上周确定的页面草图、地点基础数据表和老师对演示流程提出的反馈。任务拆分时先把必须演示的用户路径列为最高优先级，再处理页面细节和异常提示。
团队把工作拆成首页信息架构、搜索联想、详情页字段、收藏状态和路线提示五组任务，并约定每个任务完成后都需要提交截图、提交记录和一段简短说明，便于周末集中验收。

四、本周完成事项
本周已经完成首页分类卡片、地点列表、关键词搜索和地点详情页的主要交互。首页能够按教学楼、生活服务、实验室和公共设施分类展示地点，搜索页可以根据关键字返回匹配结果，详情页能展示开放时间、位置说明和注意事项。
后端接口完成了地点列表和详情查询的 mock 到本地服务切换，前端也补充了加载中、无结果和请求失败三种状态。为了保证演示稳定，还增加了基础缓存策略，避免重复进入页面时出现明显闪烁。
除功能开发外，本周还整理了测试数据和演示脚本，把关键页面截图归档到迭代目录中，并在飞书看板上标记了已完成、待复核和延期处理三类任务状态。

五、阶段成果
阶段成果是形成了一个可以端到端演示的校园导览小程序版本。用户可以从首页选择分类，进入列表后查看地点信息，也可以通过搜索直接定位教学楼或实验室，整体路径已经能够覆盖课程设计展示的核心要求。
从验收角度看，本周交付物包括可运行前端页面、本地接口服务、地点数据样例、页面截图和任务状态记录。当前版本虽然还没有接入真实地图服务，但已经能够证明信息组织、查询逻辑和详情展示方案可行。

六、风险与改进
当前主要风险在于地点数据仍然偏少，搜索排序规则也还比较简单，后续如果数据量增加，可能出现同名地点排序不稳定或无关地点靠前的问题。
另一个风险是路线提示还停留在静态文本阶段，没有和地图组件形成联动。下周需要优先验证地图组件接入成本，并补充异常情况下的提示文案，避免演示时因为接口延迟或数据缺失影响观感。

七、下周计划
下周计划首先补齐地图组件接入和路线展示，把详情页中的位置信息与可视化地图联动起来。随后继续完善地点数据、收藏功能和搜索排序，使用户能够更快找到常用地点。
在交付准备方面，下周还需要完成一次完整演示彩排，记录操作步骤、页面截图和已知问题，并根据老师反馈决定是否压缩功能范围，保证最终提交版本稳定可讲解。
'@ | Set-Content -LiteralPath $weeklyReportPath -Encoding UTF8

  $weeklyValidation = (& (Join-Path $repoRoot 'scripts\validate-report-draft.ps1') -Path $weeklyReportPath -ReportProfileName 'weekly-report' -Format json | Out-String) | ConvertFrom-Json
  Assert-True -Condition ([bool]$weeklyValidation.passed) -Message 'Weekly report validation should pass for the weekly-report profile.'
  Assert-True -Condition ([string]$weeklyValidation.reportProfileName -eq 'weekly-report') -Message 'Weekly report validation is missing the expected profile name.'
  Assert-True -Condition ([int]$weeklyValidation.summary.sectionCount -ge 7) -Message 'Weekly report validation did not load the expected section rules.'
  Assert-True -Condition ([int]$weeklyValidation.summary.paginationRiskThresholds.longSectionChars -eq 1200) -Message 'Weekly report validation did not use the weekly pagination thresholds.'
  $results.Add('weekly-report profile validation OK') | Out-Null

  $meetingMinutesPath = Join-Path $tempRoot 'meeting-minutes.md'
  @'
会议纪要

项目名称：校园导览小程序
会议主题：第 6 周迭代评审会
记录人：李四
工号/学号：20261234
参会小组：软工 2302
主持人：王老师
纪要类型：评审会议纪要
会议日期：2026-04-13
会议地点：GitHub + 飞书会议

一、会议目标
本次会议目标是对校园导览小程序第 6 周迭代的阶段成果进行集中评审，确认哪些功能已经达到可演示标准，哪些问题需要在下周继续推进。
会议同时要求把功能完成情况、风险判断和后续责任人说清楚，避免纪要只停留在泛泛而谈的过程记录上。

二、参会信息与背景
本次评审会由课程指导老师和项目小组核心成员共同参加，会议前已经提前共享了最新页面截图、演示脚本、任务看板和已知问题清单。
参会背景是当前版本已经具备首页分类、地点搜索和详情展示的基本能力，但地图接入、排序优化和演示稳定性还需要进一步确认取舍。

三、议题与输入
本次会议输入包括本周开发提交记录、关键页面截图、已完成任务列表、未关闭问题单以及老师对答辩展示顺序提出的修改建议。
会议议题主要围绕四个方面展开，分别是当前版本能否满足演示要求、路线提示是否需要继续开发、地图组件接入成本是否可控，以及下周任务应该如何重新排序。

四、讨论过程与决议
会议先回顾了本周已完成的页面与接口联调结果，确认首页分类、关键词搜索和详情页信息展示已经达到可演示水平。随后针对路线提示和地图接入进行讨论，大家一致认为当前静态文本方案虽然能说明设计思路，但若完全不接地图组件，最终展示说服力仍然不足。
在资源评估部分，小组成员说明地图组件接入主要难点在于地点坐标整理和详情页跳转联动，而不是基础显示能力本身。主持人最终决定下周优先完成地图接入验证和一条完整路线展示链路，同时压缩非关键视觉细节，确保最终版本先满足完整演示。

五、当前结论
会议确认当前版本可以作为阶段性可演示成果继续推进，不需要推翻现有页面结构和数据组织方案。首页分类、地点列表、详情页和搜索流程将保持现有实现，只针对交互细节和演示顺序做小范围优化。
同时会议明确了下一轮迭代的核心结论，即路线展示与地图联动是当前最值得投入的能力，收藏功能和复杂排序策略暂时不作为下周必须完成项。

六、风险与争议
当前主要风险在于地图组件接入后可能暴露出地点坐标不完整、详情页跳转逻辑混乱和页面加载变慢等问题。如果这些问题集中出现，下周计划就需要重新压缩范围。
会议中的主要争议点是是否继续保留静态路线文字作为备用方案。一部分成员认为备用方案能降低演示风险，另一部分成员认为双方案并行会分散开发精力，最后决定先保留静态说明，但不再投入额外美化工作。

七、后续安排
后续安排是先由前端负责人在下周前两天完成地图组件接入验证，并输出一份包含截图和性能观察的记录；后端同步整理地点坐标和分类字段，保证前端联调时有稳定数据来源。
记录人会在本次会议后更新飞书任务看板，把地图验证、演示彩排、问题回归和答辩讲解稿拆成明确任务，并在下一次周会前汇总完成情况和残留阻塞项。
'@ | Set-Content -LiteralPath $meetingMinutesPath -Encoding UTF8

  $meetingValidation = (& (Join-Path $repoRoot 'scripts\validate-report-draft.ps1') -Path $meetingMinutesPath -ReportProfileName 'meeting-minutes' -Format json | Out-String) | ConvertFrom-Json
  Assert-True -Condition ([bool]$meetingValidation.passed) -Message 'Meeting minutes validation should pass for the meeting-minutes profile.'
  Assert-True -Condition ([string]$meetingValidation.reportProfileName -eq 'meeting-minutes') -Message 'Meeting minutes validation is missing the expected profile name.'
  Assert-True -Condition ([int]$meetingValidation.summary.sectionCount -ge 7) -Message 'Meeting minutes validation did not load the expected section rules.'
  Assert-True -Condition ([int]$meetingValidation.summary.paginationRiskThresholds.longSectionChars -eq 1000) -Message 'Meeting minutes validation did not use the meeting pagination thresholds.'
  $results.Add('meeting-minutes profile validation OK') | Out-Null

  $monthlyReportPath = Join-Path $tempRoot 'monthly-report.md'
  @'
项目月报

项目名称：校园导览小程序
月报主题：2026 年 4 月项目月报
提交人：李四
工号/学号：20261234
小组/班级：软工 2302
审核人：王老师
报告类型：项目月报
月份周期：2026-04
工作环境：GitHub + 飞书 + 本地开发环境

一、本月目标
本月目标是把校园导览小程序从可演示的页面集合推进成一个具备连续使用价值的小程序版本，重点围绕地点分类、关键词搜索、详情页信息完整性和地图接入准备四条主线展开。
在目标拆解时，团队不仅关注功能是否完成，还要求每一项月度工作都能形成可复盘的输入、执行过程、结果证据和后续决策依据，从而让月报能够同时服务教学验收和团队内部管理。

二、协作环境
本月协作环境以 GitHub 仓库、飞书任务看板、接口文档和本地微信开发者工具为主，页面截图、任务状态、接口字段定义和演示脚本都统一放在同一套协作空间中维护。
为了提高联调效率，本月前两周继续采用 mock 数据驱动前端开发，后两周逐步切换到本地接口服务，并通过固定的提测时间和问题清单同步机制减少重复沟通成本。

三、本月任务与输入
本月输入主要包括课程设计任务书、老师在阶段检查中提出的演示建议、地点基础数据表、页面原型图以及上个月遗留的搜索与导航问题清单。团队按照演示必要性和实现成本把任务分成必须完成、建议完成和可延后观察三类。
在执行节奏上，项目先完成首页信息架构和地点详情字段整理，再推进搜索、分类筛选和页面状态处理，最后把地图接入的技术评估、路线提示方案和演示脚本整合到同一个月度里程碑中。

四、本月完成事项
本月已经完成首页分类卡片、地点列表、关键词搜索、详情页信息展示、加载与空状态处理，以及一套适合课堂答辩的演示流程。首页能够按教学楼、实验室、生活服务和公共设施分类展示地点，搜索页可以根据关键字快速返回结果，详情页也补齐了开放时间、位置说明和注意事项等字段。
在工程推进方面，前端完成了页面结构重构和组件拆分，后端把地点列表、关键词搜索和详情查询接口从纯 mock 方案迁移到本地服务，保证了联调路径更加接近真实环境。团队还补充了页面截图归档、任务状态同步和问题回归机制，使每周输出都能自然汇总到本月报告中。
除功能开发外，本月还集中清理了若干影响演示稳定性的细节问题，例如重复进入页面时的闪烁、无结果状态缺少提示、以及某些地点信息字段展示不完整等。通过这些整理，当前版本已经具备更稳定的展示基础，而不是仅能勉强跑通单一路径的原型。

五、阶段成果与数据
阶段成果是形成了一套可以端到端演示的校园导览小程序版本，并积累了相对完整的页面截图、任务记录、接口字段和问题清单。当前版本已经能够支持分类浏览、关键词搜索和地点详情查看三条核心路径，说明整体信息组织与交互方案是成立的。
从月度数据角度看，本月完成了主要页面的结构定稿、三类核心接口的本地联调、若干关键截图的归档和一轮完整演示脚本的整理。虽然地图组件尚未正式接入，但已有的技术评估结果和路线提示草稿已经为下一阶段推进提供了相对清晰的投入边界和实施顺序。

六、问题与改进
当前主要问题在于地点数据规模仍然偏小，搜索排序规则还不够稳定，随着数据量上升，可能出现同名地点排序不合理或弱相关结果靠前的问题。另一个问题是地图组件虽然完成了预研，但尚未进入正式联调，导致详情页到路线展示之间仍存在体验断层。
针对这些问题，团队计划下月优先推进地图接入验证，并同步补齐地点坐标、收藏状态和搜索排序策略。与此同时，还需要进一步收敛演示范围，避免为了追求功能数量而牺牲月度版本的稳定性和可讲解性。

七、下月计划
下月计划首先完成地图组件接入和一条完整路线展示链路，把详情页中的位置信息与可视化地图联动起来，形成更完整的导览体验。随后继续完善地点数据、收藏功能和搜索排序，提升用户在真实校园场景下的查找效率。
在交付准备方面，下月还需要组织一次完整的演示彩排，记录操作顺序、页面截图、已知问题和讲解重点，并根据老师反馈决定是否继续扩功能或转入稳定性打磨，确保最终提交版本既可演示又便于答辩说明。
'@ | Set-Content -LiteralPath $monthlyReportPath -Encoding UTF8

  $monthlyValidation = (& (Join-Path $repoRoot 'scripts\validate-report-draft.ps1') -Path $monthlyReportPath -ReportProfileName 'monthly-report' -Format json | Out-String) | ConvertFrom-Json
  Assert-True -Condition ([bool]$monthlyValidation.passed) -Message 'Monthly report validation should pass for the monthly-report profile.'
  Assert-True -Condition ([string]$monthlyValidation.reportProfileName -eq 'monthly-report') -Message 'Monthly report validation is missing the expected profile name.'
  Assert-True -Condition ([int]$monthlyValidation.summary.sectionCount -ge 7) -Message 'Monthly report validation did not load the expected section rules.'
  Assert-True -Condition ([int]$monthlyValidation.summary.paginationRiskThresholds.longSectionChars -eq 1600) -Message 'Monthly report validation did not use the monthly pagination thresholds.'

  $monthlyRiskDenseBlock = ((@('本月围绕校园导览小程序的里程碑拆解、执行动作、交付证据、风险处理和协作记录持续滚动复盘。') * 40) -join '') + '见图1、图2、图3、图4、图5、图6。'
  $monthlyRiskReportPath = Join-Path $tempRoot 'monthly-report-pagination-risk.md'
  @(
    '项目月报',
    '',
    '项目名称：校园导览小程序',
    '月报主题：2026 年 4 月分页风险检查',
    '提交人：李四',
    '工号/学号：20261234',
    '小组/班级：软工 2302',
    '审核人：王老师',
    '报告类型：项目月报',
    '月份周期：2026-04',
    '工作环境：GitHub + 飞书 + 本地开发环境',
    '',
    '一、本月目标',
    '本月目标聚焦首页结构调整、搜索体验收敛、详情页字段补全和地图接入预研，确保项目既可演示也可持续迭代。',
    '',
    '二、协作环境',
    '协作环境覆盖 GitHub、飞书任务看板、接口文档和本地开发工具，保证需求、实现与验证证据处于同一条追踪链路。',
    '',
    '三、本月任务与输入',
    '本月任务输入来自课程要求、老师阶段反馈、地点基础数据和上月遗留问题清单，并按必须完成、建议完成和延后观察三类优先级推进。',
    '',
    '四、本月完成事项',
    '本月已完成首页卡片、地点筛选、详情页信息补全、搜索交互修正和答辩演示流程整理，开发推进路径基本稳定。与此同时，团队还把任务拆分、页面回归、问题归档和阶段演示材料串成同一条执行链路，保证月度完成事项既有动作描述也有证据沉淀。',
    '',
    '五、阶段成果与数据',
    $monthlyRiskDenseBlock,
    '',
    '六、问题与改进',
    '当前主要问题仍集中在地图联调尚未完成、地点数据规模偏小和搜索排序规则还需要继续收敛。下一步需要把风险按技术、数据和演示三类拆开处理，避免所有问题继续堆叠在同一轮月报里。',
    '',
    '七、下月计划',
    '下月计划优先完成地图接入验证，再补齐地点坐标、收藏状态与演示彩排材料。同时会把搜索排序、路线提示和答辩讲解脚本拆成可单独验收的小目标，降低后续迭代中的联动风险。'
  ) | Set-Content -LiteralPath $monthlyRiskReportPath -Encoding UTF8
  $monthlyRiskValidation = (& (Join-Path $repoRoot 'scripts\validate-report-draft.ps1') -Path $monthlyRiskReportPath -ReportProfileName 'monthly-report' -Format json | Out-String) | ConvertFrom-Json
  $monthlyRiskWarningCodes = @($monthlyRiskValidation.summary.warningCodes | ForEach-Object { [string]$_ })
  $monthlyRiskLongFinding = @($monthlyRiskValidation.findings | Where-Object { [string]$_.code -eq 'pagination-risk-long-section' })[0]
  Assert-True -Condition ([bool]$monthlyRiskValidation.passed) -Message 'Monthly pagination-risk fixture should pass with warnings only.'
  Assert-True -Condition ([int]$monthlyRiskValidation.summary.paginationRiskCount -ge 3) -Message 'Monthly pagination-risk fixture should surface monthly profile warnings.'
  Assert-True -Condition ($monthlyRiskWarningCodes -contains 'pagination-risk-long-section') -Message 'Monthly pagination-risk fixture should report pagination-risk-long-section.'
  Assert-True -Condition ($monthlyRiskWarningCodes -contains 'pagination-risk-dense-section-block') -Message 'Monthly pagination-risk fixture should report pagination-risk-dense-section-block.'
  Assert-True -Condition ($monthlyRiskWarningCodes -contains 'pagination-risk-figure-cluster') -Message 'Monthly pagination-risk fixture should report pagination-risk-figure-cluster.'
  Assert-True -Condition ([string]$monthlyRiskValidation.summary.paginationRiskRemediations.'pagination-risk-long-section' -match 'week-by-week') -Message 'Monthly validation summary should expose the custom long-section remediation.'
  Assert-True -Condition ([string]$monthlyRiskValidation.summary.paginationRiskRemediations.'pagination-risk-figure-cluster' -match '阶段成果与数据') -Message 'Monthly validation summary should expose the custom figure-cluster remediation.'
  Assert-True -Condition ([string]$monthlyRiskLongFinding.remediation -match 'week-by-week') -Message 'Monthly validation finding should use the custom long-section remediation.'
  $results.Add('monthly-report profile validation OK') | Out-Null

  $monthlyReportTemplateDocx = Join-Path $tempRoot 'monthly-report-template.docx'
  New-MonthlyReportTemplateDocx -Path $monthlyReportTemplateDocx
  Assert-True -Condition (Test-Path -LiteralPath $monthlyReportTemplateDocx) -Message 'Failed to create the monthly-report template fixture.'

  $meetingMinutesTemplateDocx = Join-Path $tempRoot 'meeting-minutes-template.docx'
  New-MeetingMinutesTemplateDocx -Path $meetingMinutesTemplateDocx
  Assert-True -Condition (Test-Path -LiteralPath $meetingMinutesTemplateDocx) -Message 'Failed to create the meeting-minutes template fixture.'

  $weeklyReportTemplateDocx = Join-Path $tempRoot 'weekly-report-template.docx'
  New-WeeklyReportTemplateDocx -Path $weeklyReportTemplateDocx
  Assert-True -Condition (Test-Path -LiteralPath $weeklyReportTemplateDocx) -Message 'Failed to create the weekly-report template fixture.'

  $softwareTestTemplateDocx = Join-Path $tempRoot 'software-test-template.docx'
  New-SoftwareTestTemplateDocx -Path $softwareTestTemplateDocx
  Assert-True -Condition (Test-Path -LiteralPath $softwareTestTemplateDocx) -Message 'Failed to create the software-test-report template fixture.'

  $deploymentTemplateDocx = Join-Path $tempRoot 'deployment-template.docx'
  New-DeploymentTemplateDocx -Path $deploymentTemplateDocx
  Assert-True -Condition (Test-Path -LiteralPath $deploymentTemplateDocx) -Message 'Failed to create the deployment-report template fixture.'

  $courseDesignReportPath = Join-Path $tempRoot 'course-design-report.md'
  @'
软件工程课程设计报告

课程名称：软件工程综合实践
课题名称：校园导览小程序设计
学生姓名：李四
学号：20261234
指导老师：王老师
完成时间：2026-04-08
设计地点：实验楼 A201

一、设计目标
本次课程设计的目标是完成一个面向校园访客和学生的导览小程序，使用户能够快速查看教学楼、实验室和生活服务点的位置分布。
除了完成基础的地图展示功能，还需要在交互流程中突出搜索、路线提示和常用地点收藏等核心能力，保证项目具备完整的演示价值。

二、开发环境
项目开发使用 Windows 11、Node.js、微信开发者工具和 SQLite 作为本地调试环境，前端页面采用小程序原生组件实现。
为了方便联调与演示，后端接口在本机启动测试服务，并通过模拟数据覆盖地点检索、分类筛选和详情展示等典型场景。

三、需求分析
系统需要支持地点分类浏览、关键字搜索、地点详情查看和推荐路线提示，保证新生在不熟悉校园环境时也能快速定位目标区域。
在分析过程中重点梳理了教学区、宿舍区和公共服务区三类地点信息结构，并明确了页面响应速度和信息准确性两项核心约束。

四、方案设计与实现
整体方案采用前后端分层结构，前端负责地点列表、搜索页和详情页展示，后端负责地点数据组织、关键词过滤和路线推荐结果返回。
在实现阶段先完成地点数据模型和接口约定，再逐步补齐首页分类卡片、搜索联想、详情页信息模块和收藏状态管理逻辑。
为了让演示效果更加稳定，还为主要页面增加了空状态提示、加载占位和异常请求兜底提示，避免因为数据延迟导致界面体验不完整。

五、运行结果
系统启动后可以正常展示校园地点分类首页，输入教学楼关键字后能够即时返回匹配结果，并支持点击进入地点详情页查看开放时间和相关说明。
在演示测试中，推荐路线和收藏功能都能按照预期更新界面状态，整体流程从搜索到查看结果再到返回首页保持稳定，没有出现明显的页面跳转错误。

六、问题与改进
当前版本在地点数据量进一步增大时，搜索结果排序仍然偏向简单匹配规则，缺少结合距离和使用频率的综合排序能力。
后续可以引入更细致的标签体系和缓存策略，同时补充地图组件联动能力，使路线展示、地点筛选和结果高亮之间形成更自然的交互闭环。

七、设计总结
通过这次课程设计，进一步理解了从需求分析、页面拆分到接口联调的完整实现流程，也明确了前后端边界划分对项目稳定性的影响。
项目从可运行原型逐步完善到可演示成品的过程中，最大的收获是学会了围绕用户任务链路组织设计重点，而不是只堆叠单个功能模块。
'@ | Set-Content -LiteralPath $courseDesignReportPath -Encoding UTF8

  $courseDesignMetadataPath = Join-Path $tempRoot 'course-design-metadata.json'
  @'
{
  "学生姓名": "李四",
  "学号": "20261234",
  "班级": "软工 2302",
  "指导老师": "王老师",
  "课程名称": "软件工程综合实践",
  "课题名称": "校园导览小程序设计",
  "设计类别": "课程设计",
  "完成时间": "2026-04-08",
  "设计地点": "实验楼 A201"
}
'@ | Set-Content -LiteralPath $courseDesignMetadataPath -Encoding UTF8

  $courseDesignValidation = (& (Join-Path $repoRoot 'scripts\validate-report-draft.ps1') -Path $courseDesignReportPath -ReportProfileName 'course-design-report' -Format json | Out-String) | ConvertFrom-Json
  Assert-True -Condition ([bool]$courseDesignValidation.passed) -Message 'Course-design report validation should pass for the course-design profile.'
  Assert-True -Condition ([string]$courseDesignValidation.reportProfileName -eq 'course-design-report') -Message 'Course-design report validation is missing the expected profile name.'
  Assert-True -Condition ([int]$courseDesignValidation.summary.sectionCount -ge 7) -Message 'Course-design report validation did not load the expected section rules.'

  $courseDesignStructureRiskReportPath = Join-Path $tempRoot 'course-design-structure-risk-report.md'
  @(
    '软件工程课程设计报告',
    '',
    '课程名称：软件工程综合实践',
    '课题名称：校园导览小程序设计',
    '',
    '一、设计目标',
    '本次课程设计希望完成一个能够展示校园地点、支持搜索和路线提示的小程序，并把课程中的需求分析、页面设计和接口联调过程串成完整作品。',
    '为了保证成品具备演示价值，还需要在交互流程中覆盖首页分类、地点详情和结果回跳等关键场景。',
    '',
    '二、开发环境',
    '项目开发环境包括 Windows 11、Node.js、微信开发者工具和 SQLite，本地还提供模拟接口数据用于联调和演示。',
    '通过统一前后端调试环境，可以减少页面逻辑、接口字段和测试结果之间的不一致。',
    '',
    '三、需求分析',
    '系统需要支持地点分类浏览、关键字搜索、详情展示和推荐路线提示，并兼顾页面响应速度、信息完整度和移动端交互流畅性。',
    '在分析阶段还梳理了教学区、生活区和服务区等典型使用场景，为后续方案设计提供依据。',
    '',
    '四、运行结果',
    '系统启动后能够展示首页分类卡片，输入地点关键字后可以返回匹配结果，并支持点击查看地点详情和路线提示信息。',
    '在演示测试中，页面跳转、搜索流程和收藏状态更新保持稳定，没有出现明显的白屏或异常回退。',
    '',
    '五、设计总结'
  ) | Set-Content -LiteralPath $courseDesignStructureRiskReportPath -Encoding UTF8
  $courseDesignStructureRiskValidation = (& (Join-Path $repoRoot 'scripts\validate-report-draft.ps1') -Path $courseDesignStructureRiskReportPath -ReportProfileName 'course-design-report' -Format json | Out-String) | ConvertFrom-Json
  $courseDesignStructureRiskCodes = @($courseDesignStructureRiskValidation.findings | ForEach-Object { [string]$_.code })
  Assert-True -Condition (-not [bool]$courseDesignStructureRiskValidation.passed) -Message 'Course-design structural-risk fixture should not pass validation.'
  Assert-True -Condition ([string]$courseDesignStructureRiskValidation.reportProfileName -eq 'course-design-report') -Message 'Course-design structural-risk fixture should keep the course-design profile.'
  Assert-True -Condition ($courseDesignStructureRiskCodes -contains 'missing-profile-required-heading') -Message 'Course-design structural-risk fixture should report missing-profile-required-heading.'
  Assert-True -Condition ($courseDesignStructureRiskCodes -contains 'empty-section') -Message 'Course-design structural-risk fixture should report empty-section.'
  $results.Add('course-design structural validation OK') | Out-Null

  $courseDesignTemplateDocx = Join-Path $tempRoot 'course-design-template.docx'
  New-CourseDesignTemplateDocx -Path $courseDesignTemplateDocx
  Assert-True -Condition (Test-Path -LiteralPath $courseDesignTemplateDocx) -Message 'Failed to create the course-design template fixture.'

  $courseDesignTemplateFitPath = Join-Path $tempRoot 'course-design-template-fit.json'
  & (Join-Path $repoRoot 'scripts\check-report-profile-template-fit.ps1') `
    -TemplatePath $courseDesignTemplateDocx `
    -ReportPath $courseDesignReportPath `
    -MetadataPath $courseDesignMetadataPath `
    -ReportProfileName 'course-design-report' `
    -Format json `
    -OutFile $courseDesignTemplateFitPath | Out-Null
  $courseDesignTemplateFit = (Get-Content -LiteralPath $courseDesignTemplateFitPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$courseDesignTemplateFit.reportProfileName -eq 'course-design-report') -Message 'Course-design template-fit checker is missing the expected profile name.'
  Assert-True -Condition ([int]$courseDesignTemplateFit.summary.profileChangeSuggestionCount -eq 0) -Message 'Course-design template-fit checker should not request profile changes for the built-in course-design profile.'
  Assert-True -Condition ([int]$courseDesignTemplateFit.summary.inputGapCount -eq 0) -Message 'Course-design template-fit checker should not report input gaps for the complete sample inputs.'
  Assert-True -Condition ([bool]$courseDesignTemplateFit.summary.readyForNewProfile) -Message 'Course-design template-fit checker should report that the profile is ready for reuse.'

  $courseDesignFieldMapPath = Join-Path $tempRoot 'course-design-field-map.json'
  & (Join-Path $repoRoot 'scripts\generate-docx-field-map.ps1') `
    -TemplatePath $courseDesignTemplateDocx `
    -ReportPath $courseDesignReportPath `
    -MetadataPath $courseDesignMetadataPath `
    -ReportProfileName 'course-design-report' `
    -Format json `
    -OutFile $courseDesignFieldMapPath | Out-Null
  $courseDesignFieldMapRoot = (Get-Content -LiteralPath $courseDesignFieldMapPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$courseDesignFieldMapRoot.reportProfileName -eq 'course-design-report') -Message 'Course-design field-map generator is missing the expected report profile name.'
  Assert-True -Condition ($courseDesignFieldMapRoot.fieldMap.PSObject.Properties.Name -contains '课题名称') -Message 'Course-design field-map generator did not map the project title field.'
  Assert-True -Condition ($courseDesignFieldMapRoot.fieldMap.PSObject.Properties.Name -contains '学生姓名') -Message 'Course-design field-map generator did not map the student name field.'
  Assert-True -Condition ([string]$courseDesignFieldMapRoot.fieldMap.课题名称 -eq '校园导览小程序设计') -Message 'Course-design field-map generator did not fill the project title.'
  Assert-True -Condition ([string]$courseDesignFieldMapRoot.fieldMap.学生姓名 -eq '李四') -Message 'Course-design field-map generator did not fill the student name.'
  Assert-True -Condition ([string]$courseDesignFieldMapRoot.fieldMap.设计目标.mode -eq 'after') -Message 'Course-design field-map generator should preserve the target heading and fill after it.'
  Assert-True -Condition ([string]$courseDesignFieldMapRoot.fieldMap.方案设计与实现.mode -eq 'after') -Message 'Course-design field-map generator should preserve the implementation heading and fill after it.'
  Assert-True -Condition ($courseDesignFieldMapRoot.summary.diagnosticCount -eq 0) -Message 'Course-design field-map generator should not emit diagnostics for the built-in course-design profile fixture.'

  $courseDesignFilledDocx = Join-Path $tempRoot 'course-design-template.filled.docx'
  $courseDesignFillResult = & (Join-Path $repoRoot 'scripts\apply-docx-field-map.ps1') -TemplatePath $courseDesignTemplateDocx -MappingPath $courseDesignFieldMapPath -OutPath $courseDesignFilledDocx
  Assert-True -Condition (Test-Path -LiteralPath $courseDesignFilledDocx) -Message 'Course-design field-map fill did not create the filled docx.'
  Assert-True -Condition ($courseDesignFillResult.labelFillCount -ge 6) -Message 'Course-design field-map fill applied too few label fields.'
  Assert-True -Condition ($courseDesignFillResult.blockFillCount -ge 5) -Message 'Course-design field-map fill applied too few block fills.'
  $courseDesignFilledOutline = & (Join-Path $repoRoot 'scripts\extract-docx-template.ps1') -Path $courseDesignFilledDocx -Format markdown | Out-String
  Assert-True -Condition ($courseDesignFilledOutline -match '课题名称：校园导览小程序设计') -Message 'Course-design field-map fill did not update the project title.'
  Assert-True -Condition ($courseDesignFilledOutline -match '学生姓名：李四|T1R1C2: 李四') -Message 'Course-design field-map fill did not update the student name.'
  Assert-True -Condition ($courseDesignFilledOutline -match '整体方案采用前后端分层结构') -Message 'Course-design field-map fill did not write the implementation section.'

  $courseDesignStyledDocx = Join-Path $tempRoot 'course-design-template.styled.docx'
  $courseDesignStyleResult = & (Join-Path $repoRoot 'scripts\format-docx-report-style.ps1') `
    -DocxPath $courseDesignFilledDocx `
    -OutPath $courseDesignStyledDocx `
    -Overwrite `
    -Profile auto `
    -ReportProfileName 'course-design-report'
  Assert-True -Condition (Test-Path -LiteralPath $courseDesignStyledDocx) -Message 'Course-design style formatter did not create the styled docx.'
  Assert-True -Condition ([string]$courseDesignStyleResult.reportProfileName -eq 'course-design-report') -Message 'Course-design style formatter is missing the expected report profile name.'
  Assert-True -Condition ([string]$courseDesignStyleResult.resolvedProfile -eq 'school') -Message 'Course-design style formatter should resolve auto to the course-design default style profile.'

  $courseDesignImageSpecsPath = Join-Path $tempRoot 'course-design-image-specs.json'
  @"
{
  "images": [
    {
      "path": "$($sampleImageOne.Replace('\', '\\'))",
      "section": "方案设计与实现"
    },
    {
      "path": "$($sampleImageTwo.Replace('\', '\\'))",
      "section": "运行结果"
    }
  ]
}
"@ | Set-Content -LiteralPath $courseDesignImageSpecsPath -Encoding UTF8

  $courseDesignImageMapPath = Join-Path $tempRoot 'course-design-image-map.json'
  & (Join-Path $repoRoot 'scripts\generate-docx-image-map.ps1') `
    -DocxPath $courseDesignStyledDocx `
    -ImageSpecsPath $courseDesignImageSpecsPath `
    -ReportProfileName 'course-design-report' `
    -Format json `
    -OutFile $courseDesignImageMapPath | Out-Null
  $courseDesignImageMapRoot = (Get-Content -LiteralPath $courseDesignImageMapPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$courseDesignImageMapRoot.summary.reportProfileName -eq 'course-design-report') -Message 'Course-design image-map generator is missing the expected report profile name.'
  Assert-True -Condition ([string]$courseDesignImageMapRoot.images[0].caption -eq '图1 实现过程截图') -Message 'Course-design image-map generator did not use the profile-defined default caption for the implementation section.'
  Assert-True -Condition ([string]$courseDesignImageMapRoot.images[1].caption -eq '图2 运行结果截图') -Message 'Course-design image-map generator did not use the profile-defined default caption for the result section.'

  $courseDesignImageFilledDocx = Join-Path $tempRoot 'course-design-template.images.docx'
  $courseDesignImageInsertResult = & (Join-Path $repoRoot 'scripts\insert-docx-images.ps1') -DocxPath $courseDesignStyledDocx -MappingPath $courseDesignImageMapPath -OutPath $courseDesignImageFilledDocx
  Assert-True -Condition (Test-Path -LiteralPath $courseDesignImageFilledDocx) -Message 'Course-design image insertion did not create the filled docx.'
  Assert-True -Condition ([string]$courseDesignImageInsertResult.reportProfileName -eq 'course-design-report') -Message 'Course-design image insertion is missing the expected report profile name.'
  Assert-True -Condition ([string]$courseDesignImageInsertResult.mappingInputMode -eq 'mapping-path') -Message 'Course-design image insertion should record mappingInputMode=mapping-path for image-map files.'
  Assert-True -Condition ($courseDesignImageInsertResult.insertedImageCount -eq 2) -Message 'Course-design image insertion inserted an unexpected number of images.'
  $courseDesignImageTemp = Join-Path $tempRoot 'course-design-image-inspect'
  [System.IO.Compression.ZipFile]::ExtractToDirectory($courseDesignImageFilledDocx, $courseDesignImageTemp)
  [xml]$courseDesignImageDocumentXml = [System.IO.File]::ReadAllText((Join-Path $courseDesignImageTemp 'word\document.xml'), (New-Object System.Text.UTF8Encoding($false)))
  Assert-True -Condition ($courseDesignImageDocumentXml.OuterXml -match '图1 实现过程截图') -Message 'Course-design image insertion is missing the implementation caption.'
  Assert-True -Condition ($courseDesignImageDocumentXml.OuterXml -match '图2 运行结果截图') -Message 'Course-design image insertion is missing the result caption.'
  Remove-Item -LiteralPath $courseDesignImageTemp -Recurse -Force
  $results.Add('course-design profile pipeline OK') | Out-Null

  $internshipReportPath = Join-Path $tempRoot 'internship-report.md'
  @'
专业实习报告

专业名称：软件工程
实习项目：企业门户管理后台开发
学生姓名：王敏
学号：20262345
指导教师：周老师
实习时间：2026-03-01 至 2026-03-28
实习单位：杭州云帆科技有限公司（滨江区）

一、实习目标
本次专业实习的目标是在真实企业环境中参与后台管理系统的模块开发与联调过程，理解学校课程中的需求分析、接口设计和前后端协作在实际项目中的落地方式。
除了完成指定开发任务，还需要在实习过程中学习团队的代码评审、缺陷跟踪和迭代交付流程，形成对企业项目节奏和工程规范的整体认识。

二、实习单位
实习单位为杭州云帆科技有限公司，团队主要负责企业门户与内部运营系统的研发维护，日常开发环境包括 Windows 11、Node.js、Vue 和 MySQL。
办公环境采用小组协作方式推进任务，开发期间需要通过飞书同步需求、在 Git 仓库提交分支代码，并在测试环境完成接口联调和页面验收。

三、岗位职责
实习阶段主要承担企业门户管理后台中菜单权限、公告发布和操作日志三个模块的前端联调与接口适配工作，同时需要配合后端同学完成字段校验和错误提示策略调整。
在岗位要求上，不仅要按任务单完成页面开发，还要保证表单交互清晰、接口异常可回显、提交记录可追溯，并在每周例会上汇报当前进度与遗留风险。

四、工作内容
进入项目后首先熟悉现有后台项目结构，梳理登录态校验、路由权限控制和公共请求封装的实现方式，然后在导师指导下完成公告管理列表页与详情页的改造。
在后续两周里，继续参与角色权限页面、操作日志筛选条件和批量导出功能的开发，对接了新增接口字段，并针对分页状态丢失、日期筛选不准确等问题做了多轮修复与验证。
为了保证交付质量，还配合测试同学复现缺陷，补充了按钮禁用、空状态提示和异常回显逻辑，并将关键页面的提交流程整理成操作文档，便于后续成员继续维护。

五、工作成果
通过本次实习，已经能够独立完成公告发布、日志筛选和权限配置等典型后台页面的功能修改，并能在测试环境中定位接口返回与前端展示不一致的问题。
最终提交的成果包括三个稳定可用的业务模块改造、若干条缺陷修复记录、配套的操作说明文档以及一份面向组内交接的联调注意事项清单，使项目能够更顺畅地进入后续迭代阶段。

六、遇到的问题与改进
实习初期最大的困难是对项目上下文不熟悉，看到需求单时很难快速判断应该修改哪一层代码，导致早期提交需要反复返工。
后续通过先画模块关系、再对照接口文档梳理数据流的方式，逐渐降低了理解成本，但在组件复用和跨页面状态同步方面仍然存在设计不够统一的问题，后续可以继续通过抽取公共表单配置与状态管理层来改进。

七、实习体会
这次专业实习最大的收获，是把课堂上分散学习的前端开发、数据库接口、版本管理和缺陷协作真正串成了一条完整的工程链路，理解了企业项目为什么强调规范和沟通。
相比只在课程作业中完成单点功能，真实实习更要求对任务背景、影响范围和交付质量负责，也让我明确了后续需要继续提升接口抽象能力、问题定位速度和文档表达能力。
'@ | Set-Content -LiteralPath $internshipReportPath -Encoding UTF8

  $internshipMetadataPath = Join-Path $tempRoot 'internship-metadata.json'
  @'
{
  "学生姓名": "王敏",
  "学号": "20262345",
  "班级": "软工 2303",
  "指导教师": "周老师",
  "所属专业": "软件工程",
  "项目名称": "企业门户管理后台开发",
  "实习性质": "专业实习",
  "实习时间": "2026-03-01 至 2026-03-28",
  "实习地点": "杭州云帆科技有限公司（滨江区）"
}
'@ | Set-Content -LiteralPath $internshipMetadataPath -Encoding UTF8

  $internshipValidation = (& (Join-Path $repoRoot 'scripts\validate-report-draft.ps1') -Path $internshipReportPath -ReportProfileName 'internship-report' -Format json | Out-String) | ConvertFrom-Json
  Assert-True -Condition ([bool]$internshipValidation.passed) -Message 'Internship report validation should pass for the internship profile.'
  Assert-True -Condition ([string]$internshipValidation.reportProfileName -eq 'internship-report') -Message 'Internship report validation is missing the expected profile name.'
  Assert-True -Condition ([int]$internshipValidation.summary.sectionCount -ge 7) -Message 'Internship report validation did not load the expected section rules.'

  $internshipDenseUnit = ((@(
        '实习单位的项目协作流程覆盖需求评审、接口联调、测试验收和缺陷回归，导师要求每天同步进度并把关键页面截图整理为图1、图2和图3，因此这一节故意保持长段落以触发分页风险检测。'
      ) * 18) -join '')
  $internshipStructureRiskReportPath = Join-Path $tempRoot 'internship-structure-risk-report.md'
  @(
    '专业实习报告',
    '',
    '专业名称：软件工程',
    '实习项目：企业门户管理后台开发',
    '',
    '一、实习目标',
    '本次实习目标是熟悉企业项目中的需求拆解、接口联调和缺陷跟踪流程，并把课堂学习的前端开发知识迁移到真实后台系统。开展实习前还需要明确交付边界、沟通节奏和代码提交规范。',
    '',
    '二、实习单位',
    $internshipDenseUnit,
    '',
    '五、工作成果',
    '阶段成果包括完成公告管理页面、操作日志筛选和权限菜单配置的若干改造，并配合测试同学复现和关闭了多条缺陷。相关工作也沉淀为交接说明，便于后续迭代继续维护。',
    '',
    '三、岗位职责',
    '岗位职责包括根据任务单完成页面开发、接口字段联调、异常提示处理和代码评审修改，同时需要在例会上同步进展并记录待解决风险。',
    '',
    '六、遇到的问题与改进',
    '',
    '七、实习体会',
    '__________',
    '',
    '八、实习目标',
    '重复标题段落仅用于触发结构校验。'
  ) | Set-Content -LiteralPath $internshipStructureRiskReportPath -Encoding UTF8
  $internshipStructureRiskValidation = (& (Join-Path $repoRoot 'scripts\validate-report-draft.ps1') -Path $internshipStructureRiskReportPath -ReportProfileName 'internship-report' -Format json | Out-String) | ConvertFrom-Json
  $internshipStructureRiskCodes = @($internshipStructureRiskValidation.findings | ForEach-Object { [string]$_.code })
  $internshipStructureRiskWarningCodes = @($internshipStructureRiskValidation.summary.warningCodes)
  Assert-True -Condition (-not [bool]$internshipStructureRiskValidation.passed) -Message 'Internship structural-risk fixture should not pass validation.'
  Assert-True -Condition ([string]$internshipStructureRiskValidation.reportProfileName -eq 'internship-report') -Message 'Internship structural-risk fixture should keep the internship profile.'
  Assert-True -Condition ($internshipStructureRiskCodes -contains 'missing-profile-required-heading') -Message 'Internship structural-risk fixture should report missing-profile-required-heading.'
  Assert-True -Condition ($internshipStructureRiskCodes -contains 'duplicate-section-heading') -Message 'Internship structural-risk fixture should report duplicate-section-heading.'
  Assert-True -Condition ($internshipStructureRiskCodes -contains 'section-order-anomaly') -Message 'Internship structural-risk fixture should report section-order-anomaly.'
  Assert-True -Condition ($internshipStructureRiskCodes -contains 'empty-section') -Message 'Internship structural-risk fixture should report empty-section.'
  Assert-True -Condition ($internshipStructureRiskCodes -contains 'placeholder-only-section') -Message 'Internship structural-risk fixture should report placeholder-only-section.'
  Assert-True -Condition (([int]$internshipStructureRiskValidation.summary.paginationRiskCount) -ge 3) -Message 'Internship structural-risk fixture should surface pagination risk warnings.'
  Assert-True -Condition ($internshipStructureRiskWarningCodes -contains 'pagination-risk-long-section') -Message 'Internship structural-risk fixture should report pagination-risk-long-section.'
  Assert-True -Condition ($internshipStructureRiskWarningCodes -contains 'pagination-risk-dense-section-block') -Message 'Internship structural-risk fixture should report pagination-risk-dense-section-block.'
  Assert-True -Condition ($internshipStructureRiskWarningCodes -contains 'pagination-risk-figure-cluster') -Message 'Internship structural-risk fixture should report pagination-risk-figure-cluster.'
  $results.Add('internship structural validation OK') | Out-Null

  $internshipTemplateDocx = Join-Path $tempRoot 'internship-template.docx'
  New-InternshipTemplateDocx -Path $internshipTemplateDocx
  Assert-True -Condition (Test-Path -LiteralPath $internshipTemplateDocx) -Message 'Failed to create the internship template fixture.'

  $internshipTemplateFitPath = Join-Path $tempRoot 'internship-template-fit.json'
  & (Join-Path $repoRoot 'scripts\check-report-profile-template-fit.ps1') `
    -TemplatePath $internshipTemplateDocx `
    -ReportPath $internshipReportPath `
    -MetadataPath $internshipMetadataPath `
    -ReportProfileName 'internship-report' `
    -Format json `
    -OutFile $internshipTemplateFitPath | Out-Null
  $internshipTemplateFit = (Get-Content -LiteralPath $internshipTemplateFitPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$internshipTemplateFit.reportProfileName -eq 'internship-report') -Message 'Internship template-fit checker is missing the expected profile name.'
  Assert-True -Condition ([int]$internshipTemplateFit.summary.profileChangeSuggestionCount -eq 0) -Message 'Internship template-fit checker should not request profile changes for the built-in internship profile.'
  Assert-True -Condition ([int]$internshipTemplateFit.summary.inputGapCount -eq 0) -Message 'Internship template-fit checker should not report input gaps for the complete sample inputs.'
  Assert-True -Condition ([bool]$internshipTemplateFit.summary.readyForNewProfile) -Message 'Internship template-fit checker should report that the profile is ready for reuse.'

  $internshipFieldMapPath = Join-Path $tempRoot 'internship-field-map.json'
  & (Join-Path $repoRoot 'scripts\generate-docx-field-map.ps1') `
    -TemplatePath $internshipTemplateDocx `
    -ReportPath $internshipReportPath `
    -MetadataPath $internshipMetadataPath `
    -ReportProfileName 'internship-report' `
    -Format json `
    -OutFile $internshipFieldMapPath | Out-Null
  $internshipFieldMapRoot = (Get-Content -LiteralPath $internshipFieldMapPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$internshipFieldMapRoot.reportProfileName -eq 'internship-report') -Message 'Internship field-map generator is missing the expected report profile name.'
  Assert-True -Condition ($internshipFieldMapRoot.fieldMap.PSObject.Properties.Name -contains '专业名称') -Message 'Internship field-map generator did not map the major field.'
  Assert-True -Condition ($internshipFieldMapRoot.fieldMap.PSObject.Properties.Name -contains '实习项目') -Message 'Internship field-map generator did not map the internship title field.'
  Assert-True -Condition ($internshipFieldMapRoot.fieldMap.PSObject.Properties.Name -contains '学生姓名') -Message 'Internship field-map generator did not map the student name field.'
  Assert-True -Condition ([string]$internshipFieldMapRoot.fieldMap.专业名称 -eq '软件工程') -Message 'Internship field-map generator did not fill the major field.'
  Assert-True -Condition ([string]$internshipFieldMapRoot.fieldMap.实习项目 -eq '企业门户管理后台开发') -Message 'Internship field-map generator did not fill the internship title.'
  Assert-True -Condition ([string]$internshipFieldMapRoot.fieldMap.学生姓名 -eq '王敏') -Message 'Internship field-map generator did not fill the student name.'
  Assert-True -Condition ([string]$internshipFieldMapRoot.fieldMap.实习过程与内容.mode -eq 'after') -Message 'Internship field-map generator should preserve the process heading and fill after it.'
  Assert-True -Condition ([string]$internshipFieldMapRoot.fieldMap.实习成果.mode -eq 'after') -Message 'Internship field-map generator should preserve the results heading and fill after it.'
  Assert-True -Condition ($internshipFieldMapRoot.summary.diagnosticCount -eq 0) -Message 'Internship field-map generator should not emit diagnostics for the built-in internship profile fixture.'

  $internshipFilledDocx = Join-Path $tempRoot 'internship-template.filled.docx'
  $internshipFillResult = & (Join-Path $repoRoot 'scripts\apply-docx-field-map.ps1') -TemplatePath $internshipTemplateDocx -MappingPath $internshipFieldMapPath -OutPath $internshipFilledDocx
  Assert-True -Condition (Test-Path -LiteralPath $internshipFilledDocx) -Message 'Internship field-map fill did not create the filled docx.'
  Assert-True -Condition ($internshipFillResult.labelFillCount -ge 6) -Message 'Internship field-map fill applied too few label fields.'
  Assert-True -Condition ($internshipFillResult.blockFillCount -ge 5) -Message 'Internship field-map fill applied too few block fills.'
  $internshipFilledOutline = & (Join-Path $repoRoot 'scripts\extract-docx-template.ps1') -Path $internshipFilledDocx -Format markdown | Out-String
  Assert-True -Condition ($internshipFilledOutline -match '实习项目：企业门户管理后台开发') -Message 'Internship field-map fill did not update the internship title.'
  Assert-True -Condition ($internshipFilledOutline -match '学生姓名：王敏|T1R1C2: 王敏') -Message 'Internship field-map fill did not update the student name.'
  Assert-True -Condition ($internshipFilledOutline -match '进入项目后首先熟悉现有后台项目结构') -Message 'Internship field-map fill did not write the internship process section.'

  $internshipStyledDocx = Join-Path $tempRoot 'internship-template.styled.docx'
  $internshipStyleResult = & (Join-Path $repoRoot 'scripts\format-docx-report-style.ps1') `
    -DocxPath $internshipFilledDocx `
    -OutPath $internshipStyledDocx `
    -Overwrite `
    -Profile auto `
    -ReportProfileName 'internship-report'
  Assert-True -Condition (Test-Path -LiteralPath $internshipStyledDocx) -Message 'Internship style formatter did not create the styled docx.'
  Assert-True -Condition ([string]$internshipStyleResult.reportProfileName -eq 'internship-report') -Message 'Internship style formatter is missing the expected report profile name.'
  Assert-True -Condition ([string]$internshipStyleResult.resolvedProfile -eq 'school') -Message 'Internship style formatter should resolve auto to the internship default style profile.'

  $internshipImageSpecsPath = Join-Path $tempRoot 'internship-image-specs.json'
  @"
{
  "images": [
    {
      "path": "$($sampleImageOne.Replace('\', '\\'))",
      "section": "工作内容"
    },
    {
      "path": "$($sampleImageTwo.Replace('\', '\\'))",
      "section": "工作成果"
    }
  ]
}
"@ | Set-Content -LiteralPath $internshipImageSpecsPath -Encoding UTF8

  $internshipImageMapPath = Join-Path $tempRoot 'internship-image-map.json'
  & (Join-Path $repoRoot 'scripts\generate-docx-image-map.ps1') `
    -DocxPath $internshipStyledDocx `
    -ImageSpecsPath $internshipImageSpecsPath `
    -ReportProfileName 'internship-report' `
    -Format json `
    -OutFile $internshipImageMapPath | Out-Null
  $internshipImageMapRoot = (Get-Content -LiteralPath $internshipImageMapPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$internshipImageMapRoot.summary.reportProfileName -eq 'internship-report') -Message 'Internship image-map generator is missing the expected report profile name.'
  Assert-True -Condition ([string]$internshipImageMapRoot.images[0].caption -eq '图1 实习过程截图') -Message 'Internship image-map generator did not use the profile-defined default caption for the process section.'
  Assert-True -Condition ([string]$internshipImageMapRoot.images[1].caption -eq '图2 实习成果截图') -Message 'Internship image-map generator did not use the profile-defined default caption for the results section.'

  $internshipImageFilledDocx = Join-Path $tempRoot 'internship-template.images.docx'
  $internshipImageInsertResult = & (Join-Path $repoRoot 'scripts\insert-docx-images.ps1') -DocxPath $internshipStyledDocx -MappingPath $internshipImageMapPath -OutPath $internshipImageFilledDocx
  Assert-True -Condition (Test-Path -LiteralPath $internshipImageFilledDocx) -Message 'Internship image insertion did not create the filled docx.'
  Assert-True -Condition ([string]$internshipImageInsertResult.reportProfileName -eq 'internship-report') -Message 'Internship image insertion is missing the expected report profile name.'
  Assert-True -Condition ([string]$internshipImageInsertResult.mappingInputMode -eq 'mapping-path') -Message 'Internship image insertion should record mappingInputMode=mapping-path for image-map files.'
  Assert-True -Condition ($internshipImageInsertResult.insertedImageCount -eq 2) -Message 'Internship image insertion inserted an unexpected number of images.'
  $internshipImageTemp = Join-Path $tempRoot 'internship-image-inspect'
  [System.IO.Compression.ZipFile]::ExtractToDirectory($internshipImageFilledDocx, $internshipImageTemp)
  [xml]$internshipImageDocumentXml = [System.IO.File]::ReadAllText((Join-Path $internshipImageTemp 'word\document.xml'), (New-Object System.Text.UTF8Encoding($false)))
  Assert-True -Condition ($internshipImageDocumentXml.OuterXml -match '图1 实习过程截图') -Message 'Internship image insertion is missing the process caption.'
  Assert-True -Condition ($internshipImageDocumentXml.OuterXml -match '图2 实习成果截图') -Message 'Internship image insertion is missing the results caption.'
  Remove-Item -LiteralPath $internshipImageTemp -Recurse -Force
  $results.Add('internship profile pipeline OK') | Out-Null

  $mixedGroupImageSpecsPath = Join-Path $tempRoot 'image-specs-mixed-group.json'
  @"
{
  "images": [
    {
      "path": "$($sampleImageOne.Replace('\', '\\'))",
      "section": "实验结果",
      "caption": "图1 结果A",
      "widthCm": 7.8,
      "layout": {
        "mode": "row",
        "columns": 2,
        "group": "mixed-grid"
      }
    },
    {
      "path": "$($sampleImageTwo.Replace('\', '\\'))",
      "section": "实验结果",
      "caption": "图2 结果B",
      "widthCm": 7.8,
      "layout": {
        "mode": "row",
        "columns": 2,
        "group": "mixed-grid"
      }
    },
    {
      "path": "$($sampleImageThree.Replace('\', '\\'))",
      "section": "问题分析",
      "caption": "图3 分析A",
      "widthCm": 7.8,
      "layout": {
        "mode": "row",
        "columns": 2,
        "group": "mixed-grid"
      }
    },
    {
      "path": "$($sampleImageFour.Replace('\', '\\'))",
      "section": "问题分析",
      "caption": "图4 分析B",
      "widthCm": 7.8,
      "layout": {
        "mode": "row",
        "columns": 2,
        "group": "mixed-grid"
      }
    }
  ]
}
"@ | Set-Content -LiteralPath $mixedGroupImageSpecsPath -Encoding UTF8

  $mixedGroupImageMapPath = Join-Path $tempRoot 'generated-image-map-mixed-group.json'
  & (Join-Path $repoRoot 'scripts\generate-docx-image-map.ps1') -DocxPath $coverBodyFilledDocx -ImageSpecsPath $mixedGroupImageSpecsPath -Format json -OutFile $mixedGroupImageMapPath | Out-Null
  $mixedGroupImageMapRoot = (Get-Content -LiteralPath $mixedGroupImageMapPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition (@($mixedGroupImageMapRoot.images).Count -eq 4) -Message 'Mixed-group image-map generator produced an unexpected number of images.'
  Assert-True -Condition ((@($mixedGroupImageMapRoot.images | Where-Object { $_.layout.groupAnchor -eq '实验结果' }).Count) -eq 4) -Message 'Mixed-group image-map generator should unify the row group under the shared 实验结果 groupAnchor.'
  Assert-True -Condition ((@($mixedGroupImageMapRoot.notes | Where-Object { $_ -match 'mixed-grid' }).Count) -ge 1) -Message 'Mixed-group image-map generator should explain the shared groupAnchor note.'

  $mixedGroupFilledDocx = Join-Path $tempRoot 'cover-body-template.mixed-group-images.docx'
  $mixedGroupInsertResult = & (Join-Path $repoRoot 'scripts\insert-docx-images.ps1') -DocxPath $coverBodyFilledDocx -MappingPath $mixedGroupImageMapPath -OutPath $mixedGroupFilledDocx
  Assert-True -Condition (Test-Path -LiteralPath $mixedGroupFilledDocx) -Message 'Mixed-group image insertion did not create the filled docx.'
  Assert-True -Condition ($mixedGroupInsertResult.insertedImageCount -eq 4) -Message 'Mixed-group image insertion inserted an unexpected number of images.'
  $mixedGroupInspect = Join-Path $tempRoot 'mixed-group-inspect'
  [System.IO.Compression.ZipFile]::ExtractToDirectory($mixedGroupFilledDocx, $mixedGroupInspect)
  [xml]$mixedGroupDocumentXml = [System.IO.File]::ReadAllText((Join-Path $mixedGroupInspect 'word\document.xml'), (New-Object System.Text.UTF8Encoding($false)))
  $mixedGroupNamespaceManager = New-Object System.Xml.XmlNamespaceManager($mixedGroupDocumentXml.NameTable)
  $mixedGroupNamespaceManager.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
  Assert-True -Condition (@($mixedGroupDocumentXml.SelectNodes('//w:tbl', $mixedGroupNamespaceManager)).Count -eq 2) -Message 'Mixed-group image insertion should add exactly one image layout table.'
  $mixedGroupDocumentText = $mixedGroupDocumentXml.OuterXml
  $resultHeadingIndex = $mixedGroupDocumentText.IndexOf('四. 实验结果')
  $analysisHeadingIndex = $mixedGroupDocumentText.IndexOf('五. 问题分析')
  $analysisCaptionIndex = $mixedGroupDocumentText.IndexOf('图3 分析A')
  Assert-True -Condition ($resultHeadingIndex -ge 0 -and $analysisHeadingIndex -gt $resultHeadingIndex) -Message 'Mixed-group image insertion document is missing the expected section headings.'
  Assert-True -Condition ($analysisCaptionIndex -gt $resultHeadingIndex -and $analysisCaptionIndex -lt $analysisHeadingIndex) -Message 'Mixed-group image insertion should keep the whole 2x2 block under the shared 实验结果 anchor.'
  Remove-Item -LiteralPath $mixedGroupInspect -Recurse -Force
  $results.Add('docx image-map mixed group-anchor generation OK') | Out-Null
  $results.Add('docx image insertion mixed group-anchor OK') | Out-Null

  $imageSpecsPath = Join-Path $tempRoot 'image-specs.json'
  @"
{
  "images": [
    {
      "path": "$($sampleImageOne.Replace('\', '\\'))",
      "section": "实验步骤"
    },
    {
      "path": "$($sampleImageTwo.Replace('\', '\\'))",
      "section": "实验结果"
    }
  ]
}
"@ | Set-Content -LiteralPath $imageSpecsPath -Encoding UTF8

  $generatedImageMapPath = Join-Path $tempRoot 'generated-image-map.json'
  & (Join-Path $repoRoot 'scripts\generate-docx-image-map.ps1') -DocxPath $generatedFilledDocx -ImageSpecsPath $imageSpecsPath -Format json -OutFile $generatedImageMapPath | Out-Null
  Assert-True -Condition (Test-Path -LiteralPath $generatedImageMapPath) -Message 'Image-map generator did not write the output file.'
  $generatedImageMapRoot = (Get-Content -LiteralPath $generatedImageMapPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition (@($generatedImageMapRoot.images).Count -eq 2) -Message 'Image-map generator produced an unexpected number of images.'
  Assert-True -Condition ([string]$generatedImageMapRoot.summary.reportProfileName -eq 'experiment-report') -Message 'Image-map generator is missing the expected report profile name.'
  Assert-True -Condition ([string]$generatedImageMapRoot.summary.imageInputMode -eq 'specs-path') -Message 'Image-map generator should record imageInputMode=specs-path for file-backed image specs.'
  Assert-True -Condition ([string]$generatedImageMapRoot.images[0].section -eq '实验步骤') -Message 'Image-map generator did not keep the first section.'
  Assert-True -Condition ([string]$generatedImageMapRoot.images[1].section -eq '实验结果') -Message 'Image-map generator did not keep the second section.'
  Assert-True -Condition ([string]$generatedImageMapRoot.images[0].caption -match '^图1 ') -Message 'Image-map generator did not create the first caption.'
  Assert-True -Condition ([string]$generatedImageMapRoot.images[1].caption -match '^图2 ') -Message 'Image-map generator did not create the second caption.'
  $results.Add('docx image-map generation OK') | Out-Null

  $inlineImageSpecsJson = Get-Content -LiteralPath $imageSpecsPath -Raw -Encoding UTF8
  $inlineImageMapRoot = ((& (Join-Path $repoRoot 'scripts\generate-docx-image-map.ps1') `
      -DocxPath $generatedFilledDocx `
      -ImageSpecsJson $inlineImageSpecsJson `
      -Format json) | Out-String) | ConvertFrom-Json
  Assert-True -Condition ([string]$inlineImageMapRoot.summary.imageInputMode -eq 'specs-json') -Message 'Image-map generator should record imageInputMode=specs-json for inline image specs JSON.'
  Assert-True -Condition (@($inlineImageMapRoot.images).Count -eq 2) -Message 'Inline image-map generator produced an unexpected number of images.'
  Assert-True -Condition ([string]$inlineImageMapRoot.images[0].section -eq '实验步骤') -Message 'Inline image-map generator did not keep the first section.'
  Assert-True -Condition ([string]$inlineImageMapRoot.images[1].section -eq '实验结果') -Message 'Inline image-map generator did not keep the second section.'
  $results.Add('docx image-map inline specs OK') | Out-Null

  $imagePlanMarkdown = & (Join-Path $repoRoot 'scripts\generate-docx-image-map.ps1') `
    -DocxPath $generatedFilledDocx `
    -ImagePaths $sampleImageOne,$sampleImageTwo `
    -Format markdown `
    -PlanOnly | Out-String
  Assert-True -Condition ($imagePlanMarkdown -match 'DOCX Image Plan') -Message 'Image-map planner markdown output missing header.'
  Assert-True -Condition ($imagePlanMarkdown -match 'Proposed Image Placement') -Message 'Image-map planner markdown output missing placement table.'
  Assert-True -Condition ($imagePlanMarkdown -match 'fallback-order') -Message 'Image-map planner should explain fallback-order section assignments for generic file names.'

  $imagePlanJson = (& (Join-Path $repoRoot 'scripts\generate-docx-image-map.ps1') `
      -DocxPath $generatedFilledDocx `
      -ImagePaths $sampleImageOne,$sampleImageTwo `
      -Format json `
      -PlanOnly | Out-String) | ConvertFrom-Json
  Assert-True -Condition ([bool]$imagePlanJson.summary.planOnly) -Message 'Image-map planner JSON should mark planOnly output.'
  Assert-True -Condition ([string]$imagePlanJson.summary.reportProfileName -eq 'experiment-report') -Message 'Image-map planner JSON is missing the expected report profile name.'
  Assert-True -Condition ([string]$imagePlanJson.summary.imageInputMode -eq 'image-paths') -Message 'Image-map planner JSON should record imageInputMode=image-paths for direct image lists.'
  Assert-True -Condition (@($imagePlanJson.plan).Count -eq 2) -Message 'Image-map planner JSON produced an unexpected number of plan rows.'
  Assert-True -Condition ([string]$imagePlanJson.plan[0].proposedSection -eq '实验步骤') -Message 'Image-map planner should place the first fallback image in the procedure section.'
  Assert-True -Condition ([string]$imagePlanJson.plan[1].proposedSection -eq '实验结果') -Message 'Image-map planner should place the second fallback image in the result section.'
  Assert-True -Condition (-not ($imagePlanJson.PSObject.Properties.Name -contains 'images')) -Message 'Plan-only JSON should not expose an insertion-ready images array.'
  $results.Add('docx image placement planning OK') | Out-Null

  $workspaceMediaDir = Join-Path (Split-Path -Parent $repoRoot) 'media\inbound'
  New-Item -ItemType Directory -Path $workspaceMediaDir -Force | Out-Null
  $uploadedImageSuffix = [System.Guid]::NewGuid().ToString('N')
  $uploadedImageOneName = "smoke-uploaded-result-$uploadedImageSuffix-1.png"
  $uploadedImageTwoName = "smoke-uploaded-result-$uploadedImageSuffix-2.png"
  $uploadedImageOnePath = Join-Path $workspaceMediaDir $uploadedImageOneName
  $uploadedImageTwoPath = Join-Path $workspaceMediaDir $uploadedImageTwoName
  Copy-Item -LiteralPath $sampleImageOne -Destination $uploadedImageOnePath -Force
  Copy-Item -LiteralPath $sampleImageTwo -Destination $uploadedImageTwoPath -Force
  try {
    $uploadedRelativeImageMapPath = Join-Path $tempRoot 'generated-image-map-uploaded-relative.json'
    & (Join-Path $repoRoot 'scripts\generate-docx-image-map.ps1') `
      -DocxPath $generatedFilledDocx `
      -ImagePaths ("media\inbound\{0}" -f $uploadedImageOneName),("media\inbound\{0}" -f $uploadedImageTwoName) `
      -Format json `
      -OutFile $uploadedRelativeImageMapPath | Out-Null
    $uploadedRelativeImageMapRoot = (Get-Content -LiteralPath $uploadedRelativeImageMapPath -Raw -Encoding UTF8) | ConvertFrom-Json
    Assert-True -Condition ([string]$uploadedRelativeImageMapRoot.images[0].path -eq $uploadedImageOnePath) -Message 'Image-map generator did not resolve the first uploaded relative media path against the workspace root.'
    Assert-True -Condition ([string]$uploadedRelativeImageMapRoot.images[1].path -eq $uploadedImageTwoPath) -Message 'Image-map generator did not resolve the second uploaded relative media path against the workspace root.'
    Assert-True -Condition ([string]$uploadedRelativeImageMapRoot.images[0].layout.mode -eq 'row') -Message 'Image-map generator should add row layout for same-section uploaded image paths.'
    Assert-True -Condition ([int]$uploadedRelativeImageMapRoot.images[0].layout.columns -eq 2) -Message 'Image-map generator should use 2 columns for auto row layout.'
    Assert-True -Condition ([string]$uploadedRelativeImageMapRoot.images[0].layout.group -eq [string]$uploadedRelativeImageMapRoot.images[1].layout.group) -Message 'Image-map generator should put same-section uploaded image paths in one auto row group.'

    $uploadedRelativeFilledDocx = Join-Path $tempRoot 'sample-template.uploaded-relative-images.docx'
    $uploadedRelativeInsertResult = & (Join-Path $repoRoot 'scripts\insert-docx-images.ps1') -DocxPath $generatedFilledDocx -MappingPath $uploadedRelativeImageMapPath -OutPath $uploadedRelativeFilledDocx
    Assert-True -Condition (Test-Path -LiteralPath $uploadedRelativeFilledDocx) -Message 'Uploaded relative media-path image insertion did not create the filled docx.'
    Assert-True -Condition ($uploadedRelativeInsertResult.insertedImageCount -eq 2) -Message 'Uploaded relative media-path image insertion inserted an unexpected number of images.'
  } finally {
    Remove-Item -LiteralPath $uploadedImageOnePath -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $uploadedImageTwoPath -Force -ErrorAction SilentlyContinue
  }
  $results.Add('docx image-map uploaded relative-path generation OK') | Out-Null
  $results.Add('docx image insertion uploaded relative-path OK') | Out-Null

  $wideScreenshotOne = Join-Path $tempRoot 'wide-screenshot-1.png'
  $wideScreenshotTwo = Join-Path $tempRoot 'wide-screenshot-2.png'
  New-SamplePngImage -Path $wideScreenshotOne -Text 'Wide 1' -Width 1280 -Height 720
  New-SamplePngImage -Path $wideScreenshotTwo -Text 'Wide 2' -Width 1280 -Height 720 -BackgroundHex '#FCEFD8'
  $wideScreenshotSpecsPath = Join-Path $tempRoot 'wide-screenshot-specs.json'
  @"
{
  "images": [
    {
      "path": "$($wideScreenshotOne.Replace('\', '\\'))",
      "section": "实验结果"
    },
    {
      "path": "$($wideScreenshotTwo.Replace('\', '\\'))",
      "section": "实验结果"
    }
  ]
}
"@ | Set-Content -LiteralPath $wideScreenshotSpecsPath -Encoding UTF8
  $wideScreenshotImageMapPath = Join-Path $tempRoot 'generated-wide-screenshot-image-map.json'
  & (Join-Path $repoRoot 'scripts\generate-docx-image-map.ps1') -DocxPath $generatedFilledDocx -ImageSpecsPath $wideScreenshotSpecsPath -Format json -OutFile $wideScreenshotImageMapPath | Out-Null
  $wideScreenshotMapRoot = (Get-Content -LiteralPath $wideScreenshotImageMapPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition (@($wideScreenshotMapRoot.images).Count -eq 2) -Message 'Wide screenshot image-map generator produced an unexpected number of images.'
  Assert-True -Condition ([string]$wideScreenshotMapRoot.images[0].layout.mode -eq 'row') -Message 'Wide screenshot image-map generator should still auto-row same-section screenshots.'
  Assert-True -Condition ([int]$wideScreenshotMapRoot.images[0].layout.columns -eq 2) -Message 'Wide screenshot image-map generator should keep the two-column default.'
  Assert-True -Condition ([double]$wideScreenshotMapRoot.images[0].widthCm -eq 10.5) -Message 'Wide screenshot image-map generator did not use the standard default image width.'
  Assert-True -Condition ([string]$wideScreenshotMapRoot.images[0].layout.group -eq [string]$wideScreenshotMapRoot.images[1].layout.group) -Message 'Wide screenshot image-map generator should put same-section screenshots in one auto row group.'
  $results.Add('docx image-map wide screenshot row layout OK') | Out-Null

  $sectionImageFilledDocx = Join-Path $tempRoot 'sample-template.section-images.docx'
  $sectionImageInsertResult = & (Join-Path $repoRoot 'scripts\insert-docx-images.ps1') -DocxPath $generatedFilledDocx -MappingPath $generatedImageMapPath -OutPath $sectionImageFilledDocx
  Assert-True -Condition (Test-Path -LiteralPath $sectionImageFilledDocx) -Message 'Section-based image insertion did not create the filled docx.'
  Assert-True -Condition ([string]$sectionImageInsertResult.reportProfileName -eq 'experiment-report') -Message 'Section-based image insertion is missing the expected report profile name.'
  Assert-True -Condition ($sectionImageInsertResult.insertedImageCount -eq 2) -Message 'Section-based image insertion inserted an unexpected number of images.'
  $sectionImageTemp = Join-Path $tempRoot 'section-image-inspect'
  [System.IO.Compression.ZipFile]::ExtractToDirectory($sectionImageFilledDocx, $sectionImageTemp)
  [xml]$sectionImageDocumentXml = [System.IO.File]::ReadAllText((Join-Path $sectionImageTemp 'word\document.xml'), (New-Object System.Text.UTF8Encoding($false)))
  $sectionNamespaceManager = New-Object System.Xml.XmlNamespaceManager($sectionImageDocumentXml.NameTable)
  $sectionNamespaceManager.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
  $sectionNamespaceManager.AddNamespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
  Assert-True -Condition (@($sectionImageDocumentXml.SelectNodes('//wp:inline', $sectionNamespaceManager)).Count -ge 2) -Message 'Section-based image insertion is missing expected image drawing nodes.'
  Assert-True -Condition ($sectionImageDocumentXml.OuterXml -match '图1 实验步骤截图') -Message 'Section-based image insertion is missing the first caption.'
  Assert-True -Condition ($sectionImageDocumentXml.OuterXml -match '图2 实验结果截图') -Message 'Section-based image insertion is missing the second caption.'
  Remove-Item -LiteralPath $sectionImageTemp -Recurse -Force
  $results.Add('docx image insertion by section OK') | Out-Null

  $customImageProfilePath = Join-Path $tempRoot 'custom-image-insert-profile.json'
  @'
{
  "name": "image-insert-custom-profile",
  "displayName": "自定义插图实验报告",
  "defaultStyleProfile": "compact",
  "defaultExperimentProperty": "③验证性实验",
  "metadataFields": [
    { "key": "Name", "label": "姓名" }
  ],
  "extraLabels": [],
  "sectionFields": [
    { "key": "Purpose", "heading": "实验目的", "aliases": ["实验目的"], "minChars": { "standard": 30, "full": 60 } },
    { "key": "Environment", "heading": "实验环境", "aliases": ["实验环境"], "minChars": { "standard": 30, "full": 60 } },
    { "key": "Theory", "heading": "实验原理或任务要求", "aliases": ["实验原理或任务要求"], "minChars": { "standard": 30, "full": 80 } },
    { "key": "Steps", "heading": "实验过程记录", "aliases": ["实验过程记录"], "minChars": { "standard": 60, "full": 140 } },
    { "key": "Results", "heading": "实验现象记录", "aliases": ["实验现象记录"], "minChars": { "standard": 50, "full": 120 } },
    { "key": "Analysis", "heading": "问题分析", "aliases": ["问题分析"], "minChars": { "standard": 30, "full": 80 } },
    { "key": "Summary", "heading": "实验总结", "aliases": ["实验总结"], "minChars": { "standard": 30, "full": 80 } }
  ],
  "imagePlacementDefaults": {
    "fallbackSectionOrder": ["steps", "result", "analysis"],
    "defaultCaptions": {
      "steps": "过程记录图",
      "result": "现象记录图",
      "default": "自定义实验截图"
    }
  },
  "detailProfiles": {
    "standard": { "minChars": 700, "promptGuidance": [] },
    "full": { "minChars": 1100, "promptGuidance": [] }
  }
}
'@ | Set-Content -LiteralPath $customImageProfilePath -Encoding UTF8
  $resolvedCustomImageProfilePath = (Resolve-Path -LiteralPath $customImageProfilePath).Path

  $customSectionDocx = Join-Path $tempRoot 'sample-template.custom-profile-sections.docx'
  Copy-Item -LiteralPath $sampleDocx -Destination $customSectionDocx -Force
  $customSectionArchive = [System.IO.Compression.ZipFile]::Open($customSectionDocx, [System.IO.Compression.ZipArchiveMode]::Update)
  try {
    $customSectionEntry = $customSectionArchive.GetEntry('word/document.xml')
    Assert-True -Condition ($null -ne $customSectionEntry) -Message 'Custom section fixture is missing word/document.xml before mutation.'
    $customSectionReader = New-Object System.IO.StreamReader($customSectionEntry.Open(), (New-Object System.Text.UTF8Encoding($false)))
    try {
      $customSectionDocumentText = $customSectionReader.ReadToEnd()
    } finally {
      $customSectionReader.Dispose()
    }
    $customSectionEntry.Delete()
    $customSectionDocumentText = $customSectionDocumentText -replace '实验步骤', '实验过程记录'
    $customSectionDocumentText = $customSectionDocumentText -replace '实验结果', '实验现象记录'
    $customSectionEntry = $customSectionArchive.CreateEntry('word/document.xml')
    $customSectionWriter = New-Object System.IO.StreamWriter($customSectionEntry.Open(), (New-Object System.Text.UTF8Encoding($false)))
    try {
      $customSectionWriter.Write($customSectionDocumentText)
    } finally {
      $customSectionWriter.Dispose()
    }
  } finally {
    $customSectionArchive.Dispose()
  }

  $customProfileImageSpecsPath = Join-Path $tempRoot 'custom-profile-image-specs.json'
  @"
{
  "images": [
    {
      "path": "$($sampleImageOne.Replace('\', '\\'))",
      "section": "实验过程记录"
    },
    {
      "path": "$($sampleImageTwo.Replace('\', '\\'))",
      "section": "实验现象记录"
    }
  ]
}
"@ | Set-Content -LiteralPath $customProfileImageSpecsPath -Encoding UTF8

  $customProfileImageMapPath = Join-Path $tempRoot 'custom-profile-image-map.json'
  & (Join-Path $repoRoot 'scripts\generate-docx-image-map.ps1') -DocxPath $customSectionDocx -ImageSpecsPath $customProfileImageSpecsPath -ReportProfilePath $customImageProfilePath -Format json -OutFile $customProfileImageMapPath | Out-Null
  $customProfileImageMapRoot = (Get-Content -LiteralPath $customProfileImageMapPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$customProfileImageMapRoot.summary.reportProfileName -eq 'image-insert-custom-profile') -Message 'Custom-profile image-map generator is missing the expected report profile name.'
  Assert-True -Condition ([string]$customProfileImageMapRoot.summary.reportProfilePath -eq $resolvedCustomImageProfilePath) -Message 'Custom-profile image-map generator is missing the expected report profile path.'
  Assert-True -Condition ([string]$customProfileImageMapRoot.images[0].caption -eq '图1 过程记录图') -Message 'Custom-profile image-map generator did not use the profile-defined default caption for the steps section.'
  Assert-True -Condition ([string]$customProfileImageMapRoot.images[1].caption -eq '图2 现象记录图') -Message 'Custom-profile image-map generator did not use the profile-defined default caption for the results section.'

  $customProfileFilledDocx = Join-Path $tempRoot 'sample-template.custom-profile-images.docx'
  $customProfileInsertResult = & (Join-Path $repoRoot 'scripts\insert-docx-images.ps1') -DocxPath $customSectionDocx -MappingPath $customProfileImageMapPath -OutPath $customProfileFilledDocx
  Assert-True -Condition (Test-Path -LiteralPath $customProfileFilledDocx) -Message 'Custom-profile image insertion did not create the filled docx.'
  Assert-True -Condition ([string]$customProfileInsertResult.reportProfileName -eq 'image-insert-custom-profile') -Message 'Custom-profile image insertion did not inherit the expected report profile name from the image-map summary.'
  Assert-True -Condition ([string]$customProfileInsertResult.reportProfilePath -eq $resolvedCustomImageProfilePath) -Message 'Custom-profile image insertion did not inherit the expected report profile path from the image-map summary.'
  Assert-True -Condition ($customProfileInsertResult.insertedImageCount -eq 2) -Message 'Custom-profile image insertion inserted an unexpected number of images.'
  $customProfileInspect = Join-Path $tempRoot 'custom-profile-image-inspect'
  [System.IO.Compression.ZipFile]::ExtractToDirectory($customProfileFilledDocx, $customProfileInspect)
  [xml]$customProfileDocumentXml = [System.IO.File]::ReadAllText((Join-Path $customProfileInspect 'word\document.xml'), (New-Object System.Text.UTF8Encoding($false)))
  $customProfileDocumentText = $customProfileDocumentXml.OuterXml
  $customStepsHeadingIndex = $customProfileDocumentText.IndexOf('实验过程记录')
  $customResultsHeadingIndex = $customProfileDocumentText.IndexOf('实验现象记录')
  $customStepsCaptionIndex = $customProfileDocumentText.IndexOf('图1 过程记录图')
  $customResultsCaptionIndex = $customProfileDocumentText.IndexOf('图2 现象记录图')
  Assert-True -Condition ($customStepsHeadingIndex -ge 0 -and $customResultsHeadingIndex -gt $customStepsHeadingIndex) -Message 'Custom-profile image insertion document is missing the expected custom section headings.'
  Assert-True -Condition ($customStepsCaptionIndex -gt $customStepsHeadingIndex -and $customStepsCaptionIndex -lt $customResultsHeadingIndex) -Message 'Custom-profile image insertion did not place the first image under 实验过程记录.'
  Assert-True -Condition ($customResultsCaptionIndex -gt $customResultsHeadingIndex) -Message 'Custom-profile image insertion did not place the second image under 实验现象记录.'
  Remove-Item -LiteralPath $customProfileInspect -Recurse -Force
  $results.Add('docx image insertion custom profile OK') | Out-Null

  $imageMappingFile = Join-Path $tempRoot 'image-map.json'
  @"
{
  "images": [
    {
      "anchor": "P5",
      "path": "$($sampleImageOne.Replace('\', '\\'))",
      "caption": "图1 实验目的示意图",
      "widthCm": 7.5
    },
    {
      "anchor": "T1R3C2",
      "path": "$($sampleImageTwo.Replace('\', '\\'))",
      "caption": "图2 表格截图示意",
      "widthCm": 6.0
    }
  ]
}
"@ | Set-Content -LiteralPath $imageMappingFile -Encoding UTF8

  $imageFilledDocx = Join-Path $tempRoot 'sample-template.images.docx'
  $imageInsertResult = & (Join-Path $repoRoot 'scripts\insert-docx-images.ps1') -DocxPath $sampleDocx -MappingPath $imageMappingFile -OutPath $imageFilledDocx
  Assert-True -Condition (Test-Path -LiteralPath $imageFilledDocx) -Message 'Image insertion script did not create the filled docx.'
  Assert-True -Condition ($imageInsertResult.insertedImageCount -eq 2) -Message 'Image insertion script inserted an unexpected number of images.'
  Assert-True -Condition ($imageInsertResult.insertedCaptionCount -eq 2) -Message 'Image insertion script inserted an unexpected number of captions.'

  $imageInsertTemp = Join-Path $tempRoot 'image-insert-inspect'
  [System.IO.Compression.ZipFile]::ExtractToDirectory($imageFilledDocx, $imageInsertTemp)
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $imageInsertTemp 'word\media\image1.png')) -Message 'Inserted docx is missing the first media image.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $imageInsertTemp 'word\media\image2.png')) -Message 'Inserted docx is missing the second media image.'
  [xml]$imageDocumentXml = [System.IO.File]::ReadAllText((Join-Path $imageInsertTemp 'word\document.xml'), (New-Object System.Text.UTF8Encoding($false)))
  [xml]$imageRelationshipsXml = [System.IO.File]::ReadAllText((Join-Path $imageInsertTemp 'word\_rels\document.xml.rels'), (New-Object System.Text.UTF8Encoding($false)))
  [xml]$imageContentTypesXml = [System.IO.File]::ReadAllText((Join-Path $imageInsertTemp '[Content_Types].xml'), (New-Object System.Text.UTF8Encoding($false)))
  $imageNamespaceManager = New-Object System.Xml.XmlNamespaceManager($imageDocumentXml.NameTable)
  $imageNamespaceManager.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
  $imageNamespaceManager.AddNamespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
  Assert-True -Condition (@($imageDocumentXml.SelectNodes('//wp:inline', $imageNamespaceManager)).Count -ge 2) -Message 'Inserted docx is missing expected image drawing nodes.'
  Assert-True -Condition ($imageDocumentXml.OuterXml -match '图1 实验目的示意图') -Message 'Inserted docx is missing the first image caption.'
  Assert-True -Condition ($imageDocumentXml.OuterXml -match '图2 表格截图示意') -Message 'Inserted docx is missing the second image caption.'
  Assert-True -Condition (@($imageRelationshipsXml.Relationships.Relationship | Where-Object { $_.Target -match '^media/image\d+\.png$' }).Count -ge 2) -Message 'Inserted docx is missing expected image relationships.'
  Assert-True -Condition (@($imageContentTypesXml.Types.Default | Where-Object { $_.Extension -eq 'png' -and $_.ContentType -eq 'image/png' }).Count -ge 1) -Message 'Inserted docx is missing the png content type registration.'
  $results.Add('docx image insertion OK') | Out-Null

  $inlineImageMappingJson = Get-Content -LiteralPath $generatedImageMapPath -Raw -Encoding UTF8
  $inlineImageFilledDocx = Join-Path $tempRoot 'sample-template.generated-filled.inline-images.docx'
  $inlineImageInsertResult = & (Join-Path $repoRoot 'scripts\insert-docx-images.ps1') -DocxPath $generatedFilledDocx -ImagesJson $inlineImageMappingJson -OutPath $inlineImageFilledDocx
  Assert-True -Condition (Test-Path -LiteralPath $inlineImageFilledDocx) -Message 'Inline image insertion did not create the filled docx.'
  Assert-True -Condition ([string]$inlineImageInsertResult.mappingInputMode -eq 'images-json') -Message 'Image insertion should record mappingInputMode=images-json for inline image-map JSON.'
  Assert-True -Condition ($inlineImageInsertResult.insertedImageCount -eq 2) -Message 'Inline image insertion inserted an unexpected number of images.'
  $results.Add('docx image insertion inline mapping OK') | Out-Null

  $layoutCheckPath = Join-Path $tempRoot 'sample-template.images.layout-check.json'
  & (Join-Path $repoRoot 'scripts\check-docx-layout.ps1') -DocxPath $imageFilledDocx -ExpectedImageCount 2 -ExpectedCaptionCount 2 -OutFile $layoutCheckPath | Out-Null
  $layoutCheck = (Get-Content -LiteralPath $layoutCheckPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$layoutCheck.reportProfileName -eq 'experiment-report') -Message 'Layout check is missing the expected report profile name.'
  Assert-True -Condition ([int]$layoutCheck.actual.imageDrawingCount -eq 2) -Message 'Layout check did not count the expected inserted images.'
  Assert-True -Condition ([int]$layoutCheck.actual.captionCount -eq 2) -Message 'Layout check did not count the expected figure captions.'
  Assert-True -Condition ([string]$layoutCheck.message -match 'Layout check failed') -Message 'Layout check did not include a readable failure message.'
  Assert-True -Condition (-not [bool]$layoutCheck.passed) -Message 'Layout check should fail when the image fixture still has template placeholders.'
  Assert-True -Condition (@($layoutCheck.errors | Where-Object { $_.code -eq 'remaining-placeholders' }).Count -eq 1) -Message 'Layout check did not report remaining placeholders in the image fixture.'
  $placeholderLayoutCheck = (& (Join-Path $repoRoot 'scripts\check-docx-layout.ps1') -DocxPath $sampleDocx -Format json | Out-String) | ConvertFrom-Json
  Assert-True -Condition (-not [bool]$placeholderLayoutCheck.passed) -Message 'Layout check should fail when the template still has placeholders.'
  Assert-True -Condition (@($placeholderLayoutCheck.errors | Where-Object { $_.code -eq 'remaining-placeholders' }).Count -eq 1) -Message 'Layout check did not report remaining placeholders.'
  $results.Add('docx layout check OK') | Out-Null

  $rowImageSpecsPath = Join-Path $tempRoot 'row-image-specs.json'
  @"
{
  "images": [
    {
      "path": "$($sampleImageOne.Replace('\', '\\'))",
      "section": "实验结果",
      "caption": "图1 主机 A 的 ping 测试结果",
      "widthCm": 10.2,
      "layout": {
        "mode": "row",
        "columns": 2,
        "group": "results-grid"
      }
    },
    {
      "path": "$($sampleImageTwo.Replace('\', '\\'))",
      "section": "实验结果",
      "caption": "图2 主机 B 的 ping 测试结果",
      "widthCm": 10.2,
      "layout": {
        "mode": "row",
        "columns": 2,
        "group": "results-grid"
      }
    },
    {
      "path": "$($sampleImageThree.Replace('\', '\\'))",
      "section": "实验结果",
      "caption": "图3 主机 A 的 arp -a 邻居缓存结果",
      "widthCm": 10.2,
      "layout": {
        "mode": "row",
        "columns": 2,
        "group": "results-grid"
      }
    },
    {
      "path": "$($sampleImageFour.Replace('\', '\\'))",
      "section": "实验结果",
      "caption": "图4 主机 B 的 arp -a 邻居缓存结果",
      "widthCm": 10.2,
      "layout": {
        "mode": "row",
        "columns": 2,
        "group": "results-grid"
      }
    }
  ]
}
"@ | Set-Content -LiteralPath $rowImageSpecsPath -Encoding UTF8

  $rowImageMapPath = Join-Path $tempRoot 'generated-row-image-map.json'
  & (Join-Path $repoRoot 'scripts\generate-docx-image-map.ps1') -DocxPath $generatedFilledDocx -ImageSpecsPath $rowImageSpecsPath -Format json -OutFile $rowImageMapPath | Out-Null
  $rowImageMapRoot = (Get-Content -LiteralPath $rowImageMapPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition (@($rowImageMapRoot.images).Count -eq 4) -Message 'Row image-map generator produced an unexpected number of images.'
  Assert-True -Condition ([string]$rowImageMapRoot.images[0].layout.mode -eq 'row') -Message 'Row image-map generator did not preserve the row layout mode.'
  Assert-True -Condition ([int]$rowImageMapRoot.images[0].layout.columns -eq 2) -Message 'Row image-map generator did not preserve the row layout column count.'
  $results.Add('docx image-map row layout generation OK') | Out-Null

  $rowImageFilledDocx = Join-Path $tempRoot 'sample-template.row-images.docx'
  $rowImageInsertResult = & (Join-Path $repoRoot 'scripts\insert-docx-images.ps1') -DocxPath $generatedFilledDocx -MappingPath $rowImageMapPath -OutPath $rowImageFilledDocx
  Assert-True -Condition (Test-Path -LiteralPath $rowImageFilledDocx) -Message 'Row-layout image insertion did not create the filled docx.'
  Assert-True -Condition ($rowImageInsertResult.insertedImageCount -eq 4) -Message 'Row-layout image insertion inserted an unexpected number of images.'
  Assert-True -Condition ($rowImageInsertResult.insertedCaptionCount -eq 4) -Message 'Row-layout image insertion inserted an unexpected number of captions.'

  $rowImageTemp = Join-Path $tempRoot 'row-image-inspect'
  [System.IO.Compression.ZipFile]::ExtractToDirectory($rowImageFilledDocx, $rowImageTemp)
  [xml]$rowImageDocumentXml = [System.IO.File]::ReadAllText((Join-Path $rowImageTemp 'word\document.xml'), (New-Object System.Text.UTF8Encoding($false)))
  $rowNamespaceManager = New-Object System.Xml.XmlNamespaceManager($rowImageDocumentXml.NameTable)
  $rowNamespaceManager.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
  $rowNamespaceManager.AddNamespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
  Assert-True -Condition (@($rowImageDocumentXml.SelectNodes('//w:tbl[.//wp:inline]', $rowNamespaceManager)).Count -ge 1) -Message 'Row-layout image insertion is missing the expected image table.'
  Assert-True -Condition (@($rowImageDocumentXml.SelectNodes('//wp:inline', $rowNamespaceManager)).Count -ge 4) -Message 'Row-layout image insertion is missing expected drawing nodes.'
  $rowImageWidthsCm = @($rowImageDocumentXml.SelectNodes('//wp:inline/wp:extent', $rowNamespaceManager) | ForEach-Object { [Math]::Round(([int64]$_.cx) / 360000.0, 2) })
  Assert-True -Condition (@($rowImageWidthsCm | Where-Object { $_ -gt 8.0 }).Count -eq 0) -Message 'Row-layout image insertion should cap over-wide images to the available column width.'
  Assert-True -Condition (@($rowImageWidthsCm | Where-Object { $_ -ge 10.2 }).Count -eq 0) -Message 'Row-layout image insertion did not shrink images that were too wide for two columns.'
  Assert-True -Condition ($rowImageDocumentXml.OuterXml -match '图1 主机 A 的 ping 测试结果') -Message 'Row-layout image insertion is missing the first row caption.'
  Assert-True -Condition ($rowImageDocumentXml.OuterXml -match '图4 主机 B 的 arp -a 邻居缓存结果') -Message 'Row-layout image insertion is missing the final row caption.'
  $rowImageDocumentText = $rowImageDocumentXml.OuterXml
  $resultBodyIndex = $rowImageDocumentText.IndexOf('通过 arp -a 可以看到对端主机的缓存记录')
  $firstRowCaptionIndex = $rowImageDocumentText.IndexOf('图1 主机 A 的 ping 测试结果')
  $finalRowCaptionIndex = $rowImageDocumentText.IndexOf('图4 主机 B 的 arp -a 邻居缓存结果')
  $sectionBoundaryIndex = $rowImageDocumentText.IndexOf('问题分析', $resultBodyIndex)
  if ($sectionBoundaryIndex -lt 0) {
    $sectionBoundaryIndex = $rowImageDocumentText.IndexOf('<w:sectPr', $resultBodyIndex)
  }
  Assert-True -Condition ($resultBodyIndex -ge 0 -and $firstRowCaptionIndex -gt $resultBodyIndex) -Message 'Row-layout image insertion should place section-targeted images after the section body, not immediately after the heading.'
  Assert-True -Condition ($sectionBoundaryIndex -gt $finalRowCaptionIndex) -Message 'Row-layout image insertion should keep section-targeted images before the next section boundary.'
  Remove-Item -LiteralPath $rowImageTemp -Recurse -Force
  $rowLayoutCheck = (& (Join-Path $repoRoot 'scripts\check-docx-layout.ps1') -DocxPath $rowImageFilledDocx -ExpectedImageCount 4 -ExpectedCaptionCount 4 -Format json | Out-String) | ConvertFrom-Json
  Assert-True -Condition ([bool]$rowLayoutCheck.passed) -Message 'Layout check should pass for the filled row-image fixture.'
  Assert-True -Condition ([string]$rowLayoutCheck.message -match 'Layout check passed') -Message 'Layout check did not include a readable pass message.'
  Assert-True -Condition ([int]$rowLayoutCheck.actual.imageDrawingCount -eq 4) -Message 'Layout check did not count the expected row-layout images.'
  Assert-True -Condition ([int]$rowLayoutCheck.actual.captionCount -eq 4) -Message 'Layout check did not count the expected row-layout captions.'
  Assert-True -Condition ([bool]$rowLayoutCheck.captionNumberCheck.continuous) -Message 'Layout check should report continuous row-layout captions.'

  $badCaptionDocx = Join-Path $tempRoot 'sample-template.bad-caption-number.docx'
  Copy-Item -LiteralPath $rowImageFilledDocx -Destination $badCaptionDocx -Force
  $badCaptionArchive = [System.IO.Compression.ZipFile]::Open($badCaptionDocx, [System.IO.Compression.ZipArchiveMode]::Update)
  try {
    $badCaptionDocumentEntry = $badCaptionArchive.GetEntry('word/document.xml')
    Assert-True -Condition ($null -ne $badCaptionDocumentEntry) -Message 'Bad-caption fixture is missing word/document.xml before mutation.'
    $badCaptionReader = New-Object System.IO.StreamReader($badCaptionDocumentEntry.Open(), (New-Object System.Text.UTF8Encoding($false)))
    try {
      $badCaptionDocumentText = $badCaptionReader.ReadToEnd()
    } finally {
      $badCaptionReader.Dispose()
    }
    $badCaptionDocumentEntry.Delete()
    $badCaptionDocumentText = $badCaptionDocumentText -replace '图2 主机 B 的 ping 测试结果', '图3 主机 B 的 ping 测试结果'
    $badCaptionDocumentEntry = $badCaptionArchive.CreateEntry('word/document.xml')
    $badCaptionWriter = New-Object System.IO.StreamWriter($badCaptionDocumentEntry.Open(), (New-Object System.Text.UTF8Encoding($false)))
    try {
      $badCaptionWriter.Write($badCaptionDocumentText)
    } finally {
      $badCaptionWriter.Dispose()
    }
  } finally {
    $badCaptionArchive.Dispose()
  }
  $badCaptionLayoutCheck = (& (Join-Path $repoRoot 'scripts\check-docx-layout.ps1') -DocxPath $badCaptionDocx -ExpectedImageCount 4 -ExpectedCaptionCount 4 -Format json | Out-String) | ConvertFrom-Json
  Assert-True -Condition (-not [bool]$badCaptionLayoutCheck.passed) -Message 'Layout check should fail when figure caption numbers are duplicated.'
  Assert-True -Condition (-not [bool]$badCaptionLayoutCheck.captionNumberCheck.continuous) -Message 'Layout check should report non-continuous caption numbers.'
  Assert-True -Condition (@($badCaptionLayoutCheck.errors | Where-Object { $_.code -eq 'caption-number-sequence' }).Count -eq 1) -Message 'Layout check did not report a caption-number sequence error.'
  $results.Add('docx image insertion row layout OK') | Out-Null

  $styledDocx = Join-Path $tempRoot 'sample-template.row-images.styled.docx'
  $styleResult = & (Join-Path $repoRoot 'scripts\format-docx-report-style.ps1') -DocxPath $rowImageFilledDocx -OutPath $styledDocx -Overwrite
  Assert-True -Condition (Test-Path -LiteralPath $styledDocx) -Message 'Docx style formatter did not create the styled docx.'
  Assert-True -Condition ([string]$styleResult.styleProfile -eq 'default') -Message 'Docx style formatter should default to the default style profile.'
  Assert-True -Condition ($styleResult.styledTitleCount -ge 1) -Message 'Docx style formatter did not detect the report title.'
  Assert-True -Condition ($styleResult.styledHeadingCount -ge 3) -Message 'Docx style formatter did not detect enough section headings.'
  Assert-True -Condition ($styleResult.styledBodyCount -ge 5) -Message 'Docx style formatter did not detect enough body paragraphs.'
  Assert-True -Condition ($styleResult.styledCaptionCount -ge 4) -Message 'Docx style formatter did not detect enough figure captions.'
  Assert-True -Condition ($styleResult.styledImageCount -ge 4) -Message 'Docx style formatter did not detect enough image paragraphs.'
  Assert-True -Condition ($styleResult.styledMetadataCount -ge 3) -Message 'Docx style formatter did not detect enough metadata paragraphs.'
  Assert-True -Condition ($styleResult.styledListCount -ge 2) -Message 'Docx style formatter did not detect enough numbered step paragraphs.'
  Assert-True -Condition ($styleResult.styledCommandCount -ge 6) -Message 'Docx style formatter did not detect enough command paragraphs.'
  Assert-True -Condition ($styleResult.styledTableParagraphCount -ge 4) -Message 'Docx style formatter did not style expected table paragraphs.'
  Assert-True -Condition ([int]$styleResult.appliedSettings.BodyLineTwips -eq 360) -Message 'Default style profile should keep the baseline body line spacing.'

  $styledDocxTemp = Join-Path $tempRoot 'styled-docx-inspect'
  [System.IO.Compression.ZipFile]::ExtractToDirectory($styledDocx, $styledDocxTemp)
  [xml]$styledDocumentXml = [System.IO.File]::ReadAllText((Join-Path $styledDocxTemp 'word\document.xml'), (New-Object System.Text.UTF8Encoding($false)))
  Assert-True -Condition ($styledDocumentXml.OuterXml -match 'w:jc w:val="center"') -Message 'Styled docx is missing centered paragraph formatting.'
  Assert-True -Condition ($styledDocumentXml.OuterXml -match 'w:ind w:firstLine="420"') -Message 'Styled docx is missing the expected first-line indent.'
  Assert-True -Condition ($styledDocumentXml.OuterXml -match 'w:b/?') -Message 'Styled docx is missing bold heading formatting.'
  Assert-True -Condition ($styledDocumentXml.OuterXml -match 'w:rFonts[^>]*Consolas') -Message 'Styled docx is missing Consolas command formatting.'
  Assert-True -Condition ($styledDocumentXml.OuterXml -match 'w:rFonts[^>]*黑体') -Message 'Styled docx is missing heading/title font formatting.'
  Assert-True -Condition ($styledDocumentXml.OuterXml -match 'w:rFonts[^>]*宋体') -Message 'Styled docx is missing body/caption font formatting.'
  Assert-True -Condition ($styledDocumentXml.OuterXml -match 'w:sz[^>]*w:val="24"') -Message 'Styled docx is missing 12pt body font sizing.'
  Assert-True -Condition ($styledDocumentXml.OuterXml -match 'w:sz[^>]*w:val="21"') -Message 'Styled docx is missing 10.5pt caption font sizing.'
  Assert-True -Condition ($styledDocumentXml.OuterXml -match 'w:shd[^>]*w:fill="F2F2F2"') -Message 'Styled docx is missing shaded command paragraphs.'
  Assert-True -Condition ($styledDocumentXml.OuterXml -match 'w:spacing[^>]*w:line="240"') -Message 'Styled docx is missing single-spaced command paragraphs.'
  Assert-True -Condition ($styledDocumentXml.OuterXml -match 'w:keepNext') -Message 'Styled docx is missing keep-next pagination hints.'
  Assert-True -Condition ($styledDocumentXml.OuterXml -match 'w:keepLines') -Message 'Styled docx is missing keep-lines pagination hints.'
  Assert-True -Condition ($styledDocumentXml.OuterXml -match 'w:tcMar') -Message 'Styled docx is missing table cell margin normalization.'
  Assert-True -Condition ($styledDocumentXml.OuterXml -match 'w:vAlign[^>]*w:val="top"') -Message 'Styled docx is missing top-aligned table cell formatting.'
  Remove-Item -LiteralPath $styledDocxTemp -Recurse -Force

  $compactStyledDocx = Join-Path $tempRoot 'sample-template.row-images.compact-styled.docx'
  $compactStyleResult = & (Join-Path $repoRoot 'scripts\format-docx-report-style.ps1') -DocxPath $rowImageFilledDocx -OutPath $compactStyledDocx -Overwrite -Profile compact
  Assert-True -Condition (Test-Path -LiteralPath $compactStyledDocx) -Message 'Compact style profile did not create the styled docx.'
  Assert-True -Condition ([string]$compactStyleResult.styleProfile -eq 'compact') -Message 'Compact style profile result is missing the selected profile name.'
  Assert-True -Condition ([int]$compactStyleResult.appliedSettings.BodyLineTwips -eq 320) -Message 'Compact style profile did not apply the expected tighter body line spacing.'
  Assert-True -Condition ([int]$compactStyleResult.appliedSettings.HeadingBeforeTwips -eq 80) -Message 'Compact style profile did not apply the expected heading spacing.'
  Assert-True -Condition ([int]$compactStyleResult.appliedSettings.TitleAfterTwips -eq 80) -Message 'Compact style profile did not apply the expected title spacing.'
  Assert-True -Condition ([int]$compactStyleResult.appliedSettings.BodyFontHalfPoints -eq 24) -Message 'Compact style profile did not apply the expected template-like body font size.'
  Assert-True -Condition ([int]$compactStyleResult.appliedSettings.HeadingFontHalfPoints -eq 30) -Message 'Compact style profile did not apply the expected template-like heading font size.'
  $compactStyledDocxTemp = Join-Path $tempRoot 'compact-styled-docx-inspect'
  [System.IO.Compression.ZipFile]::ExtractToDirectory($compactStyledDocx, $compactStyledDocxTemp)
  [xml]$compactStyledDocumentXml = [System.IO.File]::ReadAllText((Join-Path $compactStyledDocxTemp 'word\document.xml'), (New-Object System.Text.UTF8Encoding($false)))
  Assert-True -Condition ($compactStyledDocumentXml.OuterXml -match 'w:spacing[^>]*w:line="320"') -Message 'Compact styled docx is missing the expected compact line spacing.'
  Assert-True -Condition (-not ($compactStyledDocumentXml.OuterXml -match 'w:rFonts[^>]*(黑体|宋体|Consolas)')) -Message 'Compact styled docx should preserve template font families instead of forcing explicit fonts.'
  Assert-True -Condition (-not ($compactStyledDocumentXml.OuterXml -match 'w:sz[^>]*w:val="24"')) -Message 'Compact styled docx should preserve template font sizes instead of forcing 12pt direct sizing.'
  Assert-True -Condition (-not ($compactStyledDocumentXml.OuterXml -match 'w:keepNext')) -Message 'Compact styled docx should not force keep-next pagination hints.'
  Assert-True -Condition (-not ($compactStyledDocumentXml.OuterXml -match 'w:keepLines')) -Message 'Compact styled docx should not force keep-lines pagination hints.'
  Remove-Item -LiteralPath $compactStyledDocxTemp -Recurse -Force

  $customStyleProfilePath = Join-Path $tempRoot 'custom-style-profile.json'
  $customStyleProfile = [ordered]@{
    baseProfile = 'compact'
    settings = [ordered]@{
      BodyLineTwips = 290
      HeadingBeforeTwips = 60
      TitleAfterTwips = 50
    }
  }
  [System.IO.File]::WriteAllText($customStyleProfilePath, ($customStyleProfile | ConvertTo-Json -Depth 5), (New-Object System.Text.UTF8Encoding($true)))
  $customStyledDocx = Join-Path $tempRoot 'sample-template.row-images.custom-styled.docx'
  $customStyleResult = & (Join-Path $repoRoot 'scripts\format-docx-report-style.ps1') -DocxPath $rowImageFilledDocx -OutPath $customStyledDocx -Overwrite -ProfilePath $customStyleProfilePath -HeadingBeforeTwips 70
  Assert-True -Condition (Test-Path -LiteralPath $customStyledDocx) -Message 'Custom style profile file did not create the styled docx.'
  Assert-True -Condition ([string]$customStyleResult.requestedProfile -eq 'compact') -Message 'Custom style profile file did not supply the expected base profile.'
  Assert-True -Condition ([string]$customStyleResult.styleProfile -eq 'compact') -Message 'Custom style profile file did not resolve to the expected base profile.'
  Assert-True -Condition ([string]$customStyleResult.profilePath -eq $customStyleProfilePath) -Message 'Custom style profile result is missing the applied profile file path.'
  Assert-True -Condition ([int]$customStyleResult.appliedSettings.BodyLineTwips -eq 290) -Message 'Custom style profile file did not override body line spacing.'
  Assert-True -Condition ([int]$customStyleResult.appliedSettings.TitleAfterTwips -eq 50) -Message 'Custom style profile file did not override title spacing.'
  Assert-True -Condition ([int]$customStyleResult.appliedSettings.HeadingBeforeTwips -eq 70) -Message 'Explicit command-line style settings should override the custom style profile file.'
  $customStyledDocxTemp = Join-Path $tempRoot 'custom-styled-docx-inspect'
  [System.IO.Compression.ZipFile]::ExtractToDirectory($customStyledDocx, $customStyledDocxTemp)
  [xml]$customStyledDocumentXml = [System.IO.File]::ReadAllText((Join-Path $customStyledDocxTemp 'word\document.xml'), (New-Object System.Text.UTF8Encoding($false)))
  Assert-True -Condition ($customStyledDocumentXml.OuterXml -match 'w:spacing[^>]*w:line="290"') -Message 'Custom styled docx is missing the expected profile-file line spacing.'
  Remove-Item -LiteralPath $customStyledDocxTemp -Recurse -Force

  $autoCompactStyledDocx = Join-Path $tempRoot 'cover-body-template.auto-compact-styled.docx'
  $autoCompactStyleResult = & (Join-Path $repoRoot 'scripts\format-docx-report-style.ps1') -DocxPath $coverBodyFilledDocx -OutPath $autoCompactStyledDocx -Overwrite -Profile auto
  Assert-True -Condition (Test-Path -LiteralPath $autoCompactStyledDocx) -Message 'Auto style profile did not create the compact-styled cover-body docx.'
  Assert-True -Condition ([string]$autoCompactStyleResult.requestedProfile -eq 'auto') -Message 'Auto compact style result is missing the requested profile.'
  Assert-True -Condition ([string]$autoCompactStyleResult.styleProfile -eq 'compact') -Message 'Auto style profile should resolve the cover-body template to compact.'
  Assert-True -Condition ([string]$autoCompactStyleResult.profileReason -match 'cover-style metadata table') -Message 'Auto compact style result is missing the expected decision reason.'
  Assert-True -Condition ([int]$autoCompactStyleResult.appliedSettings.BodyLineTwips -eq 320) -Message 'Auto compact style result did not apply compact body spacing.'
  Assert-True -Condition ([int]$autoCompactStyleResult.appliedSettings.BodyFontHalfPoints -eq 24) -Message 'Auto compact style result did not apply compact body font size.'

  $paragraphCoverDocx = Join-Path $tempRoot 'paragraph-cover-template.docx'
  New-ParagraphCoverTemplateDocx -Path $paragraphCoverDocx
  Assert-True -Condition (Test-Path -LiteralPath $paragraphCoverDocx) -Message 'Failed to create the paragraph-cover template fixture.'
  $autoSchoolStyledDocx = Join-Path $tempRoot 'paragraph-cover-template.auto-school-styled.docx'
  $autoSchoolStyleResult = & (Join-Path $repoRoot 'scripts\format-docx-report-style.ps1') -DocxPath $paragraphCoverDocx -OutPath $autoSchoolStyledDocx -Overwrite -Profile auto
  Assert-True -Condition (Test-Path -LiteralPath $autoSchoolStyledDocx) -Message 'Auto style profile did not create the school-styled paragraph-cover docx.'
  Assert-True -Condition ([string]$autoSchoolStyleResult.styleProfile -eq 'school') -Message 'Auto style profile should resolve the paragraph-cover template to school.'
  Assert-True -Condition ([string]$autoSchoolStyleResult.profileReason -match 'paragraph-based cover area') -Message 'Auto school style result is missing the expected decision reason.'
  Assert-True -Condition ([int]$autoSchoolStyleResult.appliedSettings.BodyLineTwips -eq 400) -Message 'Auto school style result did not apply school body spacing.'
  $autoSchoolStyledDocxTemp = Join-Path $tempRoot 'auto-school-styled-docx-inspect'
  [System.IO.Compression.ZipFile]::ExtractToDirectory($autoSchoolStyledDocx, $autoSchoolStyledDocxTemp)
  [xml]$autoSchoolStyledDocumentXml = [System.IO.File]::ReadAllText((Join-Path $autoSchoolStyledDocxTemp 'word\document.xml'), (New-Object System.Text.UTF8Encoding($false)))
  Assert-True -Condition ($autoSchoolStyledDocumentXml.OuterXml -match 'w:spacing[^>]*w:line="400"') -Message 'Auto school styled docx is missing the expected school line spacing.'
  Remove-Item -LiteralPath $autoSchoolStyledDocxTemp -Recurse -Force

  $customProfileStyledDocx = Join-Path $tempRoot 'sample-template.custom-profile-sections.styled.docx'
  $customProfileStyleResult = & (Join-Path $repoRoot 'scripts\format-docx-report-style.ps1') `
    -DocxPath $customSectionDocx `
    -OutPath $customProfileStyledDocx `
    -Overwrite `
    -Profile auto `
    -ReportProfilePath $customImageProfilePath
  Assert-True -Condition (Test-Path -LiteralPath $customProfileStyledDocx) -Message 'Custom-profile style formatting did not create the styled docx.'
  Assert-True -Condition ([string]$customProfileStyleResult.reportProfileName -eq 'image-insert-custom-profile') -Message 'Custom-profile style formatting is missing the expected report profile name.'
  Assert-True -Condition ([string]$customProfileStyleResult.reportProfilePath -eq $resolvedCustomImageProfilePath) -Message 'Custom-profile style formatting is missing the expected report profile path.'
  Assert-True -Condition ([string]$customProfileStyleResult.styleProfile -eq 'compact') -Message 'Custom-profile style formatting should resolve to the report profile defaultStyleProfile.'
  Assert-True -Condition ([string]$customProfileStyleResult.profileReason -match 'defaultStyleProfile') -Message 'Custom-profile style formatting is missing the expected profile-default decision reason.'
  Assert-True -Condition ($customProfileStyleResult.styledHeadingCount -ge 2) -Message 'Custom-profile style formatting did not recognize the custom section headings.'
  $results.Add('docx report style formatting OK') | Out-Null

  $buildReportOutputDir = Join-Path $tempRoot 'build-report-output'
  $buildStyleProfilePath = Join-Path $tempRoot 'build-style-profile.json'
  $buildStyleProfile = [ordered]@{
    baseProfile = 'auto'
    settings = [ordered]@{
      BodyLineTwips = 310
      CaptionAfterTwips = 30
    }
  }
  [System.IO.File]::WriteAllText($buildStyleProfilePath, ($buildStyleProfile | ConvertTo-Json -Depth 5), (New-Object System.Text.UTF8Encoding($true)))
  $buildReportResult = & (Join-Path $repoRoot 'scripts\build-report.ps1') `
    -TemplatePath $sampleDocx `
    -ReportPath $sampleReportPath `
    -MetadataPath (Join-Path $repoRoot 'examples\docx-report-metadata.json') `
    -ImageSpecsPath $rowImageSpecsPath `
    -RequirementsPath (Join-Path $repoRoot 'examples\e2e-sample-requirements.json') `
    -OutputDir $buildReportOutputDir `
    -StyleFinalDocx `
    -CreateTemplateFrameDocx `
    -StyleProfilePath $buildStyleProfilePath
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $buildReportOutputDir 'generated-field-map.json')) -Message 'build-report did not create the generated field map.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $buildReportOutputDir 'sample-template.filled.docx')) -Message 'build-report did not create the filled docx.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $buildReportOutputDir 'image-placement-plan.md')) -Message 'build-report did not create the image placement plan.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $buildReportOutputDir 'sample-template.filled.images.docx')) -Message 'build-report did not create the image-filled docx.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $buildReportOutputDir 'sample-template.filled.images.styled.docx')) -Message 'build-report did not create the styled docx.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $buildReportOutputDir 'sample-template.filled.images.styled.template-frame.docx')) -Message 'build-report did not create the template-frame docx.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $buildReportOutputDir 'layout-check.json')) -Message 'build-report did not create the layout check JSON.'
  $buildReportSummary = (Get-Content -LiteralPath (Join-Path $buildReportOutputDir 'summary.json') -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([bool]$buildReportSummary.validationPassed) -Message 'build-report summary reported a failed validation result.'
  Assert-True -Condition ($buildReportSummary.PSObject.Properties.Name -contains 'validationFindingCountsByCode') -Message 'build-report summary is missing validationFindingCountsByCode.'
  Assert-True -Condition ($buildReportSummary.PSObject.Properties.Name -contains 'validationWarningCodes') -Message 'build-report summary is missing validationWarningCodes.'
  Assert-True -Condition ($buildReportSummary.PSObject.Properties.Name -contains 'validationWarningSummary') -Message 'build-report summary is missing validationWarningSummary.'
  Assert-True -Condition ($buildReportSummary.PSObject.Properties.Name -contains 'validationPaginationRiskCount') -Message 'build-report summary is missing validationPaginationRiskCount.'
  Assert-True -Condition ($buildReportSummary.PSObject.Properties.Name -contains 'validationStructuralIssueCount') -Message 'build-report summary is missing validationStructuralIssueCount.'
  Assert-True -Condition ([int]$buildReportSummary.validationPaginationRiskCount -eq 0) -Message 'build-report summary should report zero pagination risks for the passing sample.'
  Assert-True -Condition ([int]$buildReportSummary.validationStructuralIssueCount -eq 0) -Message 'build-report summary should report zero structural issues for the passing sample.'
  Assert-True -Condition (@($buildReportSummary.validationWarningCodes).Count -eq 0) -Message 'build-report summary should not report warning codes for the passing sample.'
  Assert-True -Condition (@($buildReportSummary.validationWarningSummary).Count -eq 0) -Message 'build-report summary should not report warning details for the passing sample.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$buildReportSummary.imagePlanPath)) -Message 'build-report summary is missing a readable image-plan path.'
  Assert-True -Condition ([int]$buildReportSummary.imagePlanLowConfidenceCount -eq 0) -Message 'build-report summary reported an unexpected low-confidence image-plan count.'
  Assert-True -Condition (-not [bool]$buildReportSummary.imagePlanNeedsReview) -Message 'build-report summary should not mark explicit image specs for manual review.'
  Assert-True -Condition ([bool]$buildReportSummary.layoutCheckPassed) -Message 'build-report summary reported a failed layout check.'
  Assert-True -Condition ([string]$buildReportSummary.layoutCheckMessage -match 'Layout check passed') -Message 'build-report summary is missing the readable layout-check message.'
  Assert-True -Condition ([int]$buildReportSummary.expectedLayoutImageCount -eq 4) -Message 'build-report summary is missing the expected layout image count.'
  Assert-True -Condition ([int]$buildReportSummary.expectedLayoutCaptionCount -eq 4) -Message 'build-report summary is missing the expected layout caption count.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$buildReportSummary.templateFrameDocxPath)) -Message 'build-report summary is missing a readable template-frame docx path.'
  $buildReportTemplateFrameOutline = & (Join-Path $repoRoot 'scripts\extract-docx-template.ps1') -Path ([string]$buildReportSummary.templateFrameDocxPath) -Format markdown | Out-String
  Assert-True -Condition ($buildReportTemplateFrameOutline -match 'Source:') -Message 'build-report template-frame docx could not be extracted.'
  Assert-True -Condition ([string]$buildReportSummary.reportProfileName -eq 'experiment-report') -Message 'build-report summary is missing the expected report profile name.'
  Assert-True -Condition ([string]$buildReportSummary.reportInputMode -eq 'path') -Message 'build-report summary should record reportInputMode=path for file-backed reports.'
  Assert-True -Condition ([string]$buildReportSummary.metadataInputMode -eq 'path') -Message 'build-report summary should record metadataInputMode=path for metadata files.'
  Assert-True -Condition ([string]$buildReportSummary.requirementsInputMode -eq 'path') -Message 'build-report summary should record requirementsInputMode=path for requirements files.'
  Assert-True -Condition ([string]$buildReportSummary.imageInputMode -eq 'specs-path') -Message 'build-report summary should record imageInputMode=specs-path for image spec files.'
  Assert-True -Condition ([string]$buildReportSummary.requestedStyleProfile -eq 'auto') -Message 'build-report should default the requested style profile from the report profile.'
  Assert-True -Condition ([string]$buildReportSummary.styleProfilePath -eq $buildStyleProfilePath) -Message 'build-report summary is missing the custom style profile path.'
  Assert-True -Condition ([string]$buildReportSummary.styleProfile -eq 'default') -Message 'build-report summary should resolve the sample template to the default style profile.'
  Assert-True -Condition ([string]$buildReportSummary.styleProfileReason -match 'default profile') -Message 'build-report summary is missing the resolved auto-style reason.'
  Assert-True -Condition ([int]$buildReportSummary.appliedStyleSettings.BodyLineTwips -eq 310) -Message 'build-report summary is missing the overridden style settings from the custom profile file.'
  Assert-True -Condition ((Split-Path -Leaf $buildReportSummary.finalDocxPath) -eq 'sample-template.filled.images.styled.docx') -Message 'build-report summary is missing the expected final docx path.'
  $results.Add('build-report pipeline OK') | Out-Null

  $paginationRiskDenseResult = ((@(
        '实验结果表明主机 A 与主机 B 的地址配置、连通测试、ARP 缓存和截图记录均保持一致，图1 展示主机 A 的 ipconfig 输出，图2 展示主机 B 的 ping 测试输出，图3 展示 ARP 缓存核对过程，因此本段故意保持为较长密集文本以触发分页风险 warning。'
      ) * 18) -join '')
  $paginationRiskReportPath = Join-Path $tempRoot 'build-pagination-risk-report.md'
  @(
    '计算机网络实验报告',
    '',
    '课程名称：计算机网络',
    '实验名称：局域网搭建与常用 DOS 命令使用',
    '',
    '一、实验目的',
    '本实验的目的是掌握局域网中静态地址配置、基础连通性检查和 DOS 网络命令使用方法，理解地址规划、命令输出和通信结果之间的对应关系。',
    '通过记录 ipconfig、ping 和 arp 等命令结果，可以把网络配置过程与验证结论连接起来，形成可复查的实验证据。',
    '',
    '二、实验环境',
    '实验环境包括 Windows 11 主机、两台 Windows Server 虚拟机、VMware 虚拟网络和 DOS 命令窗口，虚拟机均配置在同一网段并保持固定地址。',
    '实验前确认网络适配器启用、虚拟网络模式一致、主机名和 IP 地址记录清晰，避免由于环境差异影响连通性判断。',
    '',
    '三、实验原理或任务要求',
    '同一局域网内的主机需要具备一致的网络号和正确的子网掩码，通信过程中可以通过 ICMP 回显和 ARP 地址解析观察链路是否正常。',
    '任务要求依次完成地址配置、ipconfig 参数检查、ping 连通验证和 arp 缓存查看，并结合输出解释局域网通信是否建立。',
    '',
    '四、实验步骤',
    '先为主机 A 配置 192.168.10.11 地址，为主机 B 配置 192.168.10.12 地址，并确认两台主机子网掩码均为 255.255.255.0。',
    '随后在两台主机上分别执行 ipconfig、ping 和 arp -a 命令，记录关键输出并对比地址、网关、连通状态和缓存项是否符合预期。',
    '',
    '五、实验结果',
    $paginationRiskDenseResult,
    '',
    '六、问题分析',
    '如果 ping 不通，应优先检查 IP 地址、子网掩码、虚拟网卡模式和防火墙策略，再结合 arp 输出判断是否已经完成地址解析。',
    '如果只观察单次 ping 结果而忽略 ipconfig 和 arp 信息，可能遗漏网卡选错、地址冲突或缓存未更新等问题。',
    '',
    '七、实验总结',
    '本次实验完成了局域网搭建和常用 DOS 命令验证，能够从地址配置、连通测试和缓存记录三个角度说明实验结果。',
    '通过把命令输出与配置步骤逐项对应，进一步理解了局域网通信中地址规划、协议验证和故障定位之间的关系。'
  ) | Set-Content -LiteralPath $paginationRiskReportPath -Encoding UTF8
  $customLongPaginationRemediation = 'split the walkthrough into smaller subsections'
  $customDensePaginationRemediation = 'separate command output, interpretation, and conclusion'
  $customFigurePaginationRemediation = 'move a few screenshots into neighboring sections'
  $paginationRiskRequirements = [ordered]@{
    reportProfileName = 'experiment-report'
    courseName = '计算机网络'
    experimentName = '局域网搭建与常用 DOS 命令使用'
    minChars = 700
    sections = @(
      [ordered]@{ name = '实验目的'; aliases = @('实验目的'); minChars = 30 },
      [ordered]@{ name = '实验环境'; aliases = @('实验环境', '实验设备与环境'); minChars = 30 },
      [ordered]@{ name = '实验原理或任务要求'; aliases = @('实验原理或任务要求', '实验原理', '任务要求'); minChars = 30 },
      [ordered]@{ name = '实验步骤'; aliases = @('实验步骤', '实验过程'); minChars = 60 },
      [ordered]@{ name = '实验结果'; aliases = @('实验结果', '实验现象与结果记录'); minChars = 50 },
      [ordered]@{ name = '问题分析'; aliases = @('问题分析', '结果分析'); minChars = 30 },
      [ordered]@{ name = '实验总结'; aliases = @('实验总结', '总结与思考'); minChars = 30 }
    )
    forbiddenPatterns = @('TODO', '待补充', '自行填写')
    paginationRiskRemediations = [ordered]@{
      'pagination-risk-long-section' = 'Split the walkthrough into smaller subsections before docx generation, or raise paginationRiskThresholds.longSectionChars if this lab intentionally keeps one long narrative block.'
      'pagination-risk-dense-section-block' = 'Separate command output, interpretation, and conclusion into distinct paragraphs or list items before layout, or tune denseSectionChars and denseSectionParagraphs for this report family.'
      'pagination-risk-figure-cluster' = 'Move a few screenshots into neighboring sections or group them intentionally before layout, or raise paginationRiskThresholds.figureClusterRefs if screenshot-heavy evidence is expected.'
    }
  }
  $paginationRiskRequirementsJson = $paginationRiskRequirements | ConvertTo-Json -Depth 8

  $buildReportWarningOutputDir = Join-Path $tempRoot 'build-report-warning-output'
  & (Join-Path $repoRoot 'scripts\build-report.ps1') `
    -TemplatePath $sampleDocx `
    -ReportPath $paginationRiskReportPath `
    -MetadataPath (Join-Path $repoRoot 'examples\docx-report-metadata.json') `
    -ImageSpecsPath $rowImageSpecsPath `
    -RequirementsJson $paginationRiskRequirementsJson `
    -OutputDir $buildReportWarningOutputDir `
    -StyleFinalDocx `
    -StyleProfilePath $buildStyleProfilePath | Out-Null
  $buildReportWarningSummary = (Get-Content -LiteralPath (Join-Path $buildReportWarningOutputDir 'summary.json') -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-ValidationPaginationRiskSummary -Summary $buildReportWarningSummary -MessagePrefix 'build-report warning summary' -ExpectedLongRemediationPattern $customLongPaginationRemediation -ExpectedDenseRemediationPattern $customDensePaginationRemediation -ExpectedFigureRemediationPattern $customFigurePaginationRemediation
  Assert-True -Condition ([string]$buildReportWarningSummary.requirementsInputMode -eq 'inline') -Message 'build-report warning summary should record requirementsInputMode=inline.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$buildReportWarningSummary.validationPath)) -Message 'build-report warning summary should include a readable validation path.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$buildReportWarningSummary.finalDocxPath)) -Message 'build-report warning summary final docx path should exist.'
  $results.Add('build-report validation warning propagation OK') | Out-Null

  $buildReportInlineOutputDir = Join-Path $tempRoot 'build-report-inline-output'
  $buildReportInlineMetadataJson = Get-Content -LiteralPath (Join-Path $repoRoot 'examples\docx-report-metadata.json') -Raw -Encoding UTF8
  $buildReportInlineRequirementsJson = Get-Content -LiteralPath (Join-Path $repoRoot 'examples\e2e-sample-requirements.json') -Raw -Encoding UTF8
  $buildReportInlineImageSpecsJson = Get-Content -LiteralPath $rowImageSpecsPath -Raw -Encoding UTF8
  $buildReportInlineResult = & (Join-Path $repoRoot 'scripts\build-report.ps1') `
    -TemplatePath $sampleDocx `
    -ReportPath $sampleReportPath `
    -MetadataJson $buildReportInlineMetadataJson `
    -ImageSpecsJson $buildReportInlineImageSpecsJson `
    -RequirementsJson $buildReportInlineRequirementsJson `
    -OutputDir $buildReportInlineOutputDir `
    -StyleFinalDocx `
    -StyleProfilePath $buildStyleProfilePath
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $buildReportInlineOutputDir 'generated-field-map.json')) -Message 'Inline build-report did not create the generated field map.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $buildReportInlineOutputDir 'sample-template.filled.images.styled.docx')) -Message 'Inline build-report did not create the styled docx.'
  $buildReportInlineSummary = (Get-Content -LiteralPath (Join-Path $buildReportInlineOutputDir 'summary.json') -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$buildReportInlineSummary.reportInputMode -eq 'path') -Message 'Inline build-report should still record reportInputMode=path for file-backed reports.'
  Assert-True -Condition ([string]$buildReportInlineSummary.metadataInputMode -eq 'inline') -Message 'Inline build-report should record metadataInputMode=inline for inline metadata JSON.'
  Assert-True -Condition ([string]$buildReportInlineSummary.requirementsInputMode -eq 'inline') -Message 'Inline build-report should record requirementsInputMode=inline for inline requirements JSON.'
  Assert-True -Condition ([string]$buildReportInlineSummary.imageInputMode -eq 'specs-json') -Message 'Inline build-report should record imageInputMode=specs-json for inline image specs JSON.'
  Assert-True -Condition ([bool]$buildReportInlineSummary.validationPassed) -Message 'Inline build-report summary reported a failed validation result.'
  Assert-True -Condition ($buildReportInlineSummary.PSObject.Properties.Name -contains 'validationFindingCountsByCode') -Message 'Inline build-report summary is missing validationFindingCountsByCode.'
  Assert-True -Condition ($buildReportInlineSummary.PSObject.Properties.Name -contains 'validationWarningSummary') -Message 'Inline build-report summary is missing validationWarningSummary.'
  Assert-True -Condition ([bool]$buildReportInlineSummary.layoutCheckPassed) -Message 'Inline build-report summary reported a failed layout check.'
  Assert-True -Condition ([int]$buildReportInlineSummary.expectedLayoutImageCount -eq 4) -Message 'Inline build-report summary is missing the expected layout image count.'
  Assert-True -Condition ([int]$buildReportInlineSummary.expectedLayoutCaptionCount -eq 4) -Message 'Inline build-report summary is missing the expected layout caption count.'
  Assert-True -Condition ([string]$buildReportInlineSummary.styleProfile -eq 'default') -Message 'Inline build-report summary should resolve the sample template to the default style profile.'
  $results.Add('build-report inline inputs OK') | Out-Null

  $preparedSummaryMockReportPath = Join-Path $tempRoot 'course-design-generated-report.txt'
  @'
软件工程课程设计报告

课程名称：软件工程综合实践
课题名称：校园导览小程序设计
学生姓名：李四
学号：20261234
指导老师：王老师
完成时间：2026-04-08
设计地点：实验楼 A201

一、设计目标
本次课程设计面向新生入校后对校园空间陌生、目标地点分散、路线信息不明确的典型问题，设计并实现一个聚焦校园导览场景的小程序系统。系统目标不仅是展示地点列表，而是围绕“搜索地点、查看详情、获得路线、完成到达”这一条连续任务链路组织页面能力，保证用户在教学楼、实验楼、宿舍区和公共服务区之间切换时可以快速完成信息查询与路径判断。为了让课程设计报告体现完整的工程思路，目标部分还强调了界面清晰度、检索速度、地点信息准确性、可演示性和后续扩展能力五项约束，使小程序既能用于课堂答辩展示，也能作为后续校园导览产品原型继续迭代。

二、开发环境
项目开发环境采用 Windows 11 作为主机系统，前端使用微信开发者工具完成小程序页面开发与调试，后端接口模拟层采用 Node.js 运行时组织数据与逻辑，地点数据使用 SQLite 进行本地持久化存储。为了提升课程设计阶段的联调效率，项目额外配置了接口日志输出、静态资源目录、假数据回放脚本和本地构建命令，保证校园导览小程序在离线演示时依然能够稳定展示搜索、筛选、详情与路线提示结果。开发环境选择的核心考虑是学习成本适中、调试链路清晰、便于演示答辩时快速复现，因此页面样式、数据脚本和接口调试流程都围绕“小程序可重复运行、校园导览能力可直接观察”这一目标来组织。

三、需求分析
在需求分析阶段，首先从校园导览的真实使用场景出发，将用户需求拆分为地点检索、分类浏览、详情查看、推荐路线、收藏常用地点和异常提示六类核心能力。对新生用户而言，最重要的是在不熟悉校园布局的情况下，通过输入教学楼、食堂、图书馆或宿舍关键词快速找到目标位置，并在详情页看到楼宇简介、开放时间和相邻地标，从而降低迷路概率。对演示者而言，系统还需要在小程序首页突出搜索入口、分类卡片和推荐模块，让课程设计答辩时能够在短时间内清楚展示校园导览的主要流程。结合这些场景分析后，需求部分进一步明确了性能和准确性要求，即关键词搜索应尽量减少无关结果，路线提示应能给出可理解的步骤描述，收藏状态应在切换页面后保持一致，保证小程序不是单纯的页面堆叠，而是具备连续可用性的校园导览工具。

四、方案设计与实现
系统方案采用前后端分层结构。前端小程序负责首页分类导航、搜索结果列表、地点详情页、收藏状态展示和路线提示入口；后端数据层负责地点数据组织、关键词过滤、分类映射和路线推荐结果拼装。首页通过显眼的搜索框和功能卡片承接用户的第一步操作，搜索模块支持按关键词匹配地点名称、标签和描述字段，列表页增加了分类筛选与空结果提示，详情页则集中展示地点介绍、楼层信息、附近地标和收藏按钮，使校园导览链路在页面层面保持连贯。实现过程中，为了提升小程序在校园导览场景下的响应速度，项目对热门地点列表进行了本地缓存，对重复搜索结果进行了简单去重，并把路线提示描述抽象为可复用的数据结构，便于后续接入真实地图接口。课程设计实现部分还记录了组件拆分、接口模拟、状态同步、日志调试和异常提示的具体处理方式，说明系统不仅实现了可视化界面，还在工程组织、代码可维护性和演示稳定性方面做了针对性设计。

五、运行结果
完成编码与联调后，校园导览小程序已经能够稳定展示首页推荐地点、分类入口、关键词搜索结果、地点详情信息和基础路线提示流程。在实际演示中，输入“图书馆”“实验楼”“食堂”等关键词后，系统都能在较短时间内返回匹配结果，并在详情页正确显示地点简介、位置说明与收藏状态；当用户点击路线提示入口时，页面可以展示从当前位置到目标区域的文字化引导，满足课程设计对“能看、能查、能演示”的要求。运行结果部分还验证了收藏功能在页面切换后的状态保持、空结果场景下的提示信息和日志输出是否清晰，说明小程序不仅能够完成校园导览的核心流程，还对常见交互边界做了基础覆盖。整体来看，系统已经达到课程设计答辩所需的可运行、可观察、可解释状态。

六、问题与改进
虽然当前版本已经具备校园导览小程序的核心能力，但在路线推荐精度、地点数据规模和交互细节上仍存在继续优化空间。首先，当前路线提示主要依赖预设文本和简化规则，尚未接入真实地图服务，因此在复杂路径场景下缺少更细粒度的导航能力；其次，地点数据仍以课程设计阶段整理的样例数据为主，覆盖面有限，当校园导览扩展到更多教学楼、实验室和服务窗口时，需要进一步补充与维护。除此之外，页面交互也可以继续增强，例如为小程序加入最近搜索记录、按学院分组的地点入口、收藏夹批量管理和更明显的异常反馈。后续改进的方向应当继续围绕校园导览这一核心任务链路展开，而不是单纯堆叠功能项，确保新增能力能够真实提升用户在找地点、看详情和判断路线时的效率。

七、设计总结
通过本次课程设计，进一步理解了从需求分析、页面拆分、数据建模、接口模拟到联调演示的完整实现流程，也更加明确了“以用户任务链路为中心”对小程序设计的重要性。校园导览场景看似简单，但真正落地时需要同时处理搜索效率、信息组织、交互连续性、演示稳定性和后续扩展性等多方面问题，因此课程设计过程不仅锻炼了编码能力，也训练了围绕目标场景拆解需求和验证结果的能力。最终完成的校园导览小程序虽然仍有改进空间，但已经形成了一套结构清晰、功能闭环明确、适合课堂展示和后续迭代的实现方案。整个过程最大的收获，是学会了把课程设计报告中的分析、设计、实现、结果和改进真正对应到一个可运行的小程序系统上，而不是停留在概念层面的描述。
'@ | Set-Content -LiteralPath $preparedSummaryMockReportPath -Encoding UTF8

  $preparedSummaryBuildOutputDir = Join-Path $tempRoot 'prepared-summary-url-build-output'
  & (Join-Path $repoRoot 'scripts\build-report-from-url.ps1') `
    -TemplatePath $courseDesignTemplateDocx `
    -PreparedInputsSummaryPath $reportInputsSummaryPath `
    -ImageSpecsPath $courseDesignImageSpecsPath `
    -OutputDir $preparedSummaryBuildOutputDir `
    -StyleProfile auto `
    -CreateTemplateFrameDocx `
    -PreGeneratedReportPath $preparedSummaryMockReportPath `
    -SkipSessionReset | Out-Null
  $preparedSummaryBuildSummaryPath = Join-Path $preparedSummaryBuildOutputDir 'url-build-summary.json'
  Assert-True -Condition (Test-Path -LiteralPath $preparedSummaryBuildSummaryPath) -Message 'Prepared-summary URL wrapper did not create the wrapper summary.'
  $preparedSummaryBuildSummary = (Get-Content -LiteralPath $preparedSummaryBuildSummaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$preparedSummaryBuildSummary.reportProfileName -eq 'course-design-report') -Message 'Prepared-summary URL wrapper should inherit the course-design profile.'
  Assert-True -Condition ([string]$preparedSummaryBuildSummary.reportProfileDisplayName -eq '课程设计报告') -Message 'Prepared-summary URL wrapper should inherit the course-design display name.'
  Assert-True -Condition ([string]$preparedSummaryBuildSummary.detailLevel -eq 'full') -Message 'Prepared-summary URL wrapper should preserve the prepared-summary detail level.'
  Assert-True -Condition ([string]$preparedSummaryBuildSummary.generationMode -eq 'replay') -Message 'Prepared-summary URL wrapper should mark the report generation mode as replay.'
  Assert-True -Condition ([string]$preparedSummaryBuildSummary.reportInputsSummaryPath -eq $reportInputsSummaryPath) -Message 'Prepared-summary URL wrapper should keep the original input summary path.'
  Assert-True -Condition ([string]$preparedSummaryBuildSummary.buildReportInputMode -eq 'path') -Message 'Prepared-summary URL wrapper should expose buildReportInputMode=path from the downstream build summary.'
  Assert-True -Condition ([string]$preparedSummaryBuildSummary.buildMetadataInputMode -eq 'path') -Message 'Prepared-summary URL wrapper should expose buildMetadataInputMode=path from the downstream build summary.'
  Assert-True -Condition ([string]$preparedSummaryBuildSummary.buildRequirementsInputMode -eq 'path') -Message 'Prepared-summary URL wrapper should expose buildRequirementsInputMode=path from the downstream build summary.'
  Assert-True -Condition ([string]$preparedSummaryBuildSummary.buildImageInputMode -eq 'specs-path') -Message 'Prepared-summary URL wrapper should expose buildImageInputMode=specs-path from the downstream build summary.'
  Assert-True -Condition ([bool]$preparedSummaryBuildSummary.validationPassed) -Message 'Prepared-summary URL wrapper should expose validationPassed from the downstream build summary.'
  Assert-True -Condition ([int]$preparedSummaryBuildSummary.validationWarningCount -eq 0) -Message 'Prepared-summary URL wrapper should expose zero validation warnings for the passing fixture.'
  Assert-True -Condition ([int]$preparedSummaryBuildSummary.validationPaginationRiskCount -eq 0) -Message 'Prepared-summary URL wrapper should expose zero pagination risks for the passing fixture.'
  Assert-True -Condition ([int]$preparedSummaryBuildSummary.validationStructuralIssueCount -eq 0) -Message 'Prepared-summary URL wrapper should expose zero structural issues for the passing fixture.'
  Assert-True -Condition (@($preparedSummaryBuildSummary.validationWarningCodes).Count -eq 0) -Message 'Prepared-summary URL wrapper should expose an empty validationWarningCodes array for the passing fixture.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$preparedSummaryBuildSummary.pipelineTracePath)) -Message 'Prepared-summary URL wrapper should create a pipeline-trace JSON.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$preparedSummaryBuildSummary.pipelineTraceMarkdownPath)) -Message 'Prepared-summary URL wrapper should create a pipeline-trace markdown file.'
  Assert-True -Condition ([string]$preparedSummaryBuildSummary.requestedCourseName -eq '软件工程综合实践') -Message 'Prepared-summary URL wrapper lost the requested course name from the prepared summary.'
  Assert-True -Condition ([string]$preparedSummaryBuildSummary.requestedExperimentName -eq '校园导览小程序设计') -Message 'Prepared-summary URL wrapper lost the requested title from the prepared summary.'
  Assert-True -Condition ([string]$preparedSummaryBuildSummary.preGeneratedReportPath -eq $preparedSummaryMockReportPath) -Message 'Prepared-summary URL wrapper summary is missing the explicit replay report path.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$preparedSummaryBuildSummary.rawReportPath)) -Message 'Prepared-summary URL wrapper did not write the raw report file.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$preparedSummaryBuildSummary.cleanedReportPath)) -Message 'Prepared-summary URL wrapper did not write the cleaned report file.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$preparedSummaryBuildSummary.finalDocxPath)) -Message 'Prepared-summary URL wrapper final docx path does not exist.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$preparedSummaryBuildSummary.templateFrameDocxPath)) -Message 'Prepared-summary URL wrapper template-frame docx path does not exist.'
  $preparedSummaryTrace = (Get-Content -LiteralPath ([string]$preparedSummaryBuildSummary.pipelineTracePath) -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$preparedSummaryTrace.wrapper.generationMode -eq 'replay') -Message 'Prepared-summary URL pipeline trace should keep generationMode=replay.'
  Assert-True -Condition ([string]$preparedSummaryTrace.build.reportInputMode -eq 'path') -Message 'Prepared-summary URL pipeline trace should keep build.reportInputMode=path.'
  Assert-True -Condition ([string]$preparedSummaryTrace.build.imageInputMode -eq 'specs-path') -Message 'Prepared-summary URL pipeline trace should keep build.imageInputMode=specs-path.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$preparedSummaryTrace.artifacts.templateFrameDocxPath)) -Message 'Prepared-summary URL pipeline trace should expose the copied template-frame docx.'
  Assert-True -Condition ([bool]$preparedSummaryTrace.build.validationPassed) -Message 'Prepared-summary URL pipeline trace should expose validationPassed.'
  Assert-True -Condition ([int]$preparedSummaryTrace.build.validationPaginationRiskCount -eq 0) -Message 'Prepared-summary URL pipeline trace should expose zero pagination risks.'
  $preparedSummaryTraceMarkdown = Get-Content -LiteralPath ([string]$preparedSummaryBuildSummary.pipelineTraceMarkdownPath) -Raw -Encoding UTF8
  Assert-True -Condition ($preparedSummaryTraceMarkdown -match 'Generation mode: replay') -Message 'Prepared-summary URL pipeline markdown should include the replay generation mode.'
  Assert-True -Condition ($preparedSummaryTraceMarkdown -match 'Image input mode: specs-path') -Message 'Prepared-summary URL pipeline markdown should include the build image input mode.'
  Assert-True -Condition ($preparedSummaryTraceMarkdown -match 'Validation passed: True') -Message 'Prepared-summary URL pipeline markdown should include validation status.'
  Assert-True -Condition ($preparedSummaryTraceMarkdown -match 'Pagination risks: 0') -Message 'Prepared-summary URL pipeline markdown should include pagination risk count.'
  $preparedSummaryCleanedReport = Get-Content -LiteralPath ([string]$preparedSummaryBuildSummary.cleanedReportPath) -Raw -Encoding UTF8
  Assert-True -Condition ($preparedSummaryCleanedReport -match '方案设计与实现') -Message 'Prepared-summary URL wrapper cleaned report is missing the expected implementation heading.'
  $results.Add('build-report-from-url prepared summary OK') | Out-Null

  $monthlyPreparedSummaryDir = Join-Path $repoRoot 'examples\monthly-report-prepared'
  $monthlyPreparedSummaryPath = Join-Path $monthlyPreparedSummaryDir 'report-inputs-summary.json'
  $monthlyPreparedPromptPath = Join-Path $monthlyPreparedSummaryDir 'prompt.txt'
  $monthlyPreparedMetadataPath = Join-Path $monthlyPreparedSummaryDir 'metadata.auto.json'
  $monthlyPreparedRequirementsPath = Join-Path $monthlyPreparedSummaryDir 'requirements.auto.json'
  $monthlyPreparedDefaultsPath = Join-Path $monthlyPreparedSummaryDir 'defaults.snapshot.json'
  $monthlyPreparedReferencePath = Join-Path $monthlyPreparedSummaryDir 'references\project-context.txt'
  $monthlyPreparedReplayReportPath = Join-Path $monthlyPreparedSummaryDir 'report.replay.txt'
  foreach ($fixturePath in @(
      $monthlyPreparedSummaryPath,
      $monthlyPreparedPromptPath,
      $monthlyPreparedMetadataPath,
      $monthlyPreparedRequirementsPath,
      $monthlyPreparedDefaultsPath,
      $monthlyPreparedReferencePath,
      $monthlyPreparedReplayReportPath
    )) {
    Assert-True -Condition (Test-Path -LiteralPath $fixturePath -PathType Leaf) -Message ("Monthly prepared-summary fixture is missing: {0}" -f $fixturePath)
  }

  $monthlyPreparedBuildOutputDir = Join-Path $tempRoot 'monthly-prepared-summary-url-build-output'
  & (Join-Path $repoRoot 'scripts\build-report-from-url.ps1') `
    -TemplatePath $monthlyReportTemplateDocx `
    -PreparedInputsSummaryPath $monthlyPreparedSummaryPath `
    -OutputDir $monthlyPreparedBuildOutputDir `
    -StyleProfile auto `
    -PreGeneratedReportPath $monthlyPreparedReplayReportPath `
    -SkipSessionReset | Out-Null
  $monthlyPreparedBuildSummaryPath = Join-Path $monthlyPreparedBuildOutputDir 'url-build-summary.json'
  Assert-True -Condition (Test-Path -LiteralPath $monthlyPreparedBuildSummaryPath) -Message 'Monthly prepared-summary URL wrapper did not create the wrapper summary.'
  $monthlyPreparedBuildSummary = (Get-Content -LiteralPath $monthlyPreparedBuildSummaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$monthlyPreparedBuildSummary.reportProfileName -eq 'monthly-report') -Message 'Monthly prepared-summary URL wrapper should inherit the monthly-report profile.'
  Assert-True -Condition ([string]$monthlyPreparedBuildSummary.reportProfileDisplayName -eq '月报') -Message 'Monthly prepared-summary URL wrapper should inherit the monthly-report display name.'
  Assert-True -Condition ([string]$monthlyPreparedBuildSummary.reportProfilePath -eq (Join-Path $repoRoot 'profiles\monthly-report.json')) -Message 'Monthly prepared-summary URL wrapper should resolve the relative reportProfilePath.'
  Assert-True -Condition ([string]$monthlyPreparedBuildSummary.detailLevel -eq 'full') -Message 'Monthly prepared-summary URL wrapper should preserve the prepared-summary detail level.'
  Assert-True -Condition ([string]$monthlyPreparedBuildSummary.generationMode -eq 'replay') -Message 'Monthly prepared-summary URL wrapper should mark the report generation mode as replay.'
  Assert-True -Condition ([string]$monthlyPreparedBuildSummary.reportInputsSummaryPath -eq $monthlyPreparedSummaryPath) -Message 'Monthly prepared-summary URL wrapper should keep the original input summary path.'
  Assert-True -Condition ([string]$monthlyPreparedBuildSummary.promptPath -eq $monthlyPreparedPromptPath) -Message 'Monthly prepared-summary URL wrapper should resolve the relative prompt path.'
  Assert-True -Condition ([string]$monthlyPreparedBuildSummary.metadataPath -eq $monthlyPreparedMetadataPath) -Message 'Monthly prepared-summary URL wrapper should resolve the relative metadata path.'
  Assert-True -Condition ([string]$monthlyPreparedBuildSummary.requirementsPath -eq $monthlyPreparedRequirementsPath) -Message 'Monthly prepared-summary URL wrapper should resolve the relative requirements path.'
  Assert-True -Condition ([string]$monthlyPreparedBuildSummary.defaultsPath -eq $monthlyPreparedDefaultsPath) -Message 'Monthly prepared-summary URL wrapper should resolve the relative defaults path.'
  Assert-True -Condition (@($monthlyPreparedBuildSummary.referenceTextPaths).Count -eq 1) -Message 'Monthly prepared-summary URL wrapper should keep one resolved reference text path.'
  Assert-True -Condition ([string]@($monthlyPreparedBuildSummary.referenceTextPaths)[0] -eq $monthlyPreparedReferencePath) -Message 'Monthly prepared-summary URL wrapper should resolve the relative reference text path.'
  Assert-True -Condition (@($monthlyPreparedBuildSummary.fetchedReferenceTextPaths).Count -eq 0) -Message 'Monthly prepared-summary URL wrapper should keep fetchedReferenceTextPaths empty for local fixtures.'
  Assert-True -Condition ([string]$monthlyPreparedBuildSummary.requestedCourseName -eq 'OpenClaw 实验报告技能仓库') -Message 'Monthly prepared-summary URL wrapper lost the requested project name.'
  Assert-True -Condition ([string]$monthlyPreparedBuildSummary.requestedExperimentName -eq '2026 年 4 月升级月报') -Message 'Monthly prepared-summary URL wrapper lost the requested monthly title.'
  Assert-True -Condition ([string]$monthlyPreparedBuildSummary.preGeneratedReportPath -eq $monthlyPreparedReplayReportPath) -Message 'Monthly prepared-summary URL wrapper summary is missing the replay report path.'
  Assert-True -Condition ([string]$monthlyPreparedBuildSummary.buildReportInputMode -eq 'path') -Message 'Monthly prepared-summary URL wrapper should expose buildReportInputMode=path.'
  Assert-True -Condition ([string]$monthlyPreparedBuildSummary.buildMetadataInputMode -eq 'path') -Message 'Monthly prepared-summary URL wrapper should expose buildMetadataInputMode=path.'
  Assert-True -Condition ([string]$monthlyPreparedBuildSummary.buildRequirementsInputMode -eq 'path') -Message 'Monthly prepared-summary URL wrapper should expose buildRequirementsInputMode=path.'
  Assert-True -Condition ([bool]$monthlyPreparedBuildSummary.validationPassed) -Message 'Monthly prepared-summary URL wrapper should expose validationPassed=true.'
  Assert-True -Condition ([int]$monthlyPreparedBuildSummary.validationWarningCount -eq 0) -Message 'Monthly prepared-summary URL wrapper should expose zero validation warnings.'
  Assert-True -Condition ([int]$monthlyPreparedBuildSummary.validationPaginationRiskCount -eq 0) -Message 'Monthly prepared-summary URL wrapper should expose zero pagination risks.'
  Assert-True -Condition ([int]$monthlyPreparedBuildSummary.validationStructuralIssueCount -eq 0) -Message 'Monthly prepared-summary URL wrapper should expose zero structural issues.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$monthlyPreparedBuildSummary.pipelineTracePath)) -Message 'Monthly prepared-summary URL wrapper should create a pipeline-trace JSON.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$monthlyPreparedBuildSummary.pipelineTraceMarkdownPath)) -Message 'Monthly prepared-summary URL wrapper should create a pipeline-trace markdown file.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$monthlyPreparedBuildSummary.rawReportPath)) -Message 'Monthly prepared-summary URL wrapper did not write the raw report file.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$monthlyPreparedBuildSummary.cleanedReportPath)) -Message 'Monthly prepared-summary URL wrapper did not write the cleaned report file.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$monthlyPreparedBuildSummary.finalDocxPath)) -Message 'Monthly prepared-summary URL wrapper final docx path does not exist.'
  $monthlyPreparedTrace = (Get-Content -LiteralPath ([string]$monthlyPreparedBuildSummary.pipelineTracePath) -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$monthlyPreparedTrace.wrapper.generationMode -eq 'replay') -Message 'Monthly prepared-summary URL pipeline trace should keep generationMode=replay.'
  Assert-True -Condition ([string]$monthlyPreparedTrace.artifacts.promptPath -eq $monthlyPreparedPromptPath) -Message 'Monthly prepared-summary URL pipeline trace should expose the resolved prompt path.'
  Assert-True -Condition ([bool]$monthlyPreparedTrace.build.validationPassed) -Message 'Monthly prepared-summary URL pipeline trace should expose validationPassed.'
  Assert-True -Condition ([int]$monthlyPreparedTrace.build.validationPaginationRiskCount -eq 0) -Message 'Monthly prepared-summary URL pipeline trace should expose zero pagination risks.'
  $monthlyPreparedTraceMarkdown = Get-Content -LiteralPath ([string]$monthlyPreparedBuildSummary.pipelineTraceMarkdownPath) -Raw -Encoding UTF8
  Assert-True -Condition ($monthlyPreparedTraceMarkdown -match 'Generation mode: replay') -Message 'Monthly prepared-summary URL pipeline markdown should include the replay generation mode.'
  Assert-True -Condition ($monthlyPreparedTraceMarkdown -match 'Validation passed: True') -Message 'Monthly prepared-summary URL pipeline markdown should include validation status.'
  $monthlyPreparedCleanedReport = Get-Content -LiteralPath ([string]$monthlyPreparedBuildSummary.cleanedReportPath) -Raw -Encoding UTF8
  Assert-True -Condition ($monthlyPreparedCleanedReport -match '阶段成果与数据') -Message 'Monthly prepared-summary URL wrapper cleaned report is missing the expected results heading.'
  $results.Add('build-report-from-url monthly prepared summary OK') | Out-Null

  $meetingPreparedSummaryDir = Join-Path $repoRoot 'examples\meeting-minutes-prepared'
  $meetingPreparedSummaryPath = Join-Path $meetingPreparedSummaryDir 'report-inputs-summary.json'
  $meetingPreparedPromptPath = Join-Path $meetingPreparedSummaryDir 'prompt.txt'
  $meetingPreparedMetadataPath = Join-Path $meetingPreparedSummaryDir 'metadata.auto.json'
  $meetingPreparedRequirementsPath = Join-Path $meetingPreparedSummaryDir 'requirements.auto.json'
  $meetingPreparedDefaultsPath = Join-Path $meetingPreparedSummaryDir 'defaults.snapshot.json'
  $meetingPreparedReferencePath = Join-Path $meetingPreparedSummaryDir 'references\project-context.txt'
  $meetingPreparedReplayReportPath = Join-Path $meetingPreparedSummaryDir 'report.replay.txt'
  foreach ($fixturePath in @(
      $meetingPreparedSummaryPath,
      $meetingPreparedPromptPath,
      $meetingPreparedMetadataPath,
      $meetingPreparedRequirementsPath,
      $meetingPreparedDefaultsPath,
      $meetingPreparedReferencePath,
      $meetingPreparedReplayReportPath
    )) {
    Assert-True -Condition (Test-Path -LiteralPath $fixturePath -PathType Leaf) -Message ("Meeting-minutes prepared-summary fixture is missing: {0}" -f $fixturePath)
  }

  $meetingPreparedBuildOutputDir = Join-Path $tempRoot 'meeting-prepared-summary-url-build-output'
  & (Join-Path $repoRoot 'scripts\build-report-from-url.ps1') `
    -TemplatePath $meetingMinutesTemplateDocx `
    -PreparedInputsSummaryPath $meetingPreparedSummaryPath `
    -OutputDir $meetingPreparedBuildOutputDir `
    -StyleProfile auto `
    -PreGeneratedReportPath $meetingPreparedReplayReportPath `
    -SkipSessionReset | Out-Null
  $meetingPreparedBuildSummaryPath = Join-Path $meetingPreparedBuildOutputDir 'url-build-summary.json'
  Assert-True -Condition (Test-Path -LiteralPath $meetingPreparedBuildSummaryPath) -Message 'Meeting-minutes prepared-summary URL wrapper did not create the wrapper summary.'
  $meetingPreparedBuildSummary = (Get-Content -LiteralPath $meetingPreparedBuildSummaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$meetingPreparedBuildSummary.reportProfileName -eq 'meeting-minutes') -Message 'Meeting-minutes prepared-summary URL wrapper should inherit the meeting-minutes profile.'
  Assert-True -Condition ([string]$meetingPreparedBuildSummary.reportProfileDisplayName -eq '会议纪要') -Message 'Meeting-minutes prepared-summary URL wrapper should inherit the meeting-minutes display name.'
  Assert-True -Condition ([string]$meetingPreparedBuildSummary.reportProfilePath -eq (Join-Path $repoRoot 'profiles\meeting-minutes.json')) -Message 'Meeting-minutes prepared-summary URL wrapper should resolve the relative reportProfilePath.'
  Assert-True -Condition ([string]$meetingPreparedBuildSummary.detailLevel -eq 'full') -Message 'Meeting-minutes prepared-summary URL wrapper should preserve the prepared-summary detail level.'
  Assert-True -Condition ([string]$meetingPreparedBuildSummary.generationMode -eq 'replay') -Message 'Meeting-minutes prepared-summary URL wrapper should mark the report generation mode as replay.'
  Assert-True -Condition ([string]$meetingPreparedBuildSummary.reportInputsSummaryPath -eq $meetingPreparedSummaryPath) -Message 'Meeting-minutes prepared-summary URL wrapper should keep the original input summary path.'
  Assert-True -Condition ([string]$meetingPreparedBuildSummary.promptPath -eq $meetingPreparedPromptPath) -Message 'Meeting-minutes prepared-summary URL wrapper should resolve the relative prompt path.'
  Assert-True -Condition ([string]$meetingPreparedBuildSummary.metadataPath -eq $meetingPreparedMetadataPath) -Message 'Meeting-minutes prepared-summary URL wrapper should resolve the relative metadata path.'
  Assert-True -Condition ([string]$meetingPreparedBuildSummary.requirementsPath -eq $meetingPreparedRequirementsPath) -Message 'Meeting-minutes prepared-summary URL wrapper should resolve the relative requirements path.'
  Assert-True -Condition ([string]$meetingPreparedBuildSummary.defaultsPath -eq $meetingPreparedDefaultsPath) -Message 'Meeting-minutes prepared-summary URL wrapper should resolve the relative defaults path.'
  Assert-True -Condition (@($meetingPreparedBuildSummary.referenceTextPaths).Count -eq 1) -Message 'Meeting-minutes prepared-summary URL wrapper should keep one resolved reference text path.'
  Assert-True -Condition ([string]@($meetingPreparedBuildSummary.referenceTextPaths)[0] -eq $meetingPreparedReferencePath) -Message 'Meeting-minutes prepared-summary URL wrapper should resolve the relative reference text path.'
  Assert-True -Condition (@($meetingPreparedBuildSummary.fetchedReferenceTextPaths).Count -eq 0) -Message 'Meeting-minutes prepared-summary URL wrapper should keep fetchedReferenceTextPaths empty for local fixtures.'
  Assert-True -Condition ([string]$meetingPreparedBuildSummary.requestedCourseName -eq 'OpenClaw 实验报告技能仓库') -Message 'Meeting-minutes prepared-summary URL wrapper lost the requested project name.'
  Assert-True -Condition ([string]$meetingPreparedBuildSummary.requestedExperimentName -eq '2026 年 4 月路线评审会') -Message 'Meeting-minutes prepared-summary URL wrapper lost the requested meeting topic.'
  Assert-True -Condition ([string]$meetingPreparedBuildSummary.preGeneratedReportPath -eq $meetingPreparedReplayReportPath) -Message 'Meeting-minutes prepared-summary URL wrapper summary is missing the replay report path.'
  Assert-True -Condition ([string]$meetingPreparedBuildSummary.buildReportInputMode -eq 'path') -Message 'Meeting-minutes prepared-summary URL wrapper should expose buildReportInputMode=path.'
  Assert-True -Condition ([string]$meetingPreparedBuildSummary.buildMetadataInputMode -eq 'path') -Message 'Meeting-minutes prepared-summary URL wrapper should expose buildMetadataInputMode=path.'
  Assert-True -Condition ([string]$meetingPreparedBuildSummary.buildRequirementsInputMode -eq 'path') -Message 'Meeting-minutes prepared-summary URL wrapper should expose buildRequirementsInputMode=path.'
  Assert-True -Condition ([string]$meetingPreparedBuildSummary.buildImageInputMode -eq 'none') -Message 'Meeting-minutes prepared-summary URL wrapper should expose buildImageInputMode=none.'
  Assert-True -Condition ([bool]$meetingPreparedBuildSummary.validationPassed) -Message 'Meeting-minutes prepared-summary URL wrapper should expose validationPassed=true.'
  Assert-True -Condition ([int]$meetingPreparedBuildSummary.validationWarningCount -eq 0) -Message 'Meeting-minutes prepared-summary URL wrapper should expose zero validation warnings.'
  Assert-True -Condition ([int]$meetingPreparedBuildSummary.validationPaginationRiskCount -eq 0) -Message 'Meeting-minutes prepared-summary URL wrapper should expose zero pagination risks.'
  Assert-True -Condition ([int]$meetingPreparedBuildSummary.validationStructuralIssueCount -eq 0) -Message 'Meeting-minutes prepared-summary URL wrapper should expose zero structural issues.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$meetingPreparedBuildSummary.pipelineTracePath)) -Message 'Meeting-minutes prepared-summary URL wrapper should create a pipeline-trace JSON.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$meetingPreparedBuildSummary.pipelineTraceMarkdownPath)) -Message 'Meeting-minutes prepared-summary URL wrapper should create a pipeline-trace markdown file.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$meetingPreparedBuildSummary.rawReportPath)) -Message 'Meeting-minutes prepared-summary URL wrapper did not write the raw report file.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$meetingPreparedBuildSummary.cleanedReportPath)) -Message 'Meeting-minutes prepared-summary URL wrapper did not write the cleaned report file.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$meetingPreparedBuildSummary.finalDocxPath)) -Message 'Meeting-minutes prepared-summary URL wrapper final docx path does not exist.'
  $meetingPreparedTrace = (Get-Content -LiteralPath ([string]$meetingPreparedBuildSummary.pipelineTracePath) -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$meetingPreparedTrace.wrapper.generationMode -eq 'replay') -Message 'Meeting-minutes prepared-summary URL pipeline trace should keep generationMode=replay.'
  Assert-True -Condition ([string]$meetingPreparedTrace.artifacts.promptPath -eq $meetingPreparedPromptPath) -Message 'Meeting-minutes prepared-summary URL pipeline trace should expose the resolved prompt path.'
  Assert-True -Condition ([string]$meetingPreparedTrace.build.reportInputMode -eq 'path') -Message 'Meeting-minutes prepared-summary URL pipeline trace should keep build.reportInputMode=path.'
  Assert-True -Condition ([string]$meetingPreparedTrace.build.imageInputMode -eq 'none') -Message 'Meeting-minutes prepared-summary URL pipeline trace should keep build.imageInputMode=none.'
  Assert-True -Condition ([bool]$meetingPreparedTrace.build.validationPassed) -Message 'Meeting-minutes prepared-summary URL pipeline trace should expose validationPassed.'
  Assert-True -Condition ([int]$meetingPreparedTrace.build.validationPaginationRiskCount -eq 0) -Message 'Meeting-minutes prepared-summary URL pipeline trace should expose zero pagination risks.'
  $meetingPreparedTraceMarkdown = Get-Content -LiteralPath ([string]$meetingPreparedBuildSummary.pipelineTraceMarkdownPath) -Raw -Encoding UTF8
  Assert-True -Condition ($meetingPreparedTraceMarkdown -match 'Generation mode: replay') -Message 'Meeting-minutes prepared-summary URL pipeline markdown should include the replay generation mode.'
  Assert-True -Condition ($meetingPreparedTraceMarkdown -match 'Image input mode: none') -Message 'Meeting-minutes prepared-summary URL pipeline markdown should include the build image input mode.'
  Assert-True -Condition ($meetingPreparedTraceMarkdown -match 'Validation passed: True') -Message 'Meeting-minutes prepared-summary URL pipeline markdown should include validation status.'
  $meetingPreparedCleanedReport = Get-Content -LiteralPath ([string]$meetingPreparedBuildSummary.cleanedReportPath) -Raw -Encoding UTF8
  Assert-True -Condition ($meetingPreparedCleanedReport -match '讨论过程与决议') -Message 'Meeting-minutes prepared-summary URL wrapper cleaned report is missing the expected discussion heading.'
  $results.Add('build-report-from-url meeting-minutes prepared summary OK') | Out-Null

  $weeklyPreparedSummaryDir = Join-Path $repoRoot 'examples\weekly-report-prepared'
  $weeklyPreparedSummaryPath = Join-Path $weeklyPreparedSummaryDir 'report-inputs-summary.json'
  $weeklyPreparedPromptPath = Join-Path $weeklyPreparedSummaryDir 'prompt.txt'
  $weeklyPreparedMetadataPath = Join-Path $weeklyPreparedSummaryDir 'metadata.auto.json'
  $weeklyPreparedRequirementsPath = Join-Path $weeklyPreparedSummaryDir 'requirements.auto.json'
  $weeklyPreparedDefaultsPath = Join-Path $weeklyPreparedSummaryDir 'defaults.snapshot.json'
  $weeklyPreparedReferencePath = Join-Path $weeklyPreparedSummaryDir 'references\project-context.txt'
  $weeklyPreparedReplayReportPath = Join-Path $weeklyPreparedSummaryDir 'report.replay.txt'
  foreach ($fixturePath in @(
      $weeklyPreparedSummaryPath,
      $weeklyPreparedPromptPath,
      $weeklyPreparedMetadataPath,
      $weeklyPreparedRequirementsPath,
      $weeklyPreparedDefaultsPath,
      $weeklyPreparedReferencePath,
      $weeklyPreparedReplayReportPath
    )) {
    Assert-True -Condition (Test-Path -LiteralPath $fixturePath -PathType Leaf) -Message ("Weekly prepared-summary fixture is missing: {0}" -f $fixturePath)
  }

  $weeklyPreparedBuildOutputDir = Join-Path $tempRoot 'weekly-prepared-summary-url-build-output'
  & (Join-Path $repoRoot 'scripts\build-report-from-url.ps1') `
    -TemplatePath $weeklyReportTemplateDocx `
    -PreparedInputsSummaryPath $weeklyPreparedSummaryPath `
    -OutputDir $weeklyPreparedBuildOutputDir `
    -StyleProfile auto `
    -PreGeneratedReportPath $weeklyPreparedReplayReportPath `
    -SkipSessionReset | Out-Null
  $weeklyPreparedBuildSummaryPath = Join-Path $weeklyPreparedBuildOutputDir 'url-build-summary.json'
  Assert-True -Condition (Test-Path -LiteralPath $weeklyPreparedBuildSummaryPath) -Message 'Weekly prepared-summary URL wrapper did not create the wrapper summary.'
  $weeklyPreparedBuildSummary = (Get-Content -LiteralPath $weeklyPreparedBuildSummaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$weeklyPreparedBuildSummary.reportProfileName -eq 'weekly-report') -Message 'Weekly prepared-summary URL wrapper should inherit the weekly-report profile.'
  Assert-True -Condition ([string]$weeklyPreparedBuildSummary.reportProfileDisplayName -eq '周报') -Message 'Weekly prepared-summary URL wrapper should inherit the weekly-report display name.'
  Assert-True -Condition ([string]$weeklyPreparedBuildSummary.reportProfilePath -eq (Join-Path $repoRoot 'profiles\weekly-report.json')) -Message 'Weekly prepared-summary URL wrapper should resolve the relative reportProfilePath.'
  Assert-True -Condition ([string]$weeklyPreparedBuildSummary.detailLevel -eq 'full') -Message 'Weekly prepared-summary URL wrapper should preserve the prepared-summary detail level.'
  Assert-True -Condition ([string]$weeklyPreparedBuildSummary.generationMode -eq 'replay') -Message 'Weekly prepared-summary URL wrapper should mark the report generation mode as replay.'
  Assert-True -Condition ([string]$weeklyPreparedBuildSummary.reportInputsSummaryPath -eq $weeklyPreparedSummaryPath) -Message 'Weekly prepared-summary URL wrapper should keep the original input summary path.'
  Assert-True -Condition ([string]$weeklyPreparedBuildSummary.promptPath -eq $weeklyPreparedPromptPath) -Message 'Weekly prepared-summary URL wrapper should resolve the relative prompt path.'
  Assert-True -Condition ([string]$weeklyPreparedBuildSummary.metadataPath -eq $weeklyPreparedMetadataPath) -Message 'Weekly prepared-summary URL wrapper should resolve the relative metadata path.'
  Assert-True -Condition ([string]$weeklyPreparedBuildSummary.requirementsPath -eq $weeklyPreparedRequirementsPath) -Message 'Weekly prepared-summary URL wrapper should resolve the relative requirements path.'
  Assert-True -Condition ([string]$weeklyPreparedBuildSummary.defaultsPath -eq $weeklyPreparedDefaultsPath) -Message 'Weekly prepared-summary URL wrapper should resolve the relative defaults path.'
  Assert-True -Condition (@($weeklyPreparedBuildSummary.referenceTextPaths).Count -eq 1) -Message 'Weekly prepared-summary URL wrapper should keep one resolved reference text path.'
  Assert-True -Condition ([string]@($weeklyPreparedBuildSummary.referenceTextPaths)[0] -eq $weeklyPreparedReferencePath) -Message 'Weekly prepared-summary URL wrapper should resolve the relative reference text path.'
  Assert-True -Condition (@($weeklyPreparedBuildSummary.fetchedReferenceTextPaths).Count -eq 0) -Message 'Weekly prepared-summary URL wrapper should keep fetchedReferenceTextPaths empty for local fixtures.'
  Assert-True -Condition ([string]$weeklyPreparedBuildSummary.requestedCourseName -eq 'OpenClaw 实验报告技能仓库') -Message 'Weekly prepared-summary URL wrapper lost the requested project name.'
  Assert-True -Condition ([string]$weeklyPreparedBuildSummary.requestedExperimentName -eq '2026 年 4 月第 3 周维护周报') -Message 'Weekly prepared-summary URL wrapper lost the requested weekly title.'
  Assert-True -Condition ([string]$weeklyPreparedBuildSummary.preGeneratedReportPath -eq $weeklyPreparedReplayReportPath) -Message 'Weekly prepared-summary URL wrapper summary is missing the replay report path.'
  Assert-True -Condition ([string]$weeklyPreparedBuildSummary.buildReportInputMode -eq 'path') -Message 'Weekly prepared-summary URL wrapper should expose buildReportInputMode=path.'
  Assert-True -Condition ([string]$weeklyPreparedBuildSummary.buildMetadataInputMode -eq 'path') -Message 'Weekly prepared-summary URL wrapper should expose buildMetadataInputMode=path.'
  Assert-True -Condition ([string]$weeklyPreparedBuildSummary.buildRequirementsInputMode -eq 'path') -Message 'Weekly prepared-summary URL wrapper should expose buildRequirementsInputMode=path.'
  Assert-True -Condition ([string]$weeklyPreparedBuildSummary.buildImageInputMode -eq 'none') -Message 'Weekly prepared-summary URL wrapper should expose buildImageInputMode=none.'
  Assert-True -Condition ([bool]$weeklyPreparedBuildSummary.validationPassed) -Message 'Weekly prepared-summary URL wrapper should expose validationPassed=true.'
  Assert-True -Condition ([int]$weeklyPreparedBuildSummary.validationWarningCount -eq 0) -Message 'Weekly prepared-summary URL wrapper should expose zero validation warnings.'
  Assert-True -Condition ([int]$weeklyPreparedBuildSummary.validationPaginationRiskCount -eq 0) -Message 'Weekly prepared-summary URL wrapper should expose zero pagination risks.'
  Assert-True -Condition ([int]$weeklyPreparedBuildSummary.validationStructuralIssueCount -eq 0) -Message 'Weekly prepared-summary URL wrapper should expose zero structural issues.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$weeklyPreparedBuildSummary.pipelineTracePath)) -Message 'Weekly prepared-summary URL wrapper should create a pipeline-trace JSON.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$weeklyPreparedBuildSummary.pipelineTraceMarkdownPath)) -Message 'Weekly prepared-summary URL wrapper should create a pipeline-trace markdown file.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$weeklyPreparedBuildSummary.rawReportPath)) -Message 'Weekly prepared-summary URL wrapper did not write the raw report file.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$weeklyPreparedBuildSummary.cleanedReportPath)) -Message 'Weekly prepared-summary URL wrapper did not write the cleaned report file.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$weeklyPreparedBuildSummary.finalDocxPath)) -Message 'Weekly prepared-summary URL wrapper final docx path does not exist.'
  $weeklyPreparedTrace = (Get-Content -LiteralPath ([string]$weeklyPreparedBuildSummary.pipelineTracePath) -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$weeklyPreparedTrace.wrapper.generationMode -eq 'replay') -Message 'Weekly prepared-summary URL pipeline trace should keep generationMode=replay.'
  Assert-True -Condition ([string]$weeklyPreparedTrace.artifacts.promptPath -eq $weeklyPreparedPromptPath) -Message 'Weekly prepared-summary URL pipeline trace should expose the resolved prompt path.'
  Assert-True -Condition ([string]$weeklyPreparedTrace.build.reportInputMode -eq 'path') -Message 'Weekly prepared-summary URL pipeline trace should keep build.reportInputMode=path.'
  Assert-True -Condition ([string]$weeklyPreparedTrace.build.imageInputMode -eq 'none') -Message 'Weekly prepared-summary URL pipeline trace should keep build.imageInputMode=none.'
  Assert-True -Condition ([bool]$weeklyPreparedTrace.build.validationPassed) -Message 'Weekly prepared-summary URL pipeline trace should expose validationPassed.'
  Assert-True -Condition ([int]$weeklyPreparedTrace.build.validationPaginationRiskCount -eq 0) -Message 'Weekly prepared-summary URL pipeline trace should expose zero pagination risks.'
  $weeklyPreparedTraceMarkdown = Get-Content -LiteralPath ([string]$weeklyPreparedBuildSummary.pipelineTraceMarkdownPath) -Raw -Encoding UTF8
  Assert-True -Condition ($weeklyPreparedTraceMarkdown -match 'Generation mode: replay') -Message 'Weekly prepared-summary URL pipeline markdown should include the replay generation mode.'
  Assert-True -Condition ($weeklyPreparedTraceMarkdown -match 'Image input mode: none') -Message 'Weekly prepared-summary URL pipeline markdown should include the build image input mode.'
  Assert-True -Condition ($weeklyPreparedTraceMarkdown -match 'Validation passed: True') -Message 'Weekly prepared-summary URL pipeline markdown should include validation status.'
  $weeklyPreparedCleanedReport = Get-Content -LiteralPath ([string]$weeklyPreparedBuildSummary.cleanedReportPath) -Raw -Encoding UTF8
  Assert-True -Condition ($weeklyPreparedCleanedReport -match '阶段成果') -Message 'Weekly prepared-summary URL wrapper cleaned report is missing the expected results heading.'
  $results.Add('build-report-from-url weekly prepared summary OK') | Out-Null

  $softwareTestPreparedSummaryDir = Join-Path $repoRoot 'examples\software-test-report-prepared'
  $softwareTestPreparedSummaryPath = Join-Path $softwareTestPreparedSummaryDir 'report-inputs-summary.json'
  $softwareTestPreparedPromptPath = Join-Path $softwareTestPreparedSummaryDir 'prompt.txt'
  $softwareTestPreparedMetadataPath = Join-Path $softwareTestPreparedSummaryDir 'metadata.auto.json'
  $softwareTestPreparedRequirementsPath = Join-Path $softwareTestPreparedSummaryDir 'requirements.auto.json'
  $softwareTestPreparedDefaultsPath = Join-Path $softwareTestPreparedSummaryDir 'defaults.snapshot.json'
  $softwareTestPreparedReferencePath = Join-Path $softwareTestPreparedSummaryDir 'references\project-context.txt'
  $softwareTestPreparedReplayReportPath = Join-Path $softwareTestPreparedSummaryDir 'report.replay.txt'
  foreach ($fixturePath in @(
      $softwareTestPreparedSummaryPath,
      $softwareTestPreparedPromptPath,
      $softwareTestPreparedMetadataPath,
      $softwareTestPreparedRequirementsPath,
      $softwareTestPreparedDefaultsPath,
      $softwareTestPreparedReferencePath,
      $softwareTestPreparedReplayReportPath
    )) {
    Assert-True -Condition (Test-Path -LiteralPath $fixturePath -PathType Leaf) -Message ("Software-test prepared-summary fixture is missing: {0}" -f $fixturePath)
  }

  $softwareTestPreparedBuildOutputDir = Join-Path $tempRoot 'software-test-prepared-summary-url-build-output'
  & (Join-Path $repoRoot 'scripts\build-report-from-url.ps1') `
    -TemplatePath $softwareTestTemplateDocx `
    -PreparedInputsSummaryPath $softwareTestPreparedSummaryPath `
    -OutputDir $softwareTestPreparedBuildOutputDir `
    -StyleProfile auto `
    -PreGeneratedReportPath $softwareTestPreparedReplayReportPath `
    -SkipSessionReset | Out-Null
  $softwareTestPreparedBuildSummaryPath = Join-Path $softwareTestPreparedBuildOutputDir 'url-build-summary.json'
  Assert-True -Condition (Test-Path -LiteralPath $softwareTestPreparedBuildSummaryPath) -Message 'Software-test prepared-summary URL wrapper did not create the wrapper summary.'
  $softwareTestPreparedBuildSummary = (Get-Content -LiteralPath $softwareTestPreparedBuildSummaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$softwareTestPreparedBuildSummary.reportProfileName -eq 'software-test-report') -Message 'Software-test prepared-summary URL wrapper should inherit the software-test-report profile.'
  Assert-True -Condition ([string]$softwareTestPreparedBuildSummary.reportProfileDisplayName -eq '软件测试报告') -Message 'Software-test prepared-summary URL wrapper should inherit the software-test-report display name.'
  Assert-True -Condition ([string]$softwareTestPreparedBuildSummary.reportProfilePath -eq (Join-Path $repoRoot 'profiles\software-test-report.json')) -Message 'Software-test prepared-summary URL wrapper should resolve the relative reportProfilePath.'
  Assert-True -Condition ([string]$softwareTestPreparedBuildSummary.detailLevel -eq 'full') -Message 'Software-test prepared-summary URL wrapper should preserve the prepared-summary detail level.'
  Assert-True -Condition ([string]$softwareTestPreparedBuildSummary.generationMode -eq 'replay') -Message 'Software-test prepared-summary URL wrapper should mark the report generation mode as replay.'
  Assert-True -Condition ([string]$softwareTestPreparedBuildSummary.reportInputsSummaryPath -eq $softwareTestPreparedSummaryPath) -Message 'Software-test prepared-summary URL wrapper should keep the original input summary path.'
  Assert-True -Condition ([string]$softwareTestPreparedBuildSummary.promptPath -eq $softwareTestPreparedPromptPath) -Message 'Software-test prepared-summary URL wrapper should resolve the relative prompt path.'
  Assert-True -Condition ([string]$softwareTestPreparedBuildSummary.metadataPath -eq $softwareTestPreparedMetadataPath) -Message 'Software-test prepared-summary URL wrapper should resolve the relative metadata path.'
  Assert-True -Condition ([string]$softwareTestPreparedBuildSummary.requirementsPath -eq $softwareTestPreparedRequirementsPath) -Message 'Software-test prepared-summary URL wrapper should resolve the relative requirements path.'
  Assert-True -Condition ([string]$softwareTestPreparedBuildSummary.defaultsPath -eq $softwareTestPreparedDefaultsPath) -Message 'Software-test prepared-summary URL wrapper should resolve the relative defaults path.'
  Assert-True -Condition (@($softwareTestPreparedBuildSummary.referenceTextPaths).Count -eq 1) -Message 'Software-test prepared-summary URL wrapper should keep one resolved reference text path.'
  Assert-True -Condition ([string]@($softwareTestPreparedBuildSummary.referenceTextPaths)[0] -eq $softwareTestPreparedReferencePath) -Message 'Software-test prepared-summary URL wrapper should resolve the relative reference text path.'
  Assert-True -Condition (@($softwareTestPreparedBuildSummary.fetchedReferenceTextPaths).Count -eq 0) -Message 'Software-test prepared-summary URL wrapper should keep fetchedReferenceTextPaths empty for local fixtures.'
  Assert-True -Condition ([string]$softwareTestPreparedBuildSummary.requestedCourseName -eq 'OpenClaw 实验报告技能仓库') -Message 'Software-test prepared-summary URL wrapper lost the requested course name.'
  Assert-True -Condition ([string]$softwareTestPreparedBuildSummary.requestedExperimentName -eq 'prepared replay 回归链路测试') -Message 'Software-test prepared-summary URL wrapper lost the requested test title.'
  Assert-True -Condition ([string]$softwareTestPreparedBuildSummary.preGeneratedReportPath -eq $softwareTestPreparedReplayReportPath) -Message 'Software-test prepared-summary URL wrapper summary is missing the replay report path.'
  Assert-True -Condition ([string]$softwareTestPreparedBuildSummary.buildReportInputMode -eq 'path') -Message 'Software-test prepared-summary URL wrapper should expose buildReportInputMode=path.'
  Assert-True -Condition ([string]$softwareTestPreparedBuildSummary.buildMetadataInputMode -eq 'path') -Message 'Software-test prepared-summary URL wrapper should expose buildMetadataInputMode=path.'
  Assert-True -Condition ([string]$softwareTestPreparedBuildSummary.buildRequirementsInputMode -eq 'path') -Message 'Software-test prepared-summary URL wrapper should expose buildRequirementsInputMode=path.'
  Assert-True -Condition ([string]$softwareTestPreparedBuildSummary.buildImageInputMode -eq 'none') -Message 'Software-test prepared-summary URL wrapper should expose buildImageInputMode=none.'
  Assert-True -Condition ([bool]$softwareTestPreparedBuildSummary.validationPassed) -Message 'Software-test prepared-summary URL wrapper should expose validationPassed=true.'
  Assert-True -Condition ([int]$softwareTestPreparedBuildSummary.validationWarningCount -eq 0) -Message 'Software-test prepared-summary URL wrapper should expose zero validation warnings.'
  Assert-True -Condition ([int]$softwareTestPreparedBuildSummary.validationPaginationRiskCount -eq 0) -Message 'Software-test prepared-summary URL wrapper should expose zero pagination risks.'
  Assert-True -Condition ([int]$softwareTestPreparedBuildSummary.validationStructuralIssueCount -eq 0) -Message 'Software-test prepared-summary URL wrapper should expose zero structural issues.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$softwareTestPreparedBuildSummary.pipelineTracePath)) -Message 'Software-test prepared-summary URL wrapper should create a pipeline-trace JSON.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$softwareTestPreparedBuildSummary.pipelineTraceMarkdownPath)) -Message 'Software-test prepared-summary URL wrapper should create a pipeline-trace markdown file.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$softwareTestPreparedBuildSummary.rawReportPath)) -Message 'Software-test prepared-summary URL wrapper did not write the raw report file.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$softwareTestPreparedBuildSummary.cleanedReportPath)) -Message 'Software-test prepared-summary URL wrapper did not write the cleaned report file.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$softwareTestPreparedBuildSummary.finalDocxPath)) -Message 'Software-test prepared-summary URL wrapper final docx path does not exist.'
  $softwareTestPreparedTrace = (Get-Content -LiteralPath ([string]$softwareTestPreparedBuildSummary.pipelineTracePath) -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$softwareTestPreparedTrace.wrapper.generationMode -eq 'replay') -Message 'Software-test prepared-summary URL pipeline trace should keep generationMode=replay.'
  Assert-True -Condition ([string]$softwareTestPreparedTrace.artifacts.promptPath -eq $softwareTestPreparedPromptPath) -Message 'Software-test prepared-summary URL pipeline trace should expose the resolved prompt path.'
  Assert-True -Condition ([string]$softwareTestPreparedTrace.build.reportInputMode -eq 'path') -Message 'Software-test prepared-summary URL pipeline trace should keep build.reportInputMode=path.'
  Assert-True -Condition ([string]$softwareTestPreparedTrace.build.imageInputMode -eq 'none') -Message 'Software-test prepared-summary URL pipeline trace should keep build.imageInputMode=none.'
  Assert-True -Condition ([bool]$softwareTestPreparedTrace.build.validationPassed) -Message 'Software-test prepared-summary URL pipeline trace should expose validationPassed.'
  Assert-True -Condition ([int]$softwareTestPreparedTrace.build.validationPaginationRiskCount -eq 0) -Message 'Software-test prepared-summary URL pipeline trace should expose zero pagination risks.'
  $softwareTestPreparedTraceMarkdown = Get-Content -LiteralPath ([string]$softwareTestPreparedBuildSummary.pipelineTraceMarkdownPath) -Raw -Encoding UTF8
  Assert-True -Condition ($softwareTestPreparedTraceMarkdown -match 'Generation mode: replay') -Message 'Software-test prepared-summary URL pipeline markdown should include the replay generation mode.'
  Assert-True -Condition ($softwareTestPreparedTraceMarkdown -match 'Image input mode: none') -Message 'Software-test prepared-summary URL pipeline markdown should include the build image input mode.'
  Assert-True -Condition ($softwareTestPreparedTraceMarkdown -match 'Validation passed: True') -Message 'Software-test prepared-summary URL pipeline markdown should include validation status.'
  $softwareTestPreparedCleanedReport = Get-Content -LiteralPath ([string]$softwareTestPreparedBuildSummary.cleanedReportPath) -Raw -Encoding UTF8
  Assert-True -Condition ($softwareTestPreparedCleanedReport -match '测试结果') -Message 'Software-test prepared-summary URL wrapper cleaned report is missing the expected results heading.'
  $results.Add('build-report-from-url software-test prepared summary OK') | Out-Null

  $deploymentPreparedSummaryDir = Join-Path $repoRoot 'examples\deployment-report-prepared'
  $deploymentPreparedSummaryPath = Join-Path $deploymentPreparedSummaryDir 'report-inputs-summary.json'
  $deploymentPreparedPromptPath = Join-Path $deploymentPreparedSummaryDir 'prompt.txt'
  $deploymentPreparedMetadataPath = Join-Path $deploymentPreparedSummaryDir 'metadata.auto.json'
  $deploymentPreparedRequirementsPath = Join-Path $deploymentPreparedSummaryDir 'requirements.auto.json'
  $deploymentPreparedDefaultsPath = Join-Path $deploymentPreparedSummaryDir 'defaults.snapshot.json'
  $deploymentPreparedReferencePath = Join-Path $deploymentPreparedSummaryDir 'references\project-context.txt'
  $deploymentPreparedReplayReportPath = Join-Path $deploymentPreparedSummaryDir 'report.replay.txt'
  $deploymentPreparedImageSpecsPath = Join-Path $deploymentPreparedSummaryDir 'image-specs.json'
  $deploymentPreparedEnvironmentImagePath = Join-Path $deploymentPreparedSummaryDir 'images\deploy-env.png'
  $deploymentPreparedVerifyImagePath = Join-Path $deploymentPreparedSummaryDir 'images\deploy-verify.png'
  $deploymentPreparedRollbackImagePath = Join-Path $deploymentPreparedSummaryDir 'images\deploy-rollback.png'
  foreach ($fixturePath in @(
      $deploymentPreparedSummaryPath,
      $deploymentPreparedPromptPath,
      $deploymentPreparedMetadataPath,
      $deploymentPreparedRequirementsPath,
      $deploymentPreparedDefaultsPath,
      $deploymentPreparedReferencePath,
      $deploymentPreparedReplayReportPath,
      $deploymentPreparedImageSpecsPath,
      $deploymentPreparedEnvironmentImagePath,
      $deploymentPreparedVerifyImagePath,
      $deploymentPreparedRollbackImagePath
    )) {
    Assert-True -Condition (Test-Path -LiteralPath $fixturePath -PathType Leaf) -Message ("Deployment prepared-summary fixture is missing: {0}" -f $fixturePath)
  }

  $deploymentPreparedBuildOutputDir = Join-Path $tempRoot 'deployment-prepared-summary-url-build-output'
  & (Join-Path $repoRoot 'scripts\build-report-from-url.ps1') `
    -TemplatePath $deploymentTemplateDocx `
    -PreparedInputsSummaryPath $deploymentPreparedSummaryPath `
    -ImageSpecsPath $deploymentPreparedImageSpecsPath `
    -OutputDir $deploymentPreparedBuildOutputDir `
    -StyleProfile auto `
    -PreGeneratedReportPath $deploymentPreparedReplayReportPath `
    -SkipSessionReset | Out-Null
  $deploymentPreparedBuildSummaryPath = Join-Path $deploymentPreparedBuildOutputDir 'url-build-summary.json'
  Assert-True -Condition (Test-Path -LiteralPath $deploymentPreparedBuildSummaryPath) -Message 'Deployment prepared-summary URL wrapper did not create the wrapper summary.'
  $deploymentPreparedBuildSummary = (Get-Content -LiteralPath $deploymentPreparedBuildSummaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$deploymentPreparedBuildSummary.reportProfileName -eq 'deployment-report') -Message 'Deployment prepared-summary URL wrapper should inherit the deployment-report profile.'
  Assert-True -Condition ([string]$deploymentPreparedBuildSummary.reportProfileDisplayName -eq '部署运维报告') -Message 'Deployment prepared-summary URL wrapper should inherit the deployment-report display name.'
  Assert-True -Condition ([string]$deploymentPreparedBuildSummary.reportProfilePath -eq (Join-Path $repoRoot 'profiles\deployment-report.json')) -Message 'Deployment prepared-summary URL wrapper should resolve the relative reportProfilePath.'
  Assert-True -Condition ([string]$deploymentPreparedBuildSummary.detailLevel -eq 'full') -Message 'Deployment prepared-summary URL wrapper should preserve the prepared-summary detail level.'
  Assert-True -Condition ([string]$deploymentPreparedBuildSummary.generationMode -eq 'replay') -Message 'Deployment prepared-summary URL wrapper should mark the report generation mode as replay.'
  Assert-True -Condition ([string]$deploymentPreparedBuildSummary.reportInputsSummaryPath -eq $deploymentPreparedSummaryPath) -Message 'Deployment prepared-summary URL wrapper should keep the original input summary path.'
  Assert-True -Condition ([string]$deploymentPreparedBuildSummary.promptPath -eq $deploymentPreparedPromptPath) -Message 'Deployment prepared-summary URL wrapper should resolve the relative prompt path.'
  Assert-True -Condition ([string]$deploymentPreparedBuildSummary.metadataPath -eq $deploymentPreparedMetadataPath) -Message 'Deployment prepared-summary URL wrapper should resolve the relative metadata path.'
  Assert-True -Condition ([string]$deploymentPreparedBuildSummary.requirementsPath -eq $deploymentPreparedRequirementsPath) -Message 'Deployment prepared-summary URL wrapper should resolve the relative requirements path.'
  Assert-True -Condition ([string]$deploymentPreparedBuildSummary.defaultsPath -eq $deploymentPreparedDefaultsPath) -Message 'Deployment prepared-summary URL wrapper should resolve the relative defaults path.'
  Assert-True -Condition (@($deploymentPreparedBuildSummary.referenceTextPaths).Count -eq 1) -Message 'Deployment prepared-summary URL wrapper should keep one resolved reference text path.'
  Assert-True -Condition ([string]@($deploymentPreparedBuildSummary.referenceTextPaths)[0] -eq $deploymentPreparedReferencePath) -Message 'Deployment prepared-summary URL wrapper should resolve the relative reference text path.'
  Assert-True -Condition (@($deploymentPreparedBuildSummary.fetchedReferenceTextPaths).Count -eq 0) -Message 'Deployment prepared-summary URL wrapper should keep fetchedReferenceTextPaths empty for local fixtures.'
  Assert-True -Condition ([string]$deploymentPreparedBuildSummary.requestedCourseName -eq 'OpenClaw 实验报告技能仓库') -Message 'Deployment prepared-summary URL wrapper lost the requested course name.'
  Assert-True -Condition ([string]$deploymentPreparedBuildSummary.requestedExperimentName -eq 'prepared replay 部署回放链路测试') -Message 'Deployment prepared-summary URL wrapper lost the requested deployment title.'
  Assert-True -Condition ([string]$deploymentPreparedBuildSummary.preGeneratedReportPath -eq $deploymentPreparedReplayReportPath) -Message 'Deployment prepared-summary URL wrapper summary is missing the replay report path.'
  Assert-True -Condition ([string]$deploymentPreparedBuildSummary.buildReportInputMode -eq 'path') -Message 'Deployment prepared-summary URL wrapper should expose buildReportInputMode=path.'
  Assert-True -Condition ([string]$deploymentPreparedBuildSummary.buildMetadataInputMode -eq 'path') -Message 'Deployment prepared-summary URL wrapper should expose buildMetadataInputMode=path.'
  Assert-True -Condition ([string]$deploymentPreparedBuildSummary.buildRequirementsInputMode -eq 'path') -Message 'Deployment prepared-summary URL wrapper should expose buildRequirementsInputMode=path.'
  Assert-True -Condition ([string]$deploymentPreparedBuildSummary.buildImageInputMode -eq 'specs-path') -Message 'Deployment prepared-summary URL wrapper should expose buildImageInputMode=specs-path.'
  Assert-True -Condition ([bool]$deploymentPreparedBuildSummary.validationPassed) -Message 'Deployment prepared-summary URL wrapper should expose validationPassed=true.'
  Assert-True -Condition ([int]$deploymentPreparedBuildSummary.validationWarningCount -eq 0) -Message 'Deployment prepared-summary URL wrapper should expose zero validation warnings.'
  Assert-True -Condition ([int]$deploymentPreparedBuildSummary.validationPaginationRiskCount -eq 0) -Message 'Deployment prepared-summary URL wrapper should expose zero pagination risks.'
  Assert-True -Condition ([int]$deploymentPreparedBuildSummary.validationStructuralIssueCount -eq 0) -Message 'Deployment prepared-summary URL wrapper should expose zero structural issues.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$deploymentPreparedBuildSummary.pipelineTracePath)) -Message 'Deployment prepared-summary URL wrapper should create a pipeline-trace JSON.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$deploymentPreparedBuildSummary.pipelineTraceMarkdownPath)) -Message 'Deployment prepared-summary URL wrapper should create a pipeline-trace markdown file.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$deploymentPreparedBuildSummary.buildSummaryPath)) -Message 'Deployment prepared-summary URL wrapper should keep the downstream build summary path.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$deploymentPreparedBuildSummary.imagePlanPath)) -Message 'Deployment prepared-summary URL wrapper should expose the generated image-plan path.'
  Assert-True -Condition ([int]$deploymentPreparedBuildSummary.imagePlanLowConfidenceCount -eq 0) -Message 'Deployment prepared-summary URL wrapper reported an unexpected low-confidence image-plan count.'
  Assert-True -Condition (-not [bool]$deploymentPreparedBuildSummary.imagePlanNeedsReview) -Message 'Deployment prepared-summary URL wrapper should not require manual image review for explicit image specs.'
  Assert-True -Condition ([bool]$deploymentPreparedBuildSummary.layoutCheckPassed) -Message 'Deployment prepared-summary URL wrapper reported a failed layout check.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$deploymentPreparedBuildSummary.layoutCheckPath)) -Message 'Deployment prepared-summary URL wrapper should expose the layout-check path.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$deploymentPreparedBuildSummary.rawReportPath)) -Message 'Deployment prepared-summary URL wrapper did not write the raw report file.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$deploymentPreparedBuildSummary.cleanedReportPath)) -Message 'Deployment prepared-summary URL wrapper did not write the cleaned report file.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$deploymentPreparedBuildSummary.finalDocxPath)) -Message 'Deployment prepared-summary URL wrapper final docx path does not exist.'
  $deploymentPreparedPipelineBuildSummary = (Get-Content -LiteralPath ([string]$deploymentPreparedBuildSummary.buildSummaryPath) -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$deploymentPreparedPipelineBuildSummary.reportProfileName -eq 'deployment-report') -Message 'Deployment prepared-summary downstream build summary should keep the deployment-report profile.'
  Assert-True -Condition ([string]$deploymentPreparedPipelineBuildSummary.imageInputMode -eq 'specs-path') -Message 'Deployment prepared-summary downstream build summary should keep imageInputMode=specs-path.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$deploymentPreparedPipelineBuildSummary.imagePlanPath)) -Message 'Deployment prepared-summary downstream build summary should expose the image-plan path.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$deploymentPreparedPipelineBuildSummary.imageMapPath)) -Message 'Deployment prepared-summary downstream build summary should expose the image-map path.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$deploymentPreparedPipelineBuildSummary.filledDocxWithImagesPath)) -Message 'Deployment prepared-summary downstream build summary should expose the filled-docx-with-images path.'
  Assert-True -Condition ([int]$deploymentPreparedPipelineBuildSummary.imagePlanLowConfidenceCount -eq 0) -Message 'Deployment prepared-summary downstream build summary reported an unexpected low-confidence image-plan count.'
  Assert-True -Condition (-not [bool]$deploymentPreparedPipelineBuildSummary.imagePlanNeedsReview) -Message 'Deployment prepared-summary downstream build summary should not require manual image review for explicit image specs.'
  Assert-True -Condition ([bool]$deploymentPreparedPipelineBuildSummary.layoutCheckPassed) -Message 'Deployment prepared-summary downstream build summary reported a failed layout check.'
  $deploymentPreparedTrace = (Get-Content -LiteralPath ([string]$deploymentPreparedBuildSummary.pipelineTracePath) -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$deploymentPreparedTrace.wrapper.generationMode -eq 'replay') -Message 'Deployment prepared-summary URL pipeline trace should keep generationMode=replay.'
  Assert-True -Condition ([string]$deploymentPreparedTrace.artifacts.promptPath -eq $deploymentPreparedPromptPath) -Message 'Deployment prepared-summary URL pipeline trace should expose the resolved prompt path.'
  Assert-True -Condition ([string]$deploymentPreparedTrace.build.reportInputMode -eq 'path') -Message 'Deployment prepared-summary URL pipeline trace should keep build.reportInputMode=path.'
  Assert-True -Condition ([string]$deploymentPreparedTrace.build.imageInputMode -eq 'specs-path') -Message 'Deployment prepared-summary URL pipeline trace should keep build.imageInputMode=specs-path.'
  Assert-True -Condition ([bool]$deploymentPreparedTrace.build.validationPassed) -Message 'Deployment prepared-summary URL pipeline trace should expose validationPassed.'
  Assert-True -Condition ([int]$deploymentPreparedTrace.build.validationPaginationRiskCount -eq 0) -Message 'Deployment prepared-summary URL pipeline trace should expose zero pagination risks.'
  $deploymentPreparedTraceMarkdown = Get-Content -LiteralPath ([string]$deploymentPreparedBuildSummary.pipelineTraceMarkdownPath) -Raw -Encoding UTF8
  Assert-True -Condition ($deploymentPreparedTraceMarkdown -match 'Generation mode: replay') -Message 'Deployment prepared-summary URL pipeline markdown should include the replay generation mode.'
  Assert-True -Condition ($deploymentPreparedTraceMarkdown -match 'Image input mode: specs-path') -Message 'Deployment prepared-summary URL pipeline markdown should include the build image input mode.'
  Assert-True -Condition ($deploymentPreparedTraceMarkdown -match 'Validation passed: True') -Message 'Deployment prepared-summary URL pipeline markdown should include validation status.'
  $deploymentPreparedCleanedReport = Get-Content -LiteralPath ([string]$deploymentPreparedBuildSummary.cleanedReportPath) -Raw -Encoding UTF8
  Assert-True -Condition ($deploymentPreparedCleanedReport -match '验证结果') -Message 'Deployment prepared-summary URL wrapper cleaned report is missing the expected results heading.'
  $results.Add('build-report-from-url deployment prepared summary OK') | Out-Null

  $urlWarningOutputDir = Join-Path $tempRoot 'url-warning-build-output'
  & (Join-Path $repoRoot 'scripts\build-report-from-url.ps1') `
    -TemplatePath $sampleDocx `
    -PromptText '/experiment-report 生成一份局域网搭建实验报告。' `
    -CourseName '计算机网络' `
    -ExperimentName '局域网搭建与常用 DOS 命令使用' `
    -MetadataPath (Join-Path $repoRoot 'examples\docx-report-metadata.json') `
    -ImageSpecsPath $rowImageSpecsPath `
    -RequirementsJson $paginationRiskRequirementsJson `
    -OutputDir $urlWarningOutputDir `
    -StyleProfile auto `
    -StyleProfilePath $buildStyleProfilePath `
    -PreGeneratedReportPath $paginationRiskReportPath `
    -SkipSessionReset | Out-Null
  $urlWarningSummaryPath = Join-Path $urlWarningOutputDir 'url-build-summary.json'
  Assert-True -Condition (Test-Path -LiteralPath $urlWarningSummaryPath) -Message 'URL warning wrapper did not create the wrapper summary.'
  $urlWarningSummary = (Get-Content -LiteralPath $urlWarningSummaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-ValidationPaginationRiskSummary -Summary $urlWarningSummary -MessagePrefix 'URL warning wrapper summary' -ExpectedLongRemediationPattern $customLongPaginationRemediation -ExpectedDenseRemediationPattern $customDensePaginationRemediation -ExpectedFigureRemediationPattern $customFigurePaginationRemediation
  Assert-True -Condition ([string]$urlWarningSummary.generationMode -eq 'replay') -Message 'URL warning wrapper should use replay generation mode.'
  Assert-True -Condition ([string]$urlWarningSummary.buildRequirementsInputMode -eq 'inline') -Message 'URL warning wrapper should expose buildRequirementsInputMode=inline.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$urlWarningSummary.pipelineTracePath)) -Message 'URL warning wrapper should create a pipeline trace JSON.'
  $urlWarningTrace = (Get-Content -LiteralPath ([string]$urlWarningSummary.pipelineTracePath) -Raw -Encoding UTF8) | ConvertFrom-Json
  $urlWarningTraceCodes = @($urlWarningTrace.build.validationWarningCodes | ForEach-Object { [string]$_ })
  Assert-True -Condition ([bool]$urlWarningTrace.build.validationPassed) -Message 'URL warning pipeline trace should expose validationPassed=true.'
  Assert-True -Condition ([int]$urlWarningTrace.build.validationPaginationRiskCount -ge 3) -Message 'URL warning pipeline trace should expose pagination risks.'
  Assert-True -Condition ([string]$urlWarningTrace.build.validationPaginationRiskRemediations.'pagination-risk-long-section' -match $customLongPaginationRemediation) -Message 'URL warning pipeline trace should expose the custom long-section remediation.'
  Assert-True -Condition ($urlWarningTraceCodes -contains 'pagination-risk-long-section') -Message 'URL warning pipeline trace should expose pagination-risk-long-section.'
  Assert-True -Condition ($urlWarningTraceCodes -contains 'pagination-risk-dense-section-block') -Message 'URL warning pipeline trace should expose pagination-risk-dense-section-block.'
  $urlTraceLongWarningSummary = @($urlWarningTrace.build.validationWarningSummary | Where-Object { [string]$_.code -eq 'pagination-risk-long-section' })[0]
  Assert-True -Condition (-not [string]::IsNullOrWhiteSpace([string]$urlTraceLongWarningSummary.remediation)) -Message 'URL warning pipeline trace should expose remediation guidance.'
  Assert-True -Condition ($urlWarningTraceCodes -contains 'pagination-risk-figure-cluster') -Message 'URL warning pipeline trace should expose pagination-risk-figure-cluster.'
  $urlWarningTraceMarkdown = Get-Content -LiteralPath ([string]$urlWarningSummary.pipelineTraceMarkdownPath) -Raw -Encoding UTF8
  Assert-True -Condition ($urlWarningTraceMarkdown -match 'Validation passed: True') -Message 'URL warning pipeline markdown should include validation status.'
  Assert-True -Condition ($urlWarningTraceMarkdown -match ("Pagination risks: {0}" -f [int]$urlWarningSummary.validationPaginationRiskCount)) -Message 'URL warning pipeline markdown should include pagination risk count.'
  $results.Add('build-report-from-url validation warning propagation OK') | Out-Null

  $guidedReplayOutputDir = Join-Path $tempRoot 'guided-replay-e2e-output'
  & (Join-Path $repoRoot 'scripts\run-e2e-sample.ps1') `
    -OutputDir $guidedReplayOutputDir `
    -PreGeneratedReportPath $sampleReportPath `
    -Mode guided-chat `
    -SkipInstall `
    -SkipSessionReset | Out-Null
  $guidedReplaySummaryPath = Join-Path $guidedReplayOutputDir 'summary.json'
  Assert-True -Condition (Test-Path -LiteralPath $guidedReplaySummaryPath) -Message 'Guided replay E2E did not create its summary.'
  $guidedReplaySummary = (Get-Content -LiteralPath $guidedReplaySummaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([bool]$guidedReplaySummary.passed) -Message 'Guided replay E2E summary reported a failed validation result.'
  Assert-True -Condition ([string]$guidedReplaySummary.responseFormat -eq 'gateway-chat') -Message 'Guided replay E2E should still report the guided-chat response format.'
  Assert-True -Condition ([string]$guidedReplaySummary.generationMode -eq 'replay') -Message 'Guided replay E2E should mark the report generation mode as replay.'
  Assert-True -Condition ([string]$guidedReplaySummary.preGeneratedReportPath -eq $sampleReportPath) -Message 'Guided replay E2E summary is missing the explicit replay report path.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$guidedReplaySummary.reportPath)) -Message 'Guided replay E2E report path does not exist.'
  $results.Add('run-e2e-sample guided replay OK') | Out-Null

  $feishuBuildOutputDir = Join-Path $tempRoot 'feishu-build-output'
  & (Join-Path $repoRoot 'scripts\build-report-from-feishu.ps1') `
    -TemplatePath $sampleDocx `
    -ReportPath $sampleReportPath `
    -MetadataPath (Join-Path $repoRoot 'examples\docx-report-metadata.json') `
    -ImageSpecsPath $rowImageSpecsPath `
    -RequirementsPath (Join-Path $repoRoot 'examples\e2e-sample-requirements.json') `
    -OutputDir $feishuBuildOutputDir `
    -StyleProfile auto `
    -CreateTemplateFrameDocx `
    -StyleProfilePath $buildStyleProfilePath | Out-Null
  $feishuBuildSummaryPath = Join-Path $feishuBuildOutputDir 'feishu-build-summary.json'
  Assert-True -Condition (Test-Path -LiteralPath $feishuBuildSummaryPath) -Message 'Feishu wrapper did not create the wrapper summary.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $feishuBuildOutputDir 'report.txt')) -Message 'Feishu wrapper did not copy the report body to the output root.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $feishuBuildOutputDir 'artifacts\summary.json')) -Message 'Feishu wrapper did not keep the inner build summary under artifacts.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $feishuBuildOutputDir 'artifacts\generated-field-map.json')) -Message 'Feishu wrapper did not keep generated artifacts under artifacts.'
  $feishuBuildSummary = (Get-Content -LiteralPath $feishuBuildSummaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$feishuBuildSummary.mode -eq 'local-report') -Message 'Feishu wrapper summary did not record the local-report mode.'
  Assert-True -Condition ([string]$feishuBuildSummary.generationMode -eq 'none') -Message 'Feishu wrapper summary should mark local-report runs as generationMode=none.'
  Assert-True -Condition ([string]$feishuBuildSummary.detailLevel -eq 'full') -Message 'Feishu wrapper summary did not preserve the default full detail level.'
  Assert-True -Condition ((Split-Path -Parent $feishuBuildSummary.finalDocxPath) -eq $feishuBuildOutputDir) -Message 'Feishu wrapper should copy the final docx to the output root.'
  Assert-True -Condition (Test-Path -LiteralPath $feishuBuildSummary.finalDocxPath) -Message 'Feishu wrapper summary final docx path does not exist.'
  Assert-True -Condition ((Split-Path -Parent $feishuBuildSummary.templateFrameDocxPath) -eq $feishuBuildOutputDir) -Message 'Feishu wrapper should copy the template-frame docx to the output root.'
  Assert-True -Condition (Test-Path -LiteralPath $feishuBuildSummary.templateFrameDocxPath) -Message 'Feishu wrapper summary template-frame docx path does not exist.'
  Assert-True -Condition ([string]$feishuBuildSummary.artifactsDir -eq (Join-Path $feishuBuildOutputDir 'artifacts')) -Message 'Feishu wrapper summary is missing the expected artifacts directory.'
  Assert-True -Condition ([string]$feishuBuildSummary.reportProfileName -eq 'experiment-report') -Message 'Feishu wrapper summary is missing the expected report profile name.'
  Assert-True -Condition ([string]$feishuBuildSummary.reportProfileDisplayName -eq '实验报告') -Message 'Feishu wrapper summary is missing the expected report profile display name.'
  Assert-True -Condition ([string]$feishuBuildSummary.buildReportInputMode -eq 'path') -Message 'Feishu wrapper summary should expose buildReportInputMode=path from the downstream build summary.'
  Assert-True -Condition ([string]$feishuBuildSummary.buildMetadataInputMode -eq 'path') -Message 'Feishu wrapper summary should expose buildMetadataInputMode=path from the downstream build summary.'
  Assert-True -Condition ([string]$feishuBuildSummary.buildRequirementsInputMode -eq 'path') -Message 'Feishu wrapper summary should expose buildRequirementsInputMode=path from the downstream build summary.'
  Assert-True -Condition ([string]$feishuBuildSummary.buildImageInputMode -eq 'specs-path') -Message 'Feishu wrapper summary should expose buildImageInputMode=specs-path from the downstream build summary.'
  Assert-True -Condition ([bool]$feishuBuildSummary.validationPassed) -Message 'Feishu wrapper summary should expose validationPassed from the downstream build summary.'
  Assert-True -Condition ([int]$feishuBuildSummary.validationWarningCount -eq 0) -Message 'Feishu wrapper summary should expose zero validation warnings for the passing fixture.'
  Assert-True -Condition ([int]$feishuBuildSummary.validationPaginationRiskCount -eq 0) -Message 'Feishu wrapper summary should expose zero pagination risks for the passing fixture.'
  Assert-True -Condition ([int]$feishuBuildSummary.validationStructuralIssueCount -eq 0) -Message 'Feishu wrapper summary should expose zero structural issues for the passing fixture.'
  Assert-True -Condition (@($feishuBuildSummary.validationWarningCodes).Count -eq 0) -Message 'Feishu wrapper summary should expose an empty validationWarningCodes array for the passing fixture.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$feishuBuildSummary.pipelineTracePath)) -Message 'Feishu wrapper summary should point to a pipeline-trace JSON.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$feishuBuildSummary.pipelineTraceMarkdownPath)) -Message 'Feishu wrapper summary should point to a pipeline-trace markdown file.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$feishuBuildSummary.imagePlanPath)) -Message 'Feishu wrapper summary image-plan path does not exist.'
  Assert-True -Condition ([int]$feishuBuildSummary.imagePlanLowConfidenceCount -eq 0) -Message 'Feishu wrapper summary reported an unexpected low-confidence image-plan count.'
  Assert-True -Condition (-not [bool]$feishuBuildSummary.imagePlanNeedsReview) -Message 'Feishu wrapper summary should not require manual review for explicit image specs.'
  Assert-True -Condition ([bool]$feishuBuildSummary.layoutCheckPassed) -Message 'Feishu wrapper summary reported a failed layout check.'
  Assert-True -Condition ([string]$feishuBuildSummary.layoutCheckMessage -match 'Layout check passed') -Message 'Feishu wrapper summary is missing the readable layout-check message.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$feishuBuildSummary.layoutCheckPath)) -Message 'Feishu wrapper summary layout check path does not exist.'
  $feishuPipelineTrace = (Get-Content -LiteralPath ([string]$feishuBuildSummary.pipelineTracePath) -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-True -Condition ([string]$feishuPipelineTrace.wrapper.mode -eq 'local-report') -Message 'Feishu pipeline trace should keep wrapper.mode=local-report.'
  Assert-True -Condition ([string]$feishuPipelineTrace.wrapper.generationMode -eq 'none') -Message 'Feishu pipeline trace should keep wrapper.generationMode=none.'
  Assert-True -Condition ([string]$feishuPipelineTrace.build.requirementsInputMode -eq 'path') -Message 'Feishu pipeline trace should keep build.requirementsInputMode=path.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$feishuPipelineTrace.artifacts.templateFrameDocxPath)) -Message 'Feishu pipeline trace should expose the copied template-frame docx.'
  Assert-True -Condition ([bool]$feishuPipelineTrace.build.validationPassed) -Message 'Feishu pipeline trace should expose validationPassed.'
  Assert-True -Condition ([int]$feishuPipelineTrace.build.validationPaginationRiskCount -eq 0) -Message 'Feishu pipeline trace should expose zero pagination risks.'
  $feishuPipelineTraceMarkdown = Get-Content -LiteralPath ([string]$feishuBuildSummary.pipelineTraceMarkdownPath) -Raw -Encoding UTF8
  Assert-True -Condition ($feishuPipelineTraceMarkdown -match 'Mode: local-report') -Message 'Feishu pipeline markdown should include wrapper mode.'
  Assert-True -Condition ($feishuPipelineTraceMarkdown -match 'Requirements input mode: path') -Message 'Feishu pipeline markdown should include build requirements input mode.'
  Assert-True -Condition ($feishuPipelineTraceMarkdown -match 'Validation passed: True') -Message 'Feishu pipeline markdown should include validation status.'
  Assert-True -Condition ($feishuPipelineTraceMarkdown -match 'Pagination risks: 0') -Message 'Feishu pipeline markdown should include pagination risk count.'
  $results.Add('Feishu wrapper pipeline OK') | Out-Null

  $feishuWarningOutputDir = Join-Path $tempRoot 'feishu-warning-build-output'
  & (Join-Path $repoRoot 'scripts\build-report-from-feishu.ps1') `
    -TemplatePath $sampleDocx `
    -ReportPath $paginationRiskReportPath `
    -MetadataPath (Join-Path $repoRoot 'examples\docx-report-metadata.json') `
    -ImageSpecsPath $rowImageSpecsPath `
    -RequirementsJson $paginationRiskRequirementsJson `
    -OutputDir $feishuWarningOutputDir `
    -StyleProfile auto `
    -StyleProfilePath $buildStyleProfilePath | Out-Null
  $feishuWarningSummaryPath = Join-Path $feishuWarningOutputDir 'feishu-build-summary.json'
  Assert-True -Condition (Test-Path -LiteralPath $feishuWarningSummaryPath) -Message 'Feishu warning wrapper did not create the wrapper summary.'
  $feishuWarningSummary = (Get-Content -LiteralPath $feishuWarningSummaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
  Assert-ValidationPaginationRiskSummary -Summary $feishuWarningSummary -MessagePrefix 'Feishu warning wrapper summary' -ExpectedLongRemediationPattern $customLongPaginationRemediation -ExpectedDenseRemediationPattern $customDensePaginationRemediation -ExpectedFigureRemediationPattern $customFigurePaginationRemediation
  Assert-True -Condition ([string]$feishuWarningSummary.mode -eq 'local-report') -Message 'Feishu warning wrapper should use local-report mode.'
  Assert-True -Condition ([string]$feishuWarningSummary.buildRequirementsInputMode -eq 'inline') -Message 'Feishu warning wrapper should expose buildRequirementsInputMode=inline.'
  Assert-True -Condition (Test-Path -LiteralPath ([string]$feishuWarningSummary.pipelineTracePath)) -Message 'Feishu warning wrapper should create a pipeline trace JSON.'
  $feishuWarningTrace = (Get-Content -LiteralPath ([string]$feishuWarningSummary.pipelineTracePath) -Raw -Encoding UTF8) | ConvertFrom-Json
  $feishuWarningTraceCodes = @($feishuWarningTrace.build.validationWarningCodes | ForEach-Object { [string]$_ })
  Assert-True -Condition ([bool]$feishuWarningTrace.build.validationPassed) -Message 'Feishu warning pipeline trace should expose validationPassed=true.'
  Assert-True -Condition ([int]$feishuWarningTrace.build.validationPaginationRiskCount -ge 3) -Message 'Feishu warning pipeline trace should expose pagination risks.'
  Assert-True -Condition ([string]$feishuWarningTrace.build.validationPaginationRiskRemediations.'pagination-risk-long-section' -match $customLongPaginationRemediation) -Message 'Feishu warning pipeline trace should expose the custom long-section remediation.'
  Assert-True -Condition ($feishuWarningTraceCodes -contains 'pagination-risk-long-section') -Message 'Feishu warning pipeline trace should expose pagination-risk-long-section.'
  Assert-True -Condition ($feishuWarningTraceCodes -contains 'pagination-risk-dense-section-block') -Message 'Feishu warning pipeline trace should expose pagination-risk-dense-section-block.'
  Assert-True -Condition ($feishuWarningTraceCodes -contains 'pagination-risk-figure-cluster') -Message 'Feishu warning pipeline trace should expose pagination-risk-figure-cluster.'
  $feishuTraceLongWarningSummary = @($feishuWarningTrace.build.validationWarningSummary | Where-Object { [string]$_.code -eq 'pagination-risk-long-section' })[0]
  Assert-True -Condition (-not [string]::IsNullOrWhiteSpace([string]$feishuTraceLongWarningSummary.remediation)) -Message 'Feishu warning pipeline trace should expose remediation guidance.'
  $feishuWarningTraceMarkdown = Get-Content -LiteralPath ([string]$feishuWarningSummary.pipelineTraceMarkdownPath) -Raw -Encoding UTF8
  Assert-True -Condition ($feishuWarningTraceMarkdown -match 'Validation passed: True') -Message 'Feishu warning pipeline markdown should include validation status.'
  Assert-True -Condition ($feishuWarningTraceMarkdown -match ("Pagination risks: {0}" -f [int]$feishuWarningSummary.validationPaginationRiskCount)) -Message 'Feishu warning pipeline markdown should include pagination risk count.'
  $results.Add('Feishu wrapper validation warning propagation OK') | Out-Null

  $originalWrapperAgentsHome = $env:AGENTS_HOME
  try {
    $env:AGENTS_HOME = (Join-Path $tempRoot 'wrapper-agents-home')
    $courseDesignWrapperOutputDir = Join-Path $tempRoot 'course-design-feishu-output'
    & (Join-Path $repoRoot 'scripts\build-report-from-feishu.ps1') `
      -TemplatePath $courseDesignTemplateDocx `
      -ReportPath $courseDesignReportPath `
      -MetadataPath $courseDesignMetadataPath `
      -ImageSpecsPath $courseDesignImageSpecsPath `
      -OutputDir $courseDesignWrapperOutputDir `
      -StyleProfile auto `
      -ReportProfileName 'course-design-report' | Out-Null
    $courseDesignWrapperSummaryPath = Join-Path $courseDesignWrapperOutputDir 'feishu-build-summary.json'
    Assert-True -Condition (Test-Path -LiteralPath $courseDesignWrapperSummaryPath) -Message 'Course-design Feishu wrapper did not create the wrapper summary.'
    $courseDesignWrapperSummary = (Get-Content -LiteralPath $courseDesignWrapperSummaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
    Assert-True -Condition ([string]$courseDesignWrapperSummary.reportProfileName -eq 'course-design-report') -Message 'Course-design Feishu wrapper summary is missing the expected report profile name.'
    Assert-True -Condition ([string]$courseDesignWrapperSummary.reportProfileDisplayName -eq '课程设计报告') -Message 'Course-design Feishu wrapper summary is missing the expected display name.'
    Assert-True -Condition ([string]$courseDesignWrapperSummary.generationMode -eq 'none') -Message 'Course-design Feishu local-report wrapper should mark generationMode=none.'
    Assert-True -Condition ([string]$courseDesignWrapperSummary.buildReportInputMode -eq 'path') -Message 'Course-design Feishu local-report wrapper should expose buildReportInputMode=path from the downstream build summary.'
    Assert-True -Condition ([string]$courseDesignWrapperSummary.buildMetadataInputMode -eq 'path') -Message 'Course-design Feishu local-report wrapper should expose buildMetadataInputMode=path from the downstream build summary.'
    Assert-True -Condition ([string]$courseDesignWrapperSummary.buildRequirementsInputMode -eq 'none') -Message 'Course-design Feishu local-report wrapper should expose buildRequirementsInputMode=none when no requirements input is provided.'
    Assert-True -Condition ([string]$courseDesignWrapperSummary.buildImageInputMode -eq 'specs-path') -Message 'Course-design Feishu local-report wrapper should expose buildImageInputMode=specs-path from the downstream build summary.'
    Assert-True -Condition ((Split-Path -Leaf ([string]$courseDesignWrapperSummary.defaultsPath)) -eq 'course-design-report.defaults.json') -Message 'Course-design Feishu wrapper should persist defaults under the profile-specific defaults file.'
    Assert-True -Condition (Test-Path -LiteralPath ([string]$courseDesignWrapperSummary.finalDocxPath)) -Message 'Course-design Feishu wrapper final docx path does not exist.'

    $courseDesignReplayWrapperOutputDir = Join-Path $tempRoot 'course-design-feishu-replay-output'
    & (Join-Path $repoRoot 'scripts\build-report-from-feishu.ps1') `
      -TemplatePath $courseDesignTemplateDocx `
      -CourseName '软件工程综合实践' `
      -ExperimentName '校园导览小程序设计' `
      -ImageSpecsPath $courseDesignImageSpecsPath `
      -OutputDir $courseDesignReplayWrapperOutputDir `
      -StyleProfile auto `
      -ReportProfileName 'course-design-report' `
      -PreGeneratedReportPath $preparedSummaryMockReportPath `
      -SkipSessionReset | Out-Null
    $courseDesignReplayWrapperSummaryPath = Join-Path $courseDesignReplayWrapperOutputDir 'feishu-build-summary.json'
    Assert-True -Condition (Test-Path -LiteralPath $courseDesignReplayWrapperSummaryPath) -Message 'Course-design Feishu replay wrapper did not create the wrapper summary.'
    $courseDesignReplayWrapperSummary = (Get-Content -LiteralPath $courseDesignReplayWrapperSummaryPath -Raw -Encoding UTF8) | ConvertFrom-Json
    Assert-True -Condition ([string]$courseDesignReplayWrapperSummary.mode -eq 'generated-report') -Message 'Course-design Feishu replay wrapper should record generated-report mode.'
    Assert-True -Condition ([string]$courseDesignReplayWrapperSummary.generationMode -eq 'replay') -Message 'Course-design Feishu replay wrapper should mark generationMode=replay.'
    Assert-True -Condition ([string]$courseDesignReplayWrapperSummary.buildReportInputMode -eq 'path') -Message 'Course-design Feishu replay wrapper should expose buildReportInputMode=path from the downstream build summary.'
    Assert-True -Condition ([string]$courseDesignReplayWrapperSummary.buildMetadataInputMode -eq 'path') -Message 'Course-design Feishu replay wrapper should expose buildMetadataInputMode=path from the downstream build summary.'
    Assert-True -Condition ([string]$courseDesignReplayWrapperSummary.buildRequirementsInputMode -eq 'path') -Message 'Course-design Feishu replay wrapper should expose buildRequirementsInputMode=path from the downstream build summary.'
    Assert-True -Condition ([string]$courseDesignReplayWrapperSummary.buildImageInputMode -eq 'specs-path') -Message 'Course-design Feishu replay wrapper should expose buildImageInputMode=specs-path from the downstream build summary.'
    Assert-True -Condition (Test-Path -LiteralPath ([string]$courseDesignReplayWrapperSummary.pipelineTracePath)) -Message 'Course-design Feishu replay wrapper should create a pipeline-trace JSON.'
    Assert-True -Condition (Test-Path -LiteralPath ([string]$courseDesignReplayWrapperSummary.pipelineTraceMarkdownPath)) -Message 'Course-design Feishu replay wrapper should create a pipeline-trace markdown file.'
    Assert-True -Condition ([string]$courseDesignReplayWrapperSummary.reportProfileName -eq 'course-design-report') -Message 'Course-design Feishu replay wrapper summary is missing the expected profile name.'
    Assert-True -Condition ([string]$courseDesignReplayWrapperSummary.preGeneratedReportPath -eq $preparedSummaryMockReportPath) -Message 'Course-design Feishu replay wrapper summary is missing the explicit replay report path.'
    Assert-True -Condition (Test-Path -LiteralPath ([string]$courseDesignReplayWrapperSummary.reportPath)) -Message 'Course-design Feishu replay wrapper did not copy the replayed report body.'
    Assert-True -Condition (Test-Path -LiteralPath ([string]$courseDesignReplayWrapperSummary.finalDocxPath)) -Message 'Course-design Feishu replay wrapper final docx path does not exist.'
    $courseDesignReplayTrace = (Get-Content -LiteralPath ([string]$courseDesignReplayWrapperSummary.pipelineTracePath) -Raw -Encoding UTF8) | ConvertFrom-Json
    Assert-True -Condition ([string]$courseDesignReplayTrace.wrapper.mode -eq 'generated-report') -Message 'Course-design Feishu replay pipeline trace should keep wrapper.mode=generated-report.'
    Assert-True -Condition ([string]$courseDesignReplayTrace.wrapper.generationMode -eq 'replay') -Message 'Course-design Feishu replay pipeline trace should keep wrapper.generationMode=replay.'
    Assert-True -Condition ([string]$courseDesignReplayTrace.build.imageInputMode -eq 'specs-path') -Message 'Course-design Feishu replay pipeline trace should keep build.imageInputMode=specs-path.'
    $courseDesignReplayTraceMarkdown = Get-Content -LiteralPath ([string]$courseDesignReplayWrapperSummary.pipelineTraceMarkdownPath) -Raw -Encoding UTF8
    Assert-True -Condition ($courseDesignReplayTraceMarkdown -match 'Mode: generated-report') -Message 'Course-design Feishu replay pipeline markdown should include wrapper mode.'
    Assert-True -Condition ($courseDesignReplayTraceMarkdown -match 'Generation mode: replay') -Message 'Course-design Feishu replay pipeline markdown should include replay generation mode.'
  } finally {
    $env:AGENTS_HOME = $originalWrapperAgentsHome
  }
  $results.Add('course-design Feishu wrapper OK') | Out-Null

  $templateExamplesDir = Join-Path $repoRoot 'examples\report-templates'
  $templateExamplesReadmePath = Join-Path $templateExamplesDir 'README.md'
  $templateExamplesSummaryPath = Join-Path $templateExamplesDir 'report-template-examples-summary.json'
  $templateExampleFiles = @(
    'course-design-report-template.docx',
    'deployment-report-template.docx',
    'experiment-report-template.docx',
    'internship-report-template.docx',
    'meeting-minutes-template.docx',
    'monthly-report-template.docx',
    'software-test-report-template.docx',
    'weekly-report-template.docx'
  )
  foreach ($fixturePath in @($templateExamplesReadmePath, $templateExamplesSummaryPath) + @($templateExampleFiles | ForEach-Object { Join-Path $templateExamplesDir $_ })) {
    Assert-True -Condition (Test-Path -LiteralPath $fixturePath -PathType Leaf) -Message ("Report-template example fixture is missing: {0}" -f $fixturePath)
  }
  $templateExamplesOutputDir = Join-Path $tempRoot 'report-template-examples-regenerated'
  $templateExamplesGeneration = & (Join-Path $repoRoot 'scripts\export-report-template-examples.ps1') -OutputDir $templateExamplesOutputDir -Overwrite
  Assert-True -Condition ([int]$templateExamplesGeneration.generatedCount -eq 8) -Message 'Template-example exporter should generate one template for each built-in report profile.'
  foreach ($templateFile in $templateExampleFiles) {
    $checkedInTemplatePath = Join-Path $templateExamplesDir $templateFile
    $generatedTemplatePath = Join-Path $templateExamplesOutputDir $templateFile
    Assert-True -Condition (Test-Path -LiteralPath $generatedTemplatePath -PathType Leaf) -Message ("Template-example exporter did not generate: {0}" -f $templateFile)
    $checkedInOutline = & (Join-Path $repoRoot 'scripts\extract-docx-template.ps1') -Path $checkedInTemplatePath -Format markdown | Out-String
    $generatedOutline = & (Join-Path $repoRoot 'scripts\extract-docx-template.ps1') -Path $generatedTemplatePath -Format markdown | Out-String
    Assert-True -Condition ((Normalize-OutlineForComparison -Text $checkedInOutline) -eq (Normalize-OutlineForComparison -Text $generatedOutline)) -Message ("Checked-in template example drifted from regenerated output: {0}" -f $templateFile)
  }
  $experimentTemplateOutline = & (Join-Path $repoRoot 'scripts\extract-docx-template.ps1') -Path (Join-Path $templateExamplesDir 'experiment-report-template.docx') -Format markdown | Out-String
  Assert-True -Condition ($experimentTemplateOutline -match 'Table count: 1') -Message 'Checked-in experiment template example should contain a metadata table.'
  Assert-True -Condition ($experimentTemplateOutline -match '一、实验目的') -Message 'Checked-in experiment template example should contain numbered report headings.'
  $courseDesignTemplateOutline = & (Join-Path $repoRoot 'scripts\extract-docx-template.ps1') -Path (Join-Path $templateExamplesDir 'course-design-report-template.docx') -Format markdown | Out-String
  Assert-True -Condition ($courseDesignTemplateOutline -match 'Table count: 1') -Message 'Checked-in course-design template example should contain a metadata table.'
  Assert-True -Condition ($courseDesignTemplateOutline -match '四、方案设计与实现') -Message 'Checked-in course-design template example should contain the implementation heading.'
  $results.Add('report template examples OK') | Out-Null

  $installRoot = Join-Path $tempRoot 'install-root'
  $installTarget = Join-Path $installRoot 'skill-install'
  & (Join-Path $repoRoot 'scripts\install-skill.ps1') -TargetDir $installTarget | Out-Null
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'SKILL.md')) -Message 'Install script did not copy SKILL.md.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'CODE_OF_CONDUCT.md')) -Message 'Install script did not copy CODE_OF_CONDUCT.md.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'CONTRIBUTING.md')) -Message 'Install script did not copy CONTRIBUTING.md.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'CHANGELOG.md')) -Message 'Install script did not copy CHANGELOG.md.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'SECURITY.md')) -Message 'Install script did not copy SECURITY.md.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'SUPPORT.md')) -Message 'Install script did not copy SUPPORT.md.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'ROADMAP.md')) -Message 'Install script did not copy ROADMAP.md.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\extract-docx-template.ps1')) -Message 'Install script did not copy extractor script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\build-report.ps1')) -Message 'Install script did not copy build-report script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\build-report-from-feishu.ps1')) -Message 'Install script did not copy Feishu wrapper script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\build-report-from-url.ps1')) -Message 'Install script did not copy build-report-from-url script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\check-report-profile-template-fit.ps1')) -Message 'Install script did not copy the template-fit checker script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\check-docx-layout.ps1')) -Message 'Install script did not copy layout check script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\convert-docx-template-frame.ps1')) -Message 'Install script did not copy template-frame converter script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\fetch-web-article.ps1')) -Message 'Install script did not copy web fetch script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\format-docx-report-style.ps1')) -Message 'Install script did not copy style formatter script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\generate-docx-field-map.ps1')) -Message 'Install script did not copy field-map generator script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\generate-docx-image-map.ps1')) -Message 'Install script did not copy image-map generator script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\generate-report-inputs.ps1')) -Message 'Install script did not copy report-input generation script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\export-report-template-examples.ps1')) -Message 'Install script did not copy the template-example exporter script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\insert-docx-images.ps1')) -Message 'Install script did not copy image insertion script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\new-report-profile.ps1')) -Message 'Install script did not copy profile scaffold generator script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\prepare-report-prompt.ps1')) -Message 'Install script did not copy prompt preparation script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\report-defaults.ps1')) -Message 'Install script did not copy the report-defaults helper script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\report-profiles.ps1')) -Message 'Install script did not copy the report-profiles helper script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\run-profile-preset-samples.ps1')) -Message 'Install script did not copy the profile preset sample runner script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\apply-docx-field-map.ps1')) -Message 'Install script did not copy fill script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\validate-report-draft.ps1')) -Message 'Install script did not copy validation script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\validate-report-profiles.ps1')) -Message 'Install script did not copy profile validation script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'scripts\run-e2e-sample.ps1')) -Message 'Install script did not copy e2e script.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'profiles\experiment-report.json')) -Message 'Install script did not copy the experiment-report profile.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'profiles\course-design-report.json')) -Message 'Install script did not copy the course-design-report profile.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'profiles\internship-report.json')) -Message 'Install script did not copy the internship-report profile.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'profiles\software-test-report.json')) -Message 'Install script did not copy the software-test-report profile.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'profiles\deployment-report.json')) -Message 'Install script did not copy the deployment-report profile.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'profiles\weekly-report.json')) -Message 'Install script did not copy the weekly-report profile.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'profiles\meeting-minutes.json')) -Message 'Install script did not copy the meeting-minutes profile.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'profiles\monthly-report.json')) -Message 'Install script did not copy the monthly-report profile.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'profiles\report-profile.schema.json')) -Message 'Install script did not copy the report profile schema.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\feishu-uploaded-images-docx-prompt.md')) -Message 'Install script did not copy the Feishu uploaded-images prompt example.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\local-uploaded-images-docx-prompt.md')) -Message 'Install script did not copy the local uploaded-images prompt example.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\profile-presets\README.md')) -Message 'Install script did not copy the profile preset README.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\profile-presets\weekly-report.json')) -Message 'Install script did not copy the weekly profile preset example.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\profile-presets\meeting-minutes.json')) -Message 'Install script did not copy the meeting-minutes profile preset example.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\profile-presets\monthly-report.json')) -Message 'Install script did not copy the monthly-report profile preset example.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\report-templates\README.md')) -Message 'Install script did not copy the report-template example README.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\report-templates\experiment-report-template.docx')) -Message 'Install script did not copy the experiment template example.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\report-templates\course-design-report-template.docx')) -Message 'Install script did not copy the course-design template example.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\report-templates\weekly-report-template.docx')) -Message 'Install script did not copy the weekly template example.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\deployment-report-prepared\report-inputs-summary.json')) -Message 'Install script did not copy the deployment prepared-summary fixture.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\deployment-report-prepared\report.replay.txt')) -Message 'Install script did not copy the deployment replay report fixture.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\deployment-report-prepared\image-specs.json')) -Message 'Install script did not copy the deployment replay image specs fixture.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\deployment-report-prepared\images\deploy-env.png')) -Message 'Install script did not copy the deployment environment screenshot fixture.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\deployment-report-prepared\images\deploy-verify.png')) -Message 'Install script did not copy the deployment verification screenshot fixture.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\deployment-report-prepared\images\deploy-rollback.png')) -Message 'Install script did not copy the deployment rollback screenshot fixture.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\software-test-report-prepared\report-inputs-summary.json')) -Message 'Install script did not copy the software-test prepared-summary fixture.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\software-test-report-prepared\report.replay.txt')) -Message 'Install script did not copy the software-test replay report fixture.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\weekly-report-prepared\report-inputs-summary.json')) -Message 'Install script did not copy the weekly prepared-summary fixture.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\weekly-report-prepared\report.replay.txt')) -Message 'Install script did not copy the weekly replay report fixture.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\monthly-report-prepared\report-inputs-summary.json')) -Message 'Install script did not copy the monthly prepared-summary fixture.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\monthly-report-prepared\report.replay.txt')) -Message 'Install script did not copy the monthly replay report fixture.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\meeting-minutes-prepared\report-inputs-summary.json')) -Message 'Install script did not copy the meeting-minutes prepared-summary fixture.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'examples\meeting-minutes-prepared\report.replay.txt')) -Message 'Install script did not copy the meeting-minutes replay report fixture.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget '.github\pull_request_template.md')) -Message 'Install script did not copy the PR template.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget '.github\ISSUE_TEMPLATE\bug_report.md')) -Message 'Install script did not copy the bug-report template.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget '.github\workflows\quality.yml')) -Message 'Install script did not copy the quality workflow.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget '.github\workflows\roadmap-triage.yml')) -Message 'Install script did not copy the roadmap triage workflow.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget '.github\workflows\smoke-tests.yml')) -Message 'Install script did not copy the smoke-tests workflow.'
  $results.Add('install script first install OK') | Out-Null

  & (Join-Path $repoRoot 'scripts\install-skill.ps1') -TargetDir $installTarget -Force | Out-Null
  $backupCount = @(Get-ChildItem -LiteralPath $installRoot -Filter 'skill-install.bak-*' -Force).Count
  Assert-True -Condition ($backupCount -ge 1) -Message 'Install script -Force did not create a backup directory.'
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $installTarget 'README.md')) -Message 'Install script -Force did not reinstall README.md.'
  $results.Add('install script force reinstall OK') | Out-Null

  $agentsHome = Join-Path $tempRoot 'agents-home'
  $defaultInstallTarget = Join-Path $agentsHome 'skills\experiment-report'
  $defaultBackupRoot = Join-Path $agentsHome 'skill-backups'
  & (Join-Path $repoRoot 'scripts\install-skill.ps1') -AgentsHome $agentsHome | Out-Null
  Assert-True -Condition (Test-Path -LiteralPath (Join-Path $defaultInstallTarget 'SKILL.md')) -Message 'Default install layout did not create the personal skill directory.'
  & (Join-Path $repoRoot 'scripts\install-skill.ps1') -AgentsHome $agentsHome -Force | Out-Null
  Assert-True -Condition (@(Get-ChildItem -LiteralPath $defaultBackupRoot -Filter 'experiment-report.bak-*' -Force).Count -ge 1) -Message 'Default install layout did not move backups into skill-backups.'
  Assert-True -Condition (@(Get-ChildItem -LiteralPath (Join-Path $agentsHome 'skills') -Filter 'experiment-report.bak-*' -Force).Count -eq 0) -Message 'Default install layout left backup skill directories inside the scanned skills root.'
  $results.Add('install script backup isolation OK') | Out-Null

  $resolvedOpenClaw = $null
  if (-not [string]::IsNullOrWhiteSpace($OpenClawCmd)) {
    $resolvedOpenClaw = (Resolve-Path -LiteralPath $OpenClawCmd).Path
  } else {
    foreach ($name in @('openclaw.cmd', 'openclaw')) {
      $cmd = Get-Command $name -ErrorAction SilentlyContinue
      if ($null -ne $cmd -and $cmd.Source) {
        $resolvedOpenClaw = $cmd.Source
        break
      }
    }
  }

  if ($null -ne $resolvedOpenClaw) {
    $originalOpenClawCmd = $env:OPENCLAW_CMD
    try {
      $env:OPENCLAW_CMD = $resolvedOpenClaw
      $selfCheckOutput = & (Join-Path $repoRoot 'scripts\self-check.ps1') | Out-String
      Assert-True -Condition ($selfCheckOutput -match 'OpenClaw CLI:') -Message 'self-check output missing CLI line.'
      Assert-True -Condition ($selfCheckOutput -match 'browser status:') -Message 'self-check output missing browser status section.'
      $results.Add('self-check with local OpenClaw OK') | Out-Null
    } finally {
      $env:OPENCLAW_CMD = $originalOpenClawCmd
    }
  } else {
    $results.Add('self-check skipped: OpenClaw CLI not found') | Out-Null
  }

  Write-Output 'Smoke tests passed:'
  foreach ($result in $results) {
    Write-Output ('- ' + $result)
  }
} finally {
  if (Test-Path -LiteralPath $tempRoot) {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force
  }
}

