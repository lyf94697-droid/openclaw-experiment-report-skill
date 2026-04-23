[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$DocxPath,

  [string]$OutPath,

  [switch]$Overwrite
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

function Escape-XmlText {
  param([AllowNull()][string]$Text)
  if ($null -eq $Text) {
    return ""
  }
  return [System.Security.SecurityElement]::Escape($Text)
}

function Get-NodeText {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode]$Node,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $parts = New-Object System.Collections.Generic.List[string]
  foreach ($textNode in @($Node.SelectNodes(".//w:t", $NamespaceManager))) {
    [void]$parts.Add([string]$textNode.InnerText)
  }
  return (($parts -join "") -replace "\s+", " ").Trim()
}

function Convert-ChineseOrdinalToInt {
  param([AllowNull()][string]$Text)

  $normalized = if ($null -eq $Text) { "" } else { [string]$Text }
  $normalized = $normalized.Trim()
  if ([string]::IsNullOrWhiteSpace($normalized)) {
    return $null
  }

  $lookup = @{
    "一" = 1
    "二" = 2
    "三" = 3
    "四" = 4
    "五" = 5
    "六" = 6
    "七" = 7
    "八" = 8
    "九" = 9
    "十" = 10
  }
  if ($lookup.ContainsKey($normalized)) {
    return [int]$lookup[$normalized]
  }

  return $null
}

function Get-HeadingChapterNumber {
  param([AllowNull()][string]$Text)

  $normalized = if ($null -eq $Text) { "" } else { [string]$Text }
  $normalized = $normalized.Trim()
  if ([string]::IsNullOrWhiteSpace($normalized)) {
    return $null
  }

  if ($normalized -match '^(?<num>\d+)\s*[\.．、]') {
    return [int]$matches["num"]
  }

  if ($normalized -match '^第\s*(?<num>\d+)\s*[章节]') {
    return [int]$matches["num"]
  }

  if ($normalized -match '^(?<cn>[一二三四五六七八九十]+)\s*[\.．、]') {
    return Convert-ChineseOrdinalToInt -Text $matches["cn"]
  }

  if ($normalized -match '^第\s*(?<cn>[一二三四五六七八九十]+)\s*[章节]') {
    return Convert-ChineseOrdinalToInt -Text $matches["cn"]
  }

  return $null
}

function Get-NextSubsectionNumber {
  param(
    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNode[]]$Children,

    [int]$StartIndex = -1,

    [AllowNull()]
    [System.Xml.XmlNode]$StopNode,

    [Parameter(Mandatory = $true)]
    [System.Xml.XmlNamespaceManager]$NamespaceManager,

    [int]$ChapterNumber = 4
  )

  if ($StartIndex -lt 0) {
    return 1
  }

  $maxSubsectionNumber = 0
  $pattern = '^{0}\.(?<sub>\d+)\s+' -f [regex]::Escape([string]$ChapterNumber)
  for ($i = $StartIndex + 1; $i -lt $Children.Count; $i++) {
    $child = $Children[$i]
    if ($null -ne $StopNode -and [object]::ReferenceEquals($child, $StopNode)) {
      break
    }

    if ($child.LocalName -ne "p") {
      continue
    }

    $text = Get-NodeText -Node $child -NamespaceManager $NamespaceManager
    if ([string]::IsNullOrWhiteSpace($text)) {
      continue
    }

    if ($text -match $pattern) {
      $subsectionNumber = [int]$matches["sub"]
      if ($subsectionNumber -gt $maxSubsectionNumber) {
        $maxSubsectionNumber = $subsectionNumber
      }
    }
  }

  return ($maxSubsectionNumber + 1)
}

function New-RunPropertiesXml {
  param(
    [string]$FontName = "宋体",
    [int]$SizeHalfPoints = 21,
    [switch]$Bold
  )

  $boldXml = if ($Bold) { "<w:b/>" } else { "" }
  $font = Escape-XmlText -Text $FontName
  return "<w:rPr><w:rFonts w:ascii=`"$font`" w:hAnsi=`"$font`" w:eastAsia=`"$font`"/>$boldXml<w:sz w:val=`"$SizeHalfPoints`"/><w:szCs w:val=`"$SizeHalfPoints`"/></w:rPr>"
}

function New-ParagraphXml {
  param(
    [AllowNull()][string]$Text,
    [ValidateSet("left", "center", "both")]
    [string]$Justification = "left",
    [int]$FirstLineTwips = 0,
    [int]$BeforeTwips = 0,
    [int]$AfterTwips = 0,
    [int]$LineTwips = 320,
    [string]$FontName = "宋体",
    [int]$SizeHalfPoints = 21,
    [switch]$Bold
  )

  $indentXml = if ($FirstLineTwips -gt 0) { "<w:ind w:firstLine=`"$FirstLineTwips`"/>" } else { "" }
  $rPr = New-RunPropertiesXml -FontName $FontName -SizeHalfPoints $SizeHalfPoints -Bold:$Bold
  $safeText = Escape-XmlText -Text $Text
  return @"
<w:p xmlns:w="$wordNamespace">
  <w:pPr><w:spacing w:before="$BeforeTwips" w:after="$AfterTwips" w:line="$LineTwips" w:lineRule="auto"/>$indentXml<w:jc w:val="$Justification"/></w:pPr>
  <w:r>$rPr<w:t xml:space="preserve">$safeText</w:t></w:r>
</w:p>
"@
}

function New-CellXml {
  param(
    [AllowNull()][string]$Text,
    [int]$WidthTwips = 1800,
    [switch]$Header,
    [string]$VMerge = "",
    [string]$Justification = "center"
  )

  $mergeXml = if ([string]::IsNullOrWhiteSpace($VMerge)) { "" } else { "<w:vMerge w:val=`"$VMerge`"/>" }
  $fontSize = if ($Header) { 21 } else { 20 }
  $paragraph = New-ParagraphXml -Text $Text -Justification $Justification -LineTwips 280 -FontName "宋体" -SizeHalfPoints $fontSize -Bold:$Header
  $innerParagraph = $paragraph -replace '^<w:p xmlns:w="[^"]+">', '<w:p>' -replace '</w:p>\s*$', '</w:p>'
  return @"
<w:tc>
  <w:tcPr><w:tcW w:w="$WidthTwips" w:type="dxa"/>$mergeXml<w:vAlign w:val="center"/><w:tcMar><w:top w:w="60" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tcMar></w:tcPr>
  $innerParagraph
</w:tc>
"@
}

function New-TableXml {
  param(
    [Parameter(Mandatory = $true)]
    [string[]]$Headers,

    [Parameter(Mandatory = $true)]
    [object[]]$Rows,

    [Parameter(Mandatory = $true)]
    [int[]]$Widths,

    [switch]$MergeFirstColumn
  )

  $gridXml = ($Widths | ForEach-Object { "<w:gridCol w:w=`"$_`"/>" }) -join ""
  $rowXml = New-Object System.Collections.Generic.List[string]

  $headerCells = New-Object System.Collections.Generic.List[string]
  for ($i = 0; $i -lt $Headers.Count; $i++) {
    [void]$headerCells.Add((New-CellXml -Text $Headers[$i] -WidthTwips $Widths[$i] -Header))
  }
  [void]$rowXml.Add("<w:tr><w:trPr><w:cantSplit/></w:trPr>$($headerCells -join '')</w:tr>")

  $previousFirstColumn = $null
  foreach ($row in $Rows) {
    $values = @($row)
    $cells = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $Headers.Count; $i++) {
      $value = if ($i -lt $values.Count) { [string]$values[$i] } else { "" }
      $vMerge = ""
      if ($MergeFirstColumn -and $i -eq 0) {
        if (-not [string]::IsNullOrWhiteSpace($previousFirstColumn) -and $value -eq $previousFirstColumn) {
          $vMerge = "continue"
          $value = ""
        } else {
          $vMerge = "restart"
          $previousFirstColumn = $value
        }
      }

      $justification = if ($i -eq ($Headers.Count - 1) -and $Headers.Count -gt 3) { "left" } elseif ($i -eq 2 -and $Headers.Count -eq 3) { "left" } else { "center" }
      [void]$cells.Add((New-CellXml -Text $value -WidthTwips $Widths[$i] -VMerge $vMerge -Justification $justification))
    }
    [void]$rowXml.Add("<w:tr><w:trPr><w:cantSplit/></w:trPr>$($cells -join '')</w:tr>")
  }

  return @"
<w:tbl xmlns:w="$wordNamespace">
  <w:tblPr>
    <w:tblW w:w="0" w:type="auto"/>
    <w:jc w:val="center"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
    </w:tblBorders>
    <w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar>
  </w:tblPr>
  <w:tblGrid>$gridXml</w:tblGrid>
  $($rowXml -join "`n")
</w:tbl>
"@
}

function New-TableBlockXml {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Caption,

    [Parameter(Mandatory = $true)]
    [string[]]$Headers,

    [Parameter(Mandatory = $true)]
    [object[]]$Rows,

    [Parameter(Mandatory = $true)]
    [int[]]$Widths,

    [switch]$MergeFirstColumn
  )

  return (New-ParagraphXml -Text $Caption -Justification center -BeforeTwips 100 -AfterTwips 40 -LineTwips 280 -FontName "宋体" -SizeHalfPoints 21) +
    (New-TableXml -Headers $Headers -Rows $Rows -Widths $Widths -MergeFirstColumn:$MergeFirstColumn)
}

function Get-CourseDesignTableProfile {
  param([Parameter(Mandatory = $true)][string]$DocumentText)

  if ($DocumentText -match "选课|退课|成绩提交|课程管理|授课课程") {
    return [pscustomobject]@{
      modules = @(
        @("管理员模块", "学生管理子模块", "对学生信息进行添加、修改、删除和查询"),
        @("管理员模块", "教师管理子模块", "维护教师信息并关联授课课程"),
        @("管理员模块", "课程管理子模块", "维护课程、任课教师、上课时间和地点"),
        @("教师模块", "成绩提交子模块", "查看选课学生并录入、修改课程成绩"),
        @("学生模块", "选课退课子模块", "完成课程查询、选课、退课和课表查看"),
        @("公有模块", "身份验证子模块", "完成登录、密码修改和退出系统")
      )
      database = @(
        @("1", "Student", "存储学生基本信息"),
        @("2", "Teacher", "存储教师基本信息"),
        @("3", "Users", "存储管理员账号信息"),
        @("4", "Elect", "存储学生选课信息"),
        @("5", "Course", "存储课程信息"),
        @("6", "Depart", "存储院系信息")
      )
      fieldTables = @(
        [pscustomobject]@{ name = "Student学生用户表"; rows = @(@("1","stuID","nvarchar(20) not null","学生学号","关键字"),@("2","stuPwd","nvarchar(20) not null","学生密码",""),@("3","stuName","nvarchar(20) not null","学生姓名",""),@("4","stuDepart","int","学生系院号",""),@("5","stuGrade","int","学生年级",""),@("6","stuClass","int","学生班级","")) },
        [pscustomobject]@{ name = "Teacher教师用户表"; rows = @(@("1","teaID","nvarchar(20) not null","教师编号","关键字"),@("2","teaPwd","nvarchar(20) not null","教师密码",""),@("3","teaName","nvarchar(20) not null","教师姓名",""),@("4","teaDepart","int","所属系院","")) },
        [pscustomobject]@{ name = "Course课程信息表"; rows = @(@("1","courseID","nvarchar(20) not null","课程编号","关键字"),@("2","courseName","nvarchar(50)","课程名称",""),@("3","teacherID","nvarchar(20)","任课教师编号",""),@("4","courseTime","nvarchar(50)","上课时间",""),@("5","coursePlace","nvarchar(50)","上课地点","")) },
        [pscustomobject]@{ name = "Elect选课信息表"; rows = @(@("1","electID","int not null","选课记录编号","关键字"),@("2","stuID","nvarchar(20)","学生学号","外键"),@("3","courseID","nvarchar(20)","课程编号","外键"),@("4","score","int","课程成绩","")) }
      )
    }
  }

  if ($DocumentText -match "书店|购物车|订单管理|图书管理|好书|Book图书") {
    return [pscustomobject]@{
      modules = @(
        @("用户模块", "图书浏览子模块", "展示新书、分类和图书详情"),
        @("用户模块", "购物车子模块", "维护待购买图书和数量"),
        @("用户模块", "订单管理子模块", "提交订单并查看订单状态"),
        @("评论模块", "图书评论子模块", "发布和查看图书评论"),
        @("后台模块", "图书管理子模块", "维护图书、分类、库存和上下架状态"),
        @("后台模块", "订单处理子模块", "审核订单、更新发货和销售记录")
      )
      database = @(
        @("1", "User", "存储用户账号与联系方式"),
        @("2", "Book", "存储图书基本信息"),
        @("3", "Category", "存储图书分类信息"),
        @("4", "Cart", "存储购物车明细"),
        @("5", "OrderInfo", "存储订单主表信息"),
        @("6", "Comment", "存储图书评论信息")
      )
      fieldTables = @(
        [pscustomobject]@{ name = "User用户表"; rows = @(@("1","userID","nvarchar(20) not null","用户编号","关键字"),@("2","userName","nvarchar(40)","用户姓名",""),@("3","password","nvarchar(40)","登录密码",""),@("4","phone","nvarchar(20)","联系电话","")) },
        [pscustomobject]@{ name = "Book图书表"; rows = @(@("1","bookID","nvarchar(20) not null","图书编号","关键字"),@("2","bookName","nvarchar(80)","图书名称",""),@("3","categoryID","nvarchar(20)","分类编号","外键"),@("4","price","decimal(10,2)","销售价格",""),@("5","stock","int","库存数量","")) },
        [pscustomobject]@{ name = "OrderInfo订单表"; rows = @(@("1","orderID","nvarchar(20) not null","订单编号","关键字"),@("2","userID","nvarchar(20)","用户编号","外键"),@("3","totalPrice","decimal(10,2)","订单金额",""),@("4","orderStatus","nvarchar(20)","订单状态","")) }
      )
    }
  }

  if ($DocumentText -match "校园导览|地点|路线|收藏") {
    return [pscustomobject]@{
      modules = @(
        @("用户端模块", "地点分类浏览", "展示教学楼、食堂、宿舍和公共服务点分类入口"),
        @("用户端模块", "关键词搜索", "按地点名称、标签和描述字段查询目标地点"),
        @("用户端模块", "地点详情查看", "展示地点简介、开放时间、位置说明和附近地标"),
        @("路线提示模块", "路线提示生成", "根据目标地点输出可理解的到达路径说明"),
        @("收藏模块", "常用地点收藏", "保存并展示用户常用地点"),
        @("数据管理模块", "地点数据维护", "维护地点、分类、路线和日志等基础数据")
      )
      database = @(
        @("1", "User", "存储用户账号与收藏状态"),
        @("2", "Place", "存储校园地点基本信息"),
        @("3", "Category", "存储地点分类信息"),
        @("4", "Favorite", "存储常用地点收藏记录"),
        @("5", "Route", "存储地点路线提示信息"),
        @("6", "SearchLog", "存储关键词搜索日志")
      )
      fieldTables = @(
        [pscustomobject]@{ name = "User用户表"; rows = @(@("1","userID","nvarchar(20) not null","用户编号","关键字"),@("2","userName","nvarchar(40)","用户名称",""),@("3","role","nvarchar(20)","用户角色",""),@("4","createTime","datetime","创建时间","")) },
        [pscustomobject]@{ name = "Place地点信息表"; rows = @(@("1","placeID","nvarchar(20) not null","地点编号","关键字"),@("2","placeName","nvarchar(80)","地点名称",""),@("3","categoryID","nvarchar(20)","分类编号","外键"),@("4","description","nvarchar(200)","地点简介",""),@("5","openTime","nvarchar(80)","开放时间",""),@("6","nearby","nvarchar(120)","附近地标","")) },
        [pscustomobject]@{ name = "Category地点分类表"; rows = @(@("1","categoryID","nvarchar(20) not null","分类编号","关键字"),@("2","categoryName","nvarchar(40)","分类名称",""),@("3","sortNo","int","排序号","")) },
        [pscustomobject]@{ name = "Favorite收藏记录表"; rows = @(@("1","favoriteID","int not null","收藏记录编号","关键字"),@("2","userID","nvarchar(20)","用户编号","外键"),@("3","placeID","nvarchar(20)","地点编号","外键"),@("4","createTime","datetime","收藏时间","")) },
        [pscustomobject]@{ name = "Route路线提示表"; rows = @(@("1","routeID","int not null","路线编号","关键字"),@("2","startPlace","nvarchar(80)","起点描述",""),@("3","targetPlaceID","nvarchar(20)","目标地点编号","外键"),@("4","routeText","nvarchar(300)","文字路线提示","")) }
      )
    }
  }

  return [pscustomobject]@{
    modules = @(
      @("用户模块", "信息查询子模块", "完成信息浏览、检索和详情查看"),
      @("业务模块", "核心处理子模块", "完成主要业务流程和状态更新"),
      @("数据模块", "数据维护子模块", "维护基础数据、关联数据和日志记录"),
      @("后台模块", "管理维护子模块", "完成信息维护、权限控制和异常处理")
    )
    database = @(
      @("1", "User", "存储用户账号信息"),
      @("2", "Role", "存储角色权限信息"),
      @("3", "Module", "存储功能模块信息"),
      @("4", "BusinessData", "存储核心业务数据"),
      @("5", "OperationLog", "存储操作日志")
    )
    fieldTables = @(
      [pscustomobject]@{ name = "User用户表"; rows = @(@("1","userID","nvarchar(20) not null","用户编号","关键字"),@("2","userName","nvarchar(40)","用户名称",""),@("3","password","nvarchar(40)","登录密码","")) },
      [pscustomobject]@{ name = "BusinessData业务数据表"; rows = @(@("1","dataID","nvarchar(20) not null","数据编号","关键字"),@("2","dataName","nvarchar(80)","数据名称",""),@("3","dataStatus","nvarchar(20)","数据状态","")) }
    )
  }
}

function New-CourseDesignTablesXml {
  param(
    [Parameter(Mandatory = $true)]
    [object]$Profile,

    [int]$ChapterNumber = 4,

    [int]$StartingSubsectionNumber = 1
  )

  $moduleSectionNumber = $StartingSubsectionNumber
  $databaseSectionNumber = $StartingSubsectionNumber + 1
  $fieldTableSectionNumber = $StartingSubsectionNumber + 2

  $xml = New-Object System.Collections.Generic.List[string]
  [void]$xml.Add((New-ParagraphXml -Text ("{0}.{1} 功能模块设计" -f $ChapterNumber, $moduleSectionNumber) -FirstLineTwips 420 -BeforeTwips 80 -AfterTwips 0 -LineTwips 320 -FontName "宋体" -SizeHalfPoints 21))
  [void]$xml.Add((New-TableBlockXml -Caption ("表{0}-1 功能模块表" -f $ChapterNumber) -Headers @("功能模块", "包含子功能模块", "功能") -Rows @($Profile.modules) -Widths @(1800, 2200, 4600) -MergeFirstColumn))
  [void]$xml.Add((New-ParagraphXml -Text ("{0}.{1} 数据库设计" -f $ChapterNumber, $databaseSectionNumber) -FirstLineTwips 420 -BeforeTwips 120 -AfterTwips 0 -LineTwips 320 -FontName "宋体" -SizeHalfPoints 21))
  [void]$xml.Add((New-TableBlockXml -Caption ("表{0}-2 数据库表" -f $ChapterNumber) -Headers @("序号", "数据库表", "数据表存储的内容") -Rows @($Profile.database) -Widths @(900, 2500, 5200)))
  [void]$xml.Add((New-ParagraphXml -Text ("{0}.{1} 数据库表结构" -f $ChapterNumber, $fieldTableSectionNumber) -FirstLineTwips 420 -BeforeTwips 120 -AfterTwips 0 -LineTwips 320 -FontName "宋体" -SizeHalfPoints 21))

  $tableNo = 3
  foreach ($fieldTable in @($Profile.fieldTables)) {
    $caption = "表{0}-{1} {2}" -f $ChapterNumber, $tableNo, [string]$fieldTable.name
    [void]$xml.Add((New-TableBlockXml -Caption $caption -Headers @("序号", "字段名", "字段类型", "说明", "备注") -Rows @($fieldTable.rows) -Widths @(800, 1500, 2500, 2200, 1200)))
    $tableNo++
  }

  return ($xml -join "`n")
}

function Write-OpenXmlPackage {
  param(
    [Parameter(Mandatory = $true)]
    [string]$SourceDirectory,

    [Parameter(Mandatory = $true)]
    [string]$DestinationPath
  )

  if (Test-Path -LiteralPath $DestinationPath) {
    Remove-Item -LiteralPath $DestinationPath -Force
  }

  $archive = [System.IO.Compression.ZipFile]::Open($DestinationPath, [System.IO.Compression.ZipArchiveMode]::Create)
  try {
    foreach ($file in Get-ChildItem -LiteralPath $SourceDirectory -Recurse -File) {
      $relativePath = $file.FullName.Substring($SourceDirectory.Length).TrimStart('\', '/') -replace '\\', '/'
      $entry = $archive.CreateEntry($relativePath)
      $entryStream = $entry.Open()
      try {
        $fileStream = [System.IO.File]::OpenRead($file.FullName)
        try {
          $fileStream.CopyTo($entryStream)
        } finally {
          $fileStream.Dispose()
        }
      } finally {
        $entryStream.Dispose()
      }
    }
  } finally {
    $archive.Dispose()
  }
}

$resolvedDocxPath = (Resolve-Path -LiteralPath $DocxPath).Path
if ([System.IO.Path]::GetExtension($resolvedDocxPath).ToLowerInvariant() -ne ".docx") {
  throw "Only .docx files are supported: $resolvedDocxPath"
}

if ([string]::IsNullOrWhiteSpace($OutPath)) {
  $directory = Split-Path -Parent $resolvedDocxPath
  $fileName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedDocxPath)
  $OutPath = Join-Path $directory ($fileName + ".course-tables.docx")
}

$resolvedOutPath = [System.IO.Path]::GetFullPath($OutPath)
if ((-not $Overwrite) -and (Test-Path -LiteralPath $resolvedOutPath)) {
  throw "Output file already exists: $resolvedOutPath. Re-run with -Overwrite to replace it."
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("openclaw-course-design-tables-" + [System.Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

try {
  [System.IO.Compression.ZipFile]::ExtractToDirectory($resolvedDocxPath, $tempRoot)
  $documentXmlPath = Join-Path $tempRoot "word\document.xml"
  if (-not (Test-Path -LiteralPath $documentXmlPath)) {
    throw "word/document.xml was not found in $resolvedDocxPath"
  }

  [xml]$documentXml = [System.IO.File]::ReadAllText($documentXmlPath, (New-Object System.Text.UTF8Encoding($false)))
  $namespaceManager = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
  [void]$namespaceManager.AddNamespace("w", $wordNamespace)

  $body = $documentXml.SelectSingleNode("/w:document/w:body", $namespaceManager)
  if ($null -eq $body) {
    throw "Could not locate /w:document/w:body in $resolvedDocxPath"
  }

  $documentText = Get-NodeText -Node $body -NamespaceManager $namespaceManager
  if ($documentText -match "表\d+-1\s*功能模块表|功能模块表.*数据库表结构") {
    Copy-Item -LiteralPath $resolvedDocxPath -Destination $resolvedOutPath -Force
    [pscustomobject]@{
      docxPath = $resolvedDocxPath
      outPath = $resolvedOutPath
      inserted = $false
      tableCount = 0
      reason = "course-design tables already exist"
    }
    return
  }

  $children = @($body.ChildNodes)
  $startIndex = -1
  $designHeadingText = $null
  for ($i = 0; $i -lt $children.Count; $i++) {
    if ($children[$i].LocalName -ne "p") {
      continue
    }
    $text = Get-NodeText -Node $children[$i] -NamespaceManager $namespaceManager
    if ($text -match "方案设计与实现|系统总体设计|系统设计") {
      $startIndex = $i
      $designHeadingText = $text
      break
    }
  }

  $insertBefore = $null
  if ($startIndex -ge 0) {
    for ($i = $startIndex + 1; $i -lt $children.Count; $i++) {
      if ($children[$i].LocalName -ne "p") {
        continue
      }
      $text = Get-NodeText -Node $children[$i] -NamespaceManager $namespaceManager
      if ($text -match "^(实现结果|五[\.．、]\s*实现结果|五、|问题与改进|六[\.．、]\s*问题)") {
        $insertBefore = $children[$i]
        break
      }
    }
  }

  if ($null -eq $insertBefore) {
    foreach ($child in @($body.ChildNodes)) {
      if ($child.LocalName -eq "sectPr") {
        $insertBefore = $child
        break
      }
    }
  }

  if ($null -eq $insertBefore) {
    throw "Could not find a safe insertion point for course-design tables."
  }

  $profile = Get-CourseDesignTableProfile -DocumentText $documentText
  $chapterNumber = Get-HeadingChapterNumber -Text $designHeadingText
  if ($null -eq $chapterNumber) {
    $chapterNumber = 4
  }
  $startingSubsectionNumber = Get-NextSubsectionNumber -Children $children -StartIndex $startIndex -StopNode $insertBefore -NamespaceManager $namespaceManager -ChapterNumber $chapterNumber
  $fragment = $documentXml.CreateDocumentFragment()
  $fragment.InnerXml = New-CourseDesignTablesXml -Profile $profile -ChapterNumber $chapterNumber -StartingSubsectionNumber $startingSubsectionNumber

  $insertedNodeCount = 0
  while ($fragment.HasChildNodes) {
    [void]$body.InsertBefore($fragment.FirstChild, $insertBefore)
    $insertedNodeCount++
  }

  [System.IO.File]::WriteAllText($documentXmlPath, $documentXml.OuterXml, (New-Object System.Text.UTF8Encoding($false)))
  Write-OpenXmlPackage -SourceDirectory $tempRoot -DestinationPath $resolvedOutPath

  [pscustomobject]@{
    docxPath = $resolvedDocxPath
    outPath = $resolvedOutPath
    inserted = $true
    insertedNodeCount = $insertedNodeCount
    chapterNumber = $chapterNumber
    startingSubsectionNumber = $startingSubsectionNumber
    tableCount = 2 + @($profile.fieldTables).Count
  }
} finally {
  if (Test-Path -LiteralPath $tempRoot) {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force
  }
}
