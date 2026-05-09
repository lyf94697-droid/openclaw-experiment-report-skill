# 模板填充机制

模板填充的核心思路是“先读模板，再生成映射，最后写回文档”。不要把学校模板当成普通文本拼接目标；它通常有表格、封面字段、固定标题、空白段落和正文外框。

## 输入文件

常用输入有四类：

- `TemplatePath`：学校模板或仓库示例模板，通常是 `.docx`
- `ReportPath`：已经生成或手写好的报告正文
- `MetadataPath`：姓名、学号、班级、课程名、实验名等短字段
- `ImageSpecsPath` / `ImageSpecsJson`：截图路径、插入章节、图注和布局

## 处理链路

1. 提取模板结构

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\extract-docx-template.ps1 `
  -Path ".\examples\report-templates\experiment-report-template.docx"
```

提取结果会保留段落顺序、表格单元格和疑似字段，方便判断模板真正需要填什么。

2. 生成字段映射

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\generate-docx-field-map.ps1 `
  -TemplatePath ".\examples\report-templates\experiment-report-template.docx" `
  -ReportPath ".\examples\sample-report.txt" `
  -MetadataPath ".\examples\cases\network-dos\metadata.json" `
  -OutFile ".\tests-output\generated-field-map.json"
```

字段映射会把短字段填进对应单元格，把正文段落放到模板章节后面。它不是简单替换字符串，而是尽量保留模板结构。

3. 写回模板

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\apply-docx-field-map.ps1 `
  -TemplatePath ".\examples\report-templates\experiment-report-template.docx" `
  -FieldMapPath ".\tests-output\generated-field-map.json" `
  -OutFile ".\tests-output\filled.docx"
```

日常使用不需要手动跑每一步，`build-report.ps1` 会串起来执行。

## 字段映射的三种常见写法

### 标签键映射

适合模板里有明确字段名，例如“姓名”“学号”“课程名称”：

```json
{
  "fields": {
    "姓名": "张三",
    "学号": "20260001",
    "课程名称": "计算机网络"
  }
}
```

这种方式保守，优先填空白位或占位符，不轻易覆盖已有模板文字。

### 章节段落映射

适合“实验目的”“实验步骤”“实验结果”等固定标题：

```json
{
  "sectionFields": {
    "实验步骤": {
      "mode": "after",
      "paragraphs": [
        "配置两台主机的静态地址。",
        "使用 ping 命令验证连通性。"
      ]
    }
  }
}
```

`mode: "after"` 表示保留标题，把正文写到标题后面。

### 位置键映射

适合必须精确写到某个段落或单元格的模板：

```json
{
  "positionFields": {
    "P2": "计算机网络实验报告",
    "T1R2C2": "张三"
  }
}
```

这种方式更强，但也更脆弱；模板变动后位置可能失效。

## profile 的作用

`profiles/experiment-report.json` 和 `profiles/course-design-report.json` 负责约束：

- 默认章节顺序
- 字段别名
- 图注风格
- 图片默认宽度
- 课程设计和实验报告的差异

普通实验报告默认使用 `experiment-report`。课程设计必须显式传：

```powershell
-ReportProfileName course-design-report
```

## 常见失败与处理

- metadata 缺字段：补 `metadata.json`，不要把短字段塞进正文
- 模板标题和正文标题不一致：优先补 profile alias，而不是改学校模板
- 表格单元格太短：只填短字段，正文大段落放到章节正文区域
- 生成映射置信度低：先看 diagnostics，再决定是补输入还是手工指定字段
- 插图位置不稳：改用章节 anchor，例如 `实验结果`，少用易变化的 `P8`

## 最小原则

不要为了一个模板写通用文档引擎。这个项目优先服务中文大学实验报告和课程设计报告，只把常见模板填充路径做稳。
