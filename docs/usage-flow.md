# 使用流程

这条流水线把实验报告拆成四类输入：实验要求、参考资料、真实证据、学校模板。项目的目标不是替你凭空写一份报告，而是把这些材料整理成可复查的 `docx` 交付物。

## 总流程

1. 收集输入材料
   - 实验题目、课程名、姓名、学号、班级、指导教师等基础信息
   - 实验要求、评分点、老师给的模板或优秀示例
   - 教程链接、CSDN 文章或本地参考资料
   - 真实截图、命令输出、代码片段、实验结果数据

2. 生成或准备报告正文
   - 已经有正文时，直接把正文放进 `ReportPath`
   - 只有教程链接和要求时，先通过 OpenClaw 生成正文，再进入模板填充
   - 需要避免照抄参考文章时，按 [CSDN 参考内容如何避免照抄](csdn-reference-policy.md) 处理

3. 填充学校模板
   - 先提取模板结构
   - 生成字段映射
   - 把 metadata 和正文写回 `docx`
   - 复杂模板按 [模板填充机制](template-filling.md) 做诊断

4. 嵌入截图证据
   - 用 `image-specs` 或 `image-map` 说明每张图的路径、章节、图注和布局
   - 插图后生成 `image-placement-plan.md`
   - 具体规则见 [截图证据如何嵌入报告](screenshot-evidence.md)

5. 样式收尾和检查
   - 使用 profile 收尾字体、标题、图注和版式
   - 运行 layout check，确认图宽、图注、页面外框和表格线没有明显问题
   - 关键交付建议打开 `docx` 或导出 PDF 做人工复核

## 最快验证

仓库内置一键演示，使用已有正文、模板和截图，不需要在线生成正文：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-one-click-demo.ps1
```

如果只想检查文档、案例和 JSON 是否齐全：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\check-project-readiness.ps1
```

## 已有正文的本地流程

适合你已经有一份实验报告正文，只想把它填进学校模板并插入截图：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report.ps1 `
  -TemplatePath ".\examples\report-templates\experiment-report-template.docx" `
  -ReportPath ".\examples\sample-report.txt" `
  -MetadataPath ".\examples\cases\network-dos\metadata.json" `
  -ImageSpecsPath ".\examples\cases\network-dos\image-specs.json" `
  -RequirementsPath ".\examples\cases\network-dos\requirements.json" `
  -OutputDir ".\tests-output\network-dos-demo" `
  -StyleFinalDocx `
  -StyleProfile auto
```

输出目录会包含：

- `summary.json`：本次流水线摘要
- `generated-field-map.json`：模板字段映射
- `generated-image-map.json`：图片插入映射
- `image-placement-plan.md`：截图放置计划
- `layout-check.json`：版式检查结果
- 最终 `docx`：可打开检查和提交的成品

## 从教程链接生成正文

适合只有教程链接、实验题目和学校模板的情况：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report-from-url.ps1 `
  -ReferenceUrls "https://blog.csdn.net/..." `
  -CourseName "计算机网络" `
  -ExperimentName "局域网搭建与常用 DOS 命令使用" `
  -TemplatePath "E:\reports\template.docx" `
  -StudentName "张三" `
  -StudentId "20260001" `
  -ClassName "计科 2201" `
  -ImagePaths "E:\reports\step-1.png","E:\reports\result-1.png" `
  -OutputDir "E:\reports\final-output"
```

这条路径需要 OpenClaw 和浏览器 profile 可用。教程链接只作为参考材料，真实结果必须来自你的截图、命令输出或明确提供的数据。

## 课程设计报告

课程设计报告结构和普通实验报告不同，必须显式选择 profile：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report.ps1 `
  -TemplatePath ".\examples\report-templates\course-design-report-template.docx" `
  -ReportPath ".\examples\cases\course-design-student-management\report.txt" `
  -MetadataPath ".\examples\cases\course-design-student-management\metadata.json" `
  -RequirementsPath ".\examples\cases\course-design-student-management\requirements.json" `
  -ReportProfileName course-design-report `
  -OutputDir ".\tests-output\course-design-demo" `
  -StyleFinalDocx `
  -StyleProfile auto
```

课程设计更重视需求分析、总体设计、模块实现、运行结果和总结，流程图或总体设计图默认按大图处理。

## 交付前检查

交付前建议至少完成这三件事：

- 运行 `scripts\check-project-readiness.ps1`，确认项目展示材料没有缺文件
- 跑一次 `scripts\run-one-click-demo.ps1`，确认最小演示链路可用
- 对正式报告运行 layout check，并人工打开最终 `docx` 检查图、表、标题和分页
