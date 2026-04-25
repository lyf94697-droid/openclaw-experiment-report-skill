# 一键演示样例包

这个目录专门给 `scripts/run-one-click-demo.ps1` 使用，也适合作为你替换真实材料前的参考。

## 包含内容

- `report.txt`
  样例实验报告正文

- `metadata.json`
  样例学生信息和课程信息

- `requirements.json`
  正文校验规则

## 搭配使用的其他仓库素材

- 模板：[`../report-templates/experiment-report-template.docx`](../report-templates/experiment-report-template.docx)
- 截图：[`../../demo/assets/`](../../demo/assets/)

## 最小可运行命令

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-one-click-demo.ps1
```

## 手动拆开运行

如果你想自己观察每一步，也可以直接调用：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report.ps1 `
  -TemplatePath ".\examples\report-templates\experiment-report-template.docx" `
  -ReportPath ".\examples\demo-one-click\report.txt" `
  -MetadataPath ".\examples\demo-one-click\metadata.json" `
  -RequirementsPath ".\examples\demo-one-click\requirements.json" `
  -StyleFinalDocx `
  -StyleProfile auto
```
