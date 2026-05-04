# os-process-scheduling：进程调度算法实验

这个案例适合操作系统课程里的算法实验：实现或模拟先来先服务、短作业优先、时间片轮转等调度策略，并分析等待时间、周转时间和调度结果。

## 输入文件

- `prompt.md`：生成正文时的实验事实和写作约束
- `report.txt`：可直接进入模板填充的示例正文
- `metadata.json`：学生和实验基础信息
- `requirements.json`：正文校验规则

## 可运行命令

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report.ps1 `
  -TemplatePath ".\examples\report-templates\experiment-report-template.docx" `
  -ReportPath ".\examples\cases\os-process-scheduling\report.txt" `
  -MetadataPath ".\examples\cases\os-process-scheduling\metadata.json" `
  -RequirementsPath ".\examples\cases\os-process-scheduling\requirements.json" `
  -OutputDir ".\tests-output\os-process-scheduling-case" `
  -StyleFinalDocx `
  -StyleProfile auto
```

## 适配建议

如果你有程序运行截图，可以再补一个 `image-specs.json`，把流程图放到“实验原理或任务要求”，把运行结果截图放到“实验结果”。不要把网上示例输出当成本次实验输出。
