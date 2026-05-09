# network-dos：局域网搭建与 DOS 命令实验

这个案例适合计算机网络课程里的验证性实验：配置两台主机地址，使用 `ipconfig`、`ping`、`arp` 等命令记录网络连通性和邻居缓存。

## 输入文件

- `prompt.md`：从实验要求生成正文时使用的提示词
- `metadata.json`：学生和实验基础信息
- `requirements.json`：正文校验规则和关键词
- `image-specs.json`：4 张演示截图的插入位置、图注和并排布局

## 可运行命令

这个案例可以直接复用仓库自带正文和截图跑完整 `docx` 流程：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report.ps1 `
  -TemplatePath ".\examples\report-templates\experiment-report-template.docx" `
  -ReportPath ".\examples\sample-report.txt" `
  -MetadataPath ".\examples\cases\network-dos\metadata.json" `
  -ImageSpecsPath ".\examples\cases\network-dos\image-specs.json" `
  -RequirementsPath ".\examples\cases\network-dos\requirements.json" `
  -OutputDir ".\tests-output\network-dos-case" `
  -StyleFinalDocx `
  -StyleProfile auto
```

## 预期输出

输出目录会包含最终 `docx`、字段映射、图片映射、图片放置计划和 layout check。打开最终文档时应重点检查：

- 顶部姓名、学号、课程名、实验名是否填入
- 4 张截图是否插入到实验结果附近
- 图注是否说明 `ipconfig`、`ping` 和 `arp` 的证据含义
- 并排截图是否清晰可读
