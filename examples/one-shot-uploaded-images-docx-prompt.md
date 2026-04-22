# One-Shot Uploaded Images To Docx Prompt

```text
请优先使用本仓库的本地 wrapper 连贯处理实验报告生成、模板填充、截图插入和最终排版，不要默认拆成很长的手工任务列表。

工作目录：
E:\游戏\openclaw-experiment-report-skill

默认输出目录：
E:\实验报告\新建文件夹

基础信息：
- 课程名称：计算机网络
- 模板路径：E:\实验报告\实验报告模版1.docx
- 姓名：李亦非
- 学号：244100198
- 班级：24c
- 指导教师：李老师
- 实验性质：③验证性实验
- 实验时间：2026年4月2日
- 实验地点：睿智楼四栋212

实验输入：
- 如果我提供教程链接，请把教程内容改写成实验报告正文，不要长篇照抄。
- 如果我提供已有报告正文或本地 report 文件，请优先使用已有正文。
- 如果教程标题、提示词或参考文本里能看出实验名称，可以省略 `-ExperimentName`；脚本会先自动推断，推断不到再复用最近一次保存的实验名。
- 如果我提供了实验要求、代码、截图、输出结果，以这些真实材料为准，不要编造看不到的测量值、截图内容或错误日志。

图片处理要求：
- 如果当前对话包含上传图片，先检查是否有类似 `[media attached ...]` 的真实附件路径提示。
- 如果我同时提供了本地图片路径和上传图片，请把上传图片当作语义参考，把本地路径当作最终 docx 插图文件来源。
- 如果没有拿到任何可读的真实图片路径，请明确说“当前运行时没有暴露附件文件路径，无法稳定把上传图片直接插入 docx。”不要假装插入成功。
- 判断每张图属于 `实验步骤`、`实验结果`、`问题分析` 或其他合适章节；优先依据图片可见内容，其次依据文件名，最后才按上传顺序兜底。
- 不要因为图片分配置信度低就停下来问我；低置信度时按最佳判断继续，并在最后总结里列出低置信度图片。
- 尽量为每张图写具体图注，例如“IIS 站点配置结果界面”“客户端浏览器访问网站结果”，不要只写“实验步骤截图”这类泛化图注。
- 多张同章节图片默认使用每行 2 张布局；图片很多时按章节自然分组，优先形成 2x2 小块，避免把无关章节混到同一组。
- 章节图片应放在该章节正文末尾、下一章节标题之前，不要紧跟章节标题后硬塞。
- 最终 docx 必须真正插入图片文件。

建议执行流程：
1. 先生成或准备实验报告正文。
2. 用 `scripts/generate-docx-field-map.ps1` 和 `scripts/apply-docx-field-map.ps1` 填充模板，得到已填正文的 docx。
3. 在插图前运行一次图片分配预案，只作为内部记录，不要等我确认：

   powershell -ExecutionPolicy Bypass -File .\scripts\generate-docx-image-map.ps1 `
     -DocxPath "已填正文的 docx 路径" `
     -ImagePaths "第1张真实图片路径","第2张真实图片路径" `
     -Format markdown `
     -PlanOnly `
     -OutFile "E:\实验报告\新建文件夹\image-placement-plan.md"

4. 继续生成正式 image map，不要中断等待确认：

   powershell -ExecutionPolicy Bypass -File .\scripts\generate-docx-image-map.ps1 `
     -DocxPath "已填正文的 docx 路径" `
     -ImagePaths "第1张真实图片路径","第2张真实图片路径" `
     -Format json `
     -OutFile "E:\实验报告\新建文件夹\generated-image-map.json"

5. 运行 `scripts/insert-docx-images.ps1` 插入图片。
6. 运行 `scripts/format-docx-report-style.ps1 -Profile auto` 做最终排版。
7. 运行 `scripts/check-docx-layout.ps1` 检查图片数、图注数、占位符和章节。

如果直接使用一站式 wrapper，也可以调用：

powershell -ExecutionPolicy Bypass -File .\scripts\build-report-from-feishu.ps1 `
  -ReferenceUrls "https://blog.csdn.net/你的文章链接" `
  -CourseName "计算机网络" `
  -TemplatePath "E:\实验报告\实验报告模版1.docx" `
  -StudentName "李亦非" `
  -StudentId "244100198" `
  -ClassName "24c" `
  -TeacherName "李老师" `
  -ExperimentProperty "③验证性实验" `
  -ExperimentDate "2026年4月2日" `
  -ExperimentLocation "睿智楼四栋212" `
  -ImagePaths `
    "第1张真实图片路径", `
    "第2张真实图片路径", `
    "第3张真实图片路径", `
    "第4张真实图片路径" `
  -OutputDir "E:\实验报告\新建文件夹" `
  -StyleProfile auto `
  -DetailLevel full

如果你能根据上传图片内容可靠判断章节和图注，优先生成 `-ImageSpecsJson` 或 `-ImageSpecsPath` 并传给 wrapper，这比只传 `-ImagePaths` 更稳定。

正文排版要求：
- ipconfig、ping、arp -a、PowerShell、cmd 等命令必须单独成段，不要和说明文字写在同一段。
- 实验步骤和实验结果要具体，能对应截图和验证过程。
- 问题分析写真实常见问题和排查逻辑，不要写空泛套话。
- 实验总结保持简洁，说明完成了什么、掌握了什么、后续可扩展什么。

最终回复只需要给我：
- final docx 路径
- image-placement-plan.md 路径，如果已生成
- layout-check 是否通过
- 低置信度图片分配列表，如果有
- 你实际运行过的关键命令和是否成功
```
