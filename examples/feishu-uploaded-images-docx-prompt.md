# Feishu Uploaded Images To Docx Prompt

```text
我会在当前飞书消息里直接上传实验截图，不再单独给本地图片路径。

如果你已经在之前的生成任务里设置过课程名和实验名，而且这次不变，那么 `-CourseName` 和 `-ExperimentName` 可以省略。

如果当前对话提示中已经包含类似 `[media attached ...]` 的附件路径提示，请直接从这些提示里提取图片路径，并把这些路径作为 `-ImagePaths` 传给本地脚本。不要只看图写正文，最终 docx 也要把这些附件图片真正插进去。

如果你没有拿到任何可读的附件路径，请明确说“当前运行时没有暴露附件文件路径，无法稳定把附件直接插入 docx”，不要假装已经插入成功。

工作目录：
E:\游戏\openclaw-experiment-report-skill

请直接运行：

powershell -ExecutionPolicy Bypass -File .\scripts\build-report-from-feishu.ps1 `
  -ReferenceUrls "https://blog.csdn.net/你的文章链接" `
  -CourseName "计算机网络" `
  -ExperimentName "局域网搭建与常用 DOS 命令使用" `
  -TemplatePath "E:\实验报告\实验报告模版1.docx" `
  -StudentName "李亦非" `
  -StudentId "244100198" `
  -ClassName "24c" `
  -TeacherName "李老师" `
  -ExperimentProperty "③验证性实验" `
  -ExperimentDate "2026年4月2日" `
  -ExperimentLocation "睿智楼四栋212" `
  -ImagePaths "从当前提示里的 [media attached ...] 行提取出的图片路径" `
  -OutputDir "E:\实验报告\新建文件夹" `
  -StyleProfile auto `
  -DetailLevel full

要求：
- 图1、图2属于实验步骤
- 图3、图4属于实验结果
- 四张图按 2x2 连续图片块排版
- 正文根据教程和我上传的图片来写
- 最终 docx 插图使用当前消息附件对应的图片文件
```
