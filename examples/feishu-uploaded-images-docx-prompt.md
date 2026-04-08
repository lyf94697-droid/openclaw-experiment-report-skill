# Feishu Uploaded Images To Docx Prompt

```text
我会在飞书里直接上传实验截图，不再单独给本地图片路径。

如果教程标题、提示词或参考文本里能看出实验名称，可以省略 `-ExperimentName`；脚本会先自动推断，推断不到再复用上次保存的实验名。

请先检查当前对话或最近几条图片消息里是否包含类似 `[media attached ...]` 的附件路径提示。

如果能拿到真实附件路径：
- 请把这些路径按上传顺序逐个列在 `-ImagePaths` 后面，用逗号分隔，按实际数量增删
- 运行本地脚本时保留默认图片归档行为，让脚本先把图片复制到输出目录的 `images\` 子目录
- 最终 docx 插图使用归档后的图片文件，不要只看图写正文

如果没有拿到任何可读的附件路径，请明确说：
“当前运行时没有暴露附件文件路径，无法稳定把手机/飞书上传图片直接插入 docx。”
不要假装已经插入成功。

工作目录：
E:\游戏\openclaw-experiment-report-skill

请直接运行：

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
    "从 [media attached ...] 行提取出的第 1 个真实图片路径", `
    "从 [media attached ...] 行提取出的第 2 个真实图片路径" `
  -OutputDir "E:\实验报告\新建文件夹" `
  -StyleProfile auto `
  -DetailLevel full

要求：
- 图片按上传顺序编号为图1、图2、图3……
- 如果我没有额外说明图片归属，请根据图片内容判断属于实验步骤还是实验结果
- 如果能看出每张图在做什么，请优先使用 `-ImageSpecsJson` 或 `-ImageSpecsPath` 写明每张图的 `caption` 和 `section`，不要只写“实验步骤截图”或“实验结果截图”这类泛化图注
- 如果图片数量适合分组，优先使用每行 2 张的分组布局；但宽屏虚拟机截图、浏览器截图、命令行截图如果两张一行会导致文字不清楚，请改用单张堆叠布局
- ipconfig、ping、arp -a 等 DOS/终端命令必须单独成段，不要和说明文字写在同一段，方便最终 docx 自动套用命令块排版
- 正文根据教程和我上传的图片来写
- 最终 docx 必须真正插入图片文件
```
