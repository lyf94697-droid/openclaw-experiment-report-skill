# Feishu Hybrid Images Prompt

```text
我会在飞书里直接上传实验截图，同时也会给你这些图片在我电脑上的本地路径。

如果教程标题、提示词或参考文本里能看出实验名称，可以省略 `-ExperimentName`；脚本会先自动推断，推断不到再复用上次保存的实验名。

请这样处理：
- 先以我上传的图片附件为准，识别每张图到底展示了什么内容
- 再使用本地路径作为最终 docx 插图文件来源
- 如果你看到了附件，但本地路径无法访问，请明确说出哪个路径不能访问
- 如果本地路径可访问，就直接运行本地脚本生成最终 docx

工作目录：
E:\游戏\openclaw-experiment-report-skill

请运行：

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
  -ImagePaths "E:\实验报告\step-1.png","E:\实验报告\step-2.png","E:\实验报告\result-1.png","E:\实验报告\result-2.png" `
  -OutputDir "E:\实验报告\新建文件夹" `
  -StyleProfile auto `
  -DetailLevel full

额外要求：
- 图1、图2属于实验步骤
- 图3、图4属于实验结果
- 四张图按 2x2 连续图片块排版
- 结果部分优先根据我上传的图片附件来写
- 最终插图仍使用我给出的本地路径文件
```
