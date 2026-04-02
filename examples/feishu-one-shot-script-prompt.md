# Feishu One-Shot Script Prompt

```text
请在仓库根目录直接运行本地脚本 .\scripts\build-report-from-feishu.ps1，不要自己临时拼接一长串中间 JSON。

如果你已经在之前的生成任务里设置过课程名和实验名，而且这次不变，那么 `-CourseName` 和 `-ExperimentName` 可以省略。

如果你当前工作目录不是仓库根目录，请先切换到：
E:\游戏\openclaw-experiment-report-skill

先确认下面这些路径或网址能访问；如果有任何一个不能访问，请明确说出“无法访问哪个路径或网址”，不要假装已经读取成功。

然后直接运行：

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
  -ImagePaths "E:\实验报告\step-1.png","E:\实验报告\step-2.png","E:\实验报告\result-1.png","E:\实验报告\result-2.png" `
  -OutputDir "E:\实验报告\新建文件夹" `
  -StyleProfile auto `
  -DetailLevel full

要求：
- 图1、图2属于实验步骤
- 图3、图4属于实验结果
- 四张图按 2x2 连续图片块排版
- 实验性质这一项要表现为勾选 ③验证性实验
- 结果部分以我的截图和已知事实为准
- 正文要比简略版更充实，但不要编造不存在的数据、截图细节或报错
```
