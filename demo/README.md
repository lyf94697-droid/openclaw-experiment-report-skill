# Demo

这个目录放的是对外展示友好的轻量素材，适合：

- GitHub README 预览图
- PR / Issue 里的效果说明
- 小红书封面拼图
- 抖音 / 视频号的录屏首帧和转场图

## 已包含素材

- `assets/step-network-config.png`
- `assets/step-ipconfig.png`
- `assets/result-ping.png`
- `assets/result-arp.png`

## 2x2 布局预览

| Step 1 | Step 2 |
| --- | --- |
| ![Network config](assets/step-network-config.png) | ![ipconfig result](assets/step-ipconfig.png) |
| ![ping result](assets/result-ping.png) | ![arp result](assets/result-arp.png) |

## 对应的 image specs 示例

下面这份 JSON 可以直接作为“多图分组布局”的演示样例：

```json
{
  "images": [
    {
      "path": ".\\demo\\assets\\step-network-config.png",
      "section": "实验结果",
      "caption": "图1 虚拟机网络参数配置界面",
      "layout": { "mode": "row", "columns": 2, "group": "demo-grid" }
    },
    {
      "path": ".\\demo\\assets\\step-ipconfig.png",
      "section": "实验结果",
      "caption": "图2 使用 ipconfig 查看主机地址配置",
      "layout": { "mode": "row", "columns": 2, "group": "demo-grid" }
    },
    {
      "path": ".\\demo\\assets\\result-ping.png",
      "section": "实验结果",
      "caption": "图3 主机之间的 ping 连通性测试结果",
      "layout": { "mode": "row", "columns": 2, "group": "demo-grid" }
    },
    {
      "path": ".\\demo\\assets\\result-arp.png",
      "section": "实验结果",
      "caption": "图4 arp -a 邻居缓存查看结果",
      "layout": { "mode": "row", "columns": 2, "group": "demo-grid" }
    }
  ]
}
```

## 展示建议

- GitHub：主 README 放 1 张步骤图 + 1 张结果图，完整 2x2 预览放这里
- 小红书：直接拿 2x2 拼图做封面，标题突出“自动填模板 / 自动插图 / 自动排版”
- 抖音：把 4 张图按“输入材料 -> 最终成品”的顺序做 3 到 5 秒切换

如果你想直接跑一遍成品文档，请用 [docs/one-click-demo.md](../docs/one-click-demo.md) 里的演示命令。
