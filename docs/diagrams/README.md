# 文序技术分层图

本图表达 [v0.6 目标架构](../ARCHITECTURE.md) 的逻辑分层，不代表这些模块或功能已经全部实现。Archify 版本为 2.17，图类型为 `architecture`，质量档位为 `showcase`。

![文序技术分层图](wenxu-layers.visual-check.1440x900.light.png)

## 查看与修改

- [交互 HTML](wenxu-layers.html)：下载或克隆仓库后，用浏览器打开；GitHub 文件页不直接运行 HTML。
- [可编辑 JSON 图源](wenxu-layers.architecture.json)：修改后应重新执行 Archify 的 validate、deliver、visual-check，更新图与证据。
- [浅色/深色截图联系表](wenxu-layers.visual-check.html)：本地浏览器打开，可对照两种尺寸和主题。
- [交付收据](delivery-receipt.json)与[浏览器检查记录](wenxu-layers.visual-check.json)：核对产物哈希与检查范围。

图中的八个业务模块加任务与版本管理均为逻辑职责。任务管理贯穿全程；渲染前要先进行结构与保全检查，图中的最终校验节点汇总检查结果。存储图例不代表已经选定数据库，校验图例不代表独立安全服务。

## 验证记录

原始图源、HTML 和四张架构截图按原字节归档。两份 JSON 收据中的本机绝对路径改为同目录相对路径，浏览器可执行路径移除；校验结果、测量值、产物哈希保持原记录。这是经过路径整理的归档记录。

```text
diagram_type: architecture
output: docs/diagrams/wenxu-layers.html
specification_sha256: 03d01414c0da657a6818ec63264740bf016958c8469140f535eb823a46cc2630
artifact_sha256: 00e3553f807553fa2be22801d11951147950df8538be17184711ede35bb2efee
validation: 9/9 showcase, 0 errors, 0 warnings
browser_evidence: passed
visual_review: passed
correction_rounds: 1
```

图源 3366 字节，HTML 809426 字节。原始浏览器检查在 Google Chrome 中完成：1440×900、1600×1000、1920×1080、2048×1320 四种视口均无页面溢出；首尾两个尺寸分别保存浅色、深色截图。

图像审阅已检查四张截图中节点、连线、标签和结论卡的遮挡、裁切与整体分布。自动浏览器收据的 `visualReview: pending` 是工具保留状态，独立截图复核结果为上面的 `visual_review: passed`。没有额外宣称已测试搜索、聚焦和导出等交互。

这些结果只证明图示产物及其已检查的呈现范围，不属于文序产品功能、真实稿件或 Word/WPS 兼容性验收。文序产品验收仍在 WSL/Linux 中执行。
