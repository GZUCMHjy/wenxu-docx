# POC 核心流程验证记录

## 2026-09-18 标题与字体扩展验收

同一 WSL Ubuntu/Python 环境执行 `python -m unittest discover -s tests -v`：**17 项通过**。新增覆盖标题1/2/3/9、自定义样式继承、直接大纲覆盖、中文标题样式名、无标题标记的粗体正文、空范围与非法字体/范围、宋体/黑体各作用范围、字体主题覆盖清理、无关字体 hint 保留，以及校验器拦截范围外改动。

`python tests/browser_smoke.py`：**PASS**。保留原全局格式流程与只改行距流程；新增真实浏览器选择“一级标题 → 黑体”“二级标题 → 宋体”“三级标题 → 黑体”，下载 DOCX 后独立回读检查中文/西文字体名称、字号不变、范围外段落 XML 不变及其他包部件原字节保全。选择不存在的九级标题时执行按钮禁用、旧下载消失；损坏文件在上传预检阶段阻断。

新增下载证据 `verification/Heading-1.docx`、`Heading-2.docx`、`Heading-3.docx` 和截图 `browser-heading-result.png`，完整结果见 `verification/browser-report.json`。示例文档已增加一级、二级、三级标题和各级后的正文。字体仅验证 DOCX 属性写回，不代表字体已嵌入或 WPS/Word 实机显示已经验收。

以下为初版历史记录；当前复测结果以上述 17 项及最新浏览器报告为准。

执行日期：2026-09-17。验收依据：用户最新要求的 WSL Python 运行和前端上传、固定参数修改、DOCX 实际修改流程。

## 环境与结果

- Ubuntu-24.04，WSL2，Linux `6.18.33.2-microsoft-standard-WSL2`。
- Python 3.12.3，虚拟环境 `/home/louis/.cache/vibe-edit-style/venv`。
- Streamlit 1.55.0、python-docx 1.2.0、lxml 6.0.2。
- Playwright 1.58.0，实际 Chromium 浏览器在 WSL 运行。
- Windows 请求 `http://127.0.0.1:8501/_stcore/health` 返回 `ok`。
- `python -m unittest discover -s tests -v`：`Ran 11 tests`，`OK`。
- `python tests/browser_smoke.py`：实际文件上传、控件操作和下载结束后输出 `result: PASS`。

## 浏览器结果独立检查

| 操作 | 下载结果检查 |
|---|---|
| 上传合成 DOCX，勾选字号/颜色/行距，设置 18 磅、黑色、固定 30 磅 | 下载 `sample_v1.docx`；字号 18，颜色 `000000`，行距 `w:line=600`、`w:lineRule=exact` |
| 仅启用行距，选择 2 倍 | `w:line=480`、`w:lineRule=auto`；原字号 12、原颜色 `222222` 保持 |
| 成功后修改参数 | 旧 DOCX 下载按钮消失 |
| 上传伪装为 DOCX 的损坏数据并执行 | 显示输入错误；旧下载按钮保持不可用 |
| 检查正文、强调和文件部件 | 文字顺序、加粗/斜体、缩进、段后距保持；所有非主文档 ZIP 部件逐字节一致，上传样本未改动 |

样本包含中文、混合 run、超链接、合并和嵌套表格、内嵌图片及页眉页脚。单元测试还验证：分别修改/组合修改、再次应用无变更、错误参数、损坏/超限文件、XML DTD、未支持结构和人为注入正文/段后距/图片损坏时校验失败。

页面截图人工检查：1280×960 桌面与 390×844 窄屏均可读，窄屏控件纵向排列。浏览器测试产物位于被 Git 忽略的 `verification/`：`browser-report.json`、`browser-success.png`、`browser-invalid.png`、`browser-mobile.png`、`sample_v1.docx` 和 `spacing-only.docx`。可按 README 的命令重跑生成。

本记录不代表 LibreOffice 渲染或 WPS/Word 客户端测试通过；这些尚未执行。当前只处理固定主体范围，未支持语义识别、任意复杂文档或持久化版本管理。
