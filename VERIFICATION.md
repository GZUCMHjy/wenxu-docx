# 文序 POC 验证记录

## 2026-09-20 PRD v0.5 闭环验收

所有命令、Python 进程、LibreOffice 转换和 Chromium 浏览器均在 WSL2 Ubuntu 中运行。Windows 仅提供仓库挂载和离线下载缓存，不参与应用执行或文档渲染。

### 固定环境

- WSL2 Linux `5.10.16.3-microsoft-standard-WSL2`，Ubuntu 24.04.1。
- Python 3.12.3，虚拟环境 `/home/louis/.cache/wenxu-docx/venv`。
- Streamlit 1.55.0、python-docx 1.2.0、lxml 6.0.2、PyMuPDF 1.28.2。
- LibreOffice Linux x86-64 26.2.6.3；官方发布包 SHA-256 校验通过。
- Noto Serif/Sans CJK SC 用户字体；`fc-match` 返回对应 Noto CJK 字体文件。
- Playwright 1.58.0，Chrome for Testing headless shell 145.0.7632.6。

### 单元与故障注入

命令：

```bash
/home/louis/.cache/wenxu-docx/venv/bin/python -m unittest discover -s tests -v
```

结果：`Ran 31 tests`，`OK`。

覆盖范围包括损坏/伪装/超限 DOCX、危险 XML、未支持对象、内容和媒体故障注入、标题/正文范围、样式继承、范围外属性保全、多对象多属性规则、缺对象/单位、冲突、无法解析要求、参考实例提取、报告字段、无变更和重复执行字节幂等。

### Linux 渲染检查

`examples/sample.docx` 已修正为单节 A4 纵向。WSL 内执行预检得到：16 段、2 个表格、1 张图片、1 个页眉、1 个页脚、1 个外部网页超链接，大小 38,786 字节。

LibreOffice 成功将 DOCX 转换为 1 页 PDF，PyMuPDF 成功生成 1 张 120 DPI PNG。首次渲染暴露中文字体缺失产生方框；安装 Noto CJK 并刷新 fontconfig 后重新渲染，人工查看确认页眉、标题、正文、表格、一级至三级标题和页脚中文均可读。

证据位于忽略目录 `verification/`：`render-sample-fonts.pdf` 与 `render-sample-fonts-1.png`。

### 真实浏览器主流程

命令：

```bash
export LD_LIBRARY_PATH=/home/louis/.local/opt/libasound/usr/lib/x86_64-linux-gnu
/home/louis/.cache/wenxu-docx/venv/bin/python tests/browser_smoke.py
```

结果：`PASS`。

浏览器实际完成：

1. 上传 A4 DOCX。
2. 输入一条同时包含一级标题和正文要求的文字，拆成 6 条规则并显示各自来源。
3. 确认前执行按钮禁用；确认后才可执行。
4. Linux 渲染输入和输出，下载 `sample_v1.docx`；独立回读验证一级标题为 Noto Sans CJK SC 18 磅居中，正文为 Noto Serif CJK SC 14 磅、固定行距 28 磅。
5. 检查正文字符不变、所有非 `word/document.xml` 部件逐字节不变。
6. 下载 HTML 报告，检查任务 ID、输入指纹、处理环境、规则、检查结果和保留区域字段。
7. 更改要求后确认和旧下载立即失效。
8. 以 v1 为明确输入，只把正文行距改为固定 30 磅，下载 `sample_v2.docx`；独立回读确认 v1 的标题与正文字体/字号保持。

证据：`verification/browser-report.json`、`browser-prd-result.png`、`browser-prd-mobile.png`、`sample_v1.docx`、`sample_v2.docx` 和 `sample_v1_format_report.html`。

### 30 份合成矩阵

命令：

```bash
export WENXU_LIBREOFFICE=/home/louis/.local/opt/libreoffice-26.2.6/opt/libreoffice26.2/program/soffice
/home/louis/.cache/wenxu-docx/venv/bin/python tests/poc_acceptance.py
```

结果：`PASS`，耗时 367.323 秒。

| 指标 | 结果 |
|---|---:|
| 文档数 | 30 |
| 开发集 / 留出集 | 20 / 10 |
| 合成文种 | 3 |
| 要求变更对 | 10 |
| 规则实例 | 33 |
| 内容保全率 | 100% |
| 严格文档通过率 | 100% |
| 每份相同规则连续应用 | 3 次，输出字节一致 |
| Linux 渲染 | 30/30 成功，均在 50 页以内 |

逐份指纹、分组、文种、要求对和页数在 `verification/poc-acceptance-30.json`。

## 结论边界

F01–F06 的本地 POC 软件链路和 A01–A12 的核心行为已有 WSL 自动化证据。30 份矩阵是确定性合成材料，只证明当前支持子集的引擎和流程，不代表真实业务样本验收。当前没有在 Microsoft Word 或 WPS 中执行 A13，也没有真实用户试点、商业指标或跨会话恢复证据，因此不宣称这些项目已通过。
