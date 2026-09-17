# 文序 DOCX 格式调整 POC

上传 DOCX，识别已有标题级别，按范围修改字体（宋体/黑体）、字号、颜色和行距，下载另存的新文件。Python 服务与测试运行在 WSL Ubuntu；浏览器可从 Windows 访问。

## 在 WSL 启动

在 Ubuntu 终端执行（项目路径按实际仓库位置替换）：

```bash
cd /mnt/f/LouisProject/vibe-coding-project/vibe-edit-style
uv venv --python python3 "$HOME/.cache/vibe-edit-style/venv"
uv pip install --python "$HOME/.cache/vibe-edit-style/venv/bin/python" -r requirements.txt
"$HOME/.cache/vibe-edit-style/venv/bin/python" -m streamlit run app.py
```

`uv` 用于创建 WSL 虚拟环境和安装依赖。如果使用已有的 WSL venv，也可以用该环境的 `python -m pip install -r requirements.txt`。不要复用 Windows Python 或 Windows venv。虚拟环境放在 WSL 的 Linux 文件系统中，源代码可位于 `/mnt/f/`。

打开 [http://localhost:8501](http://localhost:8501)。日常重启只需进入项目目录执行最后一行。停止服务按 Ctrl+C。服务默认只监听本机，关闭使用统计，不调用 AI 或外部文档处理服务。

## 使用

1. 可先点“下载示例文档”，或准备自己的普通 DOCX。
2. 上传不超过 10 MB 的 DOCX。
3. 查看“标题识别结果”，选择全部段落、正文、文档标题/副标题或某一级标题；可展开查看本次修改的段落。
4. 勾选需要修改的项目，再选择宋体/黑体或其他参数；未勾选的控件不会生效。
5. 点击“应用格式并生成 DOCX”，校验通过后下载结果。

字号为 6–72 磅，步长 0.5；颜色为 RGB；行距支持固定 12–72 磅以及 1、1.5、2、2.5、3 倍。12 磅对应小四、14 磅对应四号，这些是字号换算，不代表单位规范。

作用范围可选择全部段落、正文（非标题）、文档标题/副标题，以及一至九级标题。识别读取段落直接大纲级别、段落样式继承链和文档默认大纲级别；没有大纲设置时依据已有 Heading 1–9 / 标题 1–9 样式名称。直接大纲级别为 9 时视为非标题；自定义样式继承 Heading 2 时也可识别为二级。页面展示识别依据和匹配原文，空范围不能执行。

“正文”表示没有标题标记的段落，包括相应的表格单元格和超链接文字。只加粗、放大或手打“一、”的普通段落不会被猜为标题，需要先在 Word/WPS 里设置标题样式或大纲级别。本功能不创建新标题层级，也不修改标题编号。选中某一级只影响该级段落，不包含其下正文。

页眉、页脚、图片、样式表和范围外段落保留；正文内容与未选属性会在输出前比较。不会清除加粗、斜体、下划线或改写文字。宋体/黑体写入 DOCX 字体名称，不内嵌字体文件；打开文档的电脑需安装对应字体，否则办公软件可能替换显示。

下载名称为 `原名_v1.docx`，上传 `原名_v1.docx` 再修改时得到 `原名_v2.docx`。这是基于输入名称的建议，不是持久化版本库。重复下载同一个结果不产生新版本；已具备相同直接格式时返回原文件。只处理内存副本，不覆盖磁盘源文件。

修改范围、参数，更换或清除上传文件后旧结果立即失效。失败时不提供旧结果下载。

## 实现边界

Streamlit 1.55.0 + python-docx 1.2.0 + lxml 6.0.2。python-docx 提供段落/run 属性接口，写回时只替换 DOCX 包内的 `word/document.xml`，其余部件保留原字节。校验将允许变化的属性移除后比较剩余完整 XML，并逐项检查目标值。

字体同时设置 `w:rFonts` 的 `ascii`、`hAnsi`、`eastAsia`、`cs`，去除对应的主题字体引用，并保留 `hint` 等无关属性；字号同时更新 `w:sz` 与 `w:szCs`；RGB 颜色去除该颜色的主题覆盖；行距只设置 `w:line` 和 `w:lineRule`，保留段前段后与缩进。主体遍历包括超链接和表格，合并单元格不会重复修改。格式校验只允许已选范围和属性发生变化。

属性依据：[Microsoft RunFonts](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.runfonts?view=openxml-3.0.1)、[Microsoft OutlineLevel](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.outlinelevel?view=openxml-3.0.1)。

不支持 DOC、加密、宏、数字签名、正文修订/批注、域、内容控件、文本框、公式、嵌入对象与外部模板/图片关系，检测到即返回原因。普通网页超链接保留，不会访问其目标。ZIP 解压总量上限 60 MB、最多 2000 部件，并拒绝 XML DTD。

本次按用户最新的“上传 → 修改固定参数 → 文档实际修改”验收，页面提供修改摘要和 DOCX 下载。LibreOffice/PDF 分页预览、缺字体检测、Word/WPS 实机兼容性矩阵属于后续验收，当前不宣称已通过。应用不建立上传文件缓存目录；文件仅位于当前会话内存。浏览器断开后内存由 Streamlit 会话回收，进程退出后释放，不承诺关闭标签页即刻清空内存。

## 验证

在 WSL 项目目录运行：

```bash
"$HOME/.cache/vibe-edit-style/venv/bin/python" -m unittest discover -s tests -v
```

真实浏览器验收需要开发依赖。保持应用运行，在另一个 WSL 终端执行：

```bash
uv pip install --python "$HOME/.cache/vibe-edit-style/venv/bin/python" -r requirements-dev.txt
"$HOME/.cache/vibe-edit-style/venv/bin/python" -m playwright install chromium
"$HOME/.cache/vibe-edit-style/venv/bin/python" tests/browser_smoke.py
```

若 Chromium 报系统动态库缺失，按 Playwright 的错误列表安装对应 Ubuntu 库后再执行。开发依赖不参与应用运行。

测试样例含中文、一至三级标题、混合 run、超链接、合并/嵌套表格、图片与页眉页脚。浏览器检查真实上传下载，独立读取下载结果验证格式；同时覆盖分级标题字体修改、范围外段落保留、空标题范围阻断、参数/范围变更失效、只改单项和损坏输入。单元测试还覆盖九级标题、中文样式名称、自定义样式继承和直接大纲覆盖。

浏览器脚本会生成/更新合成示例 `examples/sample.docx`，将下载文件、截图与报告写入忽略的 `verification/`。这些文件仅为合成验收数据。已执行环境及结果见 [VERIFICATION.md](VERIFICATION.md)。
