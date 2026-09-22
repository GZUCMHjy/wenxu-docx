# 文序 Word 格式调整工具

v0.6 在本地处理已有 DOCX 或旧式 DOC，通过全局控件、搜索与字符选择、段落或表格栏目定位，只修改本次明确指定的格式。以 master 的产品需求和架构为基线，实施计划、契约及测试记录均保存在仓库中。

- [产品与架构入口](docs/README.md)
- [开发计划与 spec](docs/specs/v06-editing.md)
- [实施状态和验收记录](docs/verification/v06.md)
- [v0.5 历史运行说明](docs/archive/v05-runtime.md)

## WSL 启动

```bash
cd /mnt/d/vibe-project/wenxu-docx
uv venv --python python3 "$HOME/.cache/wenxu-docx/venv"
uv pip install --python "$HOME/.cache/wenxu-docx/venv/bin/python" -r requirements.txt
# 安装 Linux LibreOffice 与所需字体，或指定已有的 Linux 便携版。
export WENXU_LIBREOFFICE="$(command -v libreoffice)"
"$HOME/.cache/wenxu-docx/venv/bin/python" -m streamlit run app.py \
  --server.address 127.0.0.1 --server.port 8501 --browser.gatherUsageStats false
```

打开 [本地应用](http://localhost:8501)。现有环境有虚拟环境时，直接安装更新后的依赖即可。不要在已有环境上重建虚拟环境。程序与测试在 WSL/Linux 内运行；源代码可在 Windows 磁盘。

WSL 所在的 Windows 安装可用 WPS 时，默认通过本机自动化接口预览文档，并在 DOC 原格式副本上修改。上传后工作副本与原稿字节一致：DOC 输出 DOC，DOCX 输出 DOCX。DOC 的只读结构快照仅用于定位和校验，不作为预览、下载或回写的数据。文件仅在本机临时目录处理；原稿和实际输出分别重新打开渲染。没有 WPS 时可用 LibreOffice 处理 DOCX；DOC 原格式编辑明确提示需要本机 WPS，不自动转换。`WENXU_RENDERER=wps` 强制本机 WPS，`WENXU_RENDERER=libreoffice` 显式使用 Linux 引擎，修改后重启服务。引擎失败不会自动切换；自动化接口如果返回含其他稿件的实例，本次任务停止，不改变那些稿件。

旧版文字/参考实例工作流保留在 `legacy_app.py`，同一服务访问 [旧版入口](http://localhost:8501/?legacy=1)。旧版支持范围未扩大。

## 使用流程

1. 上传自己的 DOCX 或 DOC，保留原始字节。DOC 的原稿和工作副本分别独立渲染，比较文字、分页和全部页面像素；相同文件渲染不稳定时停止编辑。可直接下载未修改的 DOC 工作副本，文件与上传原稿相同。
2. 核对需要的段落角色和保护区。全局正文默认包含表格内正文；表头和栏目标签不自动归入正文，待归类项明确提示。
3. 全局选正文/标题/表格等范围，或局部搜索文字、选段落字符区间、栏目填写区、表格行列和合并单元格。
4. 选择本次要设置的属性。未选择的保持原样；加粗、斜体、下划线支持保持/开启/关闭。当前格式与本次目标分开展示。
5. 将控件设置、文字要求或明确采用的参考格式加入清单，查看命中位置、目标属性、局部例外、保护排除和冲突。可改目标值、取消一项或移除命中段落。
6. 点击“应用修改并检查”，系统先检查内容、对象和目标属性，再渲染前后页面。全部已满足时不生成重复版本。
7. 查看页面并确认接受版式后按上传格式下载新 DOC 或 DOCX；HTML 报告独立下载。继续调整自动基于当前版本，控件恢复保持原样；撤销回到上一个版本。

文字输入采用确定性有限语法，例如“正文宋体小四，但第三段黑体”“教学目标这一栏改黑体五号”。不完整或未理解的内容保留为待处理项，不执行正文、批注中出现的指示。

## 支持与校验边界

- 普通 DOCX / 原格式编辑的二进制 DOC，10 MB 文件上限，60 MB 解压上限，2000 部件，单节，渲染 1–50 页。
- 字符：字体（全部/仅中文/仅西文）、字号、加粗、斜体、单线下划线、RGB 颜色；可跨多个 run 精确选择。
- 段落：对齐、首行/左右缩进、段前/段后、固定/倍数行距、与下段同页。局部选字设置段落属性会显示整段影响。
- 页面：单节四边页边距。表格逻辑行列考虑横向/纵向合并；不自动重排列宽和边框。
- 图片、原生公式、批注、字段、嵌入对象、表格合并、页眉页脚按原结构/部件保留；其内部内容不开放编辑。
- 修订、文本框、内容控件、多节、宏、签名、外部模板和外链对象等未验证结构继续拒绝。

DOCX 写回只替换 `word/document.xml`。其他 ZIP 部件原字节保留；独立校验器把 run 拆分归一到字符，检查文字顺序、对象、非目标属性和实际目标值。公式、字段和嵌入对象不参与普通文本格式操作。

目标字体缺失、结构/保全失败或渲染失败均不登记可下载版本。已有字体缺失可能发生替代，报告明确披露。自动校验、当前引擎渲染、人工版式复核和其他 Word/WPS 版本兼容性是不同状态。DOC 导入不转换格式；修改后独立检查原生结构快照和嵌入对象流。保存过程中允许明确列出的修订编号等元数据更新，任何范围外内容或格式变化均阻断交付。只有所有页面的实际预览一致时才显示“逐像素一致”，该结论限定于当前引擎和 120 DPI；文字一致不代表图像、字距或版面完全一致。

## 本地验证

```bash
PY="$HOME/.cache/wenxu-docx/venv/bin/python"
"$PY" -m unittest discover -s tests -v
uv pip install --python "$PY" -r requirements-dev.txt
"$PY" -m playwright install chromium
# 保持应用运行后执行实际浏览器验证
"$PY" tests/v06_browser.py
"$PY" tests/v06_doc_browser.py
WENXU_RENDERER=libreoffice "$PY" tests/v06_import_render.py
# 本机 WPS 可用时，验证当前引擎的字体选择与下载
"$PY" tests/v06_native_browser.py
"$PY" tests/v06_native_doc.py

# 可选真实稿件验证；产物目录必须在仓库外
"$PY" tests/local_sample_acceptance.py /absolute/local/sample.doc \
  --output-dir /absolute/local/private-validation
```

项目关闭浏览器使用统计，只监听本机。稿件、只读结构索引、版本和预览保存在当前会话内存或临时目录；不调用云端模型/文档转换服务，不提供跨会话恢复。真实稿件、正文、批注、转换件及截图禁止进入 Git；仓库测试使用合成文档。私人验证脚本强制把输出保存在仓库外。
