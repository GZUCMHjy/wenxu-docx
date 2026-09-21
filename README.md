# 文序 DOCX 格式纠正 POC

本仓库实现《文序产品需求完整汇总版 v0.5》的 POC 主链路：导入自己的 DOCX，通过文字要求或参考实例生成要求表，人工核对对象、范围和规则，确认后只修改被选中的格式属性，完成内容、规则、文件和 Linux 渲染检查，再另存 DOCX 与独立 HTML 报告。

当前 POC 已覆盖 F01–F06 的软件流程。个人规范库、跨会话历史、组织发布和批量任务属于 F07–F10，按 PRD 留到第一期或第二期。

## 产品目标与技术设计

文序的下一阶段目标是：用户上传已经写好的 Word 文档，通过全局格式按钮或局部选择、文字要求，准确修改指定位置和属性，保留其余内容、格式与对象。

完整定义见 [产品与技术文档入口](docs/README.md)：

- [产品目标与需求定位](docs/PRODUCT.md)
- [v0.6 需求与验收场景](docs/REQUIREMENTS.md)
- [整体架构、分层能力与现有代码差距](docs/ARCHITECTURE.md)
- [交互架构图、JSON 图源和验证说明](docs/diagrams/README.md)

v0.6 是后续实现的目标基线。本页以下仍描述 v0.5 POC 的实际用法与限制；精确字符选择、表格栏目定位、DOC 转换和复杂对象保全等新目标尚未全部实现。

## WSL 安装与启动

所有编译、运行和验收都在 WSL Ubuntu 中执行。代码可以位于 `/mnt/d/`，虚拟环境应放在 Linux 文件系统中。

```bash
cd /mnt/d/vibe-project/wenxu-docx

# Ubuntu 中安装 Linux 渲染器和中文字体
sudo apt update
sudo apt install -y libreoffice fonts-noto-cjk

# 已安装 uv 时可跳过这一行
curl -LsSf https://astral.sh/uv/install.sh | sh

uv venv --python python3 "$HOME/.cache/wenxu-docx/venv"
uv pip install --python "$HOME/.cache/wenxu-docx/venv/bin/python" -r requirements.txt

export WENXU_LIBREOFFICE="$(command -v libreoffice)"
"$HOME/.cache/wenxu-docx/venv/bin/python" -m streamlit run app.py \
  --server.address 127.0.0.1 --server.port 8501 --browser.gatherUsageStats false
```

打开 [http://localhost:8501](http://localhost:8501)。如果没有 sudo 权限，可把官方 Linux LibreOffice 解压到用户目录，把 `WENXU_LIBREOFFICE` 指向其中的 `program/soffice`；字体可放入 `~/.local/share/fonts/` 后执行 `fc-cache -f`。

## 使用流程

1. 上传自己的 DOCX；文字要求和参考实例 DOCX 是并列入口，可单独使用或同时使用。
2. 点击“提取要求并生成核对表”。无法识别、缺对象/单位或互相冲突的内容会保留为待处理项。
3. 核对每段的标题、正文、落款或原样保留角色，再核对要求表的对象、属性、目标值、单位和来源。
4. 确认当前快照后执行。任何输入、要求或范围变化都会撤销旧确认和旧结果。
5. 检查通过后下载 `原名_vN.docx`、HTML 检查报告或诊断 JSON；可从任一会话内结果继续生成下一个未占用版本。

重复执行同一输入与规则快照会复用结果。全部要求已经满足时提示无需修改，不创建重复版本。下载不覆盖原稿；已提交、被退回和已接收是用户记录的交付观察。

## POC 支持范围

- 单个普通 DOCX，不超过 10 MB；解压后不超过 60 MB、2000 个部件。
- 单节 A4 纵向，渲染后不超过 50 页。
- 文档标题、一级至三级标题、正文、落款和用户指定的原样保留段落。
- 字体、字号、加粗、斜体、RGB 颜色、对齐、首行/左右缩进、段前/段后、固定或倍数行距、与下段同页、四边页边距。
- 已有 run、超链接和简单表格中的段落；普通图片、页眉、页脚、关系和其他包部件按原字节保全。
- 有效格式计算覆盖直接格式、字符/段落样式和样式继承。已满足目标值时不增加冗余直接格式。

POC 会阻断 DOC、损坏/加密文件、宏、数字签名、多节或非 A4 纵向文档，以及正文中的修订、批注、域、内容控件、文本框、公式、嵌入对象、外部模板或外链图片。普通网页超链接保留且不会访问目标。

结构识别以已有样式、大纲级别和文本位置为依据。弱证据会显示原文并交给用户确认，不显示伪精确置信度。POC 不改写正文、不新增或重排标题编号，也不自动套用所谓标准公文格式。

## 检查与数据处理

写回只替换 DOCX 包内 `word/document.xml`，其他部件保持原字节。校验器会：

- 比较正文字符、对象和未选择属性；
- 逐段检查确认规则的有效格式；
- 重新打开输出 DOCX；
- 用 Linux LibreOffice 转换输入和输出，并用 PyMuPDF 生成逐页预览；
- 在要求字体未安装、内容/关系丢失、规则未达到或渲染失败时阻断 DOCX 下载。

HTML 报告记录任务 ID、输入指纹、来源版本、规则与来源、应用范围、排除项、保留区域、检查结果、时间和实际处理环境。报告不会写入 DOCX 正文。

稿件、规则、版本和渲染结果仅保存在当前 Streamlit 会话内存或渲染调用的临时目录中；临时目录在调用结束时删除。关闭服务后不提供恢复能力，这是第一期 F08 的范围。

## WSL 验收

```bash
cd /mnt/d/vibe-project/wenxu-docx
PY="$HOME/.cache/wenxu-docx/venv/bin/python"

# 单元与故障注入
"$PY" -m unittest discover -s tests -v

# 真实浏览器：先保持应用运行
uv pip install --python "$PY" -r requirements-dev.txt
"$PY" -m playwright install chromium
"$PY" tests/browser_smoke.py

# 30 份合成矩阵，必须设置 Linux LibreOffice
export WENXU_LIBREOFFICE="$(command -v libreoffice)"
"$PY" tests/poc_acceptance.py
```

`browser_smoke.py` 验证真实上传、要求拆分、确认失效、Linux 渲染、v1 下载、HTML 报告和从 v1 局部调整得到 v2。`poc_acceptance.py` 验证 3 类、30 份合成文档，覆盖 10 组要求变更、20/10 开发与留出分组、三次幂等应用和逐份 Linux 渲染。运行证据写入被 Git 忽略的 `verification/`，已执行结果见 [VERIFICATION.md](VERIFICATION.md)。

合成证据不能替代真实试点材料或 Word/WPS 客户端验收。尚未实际验证的客户端和复杂文档不宣称兼容。
