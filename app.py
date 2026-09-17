"""Run inside WSL: python -m streamlit run app.py"""

import hashlib
import json
from pathlib import Path

import streamlit as st

from docx_formatting import FONT_NAMES, SCOPES, FormatError, apply_formatting, inspect_document, output_filename

st.set_page_config(page_title="文序 · DOCX 格式调整", page_icon="📄", layout="centered")
st.title("文序 · DOCX 格式调整")
st.write("上传文档，选择要调整的格式，下载新文件。")
st.caption("按标题级别或正文选择修改范围，仅修改勾选的属性；页眉页脚保持原样。")

sample = Path(__file__).parent / "examples" / "sample.docx"
if sample.exists():
    st.download_button("下载示例文档", sample.read_bytes(), "教学工作总结.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", on_click="ignore")

uploaded = st.file_uploader("上传 DOCX 文档（最大 10 MB）", type=["docx"])
source = uploaded.getvalue() if uploaded else None
source_hash = hashlib.sha256(source).hexdigest() if source else None
if st.session_state.get("inspection_hash") != source_hash:
    st.session_state["inspection_hash"] = source_hash
    st.session_state["inspection"] = []
    st.session_state["inspection_error"] = None
    if source:
        try:
            st.session_state["inspection"] = inspect_document(source)
        except FormatError as exc:
            st.session_state["inspection_error"] = str(exc)
paragraphs = st.session_state.get("inspection", [])
inspection_error = st.session_state.get("inspection_error")
scope = st.selectbox("修改范围", list(SCOPES), format_func=SCOPES.get)
selected = [p for p in paragraphs if scope == "all" or p["scope"] == scope]
if inspection_error:
    st.error(inspection_error)
elif source:
    st.caption(f"识别到 {len(paragraphs)} 个段落，本次选中 {len(selected)} 个。")
    if not selected:
        st.warning(f"没有识别到{SCOPES[scope]}，请选择其他范围，或先在 Word/WPS 中设置标题样式。")
    headings = [p for p in paragraphs if p["scope"].startswith("heading_") or p["scope"] == "title"]
    with st.expander(f"标题识别结果（{len(headings)} 段）", expanded=True):
        if headings:
            st.dataframe([{"段落": p["index"], "类型": p["kind"], "原文": p["text"][:120], "识别依据": p["source"]} for p in headings], hide_index=True, width="stretch")
        else:
            st.write("未发现标题样式或大纲级别。")
        st.caption("读取文档已有的标题样式、大纲级别及样式继承，支持一至九级。仅加粗、放大或手打序号的文字不自动判为标题。文档标题/副标题单独列出。")
    with st.expander("查看本次修改的段落"):
        st.dataframe([{"段落": p["index"], "类型": p["kind"], "原文": p["text"][:120]} for p in selected], hide_index=True, width="stretch")
st.subheader("选择修改项")
options = {}
font_left, font_right = st.columns(2)
with font_left:
    font_enabled = st.checkbox("修改字体")
    family = st.selectbox("字体", FONT_NAMES, disabled=not font_enabled)
    if font_enabled:
        options["font_family"] = family
with font_right:
    st.caption("支持宋体、黑体（含中文字体设置）。打开结果的电脑需安装对应字体，否则办公软件会使用替代字体。")
left, middle, right = st.columns(3)
with left:
    size_enabled = st.checkbox("修改字号")
    size = st.number_input("字号（磅）", min_value=6.0, max_value=72.0, value=14.0, step=0.5, disabled=not size_enabled)
    st.caption("12 磅 = 小四；14 磅 = 四号")
    if size_enabled:
        options["font_size"] = size
with middle:
    color_enabled = st.checkbox("修改字体颜色")
    color = st.color_picker("字体颜色", value="#000000", disabled=not color_enabled)
    if color_enabled:
        options["color"] = color
with right:
    spacing_enabled = st.checkbox("修改行距")
    mode = st.selectbox("行距模式", ["固定值（磅）", "倍数"], disabled=not spacing_enabled)
    if mode == "固定值（磅）":
        value = st.number_input("固定行距（磅）", min_value=12.0, max_value=72.0, value=28.0, step=0.5, disabled=not spacing_enabled)
        spacing = {"mode": "exact", "value": value}
    else:
        value = st.selectbox("行距倍数", [1.0, 1.5, 2.0, 2.5, 3.0], disabled=not spacing_enabled)
        spacing = {"mode": "multiple", "value": value}
    if spacing_enabled:
        options["line_spacing"] = spacing

fingerprint = (source_hash, uploaded.name, scope, json.dumps(options, sort_keys=True)) if source else None
# A file or option change invalidates the old result, including after a failed run.
if st.session_state.get("request") != fingerprint:
    st.session_state.pop("result", None)
    st.session_state.pop("error", None)
    st.session_state["request"] = fingerprint

if not options:
    st.info("勾选至少一项格式后即可执行。")
if st.button("应用格式并生成 DOCX", type="primary", disabled=not source or not options or bool(inspection_error) or not selected):
    st.session_state.pop("result", None)
    st.session_state.pop("error", None)
    try:
        with st.spinner("正在修改并校验文档…"):
            result, summary = apply_formatting(source, options, scope=scope)
        st.session_state["result"] = (result, summary)
    except FormatError as exc:
        st.session_state["error"] = str(exc)
    except Exception:
        st.session_state["error"] = "处理失败，未生成可下载的结果。请检查文件或更换普通 DOCX 后重试。"

if error := st.session_state.get("error"):
    st.error(error)
if result := st.session_state.get("result"):
    data, summary = result
    changed = any(summary.values())
    if changed:
        st.success("修改完成，内容与未选属性校验通过。")
    else:
        st.info("所选属性已满足，无需修改。下载将返回原文件。")
    rows = [{"修改项": "作用范围", "目标值": SCOPES[scope]}]
    if "font_family" in options:
        rows.append({"修改项": "字体", "目标值": options["font_family"]})
    if "font_size" in options:
        rows.append({"修改项": "字号", "目标值": f'{options["font_size"]:g} 磅'})
    if "color" in options:
        rows.append({"修改项": "字体颜色", "目标值": options["color"].upper()})
    if "line_spacing" in options:
        item = options["line_spacing"]
        rows.append({"修改项": "行距", "目标值": f'固定 {item["value"]:g} 磅' if item["mode"] == "exact" else f'{item["value"]:g} 倍'})
    st.table(rows)
    st.caption(f'变更了 {summary["changed_runs"]} 个文字片段、{summary["changed_paragraphs"]} 个段落行距。')
    st.download_button("下载修改后的 DOCX" if changed else "下载原 DOCX", data, output_filename(uploaded.name) if changed else uploaded.name, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", on_click="ignore")
    st.caption("可在 Word / WPS 中打开查看排版。本页面展示修改摘要，不提供分页预览。")

with st.expander("支持范围"):
    st.write("支持普通 DOCX 主体段落、表格（含合并单元格和嵌套表格）、超链接文字及原图保留。每次下载另存为新文件。")
    st.write("暂不处理带宏、加密、数字签名、修订批注、正文域、文本框、内容控件、公式和嵌入对象的文档。")
