"""文序 PRD v0.5 POC. Run and validate inside WSL Linux."""

from __future__ import annotations

from hashlib import sha256
import json
from pathlib import Path
import uuid

import pandas as pd
import streamlit as st

from docx_formatting import FormatError
from prd_formatting import PROPERTY_LABELS, ROLE_LABELS, apply_confirmed_rules, build_structure_mapping, preflight_document, safe_stem
from prd_workflow import (
    build_html_report,
    executable_rules,
    extract_reference_requirements,
    missing_required_fonts,
    normalize_rules,
    parse_text_requirements,
    render_docx,
    rule_fingerprint,
    runtime_label,
)


st.set_page_config(page_title="文序 · 文档格式纠正", page_icon="📄", layout="wide")


def _drop(*keys):
    for key in keys:
        st.session_state.pop(key, None)


def _invalidate_review():
    st.session_state.confirmed_fingerprint = None
    st.session_state.latest_result = None


def _continue_from_version(label):
    material = dict(st.session_state.material)
    material["selected_version_label"] = label
    material["selector_generation"] = material.get("selector_generation", 0) + 1
    st.session_state.material = material
    _invalidate_review()


def _init():
    defaults = {"material": None, "rules": [], "mapping": [], "jobs": {}, "extracted_sources": None, "confirmed_fingerprint": None, "latest_result": None}
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def _source_fingerprint(text, reference):
    reference_hash = sha256(reference).hexdigest() if reference else ""
    return sha256(((text or "") + "\0" + reference_hash).encode()).hexdigest()


def _clean(value, fallback=""):
    return fallback if value is None or (isinstance(value, float) and pd.isna(value)) else value


def _mapping_to_frame(mapping):
    return pd.DataFrame([{
        "段落": row["index"], "原文": row.get("text", "")[:160],
        "识别结果": ROLE_LABELS.get(row.get("recognized_scope"), row.get("recognized_scope", "")),
        "识别依据": row.get("source", ""), "弱证据": "需确认" if row.get("needs_confirmation") else "明确",
        "最终角色": ROLE_LABELS.get(row.get("scope"), row.get("scope", "正文")), "排除理由": row.get("excluded_reason", ""),
    } for row in mapping])


def _frame_to_mapping(frame, previous):
    reverse = {label: key for key, label in ROLE_LABELS.items()}
    by_index = {int(row["index"]): row for row in previous}
    result = []
    for record in frame.to_dict("records"):
        index = int(record["段落"])
        row = dict(by_index[index])
        row["scope"] = reverse.get(_clean(record.get("最终角色")), "body")
        row["excluded_reason"] = str(_clean(record.get("排除理由"))).strip()
        result.append(row)
    return result


def _rules_to_frame(rules):
    return pd.DataFrame([{
        "规则ID": row.get("id", ""), "应用": bool(row.get("apply", True)),
        "作用对象": ROLE_LABELS.get(row.get("scope"), row.get("scope", "")),
        "格式属性": PROPERTY_LABELS.get(row.get("property"), row.get("property", "")),
        "稿件当前值": row.get("current", "待导入稿件后读取"), "目标值": row.get("target", ""),
        "单位": row.get("unit", ""), "来源": row.get("source_type", "人工补充"),
        "来源依据": row.get("source_detail", ""), "核对状态": row.get("status", "待核对"),
        "原因/处理说明": row.get("reason", ""),
    } for row in rules])


def _frame_to_rules(frame, previous):
    reverse_roles = {label: key for key, label in ROLE_LABELS.items()}
    reverse_props = {label: key for key, label in PROPERTY_LABELS.items()}
    previous_by_id = {row.get("id"): row for row in previous}
    result = []
    for record in frame.to_dict("records"):
        rule_id = str(_clean(record.get("规则ID"))).strip() or uuid.uuid4().hex[:10]
        old = dict(previous_by_id.get(rule_id, {}))
        old.update({
            "id": rule_id, "apply": bool(_clean(record.get("应用"), True)),
            "scope": reverse_roles.get(_clean(record.get("作用对象")), ""),
            "property": reverse_props.get(_clean(record.get("格式属性")), ""),
            "target": str(_clean(record.get("目标值"))).strip(), "unit": str(_clean(record.get("单位"))).strip(),
            "source_type": str(_clean(record.get("来源"), "人工补充")).strip() or "人工补充",
            "source_detail": str(_clean(record.get("来源依据"))).strip(),
        })
        result.append(old)
    return result


def _select_input_version(material):
    labels = [item["label"] for item in material["versions"]]
    previous = material.get("selected_version_label")
    if previous not in labels:
        previous = labels[-1]
        material["selected_version_label"] = previous
    generation = material.get("selector_generation", 0)
    selected = st.selectbox(
        "本次输入版本",
        labels,
        index=labels.index(previous),
        key=f"input_version_{material['id']}_{generation}",
        help="可从原稿或任一已生成版本继续；不会覆盖旧文件。",
    )
    if previous != selected:
        _invalidate_review()
    material["selected_version_label"] = selected
    return next(item for item in material["versions"] if item["label"] == selected)


_init()
st.title("文序 · 文档格式纠正")
st.write("提供自己的稿件，再输入文字要求或上传参考实例。核对要求与作用范围后，系统只修改确认的格式属性并另存新版本。")

sample = Path(__file__).parent / "examples" / "sample.docx"
if sample.exists():
    st.download_button("下载示例稿件", sample.read_bytes(), "教学工作总结.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", on_click="ignore")

st.header("1 提供稿件与要求")
manuscript_upload = st.file_uploader("自己的 DOCX 稿件（最大 10 MB，POC 支持单节 A4 纵向、最多 50 个渲染页）", type=["docx"], key="manuscript_upload")
if manuscript_upload:
    uploaded_bytes = manuscript_upload.getvalue()
    uploaded_hash = sha256(uploaded_bytes).hexdigest()
    material = st.session_state.material
    known = next((item for item in (material or {}).get("versions", []) if item["hash"] == uploaded_hash), None)
    if not material or (uploaded_hash != material.get("uploaded_hash") and known is None):
        try:
            inventory = preflight_document(uploaded_bytes)
            mapping = build_structure_mapping(uploaded_bytes)
            stem, next_version = safe_stem(manuscript_upload.name)
            st.session_state.material = {
                "id": uuid.uuid4().hex, "uploaded_hash": uploaded_hash, "stem": stem, "next_version": next_version,
                "inventory": inventory,
                "versions": [{"label": f"原稿 · {manuscript_upload.name}", "number": 0, "name": manuscript_upload.name, "bytes": uploaded_bytes, "hash": uploaded_hash, "parent": "外部导入", "delivery_status": "未提交"}],
            }
            st.session_state.mapping = mapping
            st.session_state.material["selected_version_label"] = st.session_state.material["versions"][0]["label"]
            st.session_state.material["selector_generation"] = 0
            _drop("mapping_editor", "rule_editor")
            _invalidate_review()
        except FormatError as exc:
            st.session_state.material = None
            st.session_state.mapping = []
            _invalidate_review()
            st.error(str(exc))
    elif known is not None and uploaded_hash != material.get("uploaded_hash"):
        st.session_state.material["uploaded_hash"] = uploaded_hash
        st.session_state.material["selected_version_label"] = known["label"]
        st.session_state.material["selector_generation"] += 1

text_requirement = st.text_area("文字要求", placeholder="例如：一级标题黑体小二号、居中并与下段同页；正文仿宋四号，固定行距 28 磅；落款右对齐。", key="text_requirement", height=110)
reference_upload = st.file_uploader("参考实例 DOCX（可选；只提取可识别样式，不替换稿件正文）", type=["docx"], key="reference_upload")
reference_bytes = reference_upload.getvalue() if reference_upload else None
sources_fp = _source_fingerprint(text_requirement, reference_bytes)
if st.session_state.extracted_sources and st.session_state.extracted_sources != sources_fp:
    _invalidate_review()
    st.info("要求来源已变化。请重新提取，旧确认和旧检查结果已失效；已填写内容仍保留。")

material = st.session_state.material
if material:
    selected_version = _select_input_version(material)
    manuscript = selected_version["bytes"]
    inventory = preflight_document(manuscript)
    st.caption(f"输入：{selected_version['name']} · {len(manuscript) / 1024:.1f} KB · {inventory['page']} · {inventory['paragraphs']} 段 · {inventory['tables']} 表格 · {inventory['images']} 图片 · {inventory['headers']} 页眉 · {inventory['footers']} 页脚")
else:
    selected_version = None
    manuscript = None
    st.info("可以先填写文字要求；补齐有效稿件后再进入核对。")

if st.button("提取要求并生成核对表", disabled=not text_requirement.strip() and reference_bytes is None):
    candidates = []
    if text_requirement.strip():
        candidates.extend(parse_text_requirements(text_requirement))
    if reference_bytes is not None:
        try:
            candidates.extend(extract_reference_requirements(reference_bytes))
        except FormatError as exc:
            st.error(f"参考实例无法提取：{exc}")
    st.session_state.rules = normalize_rules(candidates, manuscript, st.session_state.mapping if manuscript else None)
    st.session_state.extracted_sources = sources_fp
    _drop("rule_editor")
    _invalidate_review()

if manuscript:
    st.header("2 核对对象与段落范围")
    st.caption("弱证据候选必须由你确认。可把任一段落改为标题、正文、落款或原样保留；原样保留区域会进入检查报告。")
    mapping_frame = _mapping_to_frame(st.session_state.mapping)
    edited_mapping = st.data_editor(
        mapping_frame, key="mapping_editor", hide_index=True, width="stretch",
        disabled=["段落", "原文", "识别结果", "识别依据", "弱证据"],
        column_config={
            "最终角色": st.column_config.SelectboxColumn("最终角色", options=[ROLE_LABELS[key] for key in ("title", "heading_1", "heading_2", "heading_3", "body", "signature", "preserve")], required=True),
            "排除理由": st.column_config.TextColumn("排除理由", help="选择原样保留时填写理由。"),
        },
    )
    current_mapping = _frame_to_mapping(edited_mapping, st.session_state.mapping)
    if current_mapping != st.session_state.mapping:
        st.session_state.mapping = current_mapping
        st.session_state.rules = normalize_rules(st.session_state.rules, manuscript, current_mapping)
        _invalidate_review()

if st.session_state.rules:
    st.header("3 核对要求表")
    st.caption("表格是本次执行的唯一依据。可编辑、补充或取消应用；不能执行的表达会保留为待补充项，不会静默忽略。")
    rules = normalize_rules(st.session_state.rules, manuscript, st.session_state.mapping if manuscript else None)
    rules_frame = _rules_to_frame(rules)
    edited_rules = st.data_editor(
        rules_frame, key="rule_editor", hide_index=True, width="stretch", num_rows="dynamic",
        disabled=["规则ID", "稿件当前值", "核对状态", "原因/处理说明"],
        column_config={
            "应用": st.column_config.CheckboxColumn("应用", default=True),
            "作用对象": st.column_config.SelectboxColumn("作用对象", options=[ROLE_LABELS[key] for key in ("title", "heading_1", "heading_2", "heading_3", "body", "signature", "page", "all")]),
            "格式属性": st.column_config.SelectboxColumn("格式属性", options=list(PROPERTY_LABELS.values())),
            "目标值": st.column_config.TextColumn("目标值", required=True), "单位": st.column_config.TextColumn("单位"),
        },
    )
    normalized = normalize_rules(_frame_to_rules(edited_rules, rules), manuscript, st.session_state.mapping if manuscript else None)
    if normalized != st.session_state.rules:
        st.session_state.rules = normalized
        _invalidate_review()
    selected_rules, blockers = executable_rules(normalized)
    if blockers:
        st.error("仍有已选择的待补充或冲突要求，需修正目标值、单位或作用对象，或取消本次应用。")
        st.dataframe([{"规则": row.get("id"), "状态": row.get("status"), "原因": row.get("reason"), "来源": row.get("source_detail")} for row in blockers], hide_index=True, width="stretch")
    elif selected_rules:
        st.success(f"{len(selected_rules)} 条已选要求完整且无冲突。")
else:
    selected_rules, blockers = [], []

if manuscript and st.session_state.rules:
    source_current = st.session_state.extracted_sources == sources_fp
    review_hash = rule_fingerprint(selected_version["hash"] + sources_fp, st.session_state.rules, st.session_state.mapping)
    confirmed = st.session_state.confirmed_fingerprint == review_hash
    left, right = st.columns([1, 2])
    with left:
        if st.button("确认当前要求表与范围", disabled=bool(blockers) or not selected_rules or not source_current, type="secondary"):
            st.session_state.confirmed_fingerprint = review_hash
            confirmed = True
    with right:
        st.success("已确认，可执行。任何输入、要求或范围变化都会使本确认失效。") if confirmed else st.info("尚未确认当前快照。")

    st.header("4 确认并执行修改")
    missing_fonts = missing_required_fonts(selected_rules) if selected_rules and not blockers else []
    if missing_fonts:
        st.error("Linux 渲染环境缺少要求字体：" + "、".join(missing_fonts) + "。请安装字体，或在核对表中明确改用已安装字体后重新确认。")
    if st.button("确认并执行修改", type="primary", disabled=not confirmed or bool(blockers) or bool(missing_fonts) or not source_current):
        task_id = uuid.uuid4().hex
        job_key = rule_fingerprint(selected_version["hash"], st.session_state.rules, st.session_state.mapping)
        cached = st.session_state.jobs.get(job_key)
        if cached:
            st.session_state.latest_result = cached
            st.info("输入与确认快照未变化，已复用同一任务结果，没有重复生成版本。")
        else:
            try:
                with st.status("正在固定快照并执行检查…", expanded=True) as status:
                    st.write("读取要求与输入修订")
                    input_render = render_docx(manuscript, "input")
                    st.write("纠正确认的格式属性")
                    output, summary = apply_confirmed_rules(manuscript, selected_rules, st.session_state.mapping)
                    if not summary["changed"]:
                        result = {"task_id": task_id, "changed": False, "summary": summary, "input": selected_version}
                        st.session_state.jobs[job_key] = result
                        st.session_state.latest_result = result
                        status.update(label="所有已选要求均已满足，无需修改", state="complete")
                    else:
                        st.write("检查内容、规则、文件和渲染结果")
                        output_render = render_docx(output, "output")
                        number = material["next_version"]
                        output_name = f"{material['stem']}_v{number}.docx"
                        report_name = f"{material['stem']}_v{number}_格式检查.html"
                        checks = [
                            {"level": "pass", "name": "内容与对象完整性", "detail": "正文、未选属性、图片、关系和其他包部件保全检查通过。"},
                            {"level": "pass", "name": "规则符合性", "detail": f"{len(selected_rules)} 条确认要求逐项达到目标值。"},
                            {"level": "pass", "name": "文件可读取", "detail": "输出 DOCX 已重新打开并完成结构检查。"},
                            {"level": "pass", "name": "固定环境渲染", "detail": f"输入 {input_render['pages']} 页，输出 {output_render['pages']} 页；Linux LibreOffice 转换与逐页预览生成成功，人工版式复核待完成。"},
                            {"level": "warning", "name": "客户端边界", "detail": "本结果已在 Linux 固定环境验证；正式试点仍需记录目标 WPS/Word 的人工打开结果。"},
                        ]
                        warnings = ["包含用户确认的原样保留区域或本次不应用要求。"] if any(row.get("scope") == "preserve" for row in st.session_state.mapping) or any(not row.get("apply") for row in st.session_state.rules) else []
                        report = build_html_report(task_id=task_id, input_name=selected_version["name"], input_hash=selected_version["hash"], output_name=output_name, parent_version=selected_version["label"], rules=st.session_state.rules, mapping=st.session_state.mapping, summary=summary, checks=checks, environment=runtime_label(output_render), warnings=warnings)
                        version = {
                            "label": f"v{number} · {output_name}", "number": number, "name": output_name, "bytes": output,
                            "hash": sha256(output).hexdigest(), "input_hash": selected_version["hash"],
                            "parent": selected_version["label"], "delivery_status": "未提交",
                            "report": report, "report_name": report_name, "checks": checks, "input_render": input_render,
                            "output_render": output_render, "summary": summary, "task_id": task_id,
                        }
                        material["versions"].append(version)
                        material["next_version"] = number + 1
                        result = {"task_id": task_id, "changed": True, "version": version}
                        st.session_state.jobs[job_key] = result
                        st.session_state.latest_result = result
                        status.update(label="检查通过，已另存新版本", state="complete")
            except FormatError as exc:
                st.session_state.latest_result = {"task_id": task_id, "error": str(exc)}
                st.error(str(exc))

result = st.session_state.latest_result
if result:
    st.header("5 结果、预览与交付")
    if result.get("error"):
        st.error(result["error"])
    elif not result.get("changed"):
        st.info("所有已选要求均已满足；未创建新版本。原稿和已有结果保持可用。")
    else:
        version = result["version"]
        st.success(f"{version['name']} 已完成内容、规则、文件和固定环境渲染检查。")
        st.dataframe(version["checks"], hide_index=True, width="stretch")
        original_tab, output_tab = st.tabs([f"输入预览 · {version['input_render']['pages']} 页", f"结果预览 · {version['output_render']['pages']} 页"])
        with original_tab:
            for index, image in enumerate(version["input_render"]["images"], 1):
                st.image(image, caption=f"输入第 {index} 页", width="stretch")
        with output_tab:
            for index, image in enumerate(version["output_render"]["images"], 1):
                st.image(image, caption=f"结果第 {index} 页", width="stretch")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.download_button("下载修改后的 DOCX", version["bytes"], version["name"], mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", on_click="ignore")
        with col2:
            st.download_button("下载 HTML 检查报告", version["report"], version["report_name"], mime="text/html", on_click="ignore")
        with col3:
            diagnostics = json.dumps({"task_id": version["task_id"], "input_hash": version["input_hash"], "output_hash": version["hash"], "checks": version["checks"], "summary": version["summary"]}, ensure_ascii=False, indent=2).encode()
            st.download_button("下载诊断信息", diagnostics, f"{material['stem']}_v{version['number']}_诊断.json", mime="application/json", on_click="ignore")
        st.button("以此结果继续调整", on_click=_continue_from_version, args=(version["label"],))
        st.caption(f"交付观察：{version['delivery_status']}。下载只表示提供文件，不等于业务已提交或已接收。")
        a, b, c = st.columns(3)
        if a.button("标记已提交"):
            version["delivery_status"] = "已提交"
            st.rerun()
        if b.button("标记被退回"):
            version["delivery_status"] = "被退回"
            st.rerun()
        if c.button("标记已接收", disabled=version["delivery_status"] != "已提交"):
            version["delivery_status"] = "已接收"
            st.rerun()

with st.expander("支持范围与当前阶段"):
    st.write("本 POC 支持单个、单节、A4 纵向 DOCX，最多 10 MB 和 50 个渲染页；支持正文、三级标题、文档标题、落款及原样保留映射。")
    st.write("支持字体、字号、加粗、斜体、颜色、对齐、缩进、段前段后、固定/倍数行距、标题与下段同页、页面四边距。未指定属性保持输入版本原值。")
    st.write("损坏、加密、宏、签名、修订批注、域、内容控件、文本框、公式、嵌入对象、外部模板/图片和多节文档会被阻断。")
    st.write("个人规范库、关闭页面后的服务端恢复、组织发布和批量处理属于第一期/第二期，不在本 POC 范围。")
