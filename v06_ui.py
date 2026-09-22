"""Streamlit user interface for the shared v0.6 editing protocol."""
from __future__ import annotations

import json

import streamlit as st

from docx_formatting import FormatError
from prd_formatting import PAGE_PROPERTIES, PROPERTY_LABELS
from prd_workflow import CHINESE_SIZES, FONT_NAMES, extract_reference_requirements, normalize_rules
from v06_model import ROLE_LABELS, build_model, digest
from v06_patch import current_values
from v06_plan import build_plan, encode, intent, parse_text, resolve_scope
from v06_session import EditingSession, report_html
from v06_import import IMPORT_VERSION

LABELS = {**PROPERTY_LABELS, "font_family": "字体（中文与西文）", "font_east_asia": "字体（仅中文）", "font_western": "字体（仅西文）", "underline": "下划线"}


def _reset_draft(session):
    st.session_state.draft = []
    st.session_state.roles = {int(k): v for k, v in json.loads(session.versions[session.current].roles_json).items()}
    st.session_state.protections = json.loads(session.versions[session.current].protections_json)
    st.session_state.epoch = st.session_state.get("epoch", 0) + 1
    st.session_state.latest = None
    st.session_state.pop("reference_candidates", None)


def _format_controls(prefix, page=False, initial=None, font_families=()):
    initial = initial or {}
    allowed = list(PAGE_PROPERTIES) if page else [p for p in LABELS if p not in PAGE_PROPERTIES]
    props = st.multiselect("本次设置的属性（其他保持原样）", allowed, default=list(initial), format_func=lambda p: LABELS[p], key=prefix+"props")
    values = {}
    for prop in props:
        label, key = LABELS[prop], prefix+prop
        if prop.startswith("font_") and prop != "font_size":
            installed = {f.casefold() for f in font_families}
            names = [f for f in FONT_NAMES if not installed or f.casefold() in installed]
            default = names[0] if names else next(iter(font_families), 'Noto Serif CJK SC')
            if default not in names:
                names.append(default)
            value = initial.get(prop, default)
            names.append("自定义字体")
            font = st.selectbox(label, names, index=names.index(value) if value in names else len(names)-1, key=key)
            values[prop] = st.text_input("字体名称", value=value if value not in names else "", key=key+"custom") if font == "自定义字体" else font
        elif prop == "font_size":
            st.caption("五号 10.5 磅 · 小四 12 磅 · 四号 14 磅；可输入其他字号")
            values[prop] = st.number_input("字号（磅）", 6.0, 72.0, float(initial.get(prop, 12)), 0.5, key=key)
        elif prop in {"bold", "italic", "underline", "keep_with_next"}:
            mode = st.selectbox(label, ["保持原样", "开启", "关闭"], index=(1 if initial[prop] else 2) if prop in initial else 0, key=key)
            if mode != "保持原样":
                values[prop] = mode == "开启"
        elif prop == "color":
            values[prop] = st.color_picker("文字颜色", initial.get(prop, "#FF0000"), key=key)
        elif prop == "alignment":
            choices = {"left": "左对齐", "center": "居中", "right": "右对齐", "justify": "两端对齐"}
            values[prop] = st.selectbox(label, list(choices), index=list(choices).index(initial.get(prop, "left")), format_func=choices.get, key=key)
        elif prop == "line_spacing":
            old = initial.get(prop, {"mode": "exact", "value": 28})
            mode = st.selectbox("行距方式", ["exact", "multiple"], index=0 if old["mode"] == "exact" else 1, format_func=lambda m: "固定磅值" if m == "exact" else "倍数", key=key+"mode")
            low, high = (6.0, 144.0) if mode == "exact" else (0.5, 5.0)
            values[prop] = {"mode": mode, "value": st.number_input("行距值", low, high, float(old["value"]) if mode == old["mode"] else low, 0.5, key=key+mode)}
        else:
            values[prop] = st.number_input(label+"（磅）", 0.0, 720.0, float(initial.get(prop, 0)), 0.5, key=key)
    return values


def _paragraph_label(model, pid):
    p = model.row(pid)
    return f'{p["location"]} · {p["text"][:65] or "（空段）"}'


def _choose_scope(model, prefix):
    mode = st.radio("调整方式", ["全局调整", "局部调整", "页面设置"], horizontal=True, key=prefix+"mode")
    if mode == "页面设置":
        st.caption("页边距作用于单节稿件的页面；与局部字符属性独立。")
        return model.scope("page")
    if mode == "全局调整":
        kinds = {"body": "正文", "title": "文档标题", "heading_1": "一级标题", "heading_2": "二级标题", "heading_3": "三级标题", "table": "表格文字", "all": "全文文字（正文区和表格）"}
        kind = st.selectbox("修改范围", list(kinds), format_func=kinds.get, key=prefix+"global")
        if kind == "all":
            st.caption("页眉页脚、公式和嵌入对象保持原样；字段文字受保护。")
            return model.scope("all")
        if kind == "table":
            tables = sorted({c["table"] for c in model.cells})
            if not tables:
                return None
            return model.scope("table", table=st.selectbox("表格", ["all", *tables], format_func=lambda t: "全部表格" if t == "all" else f"表 {t}", key=prefix+"table"))
        include = st.checkbox("包含表格内正文", value=True, key=prefix+"include") if kind == "body" else True
        return model.scope("role", role=kind, include_tables=include)
    local = st.selectbox("定位方式", ["搜索文字", "段落与字符", "栏目填写区", "表格单元格", "表格行", "表格列"], key=prefix+"local")
    if local == "搜索文字":
        query = st.text_input("搜索原文", key=prefix+"search")
        matches = model.search(query)
        st.caption(f"找到 {len(matches)} 处；请明确选择一个或多个命中。")
        selected = st.multiselect("选择命中位置", range(len(matches)), format_func=lambda i: f"第 {i+1} 处 · {matches[i]['context']}", key=prefix+"hits")
        if st.checkbox("全部匹配", value=False, key=prefix+"allhits"):
            selected = list(range(len(matches)))
        return model.scope("selection", spans=[span for i in selected for span in matches[i]["spans"]]) if selected else None
    if local == "段落与字符":
        pid = st.selectbox("选择段落", [p["pid"] for p in model.paragraphs], format_func=lambda p: _paragraph_label(model, p), key=prefix+"pid")
        text = model.row(pid)["text"]
        st.text(text or "（空段）")
        if st.checkbox("仅选择其中的文字", key=prefix+"chars") and text:
            start, end = st.slider("字符范围（从 0 开始，右端不包含）", 0, len(text), (0, len(text)), key=prefix+f"span{pid}")
            st.write("选中：", text[start:end])
            return model.scope("selection", spans=[{"pid": pid, "start": start, "end": end}]) if start < end else None
        return model.scope("paragraphs", pids=[pid])
    if local == "栏目填写区":
        if not model.columns:
            st.info("未发现可靠栏目候选，请使用单元格或段落定位。")
            return None
        chosen = st.selectbox("选择具体栏目（默认仅填写内容）", model.columns, format_func=lambda c: f"表 {c['table']} · {c['label']} · {c['id']}", key=prefix+"column")
        st.caption(chosen["evidence"]+"；请核对下方实际段落，栏目名称不修改。")
        return model.scope("named_column", column_id=chosen["id"])
    if local == "表格单元格":
        cells = st.multiselect("选择单元格", model.cells, format_func=lambda c: f"表 {c['table']} 行 {c['row']}–{c['end_row']} 列 {c['col']}–{c['end_col']} · {c['text'][:30]}", key=prefix+"cells")
        return model.scope("cell", cell_ids=[c["id"] for c in cells]) if cells else None
    tables = sorted({c["table"] for c in model.cells})
    if not tables:
        return None
    table = st.selectbox("表格", tables, key=prefix+"localtable")
    column = local == "表格列"
    dimension = "column" if column else "row"
    maximum = max(c["end_col" if column else "end_row"] for c in model.cells if c["table"] == table)
    value = st.number_input("逻辑列号" if column else "逻辑行号", 1, maximum, 1, key=prefix+dimension)
    st.caption("跨所选行/列的合并单元格作为完整单元格纳入，下方显示实际范围。")
    return model.scope(dimension, table=table, **{dimension: int(value)})


def _preview_pages(render, label, key):
    st.caption(f"{label} · {render['pages']} 页")
    page = st.number_input(label+"页码", 1, render["pages"], 1, key=key)
    st.image(render["images"][page-1], width="stretch")


def main():
    st.set_page_config(page_title="文序 · 精确格式调整", page_icon="📄", layout="wide")
    st.title("文序 · 精确格式调整")
    st.caption("调整已有稿件的指定格式，保留文字和其他对象。文件仅在当前本地会话处理。")
    uploaded = st.file_uploader("上传自己的稿件（DOCX / DOC，最多 10 MB）", type=["docx", "doc"])
    if not uploaded:
        st.info("上传后，可直接统一正文格式，也可搜索文字、选择段落或表格填写区。")
        st.caption("v0.6 · 当前支持范围和运行说明见仓库 README；旧版参考工作流可通过 ?legacy=1 打开。")
        return
    identity = IMPORT_VERSION + digest(uploaded.getvalue()) + uploaded.name
    if st.session_state.get("upload_identity") != identity:
        st.session_state.pop("editor", None)
        try:
            with st.spinner("读取稿件并检查文档对象…"):
                session = EditingSession.load(uploaded.name, uploaded.getvalue())
            st.session_state.editor = session
            st.session_state.upload_identity = identity
            _reset_draft(session)
        except (FormatError, ValueError, OSError) as exc:
            st.error(str(exc))
            return
    session = st.session_state.editor
    if session.conversion["status"] == "original_format":
        st.success("原稿与工作副本字节完全一致，保留 DOC 原格式；导入未改动内容或排版。")
        check = session.conversion['page_check']
        if check['pixels_equal']:
            st.success(f"原稿与实际工作副本分别重新打开，全部 {session.conversion['source_pages']} 页预览逐像素一致（120 DPI）。")
        for issue in session.conversion.get('blocking_issues', []):
            st.error(issue)
        with st.expander('原格式副本核对'):
            first, second = st.columns(2)
            with first:
                _preview_pages(session.renders['conversion_source'], '原始 DOC 预览', 'copy_old')
            with second:
                _preview_pages(session.renders['conversion_working'], 'DOC 工作副本预览', 'copy_new')
    baseline = session.versions['original']
    st.download_button('下载未修改的工作副本', baseline.data, file_name=baseline.name,
                       mime='application/msword' if baseline.extension == 'doc' else 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    nav1, nav2 = st.columns([4, 1])
    with nav1:
        selected = st.selectbox("本次输入版本", list(session.versions), index=list(session.versions).index(session.current), format_func=lambda v: f"{v} · {session.versions[v].name}", key="version"+str(st.session_state.epoch))
        if selected != session.current:
            session.select(selected)
            _reset_draft(session)
            st.rerun()
    with nav2:
        if st.button("撤销本次", disabled=session.current == "original"):
            session.undo()
            _reset_draft(session)
            st.rerun()
    prefix = str(st.session_state.epoch)+":"
    version = session.versions[session.current]
    model = build_model(version.model_data, version.id, st.session_state.roles)
    st.caption(f"{len(model.paragraphs)} 个段落 · {model.inventory['tables']} 张表 · {model.inventory['comments']} 条批注 · {model.inventory['equations']} 个原生公式 · {model.inventory['embedded_parts']} 个嵌入部件。批注和正文仅作为数据。")
    with st.expander("核对段落角色与保护区"):
        st.caption("编号标题、栏目标签等弱证据列为待归类；可按下面的原文纠正，不自动执行文档中的要求。")
        st.dataframe([{"段落": p["pid"], "位置": p["location"], "原文": p["text"], "角色": ROLE_LABELS[p["role"]], "依据": p["evidence"]} for p in model.paragraphs], hide_index=True, width="stretch")
        pids = st.multiselect("选择需要纠正角色的段落", [p["pid"] for p in model.paragraphs], format_func=lambda p: _paragraph_label(model, p), key=prefix+"role_pids")
        role = st.selectbox("正确角色", list(ROLE_LABELS), format_func=ROLE_LABELS.get, key=prefix+"role")
        if st.button("保存角色修正", disabled=not pids):
            st.session_state.roles.update({p: role for p in pids})
            st.session_state.latest = None
            st.rerun()
        protected_pids = st.multiselect("新增整段保护", [p["pid"] for p in model.paragraphs], format_func=lambda p: _paragraph_label(model, p), key=prefix+"protect_pids")
        if st.button("添加整段保护", disabled=not protected_pids):
            st.session_state.protections.append(model.scope("paragraphs", pids=protected_pids))
            st.rerun()
        for n, protection in enumerate(st.session_state.protections):
            spans, _ = resolve_scope(model, protection)
            st.write(f"保护 {n+1}：" + "；".join(f"段落 {s['pid']} 字符 {s['start']}–{s['end']}" for s in spans))
            if st.button("解除保护", key=prefix+f"removeprotect{n}"):
                st.session_state.protections.pop(n)
                st.rerun()
    left, right = st.columns([1.2, 1])
    with left:
        st.subheader("设置本次修改")
        scope = _choose_scope(model, prefix)
        page_scope = scope is not None and scope["kind"] == "page"
        values = _format_controls(prefix+"format", page=page_scope, font_families=session.environment.get('font_families', ()))
        if scope:
            try:
                spans, warnings = resolve_scope(model, scope)
                st.write(f"当前命中 {len(set(s['pid'] for s in spans))} 个段落" if not page_scope else "当前命中单节页面设置")
                for warning in warnings:
                    st.warning(warning)
                with st.expander("查看命中位置和当前格式"):
                    st.dataframe([{"位置": model.row(s["pid"])["location"], "选中内容": model.row(s["pid"])["text"][s["start"]:s["end"]]} for s in spans], hide_index=True)
                    for prop in values:
                        actual = current_values(model, scope, "font_east_asia" if prop == "font_family" else prop)
                        st.write(f"{LABELS[prop]}：", "多种格式" if len(actual) > 1 else str(actual[0]) if actual else "无可修改文字")
                omit = st.multiselect("本次移除命中段落", sorted({s["pid"] for s in spans}), format_func=lambda p: _paragraph_label(model, p), key=prefix+"omit")
                if omit:
                    scope = model.scope("selection", spans=[s for s in spans if s["pid"] not in omit and s["start"] < s["end"]])
            except FormatError as exc:
                st.warning(str(exc))
        local = scope and scope["kind"] not in {"role", "all", "table", "page"}
        exception = st.checkbox("此项作为同批全局设置的局部例外", value=bool(local), disabled=not local, key=prefix+"exception"+str(bool(local)))
        if st.button("加入待应用清单", disabled=not scope or not values):
            st.session_state.draft.append(intent(scope, values, exception=exception))
            st.session_state.latest = None
            st.rerun()
        if local and st.button("将当前选区设为保护区"):
            st.session_state.protections.append(scope)
            st.rerun()
        text = st.text_area("一句话修改", placeholder="正文宋体小四，但第三段黑体；正文固定行距28磅", key=prefix+"text")
        if st.button("解析并加入清单", disabled=not text.strip()):
            st.session_state.draft.extend(parse_text(model, text))
            st.session_state.latest = None
            st.rerun()
        with st.expander("采用参考实例要求（可选）"):
            st.caption("仅显式采用的格式进入本次清单；不会复制参考正文。复杂参考稿可使用旧版支持范围内的副本。")
            reference = st.file_uploader("参考实例 DOCX", type=["docx"], key=prefix+"reference")
            if reference and st.button("提取参考候选"):
                try:
                    st.session_state.reference_candidates = normalize_rules(extract_reference_requirements(reference.getvalue()))
                except FormatError as exc:
                    st.error(str(exc))
            candidates = st.session_state.get("reference_candidates", [])
            choices = st.multiselect("明确采用的候选项", range(len(candidates)), format_func=lambda i: f"{candidates[i]['scope']} · {LABELS.get(candidates[i]['property'])} · {candidates[i]['target']}", key=prefix+"referencechoices")
            if st.button("将选中参考要求加入清单", disabled=not choices):
                for i in choices:
                    r = candidates[i]
                    target_scope = model.scope("page") if r["scope"] == "page" else model.scope("role", role=r["scope"], include_tables=True)
                    item = intent(target_scope, {r["property"]: r.get("normalized_target")}, source="用户采用的参考格式")
                    if r.get("status") not in {"待核对", "可执行", "已满足", "待修改"} and r.get("normalized_target") is None:
                        item["unparsed"] = r.get("reason", "参考值未确定")
                    st.session_state.draft.append(item)
                st.rerun()
    with right:
        st.subheader("待应用清单")
        for i, item in enumerate(st.session_state.draft):
            with st.expander(f"第 {i+1} 项 · {item['source']}", expanded=bool(item.get("unparsed"))):
                active = st.checkbox("本次应用", value=item.get("active", True), key=prefix+f"active{i}")
                item["active"] = active
                if item.get("unparsed"):
                    st.warning("未理解："+item["unparsed"]+"。请取消此项并用控件重新指定，或重新输入。")
                else:
                    st.caption("来源范围："+item["scope"]["kind"]+(" · 明确局部例外" if item.get("exception") else ""))
                    st.write({LABELS.get(k, k): v for k, v in item["set"].items()})
                    with st.container(border=True):
                        st.caption("修改本项目标值")
                        replacement = _format_controls(prefix+f"edit{i}", page=item["scope"]["kind"] == "page", initial=item["set"], font_families=session.environment.get('font_families', ()))
                        if st.button("保存本项修改", key=prefix+f"save{i}"):
                            item["set"] = replacement
                            st.session_state.latest = None
                            st.rerun()
        review = build_plan(model, st.session_state.draft, st.session_state.protections)
        if st.session_state.draft:
            for message in review.blockers:
                st.error(message)
        else:
            st.info("从控件、文字要求或参考候选添加本次修改。")
        for message in review.notices:
            st.warning(message)
        if review.excluded:
            with st.expander("本次排除与取消记录"):
                st.dataframe(review.excluded, hide_index=True)
        if review.preview:
            st.dataframe([{"段落": a["pid"], "字符区间": f"{a['start']}–{a['end']}" if a["end"] else "整段/页面", "属性": LABELS.get(a["property"], a["property"]), "目标": str(a["value"]), "来源项": str(a["sources"])} for a in review.preview], hide_index=True, width="stretch")
        if st.button("应用修改并检查", type="primary", disabled=not review.plan or session.conversion["visual_review"] == "pending" or bool(session.conversion.get("blocking_issues"))):
            try:
                with st.spinner("修改副本，校验内容和对象，生成前后预览…"):
                    result, created = session.execute(model, review.plan)
                if result is None:
                    st.info("全部目标已满足，无需修改；未创建新版本。")
                else:
                    session.select(result.id)
                    _reset_draft(session)
                    st.session_state.latest = result.id
                    st.rerun()
            except (FormatError, OSError, ValueError) as exc:
                st.error(str(exc))
        if st.button("清空待应用清单", disabled=not st.session_state.draft):
            st.session_state.draft = []
            st.session_state.latest = None
            st.session_state.epoch += 1
            st.rerun()
    st.subheader("文档预览与交付")
    if session.current != "original":
        result = session.versions[session.current]
        report = json.loads(result.report_json)
        st.success(f"{result.name} 已完成结构、目标属性、对象保全和渲染检查。")
        review_key = "reviewed:" + digest(result.data) + ":" + result.id
        if st.session_state.get(review_key):
            st.info("人工版式复核已确认；Word/WPS 客户端兼容性尚未实测。")
        else:
            st.warning("人工版式复核待完成；渲染成功不代表 Word/WPS 客户端兼容性已通过。")
        for warning in report.get("font_warnings", []):
            st.warning(warning)
        before, after = st.columns(2)
        with before:
            _preview_pages(session.preview(result.parent), "修改前", prefix+"before")
        with after:
            _preview_pages(session.preview(result.id), "修改后", prefix+"after")
        st.caption(f"页数 {report['pages_before']} → {report['pages_after']}；自然换行或分页变化不通过额外缩字消除。")
        reviewed = st.checkbox("已查看修改前后页面，接受当前版式并下载", key=review_key)
        st.download_button("下载修改后的 " + result.extension.upper(), result.data, file_name=result.name, mime="application/msword" if result.extension == "doc" else "application/vnd.openxmlformats-officedocument.wordprocessingml.document", disabled=not reviewed)
        st.download_button("下载 HTML 检查报告", report_html(result, visual_review="confirmed" if reviewed else "pending"), file_name=result.name.rsplit(".", 1)[0]+"_检查报告.html", mime="text/html")
        st.caption("下一次修改已以当前结果为输入，控件恢复保持原样；可从版本列表返回原稿或任意历史版本。")
    elif st.button("生成当前稿件页面预览"):
        try:
            with st.spinner("生成页面预览…"):
                st.session_state.original_preview = session.preview()
            st.session_state.preview_identity = identity
        except FormatError as exc:
            st.error(str(exc))
    if session.current == "original" and st.session_state.get("preview_identity") == identity:
        _preview_pages(st.session_state.original_preview, "当前稿件", prefix+"original")
    st.download_button("下载保留的原始上传文件", session.original, file_name=session.original_name)
