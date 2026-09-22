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
        roles = {p['role'] for p in model.paragraphs if p['text'].strip()}
        available = [kind for kind in kinds if kind in roles or kind == 'all'
                     or (kind == 'table' and any(c['text'].strip() for c in model.cells))]
        kind = st.selectbox("修改范围", available, format_func=kinds.get, key=prefix+"global")
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
    methods = ['搜索文字', '段落与字符']
    if model.columns:
        methods.append('栏目填写区')
    if model.cells:
        methods.extend(['表格单元格', '表格行', '表格列'])
    local = st.selectbox("定位方式", methods, key=prefix+"local")
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
        st.caption(chosen["evidence"]+"；仅修改填写内容，栏目名称保持原样。")
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
    st.caption("跨所选行/列的合并单元格作为完整单元格纳入。")
    return model.scope(dimension, table=table, **{dimension: int(value)})


def _preview_pages(render, label, key):
    st.caption(f"{label} · {render['pages']} 页")
    page = st.number_input(label+"页码", 1, render["pages"], 1, key=key)
    st.image(render["images"][page-1], width="stretch")


def _roles_and_protections(model, prefix, selected_scope):
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
        if selected_scope and selected_scope['kind'] not in {'role', 'all', 'table', 'page'}:
            if st.button('将当前选区设为保护区'):
                st.session_state.protections.append(selected_scope)
                st.rerun()
        for n, protection in enumerate(st.session_state.protections):
            spans, _ = resolve_scope(model, protection)
            st.write(f"保护 {n+1}：" + "；".join(f"段落 {s['pid']} 字符 {s['start']}–{s['end']}" for s in spans))
            if st.button("解除保护", key=prefix+f"removeprotect{n}"):
                st.session_state.protections.pop(n)
                st.rerun()


def _reference(model, prefix):
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


def _value_label(value):
    if value is None:
        return "跟随样式"
    if isinstance(value, bool):
        return "开启" if value else "关闭"
    if isinstance(value, dict):
        return f"{value['value']:g} " + ("磅" if value['mode'] == 'exact' else "倍")
    if isinstance(value, list):
        return " / ".join(_value_label(v) for v in value)
    if isinstance(value, str) and value.startswith('theme:'):
        return '跟随文档主题'
    return str(value)


def _summary(assignments):
    grouped = {}
    for item in assignments:
        for prop, value in item.get('set', {item.get('property'): item.get('value')}).items():
            if prop:
                grouped.setdefault(prop, [])
                text = _value_label(value)
                if text not in grouped[prop]:
                    grouped[prop].append(text)
    if grouped.get('font_east_asia') == grouped.get('font_western') and 'font_east_asia' in grouped:
        grouped['font_family'] = grouped.pop('font_east_asia')
        grouped.pop('font_western')
    labels = {**LABELS, 'font_family': '字体', 'font_east_asia': '中文字体', 'font_western': '西文字体'}
    return '；'.join(f"{labels.get(prop, prop)} → {'、'.join(values)}" for prop, values in grouped.items())


def _quick_controls(prefix, fonts):
    installed = {font.casefold() for font in fonts}
    common = [font for font in FONT_NAMES if font.casefold() in installed]
    if not common:
        common = list(fonts)[:8]
    values = {}
    first, second = st.columns(2)
    with first:
        font = st.selectbox('字体', ['保持原样', *common, '其他已安装字体'], key=prefix+'quick_font')
        if font == '其他已安装字体':
            font = st.selectbox('选择已安装字体', ['保持原样', *fonts], key=prefix+'other_font')
        if font != '保持原样':
            values['font_family'] = font
    with second:
        sizes = {'保持原样': None, '小四 · 12 磅': 12, '五号 · 10.5 磅': 10.5, '四号 · 14 磅': 14,
                 '三号 · 16 磅': 16, '小三 · 15 磅': 15, '小五 · 9 磅': 9, '自定义字号': 'custom'}
        selected = st.selectbox('字号', list(sizes), key=prefix+'quick_size')
        size = sizes[selected]
        if size == 'custom':
            size = st.number_input('字号（磅）', 6.0, 72.0, 12.0, 0.5, key=prefix+'quick_points')
        if size is not None:
            values['font_size'] = size
    with st.expander('更多格式：颜色、行距、段落等'):
        extra = _format_controls(prefix+'extra', font_families=fonts)
    # Script-specific settings are explicit overrides of the common font control.
    if 'font_family' in values and set(extra) & {'font_east_asia', 'font_western'}:
        font = values.pop('font_family')
        values.update(font_east_asia=font, font_western=font)
    values.update(extra)
    return values


def _go_to(label):
    st.session_state.pending_tab = label
    st.rerun()


def _execute(session, model, review):
    try:
        with st.spinner('正在应用修改并生成最新预览…'):
            result, _ = session.execute(model, review.plan)
        if result is None:
            st.info('所选范围已经是这些设置，文件没有变化。')
            return
        session.select(result.id)
        _reset_draft(session)
        _go_to('预览与下载')
    except (FormatError, OSError, ValueError) as exc:
        st.error(str(exc))


def _edit(session, model, prefix):
    st.subheader('设置本次修改')
    scope = _choose_scope(model, prefix)
    control_prefix = prefix + 'controls' + str(st.session_state.get('control_epoch', 0)) + ':'
    values = _format_controls(control_prefix, page=True) if scope and scope['kind'] == 'page' else _quick_controls(control_prefix, session.environment.get('font_families', []))
    local = bool(scope and scope['kind'] not in {'role', 'all', 'table', 'page'})
    spans = []
    warnings = []
    if scope:
        try:
            spans, warnings = resolve_scope(model, scope)
        except FormatError as exc:
            st.warning(str(exc)); scope = None
    # Widget changes enter session state before rerun, so the plan can use the
    # batch option rendered below without showing another advanced panel here.
    exception = local and st.session_state.get(prefix+'exception', True)
    if scope:
        count = len({s['pid'] for s in spans if model.row(s['pid'])['text'].strip()})
        current = ''
        if scope['kind'] != 'page':
            observed = current_values(model, scope, 'font_east_asia')
            current = ' · 当前中文字体：' + ('、'.join(_value_label(v) for v in observed[:3]) if observed else '无可修改文字')
            if len(observed) > 3:
                current += '等混合字体'
        st.caption(('作用于页面设置' if scope['kind'] == 'page' else f'已选中 {count} 段文字') + current)
    live = [intent(scope, values, exception=exception)] if scope and values else []
    combined = [*st.session_state.draft, *live]
    review = build_plan(model, combined, st.session_state.protections)
    if combined:
        st.write('本次将应用：' + _summary(review.preview))
        for message in review.blockers:
            st.error(message)
        if review.excluded:
            st.caption('部分内容已按保护设置排除，可在清单详情中查看。')
    else:
        st.caption('选择需要更改的字体或字号，然后直接应用。')
    for message in dict.fromkeys([*warnings, *review.notices]):
        if message.startswith('正文覆盖范围存在待归类段落'):
            st.caption('部分段落尚未归类，未纳入正文；可在“历史与高级”核对。')
        else:
            st.warning(message)
    blocked = bool(session.conversion.get('blocking_issues'))
    if st.button('应用并预览', type='primary', disabled=not review.plan or blocked, width='stretch'):
        _execute(session, model, review)
    # Alternate input methods and multi-rule editing are secondary entry points.
    with st.expander('用一句话设置'):
        text = st.text_area('一句话修改', placeholder='正文黑体小四；第三段红色', key=prefix+'text')
        if st.button('解析并加入清单', disabled=not text.strip()):
            st.session_state.draft.extend(parse_text(model, text)); st.rerun()
    label = f"组合多项修改（已加入 {len(st.session_state.draft)} 项）" if st.session_state.draft else '组合多项修改（可选）'
    with st.expander(label, expanded=bool(st.session_state.draft)):
        st.caption('需要一次修改多个范围时再使用清单；上方当前设置会与清单一起应用。')
        if local:
            st.checkbox('此项作为同批全局设置的局部例外', value=True, key=prefix+'exception')
        if st.button('加入待应用清单', disabled=not live):
            st.session_state.draft.extend(live)
            st.session_state.control_epoch = st.session_state.get('control_epoch', 0) + 1
            st.rerun()
        for i, item in enumerate(st.session_state.draft):
            with st.expander(f"第 {i+1} 项 · {_summary([{'set': item['set']}]) or item['source']}"):
                active = st.checkbox('本次应用', value=item.get('active', True), key=prefix+f'active{i}')
                if active != item.get('active', True):
                    item['active'] = active
                    st.rerun()
                if item.get('unparsed'):
                    st.warning('未理解：' + item['unparsed'] + '。请取消或重新输入。')
                else:
                    replacement = _format_controls(prefix+f'edit{i}', page=item['scope']['kind']=='page', initial=item['set'], font_families=session.environment.get('font_families', []))
                    if st.button('保存本项修改', key=prefix+f'save{i}'):
                        item['set'] = replacement; st.rerun()
        if st.button('清空待应用清单', disabled=not st.session_state.draft):
            st.session_state.draft = []; st.session_state.epoch += 1; st.rerun()
        with st.expander('清单详情'):
            if review.preview:
                st.dataframe([{'段落': a['pid'], '属性': LABELS.get(a['property'], a['property']), '目标': _value_label(a['value'])} for a in review.preview], hide_index=True)
            if review.excluded:
                st.dataframe(review.excluded, hide_index=True)
    return scope


def _download(version, label, disabled=False):
    st.download_button(label, version.data, file_name=version.name,
                       mime='application/msword' if version.extension == 'doc' else 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                       disabled=disabled, type='primary')


def _delivery(session, prefix, active):
    result = session.versions[session.current]
    if session.current == 'original':
        st.info('当前是未修改的原稿。应用修改后，这里会显示最新结果。')
        _download(result, '下载未修改的工作副本')
        if active:
            try:
                with st.spinner('生成原稿预览…'):
                    view = session.preview()
                _preview_pages(view, '原稿', prefix+'source_page')
            except FormatError as exc:
                st.error(str(exc))
        return
    report = json.loads(result.report_json)
    st.success('当前结果：' + result.name)
    st.write('已应用：' + _summary(report['operations']))
    changed = report.get('page_check', {}).get('differing_pages', [])
    if changed:
        st.caption('页面发生变化：' + '、'.join(str(p) for p in changed) + '。预览默认从首个变化页开始。')
    elif report.get('page_check', {}).get('pixels_equal'):
        st.info('格式属性已写入，但当前渲染画面没有变化；请检查所选字体是否支持这些文字。')
    before, after = session.preview(result.parent), session.preview(result.id)
    controls, actions = st.columns([2, 1])
    with controls:
        maximum = max(before['pages'], after['pages'])
        selected = st.number_input('预览页码', 1, maximum, min(changed[0] if changed else 1, maximum), key=prefix+'result_page')
        compare = st.checkbox('对比修改前', key=prefix+'compare')
    with actions:
        if st.button('继续修改', width='stretch'):
            _go_to('编辑')
        if st.button('撤销本次', width='stretch'):
            session.undo(); _reset_draft(session); _go_to('预览与下载')
    st.caption(f"修改前 {before['pages']} 页 → 当前 {after['pages']} 页")
    reviewed = st.checkbox('已查看预览，确认下载', key='reviewed:' + digest(result.data) + ':' + result.id)
    _download(result, '下载修改后的 ' + result.extension.upper(), disabled=not reviewed)
    if active:
        if compare:
            left, right = st.columns(2)
            with left:
                st.caption('修改前')
                if selected <= before['pages']:
                    st.image(before['images'][selected-1], width='stretch')
                else:
                    st.info('修改前没有这一页。')
            with right:
                st.caption('当前结果')
                if selected <= after['pages']:
                    st.image(after['images'][selected-1], width='stretch')
                else:
                    st.info('修改后没有这一页。')
        elif selected <= after['pages']:
            st.image(after['images'][selected-1], width='stretch')
        else:
            st.info('这一页在修改后已不存在，可开启前后对比查看。')
    with st.expander('检查详情与报告'):
        st.caption('结构、目标格式、内容与对象保全检查已通过。其他 Word/WPS 版本的显示效果未在本次验证。')
        for warning in report.get('font_warnings', []):
            st.warning(warning)
        st.download_button('下载 HTML 检查报告', report_html(result, visual_review='confirmed' if reviewed else 'pending'),
                           file_name=result.name.rsplit('.', 1)[0]+'_检查报告.html', mime='text/html')


def _advanced(session, model, prefix, active, selected_scope):
    st.subheader('历史版本')
    selected = st.selectbox('本次输入版本', list(session.versions), index=list(session.versions).index(session.current),
                            format_func=lambda v: f"{v} · {session.versions[v].name}", key=prefix+'history')
    if selected != session.current:
        session.select(selected); _reset_draft(session); _go_to('预览与下载')
    st.download_button('下载保留的原始上传文件', session.original, file_name=session.original_name)
    _roles_and_protections(model, prefix, selected_scope)
    _reference(model, prefix)
    with st.expander('文档检查与原稿核对'):
        st.write(f"{len(model.paragraphs)} 个段落 · {model.inventory['tables']} 张表 · {model.inventory['comments']} 条批注 · {model.inventory['embedded_parts']} 个嵌入部件")
        if session.conversion['status'] == 'original_format':
            st.success('原稿与未修改工作副本字节完全一致，保留 DOC 原格式。')
            if session.conversion['page_check']['pixels_equal']:
                st.caption(f"导入时全部 {session.conversion['source_pages']} 页独立预览逐像素一致（120 DPI）。")
            show = st.checkbox('查看导入时的原稿对照', key=prefix+'import_comparison')
            if show and active:
                st.caption('这里仅核对导入时的未修改文件；最新修改结果在“预览与下载”。')
                first, second = st.columns(2)
                with first:
                    _preview_pages(session.renders['conversion_source'], '原始 DOC', prefix+'import_old')
                with second:
                    _preview_pages(session.renders['conversion_working'], '未修改工作副本', prefix+'import_new')


def main():
    st.set_page_config(page_title='文序 · 精确格式调整', page_icon='📄', layout='wide')
    st.markdown('<style>.block-container{max-width:1150px;padding-top:1.5rem}h1{font-size:1.8rem!important}</style>', unsafe_allow_html=True)
    st.title('文序 · 精确格式调整')
    with st.expander('上传或更换稿件', expanded='editor' not in st.session_state):
        uploaded = st.file_uploader('上传自己的稿件（DOCX / DOC，最多 10 MB）', type=['docx', 'doc'], key='manuscript')
    if not uploaded:
        st.info('上传稿件后，选择范围和格式即可开始。文件仅在本机处理。')
        return
    identity = IMPORT_VERSION + digest(uploaded.getvalue()) + uploaded.name
    if st.session_state.get('upload_identity') != identity:
        st.session_state.pop('editor', None)
        try:
            with st.spinner('读取稿件并核对原格式…'):
                session = EditingSession.load(uploaded.name, uploaded.getvalue())
            st.session_state.editor = session
            st.session_state.upload_identity = identity
            _reset_draft(session)
            st.session_state.pending_tab = '编辑'
        except (FormatError, ValueError, OSError) as exc:
            st.error(str(exc)); return
    session = st.session_state.editor
    version = session.versions[session.current]
    st.caption(f"{uploaded.name} · 保留 {version.extension.upper()} 格式 · 当前 {'原稿' if session.current == 'original' else session.current}")
    for issue in session.conversion.get('blocking_issues', []):
        st.error(issue)
    if st.session_state.get('pending_tab'):
        st.session_state.editor_tabs = st.session_state.pop('pending_tab')
    tabs = st.tabs(['编辑', '预览与下载', '历史与高级'], key='editor_tabs', on_change='rerun')
    prefix = str(st.session_state.epoch) + ':'
    model = build_model(version.model_data, version.id, st.session_state.roles)
    with tabs[0]:
        selected_scope = _edit(session, model, prefix)
    with tabs[1]:
        _delivery(session, prefix, tabs[1].open)
    with tabs[2]:
        _advanced(session, model, prefix, tabs[2].open, selected_scope)
