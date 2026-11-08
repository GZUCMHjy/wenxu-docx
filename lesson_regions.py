"""Semantic teaching-process regions; numerical column positions are never a standard."""
from copy import deepcopy
from dataclasses import dataclass, field
import re

from docx.oxml.ns import qn
from docx_formatting import FormatError, NS

GROUP_LABELS = {'headers': '表格列头与行头', 'content': '教学内容', 'activities': '师生活动与设计意图', 'closing': '小结、作业与板书内容'}
CLOSING_LABELS = {'课堂小结', '作业布置', '板书设计'}
CLOSING_ALIASES = {**{label: label for label in CLOSING_LABELS},
                   'summary': '课堂小结', 'homework': '作业布置'}
HEADERS = {'教学内容': 'content', '师生活动': 'activity', '设计意图': 'intent',
           **{name: 'stage' for name in ('教学环节', '教学环节及时间分配', '教学环节与时间分配', '教学环节及时间', '教学环节与时间', '教学环节时间', '教学步骤', '环节及时间分配', '环节')}}
POLICIES = {'preserve', 'same_as_content'}
METADATA_LABELS = {'课程名称', '课题', '学科', '年级', '班级', '任课教师', '教师', '授课人', '授课日期', '日期', '课时', '教研组', '审核人', '签字'}


@dataclass
class RegionMapping:
    groups: dict = field(default_factory=lambda: {key: [] for key in GROUP_LABELS})
    issues: list[str] = field(default_factory=list)
    preserved_headings: list[dict] = field(default_factory=list)
    tables: list[int] = field(default_factory=list)
    ignored_tables: list[int] = field(default_factory=list)
    notices: list[str] = field(default_factory=list)


def _label(text):
    return re.sub(r'[\s（）()]', '', text).strip('：:')


def _closing(text):
    value = _label(text).casefold()
    return next((label for alias, label in CLOSING_ALIASES.items()
                 if re.fullmatch(re.escape(alias) + r'(?:\d+(?:分钟|min))?', value)), None)


def _heading(row):
    if row['role'].startswith('heading_') or row['role'] == 'title':
        return '已有标题角色；保留独立字体与字号'
    text = row['text'].strip()
    # A short Chinese section name is a visible exception, never a heading rewrite.
    match = re.fullmatch(r'(?:[一二三四五六七八九十]+、|[（(][一二三四五六七八九十]+[）)])([\u4e00-\u9fff]{2,16})', text)
    if match and not re.search(r'计算|求|解题|已知|选择|填空|判断|证明', match[1]):
        return '中文编号短标题候选；独立字号尚未确认'
    return None


def detect_regions(model):
    result = RegionMapping()
    tables = model.root.findall('.//w:tbl', NS)
    previous_grid = None
    for tid, element in enumerate(tables, 1):
        cells = [c for c in model.cells if c['table'] == tid]
        if not cells:
            result.ignored_tables.append(tid); continue
        rows = sorted({r for c in cells for r in range(c['row'], c['end_row'] + 1)})
        row_cells = {r: sorted([c for c in cells if c['row'] <= r <= c['end_row']], key=lambda c: c['col']) for r in rows}
        headers = {}
        problems = []
        suspected = False
        for row, items in row_cells.items():
            semantic = [(c, HEADERS.get(_label(c['text']))) for c in items]
            names = [name for _, name in semantic if name]
            if len(set(names) & {'content', 'activity', 'intent'}) >= 2:
                suspected = True
                if any(names.count(name) != 1 for name in ('content', 'activity', 'intent')) or len(names) != len(items) or names.count('stage') > 1:
                    problems.append(f'表 {tid} 第 {row} 行列头重复、不完整或包含未识别栏目。')
                elif any(c['end_row'] != row for c in items):
                    problems.append(f'表 {tid} 第 {row} 行列头跨行合并，需核对。')
                else:
                    headers[row] = {col: name for c, name in semantic for col in range(c['col'], c['end_col'] + 1)}
        standalone = not headers and all(len(items) == 2 and items[0]['row'] == items[0]['end_row'] == row and items[1]['row'] == items[1]['end_row'] == row and _closing(items[0]['text']) and items[0]['end_col'] + 1 == items[1]['col'] for row, items in row_cells.items())
        if not headers and not standalone and not suspected:
            width = max(c['end_col'] for c in cells)
            metadata = len({_label(c['text']) for c in cells} & METADATA_LABELS) >= 2
            if previous_grid == width and width >= 3 and not metadata:
                result.issues.append(f'表 {tid} 可能是教学过程续表，但缺少语义列头，需核对。')
            result.ignored_tables.append(tid)
            result.notices.append(f'表 {tid} 未纳入四组教学过程区域，保留原样。')
            previous_grid = None
            continue
        if element.findall('.//w:tbl', NS) or any(a.tag == qn('w:tbl') for a in element.iterancestors()):
            problems.append(f'表 {tid} 含嵌套表格，区域边界需核对。')
        if headers and len({tuple(sorted(schema.items())) for schema in headers.values()}) > 1:
            problems.append(f'表 {tid} 的重复列头改变了业务列结构，需核对。')
        assignments = {}
        active = None
        for row, items in row_cells.items():
            if row in headers:
                active = headers[row]
                for cell in items: assignments.setdefault(cell['id'], set()).add('headers')
                continue
            if not active and not standalone:
                if any(c['text'].strip() for c in items):
                    result.notices.append(f'表 {tid} 第 {row} 行位于已识别列头之前，保留原样。')
                continue
            stage_cells = [] if standalone else [c for c in items if {active.get(col) for col in range(c['col'], c['end_col'] + 1)} == {'stage'}]
            for cell in stage_cells:
                if cell['end_row'] > cell['row'] and not _closing(cell['text']) and any(alias in _label(cell['text']).casefold() for alias in CLOSING_ALIASES):
                    problems.append(f'表 {tid} 的纵向合并行头混有特殊环节名称，需核对。')
            label_cells = items[:1] if standalone else stage_cells
            labels = [c for c in label_cells if _closing(c['text'])]
            if len(labels) > 1:
                problems.append(f'表 {tid} 第 {row} 行有多个特殊行名称，需核对。')
            closing = labels[0] if len(labels) == 1 else None
            for cell in items:
                if standalone:
                    group = 'headers' if cell is items[0] else 'closing'
                elif cell is closing:
                    group = 'headers'
                else:
                    names = {active.get(col) for col in range(cell['col'], cell['end_col'] + 1)}
                    if names == {'stage'}:
                        group = 'headers'
                    elif closing and None not in names and 'stage' not in names:
                        group = 'closing'
                    elif names == {'content'}:
                        group = 'content'
                    elif names and names <= {'activity', 'intent'}:
                        group = 'activities'
                    else:
                        problems.append(f"表 {tid} 第 {row} 行的合并单元格跨越不同业务区域，需核对。")
                        continue
                assignments.setdefault(cell['id'], set()).add(group)
        for cell in cells:
            if len(assignments.get(cell['id'], set())) > 1:
                problems.append(f"表 {tid} 的纵向合并单元格跨越普通行与特殊行，需核对。")
        if problems:
            result.issues.extend(problems)
            result.ignored_tables.append(tid)
            previous_grid = None
            continue
        result.tables.append(tid)
        previous_grid = max(c['end_col'] for c in cells) if headers else None
        for cell in cells:
            groups = assignments.get(cell['id'], set())
            if groups:
                result.groups[next(iter(groups))].extend(cell['pids'])
    for pid in list(result.groups['content']):
        row = model.row(pid); reason = _heading(row)
        if reason:
            result.groups['content'].remove(pid)
            result.preserved_headings.append({key: row[key] for key in ('pid', 'text', 'location')} | {'reason': reason})
    result.groups = {key: sorted(set(pids)) for key, pids in result.groups.items()}
    result.issues = sorted(set(result.issues))
    return result


def default_settings():
    return {key: {'font_family': '宋体', 'font_size': size} for key, size in (('headers', 9), ('content', 12), ('activities', 9), ('closing', 10.5))}


def validate_settings(settings, heading_policy='preserve'):
    from v06_plan import validate_values
    if heading_policy not in POLICIES or not isinstance(settings, dict) or set(settings) != set(GROUP_LABELS):
        raise FormatError('四组区域标准或标题策略无效。')
    result = {}
    for key, values in settings.items():
        if not isinstance(values, dict) or set(values) != {'font_family', 'font_size'}:
            raise FormatError('区域标准只调整字体与字号。')
        checked = validate_values(values)
        result[key] = {'font_family': checked['font_east_asia'], 'font_size': checked['font_size']}
    return result


def region_pids(mapping, region, heading_policy='preserve'):
    if region not in GROUP_LABELS or heading_policy not in POLICIES:
        raise FormatError('教案区域或标题策略无效。')
    pids = list(mapping.groups[region])
    if region == 'content' and heading_policy == 'same_as_content':
        pids += [h['pid'] for h in mapping.preserved_headings]
    return sorted(set(pids))


def region_rules(model, settings, heading_policy='preserve', mapping=None):
    from v06_plan import intent
    settings = validate_settings(settings, heading_policy)
    mapping = mapping or detect_regions(model)
    return [{'key': 'region:' + key, **intent(model.scope('lesson_region', region=key, heading_policy=heading_policy), settings[key], source='用户教案区域标准 · ' + label)} for key, label in GROUP_LABELS.items() if region_pids(mapping, key, heading_policy)]


def review_regions(model, settings=None, heading_policy='preserve', protections=None):
    from template_workflow import TemplateReview, _evaluate
    from v06_plan import resolve_scope
    settings = validate_settings(default_settings() if settings is None else settings, heading_policy)
    mapping = detect_regions(model)
    rules = region_rules(model, settings, heading_policy, mapping)
    review = TemplateReview('lesson_regions', rules=rules, blockers=list(mapping.issues))
    review.region_mapping = mapping
    review.notices = list(mapping.notices)
    if heading_policy == 'preserve':
        review.notices += [h['location'] + '：' + h['reason'] for h in mapping.preserved_headings]
    if not mapping.tables and not mapping.issues:
        review.blockers.append('没有可靠识别到教学过程表或独立特殊行。')
    checked = []
    for scope in protections or []:
        try:
            resolve_scope(model, scope); checked.append(deepcopy(scope))
        except (FormatError, TypeError, AttributeError) as exc:
            review.blockers.append('当前保护范围无效：' + str(exc))
    return _evaluate(model, rules, checked, review)


def binding(settings=None, heading_policy='preserve'):
    from v06_model import digest
    from v06_plan import encode
    settings = validate_settings(default_settings() if settings is None else settings, heading_policy)
    definition = {'version': 1, 'settings': settings, 'heading_policy': heading_policy}
    fingerprint = digest(encode(definition).encode())
    return {'name': '教案区域标准', 'basis_kind': 'lesson_regions', 'definition_hash': fingerprint, 'fingerprint': fingerprint}
