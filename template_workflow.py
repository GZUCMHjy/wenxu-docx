"""Explicit template adoption, calibrated against the actual saved document.

Comments are inert outside TemplateWorkspace. Unknown semantics stay pending;
portable rules refer to roles/column names, never another document's offsets.
"""
from __future__ import annotations

from copy import copy, deepcopy
from dataclasses import dataclass, field
from io import BytesIO
import json
from pathlib import PurePosixPath
import posixpath
import re
from urllib.parse import unquote
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from lxml import etree

from docx_formatting import FormatError, MAIN, MAX_FILE_BYTES, NS, _paragraphs, _xml
from prd_formatting import PARAGRAPH_PROPERTIES, PAGE_PROPERTIES, _effective_paragraph_value
from v06_model import ROLE_LABELS, TEXT_TAGS, build_model, digest, node_text, run_spans
from v06_patch import effective_run
from v06_plan import build_plan, encode, intent, parse_text, resolve_scope, validate_values
from v06_session import EditingSession

SCHEMA = 1
STAGES = {'template', 'manuscript', 'after_class', 'printing'}
EXTRACT_PROPERTIES = ('font_east_asia', 'font_western', 'font_size', 'bold', 'italic', 'underline', 'color', *sorted(PARAGRAPH_PROPERTIES))
REQUIRED_PROPERTIES = {'font_east_asia', 'font_western', 'font_size', 'bold'}


def extract_comments(model):
    """Validate annotations and locate exact visible-character boundaries in XML order."""
    parts = model.parts
    unsupported = [name for name in parts if re.search(r'/(commentsExtensible|commentsIds)\.xml$', name, re.I)]
    if unsupported:
        raise FormatError('存在尚未支持的现代批注部件，不能按无旁批模板处理。')
    rels = _xml(parts.get('word/_rels/document.xml.rels', b'<Relationships/>'))
    linked = []
    for rel in rels:
        if rel.get('Type', '').endswith('/comments'):
            if rel.get('TargetMode') == 'External':
                raise FormatError('批注关系不能指向外部文件。')
            target = unquote(rel.get('Target', ''))
            linked.append(posixpath.normpath(target.lstrip('/') if target.startswith('/') else posixpath.join('word', target)))
    declared = [name for name in parts if re.search(r'/comments[^/]*\.xml$', name, re.I) and name != 'word/commentsExtended.xml']
    if len(linked) > 1 or any(name not in linked for name in declared):
        raise FormatError('批注部件与关系不一致。')
    comments = {}
    if linked:
        if linked[0] not in parts:
            raise FormatError('批注关系目标缺失。')
        root = _xml(parts[linked[0]])
        if root.tag != qn('w:comments'):
            raise FormatError('批注部件类型不支持。')
        for c in root:
            ident = c.get(qn('w:id'))
            if c.tag != qn('w:comment') or ident is None or not re.fullmatch(r'\d+', ident) or ident in comments:
                raise FormatError('批注 ID 缺失、重复或无效。')
            if any(etree.QName(n).localname in {'drawing', 'object', 'oMath', 'altChunk'} for n in c.iter()):
                raise FormatError('旁批含尚未支持的非文本要求，需人工核对。')
            comments[ident] = {'id': ident, 'key': 'comment:' + ident, 'text': '\n'.join(node_text(p) for p in c.findall('.//w:p', NS)), 'author': c.get(qn('w:author'), ''), 'revision': model.revision}
    _validate_comment_metadata(parts, rels, linked)
    markers = {qn('w:' + name): name for name in ('commentRangeStart', 'commentRangeEnd', 'commentReference')}
    for name, raw in parts.items():
        if name.startswith('word/') and name.endswith('.xml') and name not in {MAIN, *linked}:
            if any(n.tag in markers for n in _xml(raw).iter()):
                raise FormatError('正文以外的批注位置尚未支持，需人工核对。')
    paragraphs = _paragraphs(model.root)
    pids = {p: i + 1 for i, p in enumerate(paragraphs)}
    allowed = {child for p in paragraphs for run, *_ in run_spans(p) for child in run if child.tag in TEXT_TAGS}
    bounds, positions, offset, current_pid, last_pid = {}, {}, 0, None, None
    for event, node in etree.iterwalk(model.root, events=('start', 'end')):
        if node in pids:
            if event == 'start':
                current_pid = pids[node]; bounds[current_pid] = [offset, offset]
            else:
                bounds[pids[node]][1] = offset
                last_pid, current_pid = pids[node], None
        if event != 'start':
            continue
        if node.tag in markers:
            ident = node.get(qn('w:id'))
            if ident not in comments:
                raise FormatError('批注标记引用不存在的批注。')
            kind = markers[node.tag]
            located = positions.setdefault(ident, {})
            if kind in located:
                raise FormatError('批注范围或引用标记重复。')
            located[kind] = (offset, current_pid or last_pid)
        if node in allowed:
            offset += len(node_text(node))
    for ident, item in comments.items():
        marks = positions.get(ident, {})
        ref = marks.get('commentReference')
        start, end = marks.get('commentRangeStart'), marks.get('commentRangeEnd')
        if ref is None or bool(start) != bool(end) or (start and start[0] > end[0]):
            raise FormatError('批注引用或范围边界不完整。')
        a, b = (start[0], end[0]) if start else (ref[0], ref[0])
        spans = [{'pid': pid, 'start': max(a, x)-x, 'end': min(b, y)-x} for pid, (x, y) in bounds.items() if max(a, x) < min(b, y)]
        point_pid = (start or ref)[1]
        point_offset = a - bounds[point_pid][0] if point_pid else 0
        if a == b and point_pid:
            spans = [{'pid': point_pid, 'start': point_offset, 'end': point_offset}]
        if not spans:
            raise FormatError('批注锚点无法映射到支持的正文或表格位置。')
        item.update(spans=spans, point=a == b,
                    context='\n'.join(model.row(s['pid'])['location'] + ' · ' + model.row(s['pid'])['text'] for s in spans),
                    status='pending', unparsed='')
    return list(comments.values())


def _validate_comment_metadata(parts, rels, linked):
    """Unthreaded, unresolved Office metadata is passive; replies stay unsupported."""
    w14 = '{http://schemas.microsoft.com/office/word/2010/wordml}'
    w15 = '{http://schemas.microsoft.com/office/word/2012/wordml}'
    for filename, relation in (('commentsExtended', 'commentsExtended'), ('people', 'people')):
        name = 'word/' + filename + '.xml'
        if name not in parts:
            continue
        matches = [r for r in rels if r.get('Type', '').endswith('/' + relation)]
        if len(matches) != 1 or matches[0].get('TargetMode') == 'External' or not linked:
            raise FormatError('批注扩展元数据的关系无效。')
        target = unquote(matches[0].get('Target', ''))
        if posixpath.normpath(target.lstrip('/') if target.startswith('/') else posixpath.join('word', target)) != name:
            raise FormatError('批注扩展元数据的目标不匹配。')
        metadata = _xml(parts[name])
        if filename == 'commentsExtended':
            if metadata.tag != w15 + 'commentsEx':
                raise FormatError('批注扩展类型尚未支持。')
            anchors = []
            for c in _xml(parts[linked[0]]):
                paragraphs = c.findall('.//w:p', NS)
                para_id = paragraphs[-1].get(w14 + 'paraId') if paragraphs else None
                if not para_id or not re.fullmatch('[0-9a-fA-F]{8}', para_id):
                    raise FormatError('批注扩展缺少可靠的段落关联。')
                anchors.append(para_id.upper())
            observed = []
            for item in metadata:
                if item.tag != w15 + 'commentEx' or len(item) or set(item.attrib) - {w15 + 'paraId', w15 + 'done'} or item.get(w15 + 'done') not in {None, '0', 'false'}:
                    raise FormatError('批注含回复、已解决状态或未知扩展语义，需人工核对。')
                para_id = item.get(w15 + 'paraId', '')
                if not re.fullmatch('[0-9a-fA-F]{8}', para_id):
                    raise FormatError('批注扩展段落 ID 无效。')
                observed.append(para_id.upper())
            if len(set(anchors)) != len(anchors) or len(set(observed)) != len(observed) or set(anchors) != set(observed):
                raise FormatError('批注扩展段落关联缺失或重复。')
        else:
            if metadata.tag != w15 + 'people':
                raise FormatError('批注作者元数据类型无效。')
            for person in metadata:
                if person.tag != w15 + 'person' or set(person.attrib) - {w15 + 'author'}:
                    raise FormatError('批注作者元数据含未知语义。')
                for presence in person:
                    if presence.tag != w15 + 'presenceInfo' or len(presence) or set(presence.attrib) - {w15 + 'providerId', w15 + 'userId'}:
                        raise FormatError('批注作者状态含未知语义。')


def _parse_comment(model, comment):
    text = comment['text'].strip()
    # A labelled example occupies one clause, never the rest of the comment.
    # Later requirements keep their delimiters and must still parse completely.
    text = re.sub(r'([；;])\s*(?:例|例如)\s*[:：][^；;。\n，,]*', r'\1', text)
    role = next((role for role, label in ROLE_LABELS.items() if role != 'unknown' and re.match(re.escape(label) + r'\s*[:：]', text)), None)
    if role:
        text = re.sub(r'^' + re.escape(ROLE_LABELS[role]) + r'\s*[:：]', ROLE_LABELS[role], text, count=1)
    requests = parse_text(model, text)
    unresolved = [r.get('unparsed') for r in requests if r.get('unparsed')]
    if not requests or comment['point']:
        return None, role, '；'.join(unresolved) or ('插入点需要确认实际作用范围' if comment['point'] else '要求未理解')
    if len({encode(r['scope']) for r in requests}) != 1:
        return None, role, '多个作用范围需要明确拆分'
    values = {}
    for request in requests:
        if not request['set']:
            continue
        try:
            validated = validate_values(request['set'])
        except FormatError as exc:
            return None, role, str(exc)
        for prop, value in validated.items():
            if prop in values and values[prop] != value:
                return None, role, '同一属性要求冲突'
            values[prop] = value
    whole = all(s['start'] == 0 and s['end'] == len(model.row(s['pid'])['text']) for s in comment['spans'])
    if role and not whole:
        return None, role, '角色要求锚点未覆盖完整段落，请确认扩展范围'
    scope = requests[0]['scope']
    if not values or scope['kind'] == 'unknown':
        return None, role, '；'.join(unresolved) or '实际作用范围或属性需要确认'
    return {'key': comment['key'], **intent(scope, values, source=comment['text'])}, role, '；'.join(unresolved)


def _scope(model, scope):
    if not isinstance(scope, dict):
        raise FormatError('规则作用范围无效。')
    result = {**deepcopy(scope), 'revision': model.revision}
    resolve_scope(model, result)
    if result.get('kind') == 'named_column' and result.get('column_id'):
        result['label'] = next(c['label'] for c in model.columns if c['id'] == result['column_id'])
    return result


def _same(a, b):
    if type(a) in {float, int} and type(b) in {float, int}:
        return abs(a - b) <= .051
    if isinstance(a, dict) and isinstance(b, dict) and set(a) == set(b) == {'mode', 'value'}:
        return a['mode'] == b['mode'] and abs(a['value'] - b['value']) <= .005
    return a == b


def _value_reader(model):
    """One read-only document per review; effective properties reuse the edit engine."""
    document = Document(BytesIO(model.data))
    paragraphs = [Paragraph(p, document) for p in _paragraphs(document.element)]
    spans_by_pid = [run_spans(p._p) for p in paragraphs]
    cache = {}
    formatting_cache = {}
    defaults = document.styles.element.find('w:docDefaults/w:pPrDefault/w:pPr', NS)
    default_p = OxmlElement('w:p')
    if defaults is not None: default_p.append(deepcopy(defaults))
    default_paragraph = Paragraph(default_p, document)

    def read(scope, prop):
        spans, _ = resolve_scope(model, scope)
        if prop in PAGE_PROPERTIES:
            return [getattr(document.sections[0], prop).pt]
        values = []
        for span in spans:
            pid = span['pid']; paragraph = paragraphs[pid - 1]
            if (pid, prop) not in cache:
                if prop in PARAGRAPH_PROPERTIES:
                    ppr = paragraph._p.find(qn('w:pPr'))
                    signature = (prop, etree.tostring(ppr) if ppr is not None else b'')
                    if signature not in formatting_cache:
                        value = _effective_paragraph_value(paragraph, prop)
                        if value is None and defaults is not None:
                            value = _effective_paragraph_value(default_paragraph, prop)
                        formatting_cache[signature] = value
                    value = formatting_cache[signature]
                    cache[pid, prop] = [(0, 0, value)]
                else:
                    # Include field text: exclusions cannot count as verified targets.
                    pstyle = paragraph._p.find('w:pPr/w:pStyle', NS)
                    style_id = pstyle.get(qn('w:val')) if pstyle is not None else None
                    observed = []
                    for run, a, b, _, _ in spans_by_pid[pid - 1]:
                        if a >= b: continue
                        rpr = run.find(qn('w:rPr'))
                        signature = (prop, style_id, etree.tostring(rpr) if rpr is not None else b'')
                        if signature not in formatting_cache:
                            formatting_cache[signature] = effective_run(run, paragraph, prop, document)
                        observed.append((a, b, formatting_cache[signature]))
                    cache[pid, prop] = observed
            for a, b, value in cache[pid, prop]:
                if prop in PARAGRAPH_PROPERTIES or (a < span['end'] and b > span['start']):
                    if value not in values: values.append(value)
        return values
    return read


@dataclass
class TemplateReview:
    basis_kind: str
    comments: list[dict] = field(default_factory=list)
    rules: list[dict] = field(default_factory=list)
    blockers: list[str] = field(default_factory=list)
    manual_items: list[dict] = field(default_factory=list)
    differences: list[dict] = field(default_factory=list)
    ready: bool = False
    plan: object = None
    protections: list[dict] = field(default_factory=list)
    notices: list[str] = field(default_factory=list)
    region_mapping: object = None


def _evaluate(model, rules, protections, review, read=None):
    read = read or _value_reader(model)
    review.protections = deepcopy(protections)
    # Frozen plans sort object keys; use that same order before generating ordered
    # exclusion records so the existing engine can independently rebuild them.
    rules = json.loads(encode(rules))
    if not rules:
        if not review.blockers:
            review.blockers.append('没有可转移的明确格式规则。')
        return review
    # The unprotected plan represents every requirement, including local exceptions.
    full = build_plan(model, rules)
    review.blockers.extend(full.blockers)
    if full.plan:
        for a in full.preview:
            scope = model.scope('page') if a['pid'] == 0 else model.scope('paragraphs', pids=[a['pid']]) if a['property'] in PARAGRAPH_PROPERTIES else model.scope('selection', spans=[{k: a[k] for k in ('pid', 'start', 'end')}])
            observed = read(scope, a['property'])
            if not observed or not all(_same(v, a['value']) for v in observed):
                review.differences.append({**a, 'actual': observed})
        for excluded in full.excluded:
            review.blockers.append(f"段落 {excluded.get('pid')} 有字段保护，要求未完整检查。")
    restricted = build_plan(model, rules, protections)
    if full.plan:
        for excluded in restricted.excluded:
            for assignment in full.preview:
                pid, prop = assignment['pid'], assignment['property']
                if pid != excluded.get('pid') or prop != excluded.get('property'):
                    continue
                if prop in PARAGRAPH_PROPERTIES:
                    scope = model.scope('paragraphs', pids=[pid])
                else:
                    start = max(assignment['start'], excluded.get('start', 0))
                    end = min(assignment['end'], excluded.get('end', assignment['end']))
                    if start >= end: continue
                    scope = model.scope('selection', spans=[{'pid': pid, 'start': start, 'end': end}])
                observed = read(scope, prop)
                if not observed or not all(_same(v, assignment['value']) for v in observed):
                    review.blockers.append(f'保护区内的段落 {pid} 未满足 {prop}，请明确解决冲突。')
    if review.differences or restricted.plan:
        review.blockers.extend(restricted.blockers)
    review.blockers = sorted(set(review.blockers))
    review.plan = restricted.plan if not review.blockers else None
    review.ready = not review.blockers and not review.differences
    return review


class TemplateWorkspace:
    def __init__(self, session):
        self.session = session
        self.roles = {}
        self.rules = {}
        self.decisions = {}
        self.protections = []
        self.region_standard = None
        self.original_hash = digest(session.original)

    def model(self):
        version = self.session.versions[self.session.current]
        base = self.session.model()
        roles = {**base.role_overrides, **self.roles}
        try:
            for comment in extract_comments(base):
                rule, role, pending = _parse_comment(base, comment)
                if rule and role and not pending and comment['key'] not in self.decisions:
                    for span in comment['spans']:
                        if span['pid'] not in self.roles:
                            roles[span['pid']] = role
        except FormatError:
            pass  # review exposes the annotation failure; never chooses sample styles.
        model = build_model(version.model_data, version.id, roles)
        for p in model.paragraphs:
            if p['pid'] in roles and p['pid'] not in self.roles and p['pid'] not in base.role_overrides:
                p['evidence'] = '模板旁批明确角色'
        return model

    def set_roles(self, mapping):
        candidate = {**self.roles, **mapping}
        version = self.session.versions[self.session.current]
        build_model(version.model_data, version.id, candidate)
        self.roles = candidate

    def set_rule(self, key, scope, values, *, exception=False):
        if not isinstance(key, str) or not key or type(exception) is not bool:
            raise FormatError('规则标识或例外标记无效。')
        model = self.model()
        scope = _scope(model, scope)
        values = validate_values(values)
        if any((prop in PAGE_PROPERTIES) != (scope['kind'] == 'page') for prop in values):
            raise FormatError('页面属性需要明确页面范围；字符和段落属性需要内容范围。')
        if key.startswith('style:') and (scope.get('kind') != 'role' or scope.get('role') != key.removeprefix('style:')):
            raise FormatError('类别样式规则应保留对应类别范围；局部例外请另加规则。')
        if exception and scope['kind'] in {'all', 'role', 'page'}:
            raise FormatError('局部例外需要明确局部范围。')
        if key.startswith('comment:') and key not in {c['key'] for c in extract_comments(model)}:
            raise FormatError('所选批注不存在。')
        self.rules[key] = {'key': key, **intent(scope, values, source='用户明确规范', exception=exception)}
        self.decisions.pop(key, None)

    def classify_comment(self, ident, classification, *, reason, stage, scope=None):
        key = ident if str(ident).startswith('comment:') else 'comment:' + str(ident)
        if key not in {c['key'] for c in extract_comments(self.model())}:
            raise FormatError('所选批注不存在。')
        if classification not in {'manual', 'not_applicable', 'protection'} or stage not in STAGES or not isinstance(reason, str) or not reason.strip():
            raise FormatError('人工事项或不适用项必须明确分类、原因及适用阶段。')
        decision = {'classification': classification, 'reason': reason.strip(), 'stage': stage, 'verification': 'not_automatically_verified'}
        if classification == 'protection':
            decision['scope'] = _scope(self.model(), scope)
        elif classification == 'not_applicable':
            self.rules.pop(key, None)
        self.decisions[key] = decision

    def set_protections(self, scopes):
        self.protections = [_scope(self.model(), scope) for scope in scopes]

    def reset_rule(self, key):
        self.rules.pop(key, None)
        self.decisions.pop(key, None)

    def clone(self):
        candidate = copy(self)
        candidate.session = copy(self.session)
        for name in ('versions', 'renders', 'jobs', 'environment', 'conversion'):
            setattr(candidate.session, name, getattr(self.session, name).copy())
        for name in ('roles', 'rules', 'decisions', 'protections', 'region_standard'):
            setattr(candidate, name, deepcopy(getattr(self, name)))
        return candidate

    def set_region_standard(self, settings, heading_policy='preserve'):
        from lesson_regions import validate_settings
        settings = validate_settings(settings, heading_policy)
        if self.region_standard:
            previous = self.region_standard['previous_rules']
        else:
            before = self.review()
            previous = before.rules if before.ready and before.basis_kind == 'sample_styles' else []
        previous = deepcopy(previous)
        for rule in previous: rule['scope'].pop('revision', None)
        self.region_standard = {'settings': settings, 'heading_policy': heading_policy, 'previous_rules': previous}

    def _region_rules(self, model, rules, review):
        from lesson_regions import detect_regions, region_rules
        standard = self.region_standard
        mapping = detect_regions(model)
        review.region_mapping = mapping
        review.blockers.extend(mapping.issues)
        review.notices.extend(mapping.notices)
        if standard['heading_policy'] == 'preserve':
            review.notices.extend(h['location'] + '：' + h['reason'] for h in mapping.preserved_headings)
        if not mapping.tables and not mapping.issues:
            review.blockers.append('没有可靠识别到教学过程表或独立特殊行。')
        result = []
        for rule in rules:
            values = validate_values(rule['set'])
            replaced = {key: value for key, value in values.items() if key in {'font_east_asia', 'font_western', 'font_size'}}
            unchanged = {key: value for key, value in values.items() if key not in replaced}
            if unchanged:
                result.append({**rule, 'set': unchanged})
            if replaced:
                result.append({**rule, 'key': rule['key'] + ':outside_regions', 'scope': {**rule['scope'], 'exclude_lesson_regions': True, 'lesson_heading_policy': standard['heading_policy']}, 'set': replaced})
        result.extend(region_rules(model, standard['settings'], standard['heading_policy'], mapping))
        return result

    def _protections(self, model):
        return [_scope(model, p) for p in self.protections] + [_scope(model, d['scope']) for d in self.decisions.values() if d['classification'] == 'protection']

    def _sample_rules(self, model, review, read):
        rules = []
        for role in sorted({p['role'] for p in model.paragraphs if p['text'].strip()} - {'unknown'}):
            scope = model.scope('role', role=role)
            key = 'style:' + role
            explicit = self.rules.get(key)
            overrides = explicit['set'] if explicit and explicit['scope'].get('kind') == 'role' and explicit['scope'].get('role') == role else {}
            values = {}
            for prop in EXTRACT_PROPERTIES:
                observed = read(scope, prop)
                if prop in overrides:
                    values[prop] = overrides[prop]
                elif prop == 'color' and observed in ([None], ['auto']):
                    continue
                elif len(observed) == 1 and observed[0] is not None and not isinstance(observed[0], list) and not (isinstance(observed[0], str) and observed[0].startswith('theme:')):
                    try:
                        values.update(validate_values({prop: observed[0]}))
                    except FormatError:
                        review.blockers.append(f'{ROLE_LABELS[role]}的 {prop} 尚未支持：{observed}')
                elif len(observed) > 1 or prop in REQUIRED_PROPERTIES or (observed and observed != [None]):
                    review.blockers.append(f'{ROLE_LABELS[role]}的 {prop} 混合或无法确定：{observed}；请明确规则。')
            if values:
                rules.append({'key': key, **intent(scope, values, source='来自模板现有样式')})
        return rules

    def review(self):
        model = self.model()
        try:
            comments = extract_comments(model)
        except FormatError as exc:
            return TemplateReview('comment_calibrated', blockers=[str(exc)])
        review = TemplateReview('comment_calibrated' if comments else 'sample_styles', comments=comments)
        read = _value_reader(model)
        unknown = [p['pid'] for p in model.paragraphs if p['role'] == 'unknown' and p['text'].strip()]
        if unknown and not self.region_standard:
            review.blockers.append('存在待归类内容：' + '、'.join(map(str, unknown)))
        if comments:
            rules = []
            role_evidence = {}
            for comment in comments:
                key = comment['key']
                parsed, role, unresolved = _parse_comment(model, comment)
                if parsed and role and not unresolved and key not in self.decisions:
                    for span in comment['spans']:
                        role_evidence.setdefault(span['pid'], set()).add(role)
                if key in self.decisions:
                    decision = self.decisions[key]
                    comment.update(status=decision['classification'])
                    review.manual_items.append({**comment, **decision})
                    if decision['classification'] == 'manual':
                        rule = self.rules.get(key) or _parse_comment(model, comment)[0]
                        if rule: rules.append(deepcopy(rule))
                    if decision['classification'] == 'manual' and decision['stage'] == 'template':
                        review.blockers.append(f"旁批 {comment['id']} 需要先完成模板人工处理并重新导入。")
                elif key in self.rules:
                    rules.append(deepcopy(self.rules[key])); comment['status'] = 'explicit'
                else:
                    rule = parsed
                    comment.update(role=role, unparsed=unresolved, status='pending' if unresolved else 'parsed')
                    if rule: rules.append(rule)
                    if unresolved: review.blockers.append(f"旁批 {comment['id']}：{unresolved}")
            for pid, roles in role_evidence.items():
                if len(roles) > 1 and pid not in self.roles:
                    review.blockers.append(f'段落 {pid} 的旁批角色要求冲突，请明确分类。')
        else:
            rules = deepcopy(self.region_standard['previous_rules']) if self.region_standard else self._sample_rules(model, review, read)
        generated = {r['key'] for r in rules}
        for key, rule in self.rules.items():
            if key not in generated and not key.startswith('comment:'):
                rules.append(deepcopy(rule))
        for rule in rules:
            rule['scope'] = {**rule['scope'], 'revision': model.revision}
        if self.region_standard:
            rules = self._region_rules(model, rules, review)
        review.rules = rules
        protections = self._protections(model)
        return _evaluate(model, rules, protections, review, read)

    def calibrate(self, *, environment=None):
        review = self.review()
        if review.blockers:
            raise FormatError('；'.join(review.blockers))
        if not review.differences:
            return None
        previous = self.session.current
        saved = {key: getattr(self.session, key).copy() for key in ('versions', 'jobs', 'renders', 'environment')}
        next_number = self.session.next_number
        try:
            version, _ = self.session.execute(self.model(), review.plan, environment=environment)
            if version is not None:
                self.session.select(version.id)
            after = self.review()
            if after.blockers or after.differences:
                raise FormatError('保存后重新读取未满足模板要求，不能确立基准。')
        except Exception:
            for key, value in saved.items():
                setattr(self.session, key, value)
            self.session.next_number = next_number
            self.session.select(previous)
            raise
        return version

    def baseline(self, reviewed=False):
        review = self.review()
        if not review.ready:
            raise FormatError('模板尚未满足全部适用要求：' + '；'.join(review.blockers or ['仍有格式差异']))
        version = self.session.versions[self.session.current]
        if (review.basis_kind == 'comment_calibrated' or version.parent is not None or self.region_standard) and not reviewed:
            raise FormatError('请先查看实际模板预览并确认版式。')
        def portable_record(value):
            result = deepcopy(value)
            if isinstance(result, dict):
                result.pop('revision', None)
                if 'scope' in result: result['scope'].pop('revision', None)
            return result
        manifest = {'schema': SCHEMA, 'name': self.session.original_name, 'extension': version.extension,
                    'file_hash': digest(version.data), 'original_hash': self.original_hash,
                    'basis_kind': review.basis_kind, 'roles': self.model().role_overrides,
                    'rules': {k: portable_record(v) for k, v in self.rules.items()},
                    'decisions': {k: portable_record(v) for k, v in self.decisions.items()}, 'protections': [portable_record(v) for v in self.protections],
                    'derived_rules': [portable_record(r) for r in review.rules],
                    'comments': [portable_record(c) for c in review.comments], 'reviewed': bool(reviewed)}
        if self.region_standard:
            manifest['region_standard'] = deepcopy(self.region_standard)
        manifest['fingerprint'] = digest(encode(manifest).encode())
        return TemplateBaseline(version.data, encode(manifest))


@dataclass(frozen=True)
class TemplateBaseline:
    data: bytes
    manifest_json: str

    @property
    def manifest(self): return json.loads(self.manifest_json)
    @property
    def name(self): return self.manifest['name']
    @property
    def extension(self): return self.manifest['extension']
    @property
    def basis_kind(self): return self.manifest['basis_kind']
    @property
    def fingerprint(self): return self.manifest['fingerprint']
    @property
    def file_hash(self): return self.manifest['file_hash']
    @property
    def binding(self): return {key: self.manifest[key] for key in ('name', 'basis_kind', 'fingerprint', 'file_hash')}
    @property
    def rules(self): return self.manifest['derived_rules']
    @property
    def region_standard(self): return self.manifest.get('region_standard')

    def workspace(self, *, loader=EditingSession.load):
        manifest = self.manifest
        _validate_manifest(manifest, self.data)
        workspace = TemplateWorkspace(loader(manifest['name'], self.data))
        workspace.original_hash = manifest['original_hash']
        workspace.set_roles({int(k): v for k, v in manifest['roles'].items()})
        for key, rule in manifest['rules'].items():
            if rule['key'] != key:
                raise FormatError('模板规则标识不匹配。')
            workspace.set_rule(key, rule['scope'], rule['set'], exception=rule['exception'])
        for key, value in manifest['decisions'].items():
            workspace.classify_comment(key, value['classification'], reason=value['reason'], stage=value['stage'], scope=value.get('scope'))
        workspace.set_protections(manifest['protections'])
        workspace.region_standard = deepcopy(manifest.get('region_standard'))
        return workspace

    def application_review(self, model, resolutions=None, protections=None):
        resolutions = resolutions or {}
        manifest = self.manifest
        review = TemplateReview(self.basis_kind, manual_items=[{'key': key, **value} for key, value in manifest['decisions'].items()])
        target_protections = protections if protections is not None else []
        rules, protections, pending = [], [], []
        region_mapping = None
        if self.region_standard:
            from lesson_regions import detect_regions
            region_mapping = detect_regions(model)
            review.region_mapping = region_mapping
            review.blockers.extend(region_mapping.issues)
            review.notices.extend(region_mapping.notices)
            if self.region_standard['heading_policy'] == 'preserve':
                review.notices.extend(h['location'] + '：' + h['reason'] for h in region_mapping.preserved_headings)
            if not region_mapping.tables and not region_mapping.issues:
                review.blockers.append('目标稿没有可靠识别到教学过程区域。')
        if not isinstance(target_protections, list):
            review.blockers.append('当前教案的保护范围记录无效。')
        else:
            for scope in target_protections:
                try:
                    if not isinstance(scope, dict):
                        raise FormatError('当前教案的保护范围记录无效。')
                    resolve_scope(model, scope)  # Existing target scopes must belong to this version.
                    protections.append(deepcopy(scope))
                except (FormatError, KeyError, TypeError) as exc:
                    review.blockers.append('当前教案保护范围：' + str(exc))
        for row in model.paragraphs:
            if row['role'] == 'unknown' and row['text'].strip() and not self.region_standard:
                review.blockers.append(f"目标稿段落 {row['pid']} 待归类，模板覆盖范围未确定。")
        for item in manifest['derived_rules']:
            rule = deepcopy(item); scope = rule['scope']; key = rule['key']
            if self.region_standard and scope.get('kind') == 'lesson_region':
                continue  # Recreate every available group in the new document.
            if key in resolutions:
                scope = {**resolutions[key], **{k: v for k, v in scope.items() if k in {'exclude_lesson_regions', 'lesson_heading_policy'}}}
            elif scope['kind'] == 'role':
                if not any(p['role'] == scope['role'] and p['text'].strip() for p in model.paragraphs):
                    continue  # Absent roles are not invented.
            elif scope['kind'] == 'named_column':
                scope = {k: v for k, v in scope.items() if k not in {'column_id', 'revision'}}
            elif scope['kind'] not in {'all', 'page', 'lesson_region'}:
                review.blockers.append(f'{key} 是模板局部范围，请在目标稿重新选择。')
                pending.append({**rule, 'pending_scope': True})
                continue
            try:
                rule['scope'] = _scope(model, scope)
                rules.append(rule)
            except FormatError as exc:
                review.blockers.append(f'{key}：{exc}')
                pending.append({**rule, 'pending_scope': True})
        if self.region_standard:
            from lesson_regions import region_rules
            rules.extend(region_rules(model, self.region_standard['settings'], self.region_standard['heading_policy'], region_mapping))
        for number, protection in enumerate(manifest['protections']):
            key = 'protection:' + str(number)
            scope = resolutions.get(key)
            if scope is None:
                review.blockers.append(f'{key} 保护范围需在目标稿重新确认。')
            else:
                try: protections.append(_scope(model, scope))
                except FormatError as exc: review.blockers.append(str(exc))
            pending.append({'key': key, 'scope': scope or protection, 'set': {}, 'source': '模板保护范围', 'pending_scope': scope is None})
        for key, decision in manifest['decisions'].items():
            if decision['classification'] == 'protection':
                if key not in resolutions:
                    review.blockers.append(f'{key} 保护范围需在目标稿重新确认。')
                else:
                    try: protections.append(_scope(model, resolutions[key]))
                    except FormatError as exc: review.blockers.append(str(exc))
                pending.append({'key': key, 'scope': resolutions.get(key, decision['scope']), 'set': {}, 'source': '旁批保护范围', 'pending_scope': key not in resolutions})
        review.rules = rules
        result = _evaluate(model, rules, protections, review)
        result.rules += pending
        return result

    def export_bundle(self):
        _validate_manifest(self.manifest, self.data)
        output = BytesIO()
        with ZipFile(output, 'w', ZIP_DEFLATED) as archive:
            archive.writestr('template.' + self.extension, self.data)
            archive.writestr('manifest.json', self.manifest_json.encode('utf-8'))
        return output.getvalue()

    @classmethod
    def import_bundle(cls, data, *, loader=EditingSession.load):
        if not isinstance(data, bytes) or len(data) > MAX_FILE_BYTES + 2_000_000:
            raise FormatError('模板包超过支持大小。')
        try:
            with ZipFile(BytesIO(data)) as archive:
                entries = archive.infolist()
                if len(entries) != 2 or len({i.filename for i in entries}) != 2 or sum(i.file_size for i in entries) > MAX_FILE_BYTES + 2_000_000 or any(i.flag_bits & 1 for i in entries):
                    raise FormatError('模板包结构或解压大小无效。')
                if any(i.filename not in {'manifest.json', 'template.doc', 'template.docx'} for i in entries):
                    raise FormatError('模板包包含不支持的路径。')
                meta = archive.getinfo('manifest.json')
                if meta.file_size > 2_000_000:
                    raise FormatError('模板记录过大。')
                manifest = json.loads(archive.read(meta))
                extension = manifest.get('extension') if isinstance(manifest, dict) else None
                if extension not in {'doc', 'docx'}:
                    raise FormatError('模板原格式无效。')
                original = archive.read('template.' + extension)
            _validate_manifest(manifest, original)
            workspace = cls(original, encode(manifest)).workspace(loader=loader)
            recomputed = workspace.baseline(reviewed=manifest['reviewed'])
            if recomputed.manifest != manifest:
                raise FormatError('模板记录与重新提取的实际属性不一致，请重新导入原格式文件。')
            return recomputed
        except FormatError:
            raise
        except (BadZipFile, KeyError, ValueError, TypeError, OverflowError, UnicodeError, RecursionError) as exc:
            raise FormatError('模板包或校准记录损坏，请重新导入原格式文件。') from exc


def _validate_manifest(manifest, data):
    keys = {'schema', 'name', 'extension', 'file_hash', 'original_hash', 'basis_kind', 'roles', 'rules', 'decisions', 'protections', 'derived_rules', 'comments', 'reviewed', 'fingerprint'}
    if not isinstance(manifest, dict) or set(manifest) - {'region_standard'} != keys or manifest['schema'] != SCHEMA:
        raise FormatError('模板记录版本或字段无效。')
    for key in ('file_hash', 'original_hash', 'fingerprint'):
        if not isinstance(manifest[key], str) or not re.fullmatch('[0-9a-f]{64}', manifest[key]):
            raise FormatError('模板文件指纹无效。')
    content = {k: v for k, v in manifest.items() if k != 'fingerprint'}
    if digest(data) != manifest['file_hash'] or digest(encode(content).encode()) != manifest['fingerprint']:
        raise FormatError('模板文件或规则记录已变化，原基准失效。')
    if manifest['basis_kind'] not in {'sample_styles', 'comment_calibrated'} or type(manifest['reviewed']) is not bool or not isinstance(manifest['name'], str) or not manifest['name'].strip() or PurePosixPath(manifest['name']).suffix.lower() != '.' + manifest['extension']:
        raise FormatError('模板依据类型或原格式名称无效。')
    if any(not isinstance(manifest[k], dict) for k in ('roles', 'rules', 'decisions')) or any(not isinstance(manifest[k], list) for k in ('protections', 'derived_rules', 'comments')):
        raise FormatError('模板记录结构无效。')
    if 'region_standard' in manifest:
        from lesson_regions import validate_settings
        standard = manifest['region_standard']
        if not isinstance(standard, dict) or set(standard) != {'settings', 'heading_policy', 'previous_rules'} or not isinstance(standard['previous_rules'], list):
            raise FormatError('教案区域标准记录无效。')
        if validate_settings(standard['settings'], standard['heading_policy']) != standard['settings']:
            raise FormatError('教案区域标准属性无效。')
        for rule in standard['previous_rules']:
            if not isinstance(rule, dict) or not isinstance(rule.get('scope'), dict) or not isinstance(rule.get('key'), str):
                raise FormatError('区域修订前的来源规则无效。')
            validate_values(rule.get('set'))
