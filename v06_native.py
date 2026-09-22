"""DOC-native editing. OOXML snapshots are read-only indexes, never artifacts."""
import base64
from collections import Counter
from io import BytesIO
import json
import re
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree as E

from docx_formatting import FormatError, MAIN, NS, _xml
from prd_formatting import PAGE_PROPERTIES, PARAGRAPH_PROPERTIES
from v06_model import build_model, digest
from v06_office import native_office
from v06_patch import _check_plan, _semantic_tree, _verify_properties

PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
W = NS['w']
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
NATIVE_VERSION = 'native-doc-1'


def snapshot_package(flat):
    if len(flat) > 100 * 1024 * 1024:
        raise FormatError('原格式结构快照超过处理上限。')
    root = _xml(flat)
    if root.tag != '{' + PKG + '}package':
        raise FormatError('原格式结构快照无效。')
    parts = {}; types = E.Element('{' + CT + '}Types', nsmap={None: CT})
    for part in root:
        name = part.get('{' + PKG + '}name', '').lstrip('/')
        if not name or name in parts or '..' in name.split('/') or '\\' in name:
            raise FormatError('原格式结构快照部件路径无效。')
        xml = part.find('{' + PKG + '}xmlData')
        binary = part.find('{' + PKG + '}binaryData')
        if xml is not None and len(xml) == 1:
            parts[name] = E.tostring(xml[0], encoding='UTF-8', xml_declaration=True)
        elif binary is not None:
            parts[name] = base64.b64decode(''.join((binary.text or '').split()), validate=True)
        else:
            raise FormatError('原格式结构快照含不支持的部件。')
        E.SubElement(types, '{' + CT + '}Override', PartName='/' + name, ContentType=part.get('{' + PKG + '}contentType', ''))
    parts['[Content_Types].xml'] = E.tostring(types, encoding='UTF-8', xml_declaration=True)
    output = BytesIO()
    with ZipFile(output, 'w', ZIP_DEFLATED) as archive:
        for name, raw in parts.items(): archive.writestr(name, raw)
    data = output.getvalue()
    build_model(data)
    return data


def _character_map(row):
    raw = row['text']
    # Remove exactly one paragraph terminator, retaining manual line breaks.
    body = raw[:-2] if raw.endswith('\r\x07') else raw[:-1] if raw.endswith('\r') else raw
    if 'characters' in row:
        entries = row['characters']; chars = []; offset = 0; previous = row['start']
        if ''.join(c['text'] for c in entries) != raw:
            raise FormatError('原格式逐字符索引与段落文字不一致。')
        for entry in entries:
            start, end, text = entry['start'], entry['end'], entry['text']
            if start != previous or end < start or end > row['end']:
                raise FormatError('原格式逐字符位置不连续。')
            previous = end
            visible = text[:max(0, len(body) - offset)]
            offset += len(text)
            if not visible or visible == '\x01': continue
            if len(visible.encode('utf-16-le')) // 2 != end - start:
                return chars, False
            cursor = start
            for char in visible:
                width = len(char.encode('utf-16-le')) // 2
                chars.append((cursor, cursor + width, '\n' if char in '\r\x0b\x0c' else char, char))
                cursor += width
        return chars, previous == row['end']
    cursor = row['start']; chars = []
    for char in body:
        width = len(char.encode('utf-16-le')) // 2
        if char != '\x01':  # Inline object marker, never a formatting target.
            chars.append((cursor, cursor + width, '\n' if char in '\r\x0b\x0c' else char, char))
        cursor += width
    # Field codes may occupy hidden native offsets. Such paragraphs can still
    # receive paragraph formatting, but never guess character offsets in them.
    reliable = row['end'] - cursor in {0, 1, 2}
    return chars, reliable


def validate_mapping(model, rows):
    if len(rows) != len(model.paragraphs):
        raise FormatError('原格式段落映射不一致，已停止修改。')
    previous = -1
    for paragraph, row in zip(model.paragraphs, rows):
        if type(row['start']) is not int or type(row['end']) is not int or row['start'] < previous or row['end'] < row['start']:
            raise FormatError('原格式位置映射无效。')
        chars, reliable = _character_map(row)
        if ''.join(c[2] for c in chars) != paragraph['text']:
            raise FormatError('原格式文字与结构索引不一致，已停止修改。')
        if not reliable and not paragraph['field']:
            raise FormatError('原格式字符位置无法可靠映射。')
        previous = row['end']


def inspect_native_doc(data):
    office = native_office()
    if office is None:
        raise FormatError('保留 DOC 原格式编辑需要本机 WPS；未将原稿转换为其他格式。')
    _, flat, rows = office.inspect_doc(data)
    snapshot = snapshot_package(flat)
    validate_mapping(build_model(snapshot), rows)
    return snapshot, rows


def native_assignments(model, plan, rows):
    frozen = _check_plan(model, plan)
    validate_mapping(model, rows)
    result = []
    for assignment in frozen['assignments']:
        prop = assignment['property']; value = assignment['value']
        if prop in PAGE_PROPERTIES:
            result.append({'property': prop, 'value': value}); continue
        row = rows[assignment['pid'] - 1]
        if prop in PARAGRAPH_PROPERTIES:
            result.append({'property': prop, 'value': value, **{k: row[k] for k in ('start', 'end', 'text')}}); continue
        chars, reliable = _character_map(row)
        if not reliable:
            raise FormatError('此段包含不可映射的原生字段位置，未执行字符修改。')
        # Break at inline objects and any hidden gaps so font setters cannot
        # reach protected object markers between the selected ordinary text.
        groups = []
        for start, end, _, raw in chars[assignment['start']:assignment['end']]:
            if groups and groups[-1]['end'] == start:
                groups[-1]['end'] = end; groups[-1]['text'] += raw
            else:
                groups.append({'start': start, 'end': end, 'text': raw})
        result.extend({'property': prop, 'value': value, **group} for group in groups)
    return result


def _normalized_xml(name, raw, parts=None):
    root = _xml(raw)
    if name == 'word/commentsExtended.xml' and parts is not None:
        # Preserve reply associations even when the editor regenerates IDs.
        comments = _xml(parts['word/comments.xml'])
        ids = {}
        for comment in comments:
            for index, paragraph in enumerate(comment.findall('.//w:p', NS)):
                ident = paragraph.get('{' + W14 + '}paraId')
                if ident:
                    if ident in ids: raise FormatError('批注段落标识不唯一。')
                    ids[ident] = comment.get('{' + W + '}id', '') + ':' + str(index)
        w15 = 'http://schemas.microsoft.com/office/word/2012/wordml'
        for node in root:
            for local in ('paraId', 'paraIdParent'):
                attr = '{' + w15 + '}' + local
                if attr in node.attrib:
                    if node.get(attr) not in ids: raise FormatError('批注关联无法核对。')
                    node.set(attr, ids[node.get(attr)])
    volatile = {'{' + W14 + '}paraId', '{' + W14 + '}textId'} | {
        '{' + W + '}' + local for local in ('rsidR', 'rsidRPr', 'rsidP', 'rsidRDefault', 'rsidSect')}
    for node in root.iter():
        for attr in volatile: node.attrib.pop(attr, None)
    if name == MAIN:
        # Word gives character-unit indents precedence over their twip fallback.
        # A font-size edit recalculates the fallback without changing the indent.
        for indent in root.findall('.//w:ind', NS):
            for local in ('firstLine', 'hanging', 'left', 'right', 'start', 'end'):
                chars = indent.get('{' + W + '}' + local + 'Chars')
                if chars is not None and re.fullmatch(r'-?\d+', chars):
                    indent.attrib.pop('{' + W + '}' + local, None)
        # A row exception equal to the table's explicit margins is redundant.
        # Preserve different overrides, cell margins and all other properties.
        for table in root.findall('.//w:tbl', NS):
            inherited = table.find('w:tblPr/w:tblCellMar', NS)
            if inherited is None: continue
            for exception in table.findall('w:tr/w:tblPrEx', NS):
                margin = exception.find('w:tblCellMar', NS)
                if margin is not None and E.tostring(margin, method='c14n') == E.tostring(inherited, method='c14n'):
                    exception.remove(margin)
                    if not len(exception) and not exception.attrib: exception.getparent().remove(exception)
    # Saving in the native editor updates bookkeeping. Retain every content,
    # comment, relationship, font and layout property outside this exact list.
    remove = []
    if name == 'word/settings.xml':
        remove = root.findall('w:rsids', NS)
    elif name == 'docProps/app.xml':
        remove = [c for c in root if E.QName(c).localname in {'TotalTime', 'Pages'}]
    elif name == 'docProps/core.xml':
        remove = [c for c in root if c.tag in {
            '{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}lastModifiedBy',
            '{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}revision',
            '{http://purl.org/dc/terms/}modified'}]
    elif name == 'docProps/custom.xml':
        remove = [c for c in root if c.get('name') == 'ICV' and len(c) == 1 and re.fullmatch(r'[0-9A-F]{32}_\d+', c[0].text or '')]
    for node in remove: root.remove(node)
    return E.tostring(root, method='c14n')


def _ole_payload(raw):
    import olefile
    with olefile.OleFileIO(BytesIO(raw)) as ole:
        # CFB directory timestamps/sector placement can change on a native
        # save. Every actual stream byte and storage class must remain equal.
        items = [('root', ole.root.clsid)]
        for path in sorted(ole.listdir(streams=True, storages=True)):
            entry = ole.direntries[ole._find(path)]
            items.append((path, entry.entry_type, entry.clsid,
                          digest(ole.openstream(path).read()) if entry.entry_type == 2 else None))
        return items


def _font_table_equal(before, after, assignments):
    requested = {a['value'] for a in assignments if a['property'] in {'font_east_asia', 'font_western'}}
    def entries(raw):
        root = _xml(raw); fonts = Counter()
        for node in root:
            name = node.get('{' + W + '}name')
            if node.tag != '{' + W + '}font' or not name:
                raise FormatError('字体声明无法可靠核对。')
            fonts[(name, E.tostring(node, method='c14n'))] += 1
        return fonts
    old, new = entries(before), entries(after)
    # Native editors register newly requested fonts and reorder the table.
    # Existing font metadata and all unrequested font declarations stay fixed.
    return not (old - new) and all(name in requested for name, _ in (new - old))


def verify_native_result(model, snapshot, plan):
    frozen = _check_plan(model, plan); result = build_model(snapshot)
    if model.parts.keys() != result.parts.keys():
        raise FormatError('原格式保全失败：文档部件发生增删。')
    metadata = []
    for name, before in model.parts.items():
        after = result.parts[name]
        if name == MAIN or before == after: continue
        try:
            if name.startswith('word/embeddings/') and before.startswith(bytes.fromhex('D0CF11E0A1B11AE1')):
                equal = _ole_payload(before) == _ole_payload(after)
            elif name == 'word/fontTable.xml':
                equal = _font_table_equal(before, after, frozen['assignments'])
            elif name.endswith(('.xml', '.rels')):
                equal = _normalized_xml(name, before, model.parts) == _normalized_xml(name, after, result.parts)
            else:
                equal = False
        except Exception as exc:
            raise FormatError('原格式保全失败：无法核对嵌入对象或部件。') from exc
        if not equal:
            raise FormatError('原格式保全失败：非目标文档部件发生变化：' + name)
        if name.startswith('docProps/') or name == 'word/settings.xml': metadata.append(name)
    assignments = frozen['assignments']
    if _semantic_tree(_normalized_xml(MAIN, model.parts[MAIN]), assignments) != _semantic_tree(_normalized_xml(MAIN, result.parts[MAIN]), assignments):
        raise FormatError('原格式保全失败：内容、对象、范围外或非目标格式发生变化。')
    _verify_properties(snapshot, assignments)
    return {'structure': 'passed', 'properties': 'passed', 'preservation': 'passed',
            'preservation_profile': 'native-snapshot-and-ole-streams', 'save_metadata_parts': metadata,
            'render': 'not_run', 'visual_review': 'pending', 'inventory': result.inventory,
            'client_compatibility': 'current_native_engine'}


def edit_native_doc(data, model, plan, rows):
    office = native_office()
    if office is None:
        raise FormatError('本机原格式编辑器不可用；原稿和已有版本保留。')
    request = {'input_hash': digest(data), 'assignments': native_assignments(model, plan, rows)}
    output, flat, mapping = office.inspect_doc(data, request)
    snapshot = snapshot_package(flat)
    validate_mapping(build_model(snapshot), mapping)
    checks = verify_native_result(model, snapshot, plan)
    return output, snapshot, mapping, checks
