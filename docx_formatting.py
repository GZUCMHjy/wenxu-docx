"""Change selected properties and paragraph roles; retain other ZIP parts verbatim."""

from copy import deepcopy
from io import BytesIO
import math
from pathlib import PurePosixPath
import re
from zipfile import BadZipFile, ZipFile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from lxml import etree

MAX_FILE_BYTES = 10 * 1024 * 1024
MAX_UNPACKED_BYTES = 60 * 1024 * 1024
MAIN = "word/document.xml"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
PARAGRAPHS = etree.XPath("./w:body//w:p", namespaces=NS)
RUNS = etree.XPath(".//w:r[w:t or w:tab or w:br or w:cr]", namespaces=NS)
DOCX_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
SCOPES = {"all": "全部段落", "body": "正文（非标题）", "title": "文档标题/副标题"}
SCOPES.update({f"heading_{i}": f"{label}级标题" for i, label in enumerate("一二三四五六七八九", 1)})
FONT_NAMES = ("宋体", "黑体")
FONT_ATTRS = ("ascii", "hAnsi", "eastAsia", "cs")
FONT_THEMES = ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme")


class FormatError(ValueError):
    """An actionable input or verification error suitable for the UI."""


def _xml(raw):
    parser = etree.XMLParser(resolve_entities=False, load_dtd=False, no_network=True, remove_blank_text=True)
    root = etree.fromstring(raw, parser)
    if root.getroottree().docinfo.doctype:
        raise FormatError("文件包含不支持的 XML 实体声明。")
    return root


def _package(data):
    if not data or len(data) > MAX_FILE_BYTES:
        raise FormatError("请上传非空 DOCX 文件，大小不超过 10 MB。")
    try:
        with ZipFile(BytesIO(data)) as z:
            entries = z.infolist()
            names = z.namelist()
            if len(entries) > 2000 or sum(i.file_size for i in entries) > MAX_UNPACKED_BYTES:
                raise FormatError("文档解压后过大，请使用较小的文件。")
            if len(names) != len(set(names)) or any(i.flag_bits & 1 for i in entries):
                raise FormatError("文档包含重复部件或加密内容。")
            if not {MAIN, "[Content_Types].xml", "_rels/.rels"} <= set(names):
                raise FormatError("文件不是有效的 DOCX 文档。")
            parts = {name: z.read(name) for name in names}
        for name, raw in parts.items():
            if name.endswith((".xml", ".rels")):
                root = _xml(raw)
                if name.endswith(".rels"):
                    for rel in root:
                        if rel.get("TargetMode") == "External" and not rel.get("Type", "").endswith("/hyperlink"):
                            raise FormatError("暂不支持外链图片、外部模板或外部嵌入对象。")
        types = _xml(parts["[Content_Types].xml"])
        if not any(e.get("PartName") == "/" + MAIN and e.get("ContentType") == DOCX_TYPE for e in types):
            raise FormatError("仅支持普通 DOCX，不支持 DOC、宏文档或模板文件。")
        if any("vbaproject" in name.lower() or name.startswith("_xmlsignatures/") for name in parts):
            raise FormatError("暂不处理带宏或数字签名的文档。")
        root = _xml(parts[MAIN])
        if root.tag != qn("w:document") or root.find("w:body", NS) is None:
            raise FormatError("文档主体格式不受支持。")
        unsupported = {"ins", "del", "moveFrom", "moveTo", "sdt", "altChunk", "txbxContent", "object", "fldSimple", "fldChar", "footnoteReference", "endnoteReference", "commentRangeStart", "commentReference", "customXml"}
        if any(etree.QName(e).localname in unsupported or etree.QName(e).localname.endswith("PrChange") for e in root.iter() if isinstance(e.tag, str)):
            raise FormatError("暂不支持含修订、批注、文本框、域、内容控件或嵌入对象的正文。")
        if root.xpath(".//*[local-name()='oMath' or local-name()='oMathPara']"):
            raise FormatError("暂不支持公式排版。")
        settings = parts.get("word/settings.xml")
        if settings:
            for protection in _xml(settings).findall("w:documentProtection", NS):
                if protection.get(qn("w:enforcement")) in {"1", "true", "on"}:
                    raise FormatError("文档已启用编辑保护，请使用可编辑的副本。")
        return parts
    except FormatError:
        raise
    except (BadZipFile, etree.XMLSyntaxError, KeyError, RuntimeError, ValueError, OSError) as exc:
        raise FormatError("文档已损坏或不是有效 DOCX，请重新选择文件。") from exc


def _validate_options(options):
    if not isinstance(options, dict) or not options or set(options) - {"color", "font_size", "line_spacing", "font_family"}:
        raise FormatError("请至少选择一项受支持的格式修改。")
    if "font_family" in options and options["font_family"] not in FONT_NAMES:
        raise FormatError("字体仅支持宋体或黑体。")
    if "color" in options and (not isinstance(options["color"], str) or not re.fullmatch(r"#[0-9a-fA-F]{6}", options["color"])):
        raise FormatError("字体颜色必须是 #RRGGBB 格式。")
    if "font_size" in options:
        _number(options["font_size"], 6, 72, "字号", half_point=True)
    if "line_spacing" in options:
        spacing = options["line_spacing"]
        if not isinstance(spacing, dict) or set(spacing) != {"mode", "value"} or spacing["mode"] not in {"exact", "multiple"}:
            raise FormatError("行距必须选择固定值或倍数。")
        limits = (12, 72) if spacing["mode"] == "exact" else (1, 3)
        _number(spacing["value"], *limits, "行距", half_point=True)


def _number(value, minimum, maximum, label, half_point=False):
    if type(value) not in (int, float) or not math.isfinite(value) or not minimum <= value <= maximum or (half_point and value * 2 != int(value * 2)):
        raise FormatError(f"{label}须在 {minimum}–{maximum} 之间，步长为 0.5。")


def _paragraphs(root):
    # XML traversal visits merged/nested table paragraphs once, including hyperlink runs.
    return PARAGRAPHS(root)


def _runs(paragraph):
    return RUNS(paragraph)


def _paragraph_roles(root, parts):
    styles_root = _xml(parts["word/styles.xml"]) if "word/styles.xml" in parts else etree.Element("styles")
    styles = {s.get(qn("w:styleId")): s for s in styles_root.findall("w:style", NS) if s.get(qn("w:type")) == "paragraph"}
    default = next((key for key, s in styles.items() if s.get(qn("w:default")) in {"1", "true", "on"}), None)
    defaults = styles_root.find("w:docDefaults/w:pPrDefault/w:pPr/w:outlineLvl", NS)
    rows = []
    for index, p in enumerate(_paragraphs(root)):
        style_ref = p.find("w:pPr/w:pStyle", NS)
        key = style_ref.get(qn("w:val")) if style_ref is not None else default
        chain, seen = [], set()
        while key in styles:
            if key in seen:
                raise FormatError("段落样式存在循环继承，无法可靠识别标题。")
            seen.add(key)
            style = styles[key]
            name = style.find("w:name", NS)
            chain.append((key, name.get(qn("w:val"), key) if name is not None else key, style))
            base = style.find("w:basedOn", NS)
            key = base.get(qn("w:val")) if base is not None else None
        scope, evidence = "body", "无标题标记"
        outlines = [(p.find("w:pPr/w:outlineLvl", NS), "段落大纲级别")]
        outlines += [(style.find("w:pPr/w:outlineLvl", NS), f"样式 {name} 的大纲级别") for _, name, style in chain]
        outlines.append((defaults, "文档默认大纲级别"))
        for outline, origin in outlines:
            if outline is not None:
                value = outline.get(qn("w:val"), "")
                if value not in tuple(str(i) for i in range(10)):
                    raise FormatError("文档包含无效的大纲级别。")
                scope = f"heading_{int(value) + 1}" if value != "9" else "body"
                evidence = origin
                break
        else:
            # Recognize explicit style metadata, never guess from text length or bold.
            for key, name, _ in chain:
                match = re.fullmatch(r"(?:heading|标题)\s*([1-9])", name, re.IGNORECASE) or re.fullmatch(r"Heading([1-9])", key, re.IGNORECASE)
                if match:
                    scope, evidence = f"heading_{match[1]}", f"标题样式 {name}"
                    break
                if name.casefold() in {"title", "subtitle", "标题", "副标题"} or key in {"Title", "Subtitle"}:
                    scope, evidence = "title", f"样式 {name}"
                    break
        rows.append({"index": index + 1, "scope": scope, "kind": SCOPES[scope], "text": "".join(p.itertext(tag=qn("w:t"))), "source": evidence})
    return rows


def inspect_document(docx_bytes):
    """Return roles and evidence for all main-story paragraphs in document order."""
    parts = _package(docx_bytes)
    return _paragraph_roles(_xml(parts[MAIN]), parts)


def _selected_indices(root, parts, scope):
    if not isinstance(scope, str) or scope not in SCOPES:
        raise FormatError("请选择有效的修改范围。")
    selected = set(range(len(_paragraphs(root)))) if scope == "all" else {i for i, row in enumerate(_paragraph_roles(root, parts)) if row["scope"] == scope}
    if not selected:
        raise FormatError(f"没有识别到{SCOPES[scope]}，请检查文档的标题样式或选择其他范围。")
    return selected


def _masked(root, options, selected):
    """Remove ONLY permitted changes, then compare the entire remaining XML tree."""
    root = deepcopy(root)
    for index, p in enumerate(_paragraphs(root)):
        if index not in selected:
            continue
        for r in _runs(p):
            rpr = r.find("w:rPr", NS)
            if rpr is not None:
                tags = (["color"] if "color" in options else []) + (["sz", "szCs"] if "font_size" in options else [])
                for tag in tags:
                    for prop in rpr.findall("w:" + tag, NS):
                        rpr.remove(prop)
                fonts = rpr.find("w:rFonts", NS)
                if fonts is not None and "font_family" in options:
                    for attr in FONT_ATTRS + FONT_THEMES:
                        fonts.attrib.pop(qn("w:" + attr), None)
                    if not len(fonts) and not fonts.attrib:
                        rpr.remove(fonts)
                if not len(rpr) and not rpr.attrib:
                    r.remove(rpr)
        ppr = p.find("w:pPr", NS)
        if ppr is not None:
            spacing = ppr.find("w:spacing", NS)
            if spacing is not None and "line_spacing" in options:
                for attr in ("line", "lineRule"):
                    spacing.attrib.pop(qn("w:" + attr), None)
                if not len(spacing) and not spacing.attrib:
                    ppr.remove(spacing)
            if not len(ppr) and not ppr.attrib:
                p.remove(ppr)
    return etree.tostring(root, method="c14n")


def validate_result(source, output, options, scope="all"):
    _validate_options(options)
    before, after = _package(source), _package(output)
    if before.keys() != after.keys() or any(before[n] != after[n] for n in before if n != MAIN):
        raise FormatError("校验失败：文档附件、关系或其他部件发生了变化。")
    a, b = _xml(before[MAIN]), _xml(after[MAIN])
    selected = _selected_indices(a, before, scope)
    if _masked(a, options, selected) != _masked(b, options, selected):
        raise FormatError("校验失败：正文或未选择的属性发生了变化。")
    for index, p in enumerate(_paragraphs(b)):
        if index not in selected:
            continue
        if "line_spacing" in options:
            target = options["line_spacing"]
            exact = target["mode"] == "exact"
            s = p.find("w:pPr/w:spacing", NS)
            if s is None or s.get(qn("w:line")) != str(round(target["value"] * (20 if exact else 240))) or s.get(qn("w:lineRule"), "auto") != ("exact" if exact else "auto"):
                raise FormatError("校验失败：行距没有达到所选值。")
        for r in _runs(p):
            if "font_family" in options:
                fonts = r.find("w:rPr/w:rFonts", NS)
                if fonts is None or any(fonts.get(qn("w:" + attr)) != options["font_family"] for attr in FONT_ATTRS) or any(qn("w:" + attr) in fonts.attrib for attr in FONT_THEMES):
                    raise FormatError("校验失败：字体没有达到所选值。")
            if "font_size" in options:
                for tag in ("sz", "szCs"):
                    s = r.find("w:rPr/w:" + tag, NS)
                    if s is None or s.get(qn("w:val")) != str(round(options["font_size"] * 2)):
                        raise FormatError("校验失败：字号没有达到所选值。")
            if "color" in options:
                c = r.find("w:rPr/w:color", NS)
                if c is None or c.get(qn("w:val")) != options["color"][1:].upper() or any(qn("w:" + attr) in c.attrib for attr in ("themeColor", "themeTint", "themeShade")):
                    raise FormatError("校验失败：字体颜色没有达到所选值。")
    try:
        Document(BytesIO(output))
    except Exception as exc:
        raise FormatError("校验失败：生成的 DOCX 无法重新读取。") from exc


def apply_formatting(docx_bytes, options, scope="all"):
    _validate_options(options)
    parts = _package(docx_bytes)
    try:
        doc = Document(BytesIO(docx_bytes))
    except Exception as exc:
        raise FormatError("无法读取此 DOCX，请检查文档是否完整。") from exc
    summary = {"changed_paragraphs": 0, "changed_runs": 0}
    selected = _selected_indices(_xml(parts[MAIN]), parts, scope)
    for index, p in enumerate(_paragraphs(doc.element)):
        if index not in selected:
            continue
        if "line_spacing" in options:
            old = etree.tostring(p.find(qn("w:pPr"))) if p.find(qn("w:pPr")) is not None else b""
            spacing = options["line_spacing"]
            Paragraph(p, doc).paragraph_format.line_spacing = Pt(spacing["value"]) if spacing["mode"] == "exact" else float(spacing["value"])
            if etree.tostring(p.find(qn("w:pPr"))) != old:
                summary["changed_paragraphs"] += 1
        for r in _runs(p):
            old = etree.tostring(r)
            font = Run(r, doc).font
            if "font_family" in options:
                fonts = r.get_or_add_rPr().get_or_add_rFonts()
                for attr in FONT_ATTRS:
                    fonts.set(qn("w:" + attr), options["font_family"])
                for attr in FONT_THEMES:
                    fonts.attrib.pop(qn("w:" + attr), None)
            if "font_size" in options:
                font.size = Pt(options["font_size"])
                rpr = r.get_or_add_rPr()
                cs = rpr.find(qn("w:szCs"))
                if cs is None:
                    cs = OxmlElement("w:szCs")
                    rpr.find(qn("w:sz")).addnext(cs)
                cs.set(qn("w:val"), str(round(options["font_size"] * 2)))
            if "color" in options:
                font.color.rgb = RGBColor.from_string(options["color"][1:].upper())
                color = r.find("w:rPr/w:color", NS)
                for attr in ("themeColor", "themeTint", "themeShade"):
                    color.attrib.pop(qn("w:" + attr), None)
            if etree.tostring(r) != old:
                summary["changed_runs"] += 1
    if not any(summary.values()):
        return docx_bytes, summary
    output = BytesIO()
    # Only the main document is replaced. Images, headers, footers, styles and relationships
    # retain their original bytes, including parts python-docx doesn't model.
    with ZipFile(BytesIO(docx_bytes)) as src, ZipFile(output, "w") as dst:
        dst.comment = src.comment
        for entry in src.infolist():
            raw = etree.tostring(doc.element, xml_declaration=True, encoding="UTF-8", standalone=True) if entry.filename == MAIN else src.read(entry.filename)
            dst.writestr(entry, raw)
    result = output.getvalue()
    validate_result(docx_bytes, result, options, scope=scope)
    return result, summary


def output_filename(name):
    name = PurePosixPath(name.replace("\\", "/")).name
    stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name.rsplit(".", 1)[0]).strip(" .") or "文档"
    version = re.search(r"_v(\d+)$", stem)
    number = int(version.group(1)) + 1 if version else 1
    if version:
        stem = stem[:version.start()]
    return f"{stem[:100]}_v{number}.docx"
