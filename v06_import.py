"""Compatibility repair at the DOC import boundary, before the editing baseline."""
from collections import Counter
from copy import deepcopy
from fractions import Fraction
from io import BytesIO
from zipfile import ZipFile

from lxml import etree

from docx_formatting import FormatError, _package, _xml

IMPORT_VERSION = "v06-import-4-original-format"
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}


def _tag(name):
    prefix, local = name.split(":")
    return "{" + NS[prefix] + "}" + local


def _required(node, path):
    found = node.find(path, NS)
    if found is None:
        raise ValueError("missing " + path)
    return found


def _pair(node, keys):
    return tuple(int(node.attrib[key]) for key in keys)


def _axis_aligned(transform):
    if int(transform.get("rot", "0")) or any(transform.get(k, "0") not in {"0", "false"} for k in ("flipH", "flipV")):
        raise ValueError("rotated or flipped group picture")


def _expand_group(group, next_id):
    # LO can read the WMF pictures in a DOC but lose the entire group when
    # reopening its own DOCX export. Keep the actual vector media and mapping;
    # only flatten supported picture-only, non-wrapping floating groups.
    graphic_data = group.getparent()
    graphic = graphic_data.getparent()
    anchor = graphic.getparent()
    drawing = anchor.getparent()
    if (graphic_data.tag != _tag("a:graphicData") or graphic.tag != _tag("a:graphic")
            or anchor.tag != _tag("wp:anchor") or drawing.tag != _tag("w:drawing")
            or len(graphic_data) != 1 or len(drawing) != 1):
        raise ValueError("unsupported group owner")
    if anchor.get("simplePos", "0") not in {"0", "false"}:
        raise ValueError("simple position")
    if any(etree.QName(child).localname in {"sizeRelH", "sizeRelV"} for child in anchor):
        raise ValueError("relative group sizing")
    _required(anchor, "wp:wrapNone")
    extent = _pair(_required(anchor, "wp:extent"), ("cx", "cy"))
    position = [_required(anchor, "wp:positionH/wp:posOffset"), _required(anchor, "wp:positionV/wp:posOffset")]
    origin = [int(p.text) for p in position]
    effects = anchor.find("wp:effectExtent", NS)
    if effects is not None and any(int(v) for v in effects.attrib.values()):
        raise ValueError("group effects")
    properties = _required(group, "wpg:grpSpPr")
    transform = _required(properties, "a:xfrm")
    if len(properties) != 1:
        raise ValueError("group effects or fills")
    _axis_aligned(transform)
    if _pair(_required(transform, "a:off"), ("x", "y")) != (0, 0):
        raise ValueError("nonzero outer group origin")
    if any(n <= 0 for n in _pair(_required(transform, "a:ext"), ("cx", "cy"))):
        raise ValueError("invalid group extent")
    child_origin = _pair(_required(transform, "a:chOff"), ("x", "y"))
    child_extent = _pair(_required(transform, "a:chExt"), ("cx", "cy"))
    if any(n <= 0 for n in (*extent, *child_extent)):
        raise ValueError("invalid extent")
    scale = [Fraction(e, c) for e, c in zip(extent, child_extent)]
    pictures = group.findall("pic:pic", NS)
    if not pictures or any(c.tag not in {_tag("wpg:cNvGrpSpPr"), _tag("wpg:grpSpPr"), _tag("pic:pic")} for c in group):
        raise ValueError("group is not picture-only")
    owner = drawing
    if drawing.getparent().tag == _tag("mc:Choice"):
        choice = drawing.getparent()
        owner = choice.getparent()
        if (owner.tag != _tag("mc:AlternateContent") or len(choice) != 1
                or len(owner.findall("mc:Choice", NS)) != 1):
            raise ValueError("ambiguous alternate content")
    if owner.getparent().tag != _tag("w:r"):
        raise ValueError("unsupported anchor location")
    replacements = []
    for picture in pictures:
        child_transform = _required(picture, "pic:spPr/a:xfrm")
        _axis_aligned(child_transform)
        offset = _pair(_required(child_transform, "a:off"), ("x", "y"))
        size = _pair(_required(child_transform, "a:ext"), ("cx", "cy"))
        dimensions = [round(s * k) for s, k in zip(size, scale)]
        if any(d <= 0 for d in dimensions):
            raise ValueError("invalid picture size")
        # Build a standard anchored picture, retaining media relationships,
        # crop/stretch, shape styling, anchor paragraph and stacking order.
        new_drawing = deepcopy(drawing)
        new_anchor = new_drawing[0]
        for axis, delta in zip(("H", "V"), (round((offset[i] - child_origin[i]) * scale[i]) + origin[i] for i in range(2))):
            _required(new_anchor, "wp:position" + axis + "/wp:posOffset").text = str(delta)
        _required(new_anchor, "wp:extent").attrib.update({"cx": str(dimensions[0]), "cy": str(dimensions[1])})
        new_picture = deepcopy(picture)
        _required(new_picture, "pic:spPr/a:xfrm/a:off").attrib.update({"x": "0", "y": "0"})
        _required(new_picture, "pic:spPr/a:xfrm/a:ext").attrib.update({"cx": str(dimensions[0]), "cy": str(dimensions[1])})
        picture_properties = _required(new_picture, "pic:nvPicPr/pic:cNvPr")
        next_id += 1
        picture_properties.set("id", str(next_id))
        doc_properties = _required(new_anchor, "wp:docPr")
        doc_properties.set("id", str(next_id))
        doc_properties.set("name", picture_properties.get("name", "Imported picture"))
        new_graphic_data = _required(new_anchor, "a:graphic/a:graphicData")
        new_graphic_data.clear()
        new_graphic_data.set("uri", NS["pic"])
        new_graphic_data.append(new_picture)
        replacements.append(new_drawing)
    return owner, replacements, next_id


def repair_imported_picture_groups(data):
    """Only for a freshly converted DOC; never silently rewrite uploaded DOCX."""
    parts = _package(data, preservation_profile=True)
    changed, repairs = {}, []
    for name, raw in parts.items():
        if not name.startswith("word/") or not name.endswith(".xml"):
            continue
        root = _xml(raw)
        groups = root.findall(".//wpg:wgp", NS)
        if not groups:
            continue
        try:
            next_id = max([int(e.get("id", "0")) for e in root.iter() if e.tag in {_tag("wp:docPr"), _tag("pic:cNvPr")}] + [0])
            for group in groups:
                owner, replacements, next_id = _expand_group(group, next_id)
                parent, index = owner.getparent(), owner.getparent().index(owner)
                for offset, replacement in enumerate(replacements):
                    parent.insert(index + offset, replacement)
                parent.remove(owner)
                repairs.append({"part": name, "kind": "floating_picture_group", "pictures": len(replacements), "media_bytes_preserved": True})
        except (ValueError, KeyError, TypeError, AttributeError) as exc:
            raise FormatError("DOC 含暂不能保证原位置的组合图形，已停止转换。请用原编辑器另存为 DOCX 后导入；原稿未修改。") from exc
        changed[name] = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
    if not changed:
        return data, repairs
    output = BytesIO()
    with ZipFile(BytesIO(data)) as old, ZipFile(output, "w") as new:
        new.comment = old.comment
        for entry in old.infolist():
            new.writestr(entry, changed.get(entry.filename, parts[entry.filename]))
    return output.getvalue(), repairs


def compare_rendered_text(before, after):
    """Extraction evidence, not a claim of complete visual/content equivalence."""
    old, new = ("".join(render["text"].split()) for render in (before, after))
    missing = Counter(old) - Counter(new)
    added = Counter(new) - Counter(old)
    return {"sequence_equal": old == new, "missing_count": sum(missing.values()), "added_count": sum(added.values()),
            "missing_characters": dict(missing), "added_characters": dict(added)}


def compare_rendered_pages(before, after):
    """Compare actual rasters, including pictures with no extractable text."""
    from PIL import Image, ImageChops
    old, new = before['images'], after['images']
    count_equal = before['pages'] == after['pages']
    complete = len(old) == before['pages'] and len(new) == after['pages']
    checks = []
    for index, (a, b) in enumerate(zip(old, new), 1):
        with Image.open(BytesIO(a)) as left, Image.open(BytesIO(b)) as right:
            left, right = left.convert('RGB'), right.convert('RGB')
            same_size = left.size == right.size
            box, changed = None, None
            if same_size:
                delta = ImageChops.difference(left, right)
                r, g, b = delta.split()
                mask = ImageChops.lighter(ImageChops.lighter(r, g), b)
                box = mask.getbbox()
                changed = left.width * left.height - mask.histogram()[0]
            checks.append({'page': index, 'identical': same_size and changed == 0,
                           'same_size': same_size, 'changed_pixels': changed,
                           'difference_bbox': list(box) if box else None})
    identical = [p['page'] for p in checks if p['identical']]
    differing = [p['page'] for p in checks if not p['identical']]
    differing += list(range(min(len(old), len(new)) + 1, max(before['pages'], after['pages']) + 1))
    return {'page_count_equal': count_equal, 'complete': complete,
            'pixels_equal': complete and count_equal and len(identical) == before['pages'],
            'identical_pages': identical, 'differing_pages': differing, 'pages': checks}
