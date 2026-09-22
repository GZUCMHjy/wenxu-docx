"""Synthetic group geometry, original-media preservation and import-loss gates."""
from io import BytesIO
import struct
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from docx import Document
from lxml import etree as E

from docx_formatting import FormatError, MAIN, _xml
from v06_import import NS, compare_rendered_text, repair_imported_picture_groups
from v06_plan import build_plan, intent
from v06_session import EditingSession


def synthetic_wmf(text):
    """Small original vector fixture; no bytes taken from a private document."""
    def record(function, payload=b""):
        payload += b"\0" * (len(payload) % 2)
        return struct.pack("<IH", (len(payload) + 6) // 2, function) + payload
    font = struct.pack("<hhhhhBBBBBBBB32s", -80, 0, 0, 0, 400, 0, 0, 0, 0, 0, 0, 0, 0, b"DejaVu Serif")
    raw = text.encode("ascii")
    records = [record(0x0103, struct.pack("<H", 8)),  # anisotropic map mode
               record(0x020B, struct.pack("<hh", 0, 0)),
               record(0x020C, struct.pack("<hh", 120, 600)),
               record(0x0102, struct.pack("<H", 1)),
               record(0x02FB, font), record(0x012D, struct.pack("<H", 0)),
               record(0x0209, struct.pack("<I", 0)),
               record(0x0521, struct.pack("<H", len(raw)) + raw + b"\0" * (len(raw) % 2) + struct.pack("<hh", 0, 0)),
               record(0)]
    body = b"".join(records)
    return struct.pack("<HHHIHIH", 1, 9, 0x0300, (18 + len(body)) // 2, 1, max(map(len, records)) // 2, 0) + body


FORMULAS = ("x>y", "c>0", "xc>yc", "x>y", "c<0", "xc<yc")


def grouped_sample():
    doc = Document()
    doc.add_paragraph("Synthetic formula group")
    buffer = BytesIO(); doc.save(buffer)
    with ZipFile(BytesIO(buffer.getvalue())) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    root = _xml(parts[MAIN])
    run = E.SubElement(root.find("w:body/w:p", NS), "{"+NS["w"]+"}r")
    namespace_attrs = " ".join(f'xmlns:{key}="{value}"' for key, value in NS.items())
    drawing = E.fromstring(f'''<w:drawing {namespace_attrs}>
      <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="5" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
        <wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:posOffset>914400</wp:posOffset></wp:positionV>
        <wp:extent cx="2926080" cy="640080"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>
        <wp:docPr id="50" name="Synthetic group"/>
        <a:graphic><a:graphicData uri="{NS['wpg']}"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="2926080" cy="640080"/><a:chOff x="0" y="0"/><a:chExt cx="2926080" cy="640080"/></a:xfrm>
        </wpg:grpSpPr></wpg:wgp></a:graphicData></a:graphic>
      </wp:anchor></w:drawing>''')
    alternate = E.SubElement(run, "{"+NS["mc"]+"}AlternateContent")
    choice = E.SubElement(alternate, "{"+NS["mc"]+"}Choice", Requires="wpg")
    choice.append(drawing)
    E.SubElement(alternate, "{"+NS["mc"]+"}Fallback")
    group = drawing.find(".//wpg:wgp", NS)
    rels = _xml(parts['word/_rels/document.xml.rels'])
    types = _xml(parts['[Content_Types].xml'])
    E.SubElement(types, "{http://schemas.openxmlformats.org/package/2006/content-types}Default", Extension="wmf", ContentType="image/x-wmf")
    for i, formula in enumerate(FORMULAS):
        picture = E.fromstring(f'''<pic:pic {namespace_attrs} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <pic:nvPicPr><pic:cNvPr id="{51+i}" name="Synthetic formula {i+1}"/><pic:cNvPicPr/></pic:nvPicPr>
          <pic:blipFill><a:blip r:embed="formula{i}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
          <pic:spPr><a:xfrm><a:off x="{(i%3)*1005840}" y="{(i//3)*365760}"/><a:ext cx="914400" cy="274320"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="0"><a:noFill/></a:ln></pic:spPr>
        </pic:pic>''')
        group.append(picture)
        target = f'media/formula{i}.wmf'
        parts['word/'+target] = synthetic_wmf(formula)
        E.SubElement(rels, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship", Id=f'formula{i}', Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", Target=target)
    parts.update({MAIN: E.tostring(root), 'word/_rels/document.xml.rels': E.tostring(rels), '[Content_Types].xml': E.tostring(types)})
    return package(parts)


def package(parts):
    out = BytesIO()
    with ZipFile(out, 'w') as z:
        for name, data in parts.items(): z.writestr(name, data)
    return out.getvalue()


def alter(data, mutate):
    with ZipFile(BytesIO(data)) as z: parts = {n: z.read(n) for n in z.namelist()}
    root = _xml(parts[MAIN]); mutate(root); parts[MAIN] = E.tostring(root)
    return package(parts)


def rendered(text):
    from PIL import Image
    buf = BytesIO(); Image.new('RGB', (4, 4), 'white').save(buf, format='PNG')
    return {'text': text, 'pages': 1, 'images': [buf.getvalue()], 'pdf': b'test-double'}


class ImportTests(unittest.TestCase):
    def test_group_repair_preserves_media_parts_and_body(self):
        source = grouped_sample()
        output, repairs = repair_imported_picture_groups(source)
        self.assertEqual(repairs, [{'part': MAIN, 'kind': 'floating_picture_group', 'pictures': 6, 'media_bytes_preserved': True}])
        with ZipFile(BytesIO(source)) as old, ZipFile(BytesIO(output)) as new:
            self.assertEqual(old.namelist(), new.namelist())
            for name in old.namelist():
                if name != MAIN: self.assertEqual(old.read(name), new.read(name))
            root = _xml(new.read(MAIN))
            self.assertEqual(root.findall('.//wpg:wgp', NS), [])
            self.assertEqual(len(root.findall('.//wp:anchor', NS)), 6)
            self.assertEqual(root.find('.//w:t', NS).text, 'Synthetic formula group')
            ids = [e.get('id') for e in root.findall('.//wp:docPr', NS)]
            self.assertEqual(len(set(ids)), 6)
            for i, anchor in enumerate(root.findall('.//wp:anchor', NS)):
                self.assertEqual(int(anchor.find('wp:positionH/wp:posOffset', NS).text), (i%3)*1005840)
                self.assertEqual(int(anchor.find('wp:positionV/wp:posOffset', NS).text), 914400+(i//3)*365760)
                self.assertEqual(anchor.find('.//pic:spPr/a:xfrm/a:off', NS).attrib, {'x':'0','y':'0'})
        self.assertEqual(repair_imported_picture_groups(output), (output, []))

    def test_child_coordinate_translation_and_scaling(self):
        def mutate(root):
            group = root.find('.//wpg:grpSpPr/a:xfrm', NS)
            group.find('a:chOff', NS).attrib.update({'x':'100', 'y':'200'})
            group.find('a:chExt', NS).attrib.update({'cx':'1463040', 'cy':'320040'})
        output, _ = repair_imported_picture_groups(alter(grouped_sample(), mutate))
        with ZipFile(BytesIO(output)) as z: root = _xml(z.read(MAIN))
        anchor = root.find('.//wp:anchor', NS)
        self.assertEqual(anchor.find('wp:positionH/wp:posOffset', NS).text, '-200')
        self.assertEqual(anchor.find('wp:positionV/wp:posOffset', NS).text, '914000')
        self.assertEqual(anchor.find('wp:extent', NS).attrib, {'cx':'1828800','cy':'548640'})

    def test_unsupported_group_geometry_fails_closed(self):
        mutations = [lambda r: r.find('.//wpg:grpSpPr/a:xfrm', NS).set('rot', '5400000'),
                     lambda r: r.find('.//wpg:grpSpPr/a:xfrm/a:chExt', NS).set('cx', '0'),
                     lambda r: r.find('.//pic:spPr/a:xfrm', NS).set('flipH', '1'),
                     lambda r: E.SubElement(r.find('.//wpg:wgp', NS), '{'+NS['wpg']+'}wgp'),
                     lambda r: setattr(r.find('.//wp:wrapNone', NS), 'tag', '{'+NS['wp']+'}wrapSquare'),
                     lambda r: E.SubElement(r.find('.//wp:anchor', NS), '{http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing}sizeRelH')]
        for mutate in mutations:
            with self.subTest(mutation=mutate), self.assertRaises(FormatError):
                repair_imported_picture_groups(alter(grouped_sample(), mutate))

    def test_uploaded_docx_is_not_rewritten(self):
        source = grouped_sample()
        s = EditingSession.load('synthetic.docx', source, inspector=lambda *a: self.fail('DOCX must not be inspected'))
        self.assertEqual(s.versions['original'].data, source)
        self.assertEqual(s.original, source)

    def test_doc_import_retains_actual_bytes_for_both_previews(self):
        seen = []
        def render(data, extension):
            seen.append((data, extension))
            return rendered('synthetic')
        snapshot = grouped_sample()
        with patch('v06_session.validate_doc', return_value={}):
            session = EditingSession.load('synthetic.doc', b'original', inspector=lambda data: (snapshot, []), renderer=render)
        self.assertEqual(seen, [(b'original', 'doc'), (b'original', 'doc')])
        self.assertEqual(session.versions['original'].data, b'original')
        self.assertEqual(session.versions['original'].model_data, snapshot)
        self.assertEqual(session.conversion['status'], 'original_format')
        self.assertTrue(session.conversion['byte_identical'])
        self.assertTrue(session.conversion['page_check']['pixels_equal'])
        self.assertEqual(session.conversion['blocking_issues'], [])

    def test_unstable_original_render_cannot_be_accepted(self):
        views = iter([rendered('body x>y'), rendered('body')])
        with patch('v06_session.validate_doc', return_value={}):
            session = EditingSession.load('synthetic.doc', b'original', inspector=lambda data: (grouped_sample(), []), renderer=lambda *args: next(views))
        self.assertEqual(session.conversion['text_check']['missing_count'], 3)
        session.conversion['visual_review'] = 'confirmed'
        model = session.model()
        plan = build_plan(model, [intent(model.scope('all'), {'font_size': 12})]).plan
        with self.assertRaisesRegex(FormatError, '两次独立渲染不一致'):
            session.execute(model, plan, environment={'font_families': []})
        self.assertEqual(list(session.versions), ['original'])

    def test_text_extraction_order_is_not_treated_as_character_loss(self):
        evidence = compare_rendered_text(rendered('left\nright'), rendered('right left、'))
        self.assertFalse(evidence['sequence_equal'])
        self.assertEqual(evidence['missing_count'], 0)
        self.assertEqual(evidence['added_count'], 1)


if __name__ == '__main__':
    unittest.main()
