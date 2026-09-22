"""Synthetic native offsets and independent preservation fault injection."""
from copy import deepcopy
from io import BytesIO
import unittest
from unittest.mock import patch

from docx import Document
from docx.oxml.ns import qn
from lxml import etree as E

from docx_formatting import FormatError, MAIN, NS
from test_v06 import sample, pack, xml_edit
from test_v06_import import rendered
from v06_model import build_model
from v06_native import _character_map, _normalized_xml, _font_table_equal, native_assignments, verify_native_result, W14
from v06_patch import patch_document
from v06_plan import build_plan, intent
from v06_session import EditingSession


class NativeTests(unittest.TestCase):
    def test_utf16_and_inline_object_offsets(self):
        row = {'start': 10, 'end': 16, 'text': '甲😀\x01乙\r'}
        chars, reliable = _character_map(row)
        self.assertTrue(reliable)
        self.assertEqual(chars, [(10, 11, '甲', '甲'), (11, 13, '😀', '😀'), (14, 15, '乙', '乙')])

    def test_native_ole_and_hidden_anchor_positions_split_targets(self):
        doc = Document(); doc.add_paragraph('甲乙丙'); buf = BytesIO(); doc.save(buf)
        model = build_model(buf.getvalue())
        row = {'start': 0, 'end': 31, 'text': '甲\x01乙丙\r', 'characters': [
            {'start': 0, 'end': 1, 'text': '甲'}, {'start': 1, 'end': 27, 'text': '\x01'},
            {'start': 27, 'end': 28, 'text': '乙'}, {'start': 28, 'end': 29, 'text': ''},
            {'start': 29, 'end': 30, 'text': '丙'}, {'start': 30, 'end': 31, 'text': '\r'}]}
        plan = build_plan(model, [intent(model.scope('all'), {'color': '#FF0000'})]).plan
        assignments = native_assignments(model, plan, [row])
        self.assertEqual([(a['start'], a['end'], a['text']) for a in assignments], [(0, 1, '甲'), (27, 28, '乙'), (29, 30, '丙')])
        corrupt = deepcopy(row); corrupt['characters'][2]['start'] = 26
        with self.assertRaises(FormatError): native_assignments(model, plan, [corrupt])

    def test_cell_end_and_manual_breaks(self):
        chars, ok = _character_map({'start': 2, 'end': 6, 'text': '甲\r乙\r\x07'})
        self.assertTrue(ok); self.assertEqual(''.join(c[2] for c in chars), '甲\n乙')

    def test_native_preservation_rejects_text_objects_and_untargeted_properties(self):
        model = build_model(sample())
        plan = build_plan(model, [intent(model.scope('selection', spans=[{'pid': 2, 'start': 0, 'end': 2}]), {'color': '#FF0000'})]).plan
        output, _ = patch_document(model, plan)
        self.assertEqual(verify_native_result(model, output, plan)['preservation'], 'passed')
        corruptions = [pack(output, {'word/embeddings/synthetic.bin': b'changed'}),
                       xml_edit(output, lambda root: setattr(root.find('.//w:t', NS), 'text', 'changed')),
                       xml_edit(output, lambda root: E.SubElement(root.find('.//w:pPr', NS), qn('w:keepNext')))]
        for corrupted in corruptions:
            with self.subTest(), self.assertRaises(FormatError): verify_native_result(model, corrupted, plan)

    def test_comment_id_regeneration_preserves_association_not_just_count(self):
        w15 = 'http://schemas.microsoft.com/office/word/2012/wordml'
        def parts(first, second):
            comments = f'<w:comments xmlns:w="{NS["w"]}" xmlns:w14="{W14}"><w:comment w:id="1"><w:p w14:paraId="{first}"/></w:comment><w:comment w:id="2"><w:p w14:paraId="{second}"/></w:comment></w:comments>'.encode()
            extended = f'<w15:commentsEx xmlns:w15="{w15}"><w15:commentEx w15:paraId="{second}" w15:paraIdParent="{first}" w15:done="0"/></w15:commentsEx>'.encode()
            return {'word/comments.xml': comments, 'word/commentsExtended.xml': extended}
        before, after = parts('AAA', 'BBB'), parts('CCC', 'DDD')
        name = 'word/commentsExtended.xml'
        normalized = _normalized_xml(name, before[name], before)
        self.assertEqual(normalized, _normalized_xml(name, after[name], after))
        damaged = after[name].replace(b'paraIdParent="CCC"', b'paraIdParent="DDD"')
        self.assertNotEqual(normalized, _normalized_xml(name, damaged, after))

    def test_only_save_metadata_can_be_normalized(self):
        ns = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
        a = f'<cp:coreProperties xmlns:cp="{ns}"><cp:revision>1</cp:revision><cp:keywords>keep</cp:keywords></cp:coreProperties>'.encode()
        b = a.replace(b'>1<', b'>2<')
        self.assertEqual(_normalized_xml('docProps/core.xml', a), _normalized_xml('docProps/core.xml', b))
        self.assertNotEqual(_normalized_xml('docProps/core.xml', a), _normalized_xml('docProps/core.xml', b.replace(b'keep', b'lost')))

    def test_character_indents_preserve_units_and_explicit_point_indents(self):
        a = f'<w:document xmlns:w="{NS["w"]}"><w:ind w:firstLine="480" w:firstLineChars="200"/><w:ind w:left="720"/></w:document>'.encode()
        same = a.replace(b'firstLine="480"', b'firstLine="600"')
        self.assertEqual(_normalized_xml(MAIN, a), _normalized_xml(MAIN, same))
        for damaged in (same.replace(b'Chars="200"', b'Chars="300"'), same.replace(b'left="720"', b'left="360"')):
            self.assertNotEqual(_normalized_xml(MAIN, a), _normalized_xml(MAIN, damaged))

    def test_redundant_row_margins_only(self):
        margin = '<w:tblCellMar><w:left w:type="dxa" w:w="108"/></w:tblCellMar>'
        a = f'<w:document xmlns:w="{NS["w"]}"><w:tbl><w:tblPr>{margin}</w:tblPr><w:tr><w:tc/></w:tr></w:tbl></w:document>'.encode()
        same = a.replace(b'<w:tr>', ('<w:tr><w:tblPrEx>' + margin + '</w:tblPrEx>').encode())
        self.assertEqual(_normalized_xml(MAIN, a), _normalized_xml(MAIN, same))
        damaged = same.replace(b'<w:tblPrEx><w:tblCellMar><w:left w:type="dxa" w:w="108"', b'<w:tblPrEx><w:tblCellMar><w:left w:type="dxa" w:w="144"')
        self.assertNotEqual(_normalized_xml(MAIN, a), _normalized_xml(MAIN, damaged))

    def test_font_table_reordering_and_only_requested_additions(self):
        before = f'<w:fonts xmlns:w="{NS["w"]}"><w:font w:name="A"/><w:font w:name="B"/></w:fonts>'.encode()
        after = before.replace(b'<w:font w:name="A"/><w:font w:name="B"/>', b'<w:font w:name="B"/><w:font w:name="C"/><w:font w:name="A"/>')
        assignment = [{'property': 'font_western', 'value': 'C'}]
        self.assertTrue(_font_table_equal(before, after, assignment))
        self.assertFalse(_font_table_equal(before, after, []))
        self.assertFalse(_font_table_equal(before, after.replace(b'<w:font w:name="A"/>', b''), assignment))

    def test_doc_noop_retains_bytes_and_native_failure_is_atomic(self):
        index = sample(False)
        initial = build_model(index)
        initial_plan = build_plan(initial, [intent(initial.scope('selection', spans=[{'pid': 2, 'start': 0, 'end': 2}]), {'font_size': 15})]).plan
        index, _ = patch_document(initial, initial_plan)
        with patch('v06_session.validate_doc', return_value={}):
            session = EditingSession.load('synthetic.doc', b'original', renderer=lambda *a: rendered('same'), inspector=lambda data: (index, []))
        model = session.model()
        plan = build_plan(model, [intent(model.scope('selection', spans=[{'pid': 2, 'start': 0, 'end': 2}]), {'font_size': 15})]).plan
        with patch('v06_session.edit_native_doc') as native:
            self.assertEqual(session.execute(model, plan, environment={'font_families': []}), (None, False))
            native.assert_not_called()
        change = build_plan(model, [intent(model.scope('all'), {'color': '#FF0000'})]).plan
        with patch('v06_session.edit_native_doc', side_effect=FormatError('native failure')):
            with self.assertRaises(FormatError): session.execute(model, change, environment={'font_families': []})
        self.assertEqual(session.versions['original'].data, b'original')
        self.assertEqual(len(session.versions), 1); self.assertEqual(session.next_number, 1)


if __name__ == '__main__': unittest.main()
