"""Synthetic template documents only; no private teaching material."""
from io import BytesIO
import json
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from lxml import etree

from docx_formatting import FormatError, MAIN, NS
from v06_model import build_model, digest
from v06_patch import current_values
from v06_session import EditingSession
from tests.test_v06 import ENV, fake_render, pack, xml_edit
from template_workflow import TemplateWorkspace, TemplateBaseline, TemplateReview, extract_comments


def specimen(texts=("合成普通段落", "另一普通段落"), bold=False):
    doc = Document()
    normal = doc.styles['Normal']
    normal.font.name = '宋体'
    fonts = normal.element.get_or_add_rPr().get_or_add_rFonts()
    for attr in ('ascii', 'hAnsi', 'cs', 'eastAsia'):
        fonts.set(qn('w:' + attr), '宋体')
    normal.font.size = Pt(12)
    size_cs = OxmlElement('w:szCs'); size_cs.set(qn('w:val'), '24')
    normal.element.get_or_add_rPr().append(size_cs)
    normal.font.bold = bold
    for text in texts:
        doc.add_paragraph(text)
    stream = BytesIO(); doc.save(stream)
    return stream.getvalue()


def commented(data, text='正文：宋体小四，不加粗', start=0, end=0, point=False):
    model = build_model(data)
    root = model.root
    paragraphs = root.findall('.//w:p', NS)
    marker = OxmlElement('w:commentRangeStart'); marker.set(qn('w:id'), '0')
    if not point:
        paragraphs[start].insert(0, marker)
        marker = OxmlElement('w:commentRangeEnd'); marker.set(qn('w:id'), '0')
        paragraphs[end].append(marker)
    run = OxmlElement('w:r')
    marker = OxmlElement('w:commentReference'); marker.set(qn('w:id'), '0')
    run.append(marker); paragraphs[end].append(run)
    comments = OxmlElement('w:comments')
    comment = OxmlElement('w:comment'); comment.set(qn('w:id'), '0')
    p = OxmlElement('w:p'); r = OxmlElement('w:r'); t = OxmlElement('w:t'); t.text = text
    r.append(t); p.append(r); comment.append(p); comments.append(comment)
    rels = etree.fromstring(model.parts['word/_rels/document.xml.rels'])
    etree.SubElement(rels, '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship', Id='comments', Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', Target='comments.xml')
    types = etree.fromstring(model.parts['[Content_Types].xml'])
    etree.SubElement(types, '{http://schemas.openxmlformats.org/package/2006/content-types}Override', PartName='/word/comments.xml', ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml')
    return pack(data, {MAIN: etree.tostring(root), 'word/comments.xml': etree.tostring(comments), 'word/_rels/document.xml.rels': etree.tostring(rels), '[Content_Types].xml': etree.tostring(types)})


def workspace(data):
    return TemplateWorkspace(EditingSession.load('synthetic.docx', data, renderer=fake_render))


class TemplateTests(unittest.TestCase):
    def test_no_comments_auto_inherited_styles_and_no_mutation(self):
        source = specimen(); w = workspace(source)
        review = w.review()
        self.assertEqual(review.blockers, [])
        self.assertTrue(review.ready)
        baseline = w.baseline()
        self.assertEqual(baseline.basis_kind, 'sample_styles')
        self.assertEqual(baseline.data, source)
        self.assertEqual(len(w.session.versions), 1)
        body = next(r for r in review.rules if r['scope'].get('role') == 'body')
        self.assertEqual(body['set']['font_east_asia'], '宋体')
        self.assertNotIn('heading_1', {r['scope'].get('role') for r in review.rules})

    def test_actual_calibration_postverify_and_source_preserved(self):
        source = commented(specimen(('1. 合成练习', '合成普通内容'), bold=True))
        w = workspace(source)
        self.assertEqual(w.model().row(1)['role'], 'body')
        self.assertTrue(w.review().differences)
        with self.assertRaises(FormatError): w.baseline(reviewed=True)
        result = w.calibrate(environment=ENV)
        self.assertEqual(w.session.current, result.id)
        self.assertEqual(w.session.original, source)
        self.assertEqual(w.review().differences, [])
        with self.assertRaises(FormatError): w.baseline()
        baseline = w.baseline(reviewed=True)
        self.assertEqual(baseline.data, result.data)
        self.assertEqual(current_values(build_model(result.data), build_model(result.data).scope('all'), 'bold'), [False])

    def test_unknown_clause_and_point_require_explicit_decision(self):
        w = workspace(commented(specimen(), '正文：宋体小四，不加粗；必须填写课程来源'))
        self.assertTrue(w.review().blockers)
        w.classify_comment('0', 'manual', reason='核对外部课程资料', stage='manuscript')
        self.assertTrue(w.review().manual_items)
        self.assertEqual(w.review().manual_items[0]['verification'], 'not_automatically_verified')
        w = workspace(commented(specimen(), point=True))
        comments = w.review().comments
        self.assertTrue(comments[0]['point'])
        self.assertTrue(comments[0]['context'])
        self.assertTrue(w.review().blockers)

    def test_bundle_roundtrip_hash_and_changed_rule_revalidation(self):
        baseline = workspace(specimen()).baseline()
        bundle = baseline.export_bundle()
        loader = lambda name, data: EditingSession.load(name, data, renderer=fake_render)
        restored = TemplateBaseline.import_bundle(bundle, loader=loader)
        self.assertEqual(restored.fingerprint, baseline.fingerprint)
        with ZipFile(BytesIO(bundle)) as z:
            parts = {name: z.read(name) for name in z.namelist()}
        manifest = json.loads(parts['manifest.json'])
        manifest['file_hash'] = '0' * 64
        parts['manifest.json'] = json.dumps(manifest).encode()
        out = BytesIO()
        with ZipFile(out, 'w') as z:
            for name, value in parts.items(): z.writestr(name, value)
        with self.assertRaises(FormatError): TemplateBaseline.import_bundle(out.getvalue(), loader=loader)

    def test_mapping_uses_new_document_roles(self):
        baseline = workspace(specimen()).baseline()
        doc = Document(BytesIO(specimen(('新稿正文',))))
        doc.add_paragraph('新稿标题', style='Heading 1')
        out = BytesIO(); doc.save(out)
        model = build_model(out.getvalue(), 'new')
        review = baseline.application_review(model)
        self.assertFalse(review.blockers)
        self.assertEqual({a['pid'] for a in review.plan.read()['assignments']}, {1})

    def test_cross_run_paragraph_table_and_external_end_marker(self):
        doc = Document(BytesIO(specimen(('ABCD', 'EFGH'))))
        doc.add_table(rows=1, cols=1).cell(0, 0).text = 'IJKL'
        out = BytesIO(); doc.save(out)
        data = commented(out.getvalue(), start=0, end=2)
        def alter(root):
            paragraphs = root.findall('.//w:p', NS)
            first = paragraphs[0]; original = first.find('w:r', NS)
            original.find('w:t', NS).text = 'AB'
            suffix = OxmlElement('w:r'); t = OxmlElement('w:t'); t.text = 'CD'; suffix.append(t)
            first.append(suffix)
            marker = first.find('w:commentRangeStart', NS); first.remove(marker)
            first.insert(first.index(suffix), marker)
            final = paragraphs[2]; marker = final.find('w:commentRangeEnd', NS); final.remove(marker)
            root.find('w:body', NS).append(marker)
        comments = extract_comments(build_model(xml_edit(data, alter)))
        self.assertEqual(comments[0]['spans'], [{'pid': 1, 'start': 2, 'end': 4}, {'pid': 2, 'start': 0, 'end': 4}, {'pid': 3, 'start': 0, 'end': 4}])
        self.assertIn('表 1', comments[0]['context'])
        self.assertTrue(workspace(xml_edit(data, alter)).review().blockers)

    def test_malformed_annotations_never_fall_back_to_sample_styles(self):
        source = commented(specimen())
        def missing_end(root):
            node = root.find('.//w:commentRangeEnd', NS); node.getparent().remove(node)
        for data in (xml_edit(source, missing_end), pack(source, {'word/comments.xml': b'<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'})):
            review = workspace(data).review()
            self.assertEqual(review.basis_kind, 'comment_calibrated')
            self.assertTrue(review.blockers)
            self.assertFalse(review.ready)

    def test_mixed_styles_resolve_only_explicit_properties(self):
        def change(root):
            r = root.findall('.//w:p', NS)[1].find('w:r', NS)
            rpr = OxmlElement('w:rPr'); bold = OxmlElement('w:b'); rpr.append(bold); r.insert(0, rpr)
        w = workspace(xml_edit(specimen(), change))
        self.assertTrue(any('bold' in b for b in w.review().blockers))
        w.set_rule('style:body', w.model().scope('role', role='body'), {'bold': False})
        self.assertFalse(w.review().blockers)
        self.assertTrue(w.review().differences)
        w.calibrate(environment=ENV)
        baseline = w.baseline(reviewed=True)
        self.assertEqual(baseline.rules[0]['set']['font_east_asia'], '宋体')
        self.assertEqual(baseline.rules[0]['set']['bold'], False)

    def test_unknown_roles_and_theme_fonts_remain_pending(self):
        w = workspace(specimen(('1. 合成练习',)))
        self.assertTrue(w.review().blockers)
        w.set_roles({1: 'body'})
        self.assertTrue(w.review().ready)
        doc = Document(); doc.add_paragraph('未声明有效字体')
        out = BytesIO(); doc.save(out)
        self.assertTrue(any('font_' in x for x in workspace(out.getvalue()).review().blockers))

    def test_conflict_exception_and_protected_mismatch(self):
        w = workspace(commented(specimen()))
        m = w.model()
        w.set_rule('local', m.scope('paragraphs', pids=[2]), {'bold': True})
        self.assertTrue(any('冲突' in x for x in w.review().blockers))
        w.set_rule('local', m.scope('paragraphs', pids=[2]), {'bold': True}, exception=True)
        self.assertFalse(w.review().blockers)
        w.set_protections([m.scope('paragraphs', pids=[2])])
        self.assertTrue(any('保护' in x for x in w.review().blockers))
        with self.assertRaises(FormatError): w.baseline(reviewed=True)
        w.set_protections([])
        w.calibrate(environment=ENV)
        baseline = w.baseline(reviewed=True)
        self.assertTrue(baseline.application_review(build_model(specimen())).blockers)
        target = build_model(specimen())
        self.assertFalse(baseline.application_review(target, {'local': target.scope('paragraphs', pids=[2])}).blockers)

    def test_manual_clause_does_not_drop_understood_formatting(self):
        w = workspace(commented(specimen(bold=True), '正文：宋体小四，不加粗；补填课程来源'))
        w.classify_comment('0', 'manual', reason='正式稿填入实际来源', stage='manuscript')
        self.assertTrue(w.review().differences)
        with self.assertRaises(FormatError): w.baseline(reviewed=True)
        w.calibrate(environment=ENV)
        baseline = w.baseline(reviewed=True)
        self.assertEqual(baseline.rules[0]['set']['bold'], False)
        restored = TemplateBaseline.import_bundle(baseline.export_bundle(), loader=lambda n, d: EditingSession.load(n, d, renderer=fake_render))
        self.assertEqual(restored.fingerprint, baseline.fingerprint)

    def test_template_stage_manual_requirement_still_blocks(self):
        w = workspace(commented(specimen(), '正文：宋体小四；删除第二页'))
        w.classify_comment('0', 'manual', reason='模板本身需要处理', stage='template')
        self.assertTrue(w.review().blockers)
        with self.assertRaises(FormatError): w.baseline(reviewed=True)

    def test_explicit_protection_preserved_and_requires_target_mapping(self):
        w = workspace(commented(specimen(), '整页保持不变'))
        w.classify_comment('0', 'protection', reason='已确认全部内容属于保护页', stage='template', scope=w.model().scope('all'))
        w.set_rule('body', w.model().scope('role', role='body'), {'bold': False})
        baseline = w.baseline(reviewed=True)
        target = build_model(specimen())
        self.assertTrue(baseline.application_review(target).blockers)
        self.assertFalse(baseline.application_review(target, {'comment:0': target.scope('all')}).blockers)

    def test_recomputed_checksum_cannot_bypass_actual_property_validation(self):
        baseline = workspace(commented(specimen())).baseline(reviewed=True)
        manifest = baseline.manifest
        manifest['derived_rules'][0]['set']['bold'] = True
        manifest['fingerprint'] = digest(json.dumps({k:v for k,v in manifest.items() if k != 'fingerprint'}, ensure_ascii=False, sort_keys=True, separators=(',', ':')).encode())
        forged = TemplateBaseline(baseline.data, json.dumps(manifest)).export_bundle()
        with self.assertRaises(FormatError):
            TemplateBaseline.import_bundle(forged, loader=lambda n,d: EditingSession.load(n,d,renderer=fake_render))

    def test_rule_change_changes_fingerprint_without_mutating_old_baseline(self):
        w = workspace(specimen())
        old = w.baseline()
        w.set_rule('style:body', w.model().scope('role', role='body'), {'bold': True})
        with self.assertRaises(FormatError): w.baseline()
        w.calibrate(environment=ENV)
        new = w.baseline(reviewed=True)
        self.assertNotEqual(old.fingerprint, new.fingerprint)
        self.assertEqual(old.data, w.session.original)

    def test_structural_heading_requirement_is_not_visual_formatting(self):
        w = workspace(commented(specimen(), '一级标题：宋体四号，加粗，并应用Word标题样式'))
        self.assertTrue(w.review().blockers)
        with self.assertRaises(FormatError): w.baseline(reviewed=True)

    def test_paragraph_defaults_are_extracted(self):
        w = workspace(specimen())
        values = w.review().rules[0]['set']
        self.assertIn('space_after', values)
        self.assertIn('line_spacing', values)

    def test_passive_modern_metadata_supported_but_replies_blocked(self):
        data = commented(specimen())
        parts = build_model(data).parts
        w14 = '{http://schemas.microsoft.com/office/word/2010/wordml}'
        w15 = '{http://schemas.microsoft.com/office/word/2012/wordml}'
        comments = etree.fromstring(parts['word/comments.xml'])
        comments.find('.//w:p', NS).set(w14 + 'paraId', '00000001')
        extended = etree.Element(w15 + 'commentsEx')
        info = etree.SubElement(extended, w15 + 'commentEx', {w15 + 'paraId': '00000001', w15 + 'done': '0'})
        people = etree.Element(w15 + 'people')
        person = etree.SubElement(people, w15 + 'person', {w15 + 'author': 'Synthetic'})
        etree.SubElement(person, w15 + 'presenceInfo', {w15 + 'providerId': 'None', w15 + 'userId': 'Synthetic'})
        rels = etree.fromstring(parts['word/_rels/document.xml.rels'])
        types = etree.fromstring(parts['[Content_Types].xml'])
        for name in ('commentsExtended', 'people'):
            etree.SubElement(rels, '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship', Id=name, Type='http://schemas.microsoft.com/office/2011/relationships/' + name, Target=name + '.xml')
            etree.SubElement(types, '{http://schemas.openxmlformats.org/package/2006/content-types}Override', PartName='/word/' + name + '.xml', ContentType='application/vnd.ms-word.' + name + '+xml')
        changed = {'word/comments.xml': etree.tostring(comments), 'word/commentsExtended.xml': etree.tostring(extended), 'word/people.xml': etree.tostring(people), 'word/_rels/document.xml.rels': etree.tostring(rels), '[Content_Types].xml': etree.tostring(types)}
        w = workspace(pack(data, changed))
        self.assertTrue(w.review().ready)
        info.set(w15 + 'paraIdParent', '00000002')
        changed['word/commentsExtended.xml'] = etree.tostring(extended)
        self.assertTrue(workspace(pack(data, changed)).review().blockers)
        info.attrib.pop(w15 + 'paraIdParent'); info.set(w15 + 'paraId', '00000002')
        changed['word/commentsExtended.xml'] = etree.tostring(extended)
        self.assertTrue(workspace(pack(data, changed)).review().blockers)

    def test_native_baseline_exports_actual_doc_bytes_not_analysis_snapshot(self):
        from v06_session import Version
        native = b'synthetic-original-format-bytes'
        snapshot = specimen()
        def load(name, data):
            return EditingSession(name, data, {'original': Version('original', name, data, None, analysis_data=snapshot)}, renderer=fake_render)
        w = TemplateWorkspace(load('synthetic.doc', native))
        baseline = w.baseline()
        self.assertEqual(baseline.data, native)
        with ZipFile(BytesIO(baseline.export_bundle())) as z:
            self.assertEqual(z.read('template.doc'), native)
        restored = TemplateBaseline.import_bundle(baseline.export_bundle(), loader=load)
        self.assertEqual(restored.data, native)

    def test_failed_postverification_rolls_back_registered_version_and_cache(self):
        w = workspace(commented(specimen(bold=True)))
        review = w.review()
        before = (w.session.next_number, dict(w.session.jobs), dict(w.session.renders), dict(w.session.environment))
        failed = TemplateReview('comment_calibrated', blockers=['synthetic post-save failure'])
        with patch.object(w, 'review', side_effect=[review, failed]):
            with self.assertRaises(FormatError): w.calibrate(environment=ENV)
        self.assertEqual(list(w.session.versions), ['original'])
        self.assertEqual(w.session.current, 'original')
        self.assertEqual((w.session.next_number, w.session.jobs, w.session.renders, w.session.environment), before)

    def test_reset_resolution_restores_pending_requirement(self):
        w = workspace(commented(specimen(), point=True))
        w.set_rule('comment:0', w.model().scope('role', role='body'), {'bold': False})
        self.assertTrue(w.review().ready)
        w.reset_rule('comment:0')
        self.assertTrue(w.review().blockers)

    def test_named_column_id_is_transferred_by_label(self):
        doc = Document(BytesIO(specimen()))
        table = doc.add_table(rows=1, cols=2)
        table.cell(0, 0).text = '合成栏目'; table.cell(0, 1).text = '栏目正文'
        out = BytesIO(); doc.save(out)
        w = workspace(commented(out.getvalue(), '栏目规则待明确'))
        model = w.model(); column = next(c for c in model.columns if c['label'] == '合成栏目')
        w.set_roles({pid: 'label' for pid in column['label_pids']})
        w.set_rule('comment:0', model.scope('named_column', column_id=column['id']), {'bold': False})
        baseline = w.baseline(reviewed=True)
        self.assertEqual(baseline.rules[0]['scope']['label'], '合成栏目')
        target = build_model(out.getvalue(), role_overrides={pid: 'label' for pid in column['label_pids']})
        self.assertFalse(baseline.application_review(target).blockers)

    def test_explicit_rgb_template_color_is_transferred(self):
        from docx.shared import RGBColor
        doc = Document(BytesIO(specimen()))
        doc.styles['Normal'].font.color.rgb = RGBColor.from_string('CC1122')
        out = BytesIO(); doc.save(out)
        baseline = workspace(out.getvalue()).baseline()
        self.assertEqual(baseline.rules[0]['set']['color'], '#CC1122')
        review = baseline.application_review(build_model(specimen()))
        self.assertTrue(any(d['property'] == 'color' for d in review.differences))

    def test_matching_protected_text_allows_repair_of_unprotected_text(self):
        doc = Document(BytesIO(specimen(('AB',))))
        doc.paragraphs[0].add_run('CD').bold = True
        out = BytesIO(); doc.save(out)
        w = workspace(commented(out.getvalue()))
        w.set_protections([w.model().scope('selection', spans=[{'pid': 1, 'start': 0, 'end': 2}])])
        self.assertFalse(w.review().blockers)
        self.assertTrue(w.review().differences)
        w.calibrate(environment=ENV)
        self.assertTrue(w.review().ready)

    def test_example_clause_never_discards_later_requirements(self):
        for delimiter in ('；', '，', '。', '\n'):
            w = workspace(commented(specimen(), '正文：宋体小四，不加粗；例如：普通段落' + delimiter + '必须删除第二页'))
            with self.subTest(delimiter=delimiter):
                review = w.review()
                self.assertTrue(review.blockers)
                self.assertIn('删除第二页', review.comments[0]['unparsed'])
                with self.assertRaises(FormatError): w.baseline(reviewed=True)

    def test_application_retains_existing_target_protections(self):
        baseline = workspace(commented(specimen())).baseline(reviewed=True)
        target = build_model(specimen(bold=True), 'target')
        protected = target.scope('paragraphs', pids=[1])
        review = baseline.application_review(target, protections=[protected])
        self.assertTrue(any('保护' in b for b in review.blockers))
        self.assertIsNone(review.plan)
        matching = build_model(specimen(), 'matching')
        review = baseline.application_review(matching, protections=[matching.scope('paragraphs', pids=[1])])
        self.assertFalse(review.blockers)
        self.assertEqual(review.plan.read()['protections'], [matching.scope('paragraphs', pids=[1])])
        stale = baseline.application_review(matching, protections=[protected])
        self.assertTrue(stale.blockers)

    def test_noop_fully_protected_application_retains_resolved_protections(self):
        w = workspace(commented(specimen()))
        w.set_protections([w.model().scope('all')])
        baseline = w.baseline(reviewed=True)
        target = build_model(specimen(), 'target')
        protected = target.scope('all')
        review = baseline.application_review(target, resolutions={'protection:0': protected})
        self.assertTrue(review.ready)
        self.assertFalse(review.blockers)
        self.assertIsNone(review.plan)
        self.assertEqual(review.protections, [protected])


if __name__ == '__main__':
    unittest.main()
