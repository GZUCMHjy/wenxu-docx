"""Synthetic teaching-process tables; no private source documents."""
from io import BytesIO
import unittest

from docx import Document
from docx_formatting import FormatError
from v06_model import build_model
from v06_patch import patch_document, verify_result
from v06_plan import build_plan, intent, resolve_scope
from v06_session import EditingSession
from template_workflow import TemplateBaseline
from tests.test_template_workflow import specimen, workspace, commented
from tests.test_v06 import ENV, fake_render
from lesson_regions import detect_regions, default_settings, review_regions, binding


def lesson(order=('stage', 'content', 'activity', 'intent'), repeat=False):
    doc = Document(BytesIO(specimen(('合成教案标题',))))
    names = {'stage': '教学环节及时间分配', 'content': '教学内容', 'activity': '师生活动', 'intent': '设计意图'}
    table = doc.add_table(rows=4 + int(repeat), cols=len(order))
    for col, key in enumerate(order): table.cell(0, col).text = names[key]
    for row, label in ((1, '新课探究'), (2, '课堂小结'), (3, '作业布置')):
        for col, key in enumerate(order): table.cell(row, col).text = label if key == 'stage' else f'{label}-{key}填写'
    if repeat:
        for col, key in enumerate(order): table.cell(4, col).text = names[key]
    return doc


def data(doc):
    stream = BytesIO(); doc.save(stream); return stream.getvalue()


class LessonRegionTests(unittest.TestCase):
    def test_english_closing_labels_use_special_row_sizes_and_keep_text(self):
        from v06_patch import current_values
        for summary, homework in [('Summary\n（6min）', 'Homework（2 min）'), ('SUMMARY', 'homework')]:
            with self.subTest(summary=summary):
                doc = lesson(); table = doc.tables[0]
                for row, label in ((2, summary), (3, homework)):
                    table.cell(row, 0).text = label
                    table.cell(row, 1).merge(table.cell(row, 3))
                model = build_model(data(doc)); review = review_regions(model)
                self.assertFalse(review.blockers)
                self.assertEqual(len(review.region_mapping.groups['closing']), 6)
                output, _ = patch_document(model, review.plan)
                verify_result(model, output, review.plan)
                after = build_model(output)
                self.assertEqual([p['text'] for p in model.paragraphs], [p['text'] for p in after.paragraphs])
                self.assertEqual(current_values(after, after.scope('lesson_region', region='closing'), 'font_size'), [10.5])
                self.assertEqual(current_values(after, after.scope('lesson_region', region='headers'), 'font_size'), [9])

    def test_english_stage_labels_do_not_relax_ambiguous_merge_guard(self):
        doc = lesson(); table = doc.tables[0]
        table.cell(2, 0).text = 'Summary（6 min）'
        table.cell(1, 0).merge(table.cell(2, 0))
        review = review_regions(build_model(data(doc)))
        self.assertTrue(any('纵向合并行头' in b for b in review.blockers))
        doc = lesson(); table = doc.tables[0]
        table.cell(2, 0).text = 'Homework review'
        table.cell(2, 1).merge(table.cell(2, 3))
        review = review_regions(build_model(data(doc)))
        self.assertTrue(any('跨越不同业务区域' in b for b in review.blockers))
        self.assertFalse(any('没有可靠识别' in b or '没有可转移' in b for b in review.blockers))
        self.assertIsNone(review.plan)

    def test_english_word_in_content_without_stage_column_is_not_a_row_label(self):
        doc = lesson(('content', 'activity', 'intent'))
        doc.tables[0].cell(1, 0).text = 'Summary'
        model = build_model(data(doc)); mapped = detect_regions(model)
        pid = next(c['pids'][0] for c in model.cells if c['row'] == 2 and c['col'] == 1)
        self.assertIn(pid, mapped.groups['content'])
        self.assertNotIn(pid, mapped.groups['headers'])

    def test_semantic_reordered_columns_and_repeated_headers(self):
        model = build_model(data(lesson(('stage', 'intent', 'content', 'activity'), repeat=True)))
        mapped = detect_regions(model)
        self.assertFalse(mapped.issues)
        texts = lambda group: [model.row(pid)['text'] for pid in mapped.groups[group]]
        self.assertEqual(texts('content'), ['新课探究-content填写'])
        self.assertEqual(set(texts('activities')), {'新课探究-activity填写', '新课探究-intent填写'})
        self.assertEqual(len(texts('closing')), 6)
        self.assertEqual(texts('headers').count('教学内容'), 2)

    def test_horizontal_closing_merge_overrides_general_columns(self):
        doc = lesson(); doc.tables[0].cell(2, 1).merge(doc.tables[0].cell(2, 3))
        model = build_model(data(doc)); mapped = detect_regions(model)
        self.assertFalse(mapped.issues)
        merged = next(c for c in model.cells if c['row'] == 3 and c['col'] == 2)
        self.assertTrue(set(merged['pids']) <= set(mapped.groups['closing']))
        review = review_regions(model)
        self.assertFalse(review.blockers)
        output, changes = patch_document(model, review.plan)
        self.assertTrue(changes)
        verify_result(model, output, review.plan)
        self.assertEqual([p['text'] for p in build_model(output).paragraphs], [p['text'] for p in model.paragraphs])

    def test_vertical_merge_across_normal_and_closing_rows_blocks(self):
        doc = lesson(); doc.tables[0].cell(1, 1).merge(doc.tables[0].cell(2, 1))
        review = review_regions(build_model(data(doc)))
        self.assertTrue(review.blockers)
        self.assertIsNone(review.plan)

    def test_content_text_matching_closing_label_does_not_override_stage_column(self):
        doc = lesson(('content', 'activity', 'intent', 'stage'))
        doc.tables[0].cell(1, 0).text = '课堂小结'
        model = build_model(data(doc)); mapped = detect_regions(model)
        pid = next(c['pids'][0] for c in model.cells if c['row'] == 2 and c['col'] == 1)
        self.assertFalse(mapped.issues)
        self.assertIn(pid, mapped.groups['content'])
        self.assertNotIn(pid, mapped.groups['headers'])
        activity = next(c['pids'][0] for c in model.cells if c['row'] == 2 and c['col'] == 2)
        self.assertIn(activity, mapped.groups['activities'])

    def test_vertical_stage_merge_mixing_normal_and_closing_semantics_blocks(self):
        doc = lesson(); doc.tables[0].cell(1, 0).merge(doc.tables[0].cell(2, 0))
        review = review_regions(build_model(data(doc)))
        self.assertTrue(review.blockers)
        self.assertIsNone(review.plan)

    def test_internal_headings_preserved_and_exercises_included(self):
        doc = lesson(); cell = doc.tables[0].cell(1, 1)
        cell.text = '1. 计算下列各题'
        cell.add_paragraph('一、概念形成')
        cell.add_paragraph('显式标题', style='Heading 2')
        model = build_model(data(doc)); mapped = detect_regions(model)
        self.assertEqual(len(mapped.preserved_headings), 2)
        self.assertIn('1. 计算下列各题', [model.row(pid)['text'] for pid in mapped.groups['content']])
        review = review_regions(model, heading_policy='same_as_content')
        self.assertFalse(review.blockers)
        self.assertTrue({h['pid'] for h in mapped.preserved_headings} <= {a['pid'] for a in review.plan.read()['assignments']})

    def test_unrelated_cover_table_untouched(self):
        doc = lesson(); cover = doc.add_table(rows=1, cols=2)
        cover.cell(0, 0).text = '任课教师'; cover.cell(0, 1).text = '合成人名'
        model = build_model(data(doc)); mapped = detect_regions(model)
        self.assertIn(2, mapped.ignored_tables)
        assigned = {a['pid'] for a in review_regions(model).plan.read()['assignments']}
        self.assertTrue(all(p['pid'] not in assigned for p in model.paragraphs if p['table'] == 2))

    def test_same_width_metadata_ignored_but_headerless_continuation_blocks(self):
        doc = lesson(); table = doc.add_table(rows=1, cols=4)
        for i, label in enumerate(('学科', '年级', '任课教师', '日期')): table.cell(0, i).text = label
        mapping = detect_regions(build_model(data(doc)))
        self.assertFalse(mapping.issues)
        self.assertEqual(mapping.ignored_tables, [2])
        for i in range(4): table.cell(0, i).text = '续表正文' + str(i)
        self.assertTrue(review_regions(build_model(data(doc))).blockers)

    def test_settings_cannot_modify_unrequested_properties(self):
        settings = default_settings(); settings['content']['bold'] = True
        with self.assertRaises(FormatError): review_regions(build_model(data(lesson())), settings)

    def test_multiple_tables_remap_independently_and_standalone_special_rows(self):
        doc = lesson()
        other = lesson(('content', 'intent', 'activity'))
        from copy import deepcopy
        doc.element.body.insert(-1, deepcopy(other.tables[0]._tbl))
        special = doc.add_table(rows=2, cols=2)
        special.cell(0, 0).text = '板书设计'; special.cell(0, 1).text = '结构图'
        special.cell(1, 0).text = '作业布置'; special.cell(1, 1).text = '复习任务'
        model = build_model(data(doc)); mapped = detect_regions(model)
        self.assertFalse(mapped.issues)
        self.assertEqual(mapped.tables, [1, 2, 3])
        self.assertTrue(all(pid in mapped.groups['closing'] for c in model.cells if c['table'] == 3 and c['col'] == 2 for pid in c['pids']))

    def test_duplicate_headers_nested_tables_and_normal_cross_region_merge_block(self):
        variants = []
        doc = lesson(); doc.tables[0].cell(0, 3).text = '教学内容'; variants.append(doc)
        doc = lesson(); doc.tables[0].cell(1, 1).add_table(rows=1, cols=1).cell(0, 0).text = '嵌套正文'; variants.append(doc)
        doc = lesson(); doc.tables[0].cell(1, 1).merge(doc.tables[0].cell(1, 2)); variants.append(doc)
        for index, doc in enumerate(variants):
            with self.subTest(index=index):
                review = review_regions(build_model(data(doc)))
                self.assertTrue(review.blockers)
                self.assertIsNone(review.plan)

    def test_portable_scope_uses_new_document_positions(self):
        first = build_model(data(lesson()))
        doc = lesson(('intent', 'activity', 'content', 'stage'))
        doc.paragraphs[0].insert_paragraph_before('新增封面段落')
        second = build_model(data(doc))
        scope = second.scope('lesson_region', region='content', heading_policy='preserve')
        spans, _ = resolve_scope(second, scope)
        self.assertEqual([second.row(s['pid'])['text'] for s in spans], ['新课探究-content填写'])
        self.assertNotEqual([s['pid'] for s in spans], detect_regions(first).groups['content'])

    def test_protected_region_difference_blocks_and_baseline_identity_changes(self):
        model = build_model(data(lesson()))
        protection = model.scope('lesson_region', region='headers')
        review = review_regions(model, protections=[protection])
        self.assertTrue(review.blockers)
        self.assertIsNone(review.plan)
        settings = default_settings(); original = binding(settings)
        settings['headers']['font_size'] = 10
        self.assertNotEqual(binding(settings)['fingerprint'], original['fingerprint'])
        self.assertNotIn('file_hash', original)

    def test_imported_template_revision_is_actual_file_and_roundtrips(self):
        source = data(lesson())
        w = workspace(source)
        candidate = w.clone(); candidate.set_region_standard(default_settings())
        self.assertIsNone(w.region_standard)
        self.assertFalse(candidate.review().blockers)
        changed = candidate.calibrate(environment=ENV)
        self.assertIsNotNone(changed)
        baseline = candidate.baseline(reviewed=True)
        self.assertNotEqual(baseline.data, source)
        self.assertEqual(w.session.current, 'original')
        loader = lambda name, value: EditingSession.load(name, value, renderer=fake_render)
        restored = TemplateBaseline.import_bundle(baseline.export_bundle(), loader=loader)
        self.assertEqual(restored.fingerprint, baseline.fingerprint)
        target = build_model(data(lesson(('content', 'activity', 'stage', 'intent'))))
        review = restored.application_review(target)
        self.assertFalse(review.blockers)
        self.assertTrue(review.differences)
        newer = restored.workspace(loader=loader)
        settings = default_settings(); settings['content']['font_size'] = 14
        newer.set_region_standard(settings); newer.calibrate(environment=ENV)
        self.assertNotEqual(newer.baseline(reviewed=True).fingerprint, baseline.fingerprint)
        self.assertEqual(restored.data, baseline.data)

    def test_new_standard_overrides_old_font_but_preserves_other_requirements(self):
        doc = lesson(); source = data(doc)
        model = build_model(source)
        first = next(p['pid'] for p in model.paragraphs if p['text'] == '新课探究-content填写')
        source = commented(source, '正文：黑体四号，加粗', start=first - 1, end=first - 1)
        w = workspace(source); w.set_region_standard(default_settings())
        self.assertFalse(w.review().blockers)
        plan = w.review().plan.read()
        values = {a['property']: a['value'] for a in plan['assignments'] if a['pid'] == first}
        self.assertEqual(values['font_east_asia'], '宋体')
        self.assertEqual(values['font_size'], 12)
        self.assertEqual(values['bold'], True)
        w.calibrate(environment=ENV)
        baseline = w.baseline(reviewed=True)
        self.assertIn('黑体四号', baseline.manifest['comments'][0]['text'])

    def test_unknown_original_comment_still_blocks_after_region_revision(self):
        w = workspace(commented(data(lesson()), '正文：宋体小四；删除第二页'))
        w.set_region_standard(default_settings())
        self.assertTrue(any('旁批' in b for b in w.review().blockers))
        with self.assertRaises(FormatError): w.baseline(reviewed=True)

    def test_remapping_prior_named_column_keeps_region_exclusion(self):
        w = workspace(data(lesson())); model = w.model(); column = model.columns[0]
        w.set_rule('prior', model.scope('named_column', column_id=column['id']), {'font_size': 14})
        w.set_region_standard(default_settings()); w.calibrate(environment=ENV)
        baseline = w.baseline(reviewed=True)
        target = build_model(data(lesson()))
        self.assertFalse(baseline.application_review(target).blockers)
        selected = target.scope('paragraphs', pids=detect_regions(target).groups['content'])
        self.assertFalse(baseline.application_review(target, resolutions={'prior:outside_regions': selected}).blockers)

    def test_preserved_heading_keeps_existing_explicit_font_requirement(self):
        from docx.shared import Pt
        from v06_patch import current_values
        doc = lesson(); p = doc.tables[0].cell(1, 1).add_paragraph('合成小标题', style='Heading 2')
        p.runs[0].font.size = Pt(16)
        w = workspace(data(doc)); model = w.model()
        pid = next(row['pid'] for row in model.paragraphs if row['text'] == '合成小标题')
        w.set_rule('old-heading-rule', model.scope('paragraphs', pids=[pid]), {'font_size': 14})
        w.set_region_standard(default_settings())
        self.assertTrue(any(a['pid'] == pid and a['property'] == 'font_size' and a['value'] == 14 for a in w.review().differences))
        w.calibrate(environment=ENV)
        self.assertEqual(current_values(w.model(), w.model().scope('paragraphs', pids=[pid]), 'font_size'), [14])
        self.assertTrue(w.review().ready)
        w.set_region_standard(default_settings(), heading_policy='same_as_content')
        self.assertFalse(w.review().blockers)
        w.calibrate(environment=ENV)
        self.assertEqual(current_values(w.model(), w.model().scope('paragraphs', pids=[pid]), 'font_size'), [12])


if __name__ == '__main__':
    unittest.main()
