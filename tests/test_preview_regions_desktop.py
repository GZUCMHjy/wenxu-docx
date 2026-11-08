"""Preview-first interaction contracts, using synthetic documents only."""
import json
import unittest
from unittest.mock import patch

from tests import test_desktop as helpers
from lesson_regions import binding, default_settings


class PreviewFirstTests(unittest.TestCase):
    setUpClass = classmethod(helpers.WindowTests.setUpClass.__func__)
    setUp = helpers.WindowTests.setUp
    wait_job = helpers.WindowTests.wait_job
    apply_font = helpers.WindowTests.apply_font
    def test_main_surface_keeps_details_closed(self):
        self.assertFalse(self.window.editor_dock.isVisible())
        self.assertFalse(self.window.scope.isVisible())
        self.assertFalse(self.window.reviewed.isVisible())
        self.assertEqual(self.window.primary.text(), '按金标准整理')

    def test_comparison_uses_single_canvas_and_save_targets_result(self):
        self.window.editor_dock.show()
        self.apply_font()
        result = self.window.state.session.versions['v1'].data
        self.window.compare.setChecked(True)
        self.assertFalse(self.window.original_view.isVisible())
        self.assertIn('原稿', self.window.current_view.title.text())
        self.assertEqual(self.window.primary.text(), '保存修改后')
        target = self.root / 'preview-result.docx'
        with patch('desktop.window.QFileDialog.getSaveFileName', return_value=(str(target), '')):
            self.window.primary.click()
        self.wait_job()
        self.assertEqual(target.read_bytes(), result)
        self.assertEqual(self.source.read_bytes(), self.data)

    def test_cancelled_save_does_not_confirm_or_mark_saved(self):
        self.window.editor_dock.show()
        self.apply_font()
        with patch('desktop.window.QFileDialog.getSaveFileName', return_value=('', '')):
            self.window.primary.click()
        self.assertNotIn('v1', self.window.state.reviewed)
        self.assertNotIn('v1', self.window.state.saved)

    def test_region_definition_is_recorded_without_fake_template_file(self):
        from v06_plan import build_plan, intent
        model = self.window.model
        plan = build_plan(model, [intent(model.scope('all'), {'font_family': '黑体'})]).plan
        version, _ = self.window.state.session.execute(model, plan, environment=helpers.ENV, template_binding=binding(default_settings()))
        recorded = json.loads(version.template_json)
        self.assertEqual(recorded['basis_kind'], 'lesson_regions')
        self.assertIn('definition_hash', recorded)
        self.assertNotIn('file_hash', recorded)

    def open_lesson(self):
        from tests.test_lesson_regions import lesson, data
        self.source.write_bytes(data(lesson()))
        self.window.open_path(self.source)
        self.wait_job()

    def test_region_edit_repeat_and_failure_keep_last_successful_standard(self):
        from v06_patch import current_values
        self.open_lesson()
        self.window.primary.click()
        self.wait_job()
        self.assertEqual(self.errors, [])
        self.assertEqual(self.window.state.session.current, 'v1')
        for region, size in [('headers', 9), ('content', 12), ('activities', 9), ('closing', 10.5)]:
            self.assertEqual(current_values(self.window.model, self.window.model.scope('lesson_region', region=region), 'font_size'), [size])
        self.window.organize(); self.wait_job()
        self.assertEqual(len(self.window.state.session.versions), 2)
        changed = default_settings(); changed['content']['font_size'] = 14
        with patch.object(self.window.state.session, 'execute', side_effect=helpers.FormatError('render failed')):
            self.window.apply_region_standard(changed); self.wait_job()
        self.assertEqual(self.window.region_settings, default_settings())
        self.assertEqual(self.window.state.session.current, 'v1')
        self.window.apply_region_standard(changed); self.wait_job()
        self.assertEqual(self.window.state.session.current, 'v2')
        self.assertEqual(self.window.region_settings, changed)

    def test_template_adjust_cancel_keeps_adopted_baseline(self):
        from template_workflow import TemplateWorkspace
        from desktop.standard_dialog import StandardDialog
        from desktop.template_dialog import TemplateDialog
        self.open_lesson()
        workspace = TemplateWorkspace(helpers.load('template.docx', self.source.read_bytes()))
        workspace.set_region_standard(default_settings())
        workspace.calibrate(environment=helpers.ENV)
        baseline = workspace.baseline(reviewed=True)
        self.window.template_workspace, self.window.template_baseline = workspace, baseline
        self.window.standard_mode = 'template'
        changed = default_settings(); changed['content']['font_size'] = 14
        with patch.object(StandardDialog, 'exec', return_value=1), patch.object(StandardDialog, 'values', return_value=(changed, 'preserve')), patch.object(TemplateDialog, 'exec', return_value=0):
            self.window.adjust_standard(); self.wait_job()
        self.assertEqual(self.errors, [])
        self.assertIs(self.window.template_baseline, baseline)
        self.assertEqual(workspace.region_standard['settings'], default_settings())
        self.assertEqual(self.window.state.session.current, 'original')
