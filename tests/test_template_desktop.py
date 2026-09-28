"""The modest desktop additions expose real ranges and preserve versions."""
from io import BytesIO
import json
from pathlib import Path
import tempfile
import time
import unittest
from unittest.mock import patch

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt
from PySide6.QtCore import Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication
from PySide6.QtWidgets import QMessageBox, QDialog, QDialogButtonBox

from desktop.controller import DocumentState
from desktop.window import MainWindow
from tests.test_desktop import load
from tests.test_v06 import ENV
from v06_patch import current_values
from v06_plan import build_plan, intent


def plain_template(size=12, comment=None):
    doc = Document()
    normal = doc.styles['Normal']
    normal.font.size = Pt(size)
    normal.font.bold = False
    size_cs = OxmlElement('w:szCs'); size_cs.set(qn('w:val'), str(size * 2))
    normal.element.get_or_add_rPr().append(size_cs)
    fonts = normal.element.get_or_add_rPr().get_or_add_rFonts()
    fonts.clear()
    for key in ('eastAsia', 'ascii', 'hAnsi', 'cs'):
        fonts.set(qn('w:' + key), '宋体')
    p = doc.add_paragraph('这是合成教学内容，用于验证模板格式。')
    doc.add_paragraph('另一段合成正文。')
    if comment:
        doc.add_comment(p.runs, comment, author='Synthetic')
    out = BytesIO(); doc.save(out)
    return out.getvalue()


class TemplateWindowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(); self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source = self.root / 'teaching.docx'; self.source.write_bytes(plain_template(15))
        self.template = self.root / 'standard.docx'; self.template.write_bytes(plain_template())
        self.window = MainWindow(loader=lambda p: DocumentState.open(p, loader=load), office_check=lambda: ENV,
                                 template_session_loader=load)
        self.errors = []; self.window.error = self.errors.append
        self.window.confirm_discard = lambda: True
        self.window.show(); self.addCleanup(self.window.close)
        self.window.open_path(self.source); self.wait()

    def wait(self):
        deadline = time.monotonic() + 15
        while self.window.job and time.monotonic() < deadline:
            self.app.processEvents(); QTest.qWait(10)
        self.assertIsNone(self.window.job)
        self.app.processEvents()

    def test_classification_changes_range_without_rewriting_original(self):
        w = self.window
        self.assertTrue(hasattr(w, 'role_panel'), 'classification panel is missing')
        tree = w.role_panel.tree
        body = next(tree.topLevelItem(i) for i in range(tree.topLevelItemCount()) if tree.topLevelItem(i).data(0, Qt.ItemDataRole.UserRole) == 'body')
        self.assertEqual(body.childCount(), 2)
        body.child(0).setSelected(True)
        w.role_panel.role.setCurrentIndex(w.role_panel.role.findData('heading_1'))
        w.role_panel.correct.click()
        self.assertEqual(w.model.row(1)['role'], 'heading_1')
        self.assertEqual(w.state.session.original, self.source.read_bytes())
        self.assertIn('一级标题', [w.scope.itemText(i) for i in range(w.scope.count())])

    def test_no_comment_template_auto_adoption_apply_and_v2_binding(self):
        w = self.window
        self.assertTrue(hasattr(w, 'load_template_path'), 'template import is missing')
        w.load_template_path(self.template); self.wait()
        self.assertEqual(self.errors, [])
        self.assertIsNotNone(w.template_baseline)
        self.assertEqual(w.template_baseline.basis_kind, 'sample_styles')
        w.apply_template(); self.wait()
        self.assertEqual(self.errors, [])
        self.assertEqual(w.state.session.current, 'v1')
        self.assertEqual(current_values(w.model, w.model.scope('role', role='body'), 'font_size'), [12])
        first = json.loads(w.state.session.versions['v1'].template_json)
        self.assertEqual(first['file_hash'], w.template_baseline.file_hash)
        w.format_panel.controls['bold'].setCurrentIndex(w.format_panel.controls['bold'].findData(True))
        w.apply(); self.wait()
        self.assertEqual(w.state.session.current, 'v2')
        self.assertEqual(json.loads(w.state.session.versions['v2'].template_json), first)
        w.undo(); self.assertEqual(w.state.session.current, 'v1')
        self.assertEqual(self.source.read_bytes(), plain_data := w.state.session.original)
        self.assertNotEqual(plain_data, w.state.session.versions['v1'].data)

    def test_failed_template_import_keeps_previous_baseline(self):
        w = self.window
        self.assertTrue(hasattr(w, 'load_template_path'), 'template import is missing')
        w.load_template_path(self.template); self.wait()
        baseline = w.template_baseline
        self.assertIsNotNone(baseline)
        w.load_template_path(self.root / 'missing.docx'); self.wait()
        self.assertIs(w.template_baseline, baseline)
        self.assertEqual(w.state.session.current, 'original')

    def test_comment_dialog_calibrates_file_before_adoption_and_bundle_reimport(self):
        from desktop.template_dialog import TemplateDialog
        from template_workflow import TemplateWorkspace
        original = plain_template(15, '正文：宋体小四，不加粗')
        workspace = TemplateWorkspace(load('annotated.docx', original))
        dialog = TemplateDialog(workspace, self.window)
        self.addCleanup(dialog.close)
        self.assertFalse(dialog.adopt_button.isEnabled())
        self.assertTrue(dialog.calibrate_button.isEnabled())
        dialog.calibrate_button.click()
        deadline = time.monotonic() + 15
        while dialog.job and time.monotonic() < deadline:
            self.app.processEvents(); QTest.qWait(10)
        self.assertIsNone(dialog.job)
        self.assertEqual(dialog.review.blockers, [])
        self.assertEqual(dialog.review.differences, [])
        self.assertFalse(dialog.adopt_button.isEnabled())
        dialog.reviewed.setChecked(True)
        self.assertTrue(dialog.adopt_button.isEnabled())
        dialog.adopt_button.click()
        self.assertIsNotNone(dialog.baseline)
        self.assertNotEqual(dialog.baseline.data, original)
        bundle = self.root / 'gold.wenxu-template'
        bundle.write_bytes(dialog.baseline.export_bundle())
        self.window.load_template_path(bundle); self.wait()
        self.assertEqual(self.errors, [])
        self.assertEqual(self.window.template_baseline.file_hash, dialog.baseline.file_hash)

    def test_unknown_comment_cannot_be_adopted_from_dialog(self):
        from desktop.template_dialog import TemplateDialog
        from template_workflow import TemplateWorkspace
        workspace = TemplateWorkspace(load('unknown.docx', plain_template(12, '此处参考另一本资料的全部规定')))
        dialog = TemplateDialog(workspace, self.window)
        self.addCleanup(dialog.close)
        dialog.reviewed.setChecked(True)
        self.assertFalse(dialog.adopt_button.isEnabled())
        self.assertFalse(dialog.calibrate_button.isEnabled())

    def test_followup_edit_retains_protection_until_explicit_release(self):
        w = self.window
        model = w.model
        plan = build_plan(model, [intent(model.scope('all'), {'bold': True})],
                          [model.scope('paragraphs', pids=[1])]).plan
        version, _ = w.state.session.execute(model, plan, environment=ENV)
        w.applied((version, w.state.session.preview(version.id)))
        w.format_panel.controls['bold'].setCurrentIndex(w.format_panel.controls['bold'].findData(True))
        w.apply(); self.wait()
        self.assertEqual(current_values(w.model, w.model.scope('paragraphs', pids=[1]), 'bold'), [False])
        self.assertEqual(w.state.session.current, 'v1')
        self.assertTrue(w.state.protected_scopes())
        w.state.release_protection(0)
        w.format_panel.controls['bold'].setCurrentIndex(w.format_panel.controls['bold'].findData(True))
        w.apply(); self.wait()
        self.assertEqual(current_values(w.model, w.model.scope('paragraphs', pids=[1]), 'bold'), [True])
        self.assertEqual(w.state.session.current, 'v2')
        w.undo()
        self.assertTrue(json.loads(w.state.session.versions['v1'].protections_json))

    def test_satisfied_template_does_not_write_defaults_or_create_version(self):
        w = self.window
        self.template.write_bytes(plain_template(15))
        w.load_template_path(self.template); self.wait()
        w.apply_template(); self.wait()
        self.assertEqual(self.errors, [])
        self.assertEqual(w.state.session.current, 'original')
        self.assertEqual(len(w.state.session.versions), 1)

    def test_satisfied_template_selection_binds_next_edit_without_mutating_history(self):
        w = self.window
        self.template.write_bytes(plain_template(15))
        w.load_template_path(self.template); self.wait()
        w.apply_template(); self.wait()
        second = self.root / 'new-standard.docx'; second.write_bytes(self.template.read_bytes())
        w.load_template_path(second); self.wait()
        expected = w.template_baseline.fingerprint
        w.apply_template(); self.wait()
        self.assertEqual(json.loads(w.state.session.versions['original'].template_json), {})
        w.format_panel.controls['bold'].setCurrentIndex(w.format_panel.controls['bold'].findData(True))
        w.apply(); self.wait()
        self.assertEqual(self.errors, [])
        self.assertEqual(json.loads(w.state.session.versions['v1'].template_json)['fingerprint'], expected)

    def test_template_and_manuscript_save_paths_protect_both_sources(self):
        from desktop.template_dialog import TemplateDialog
        from template_workflow import TemplateWorkspace
        w = self.window
        w.template_source = self.template
        original, standard = self.source.read_bytes(), self.template.read_bytes()
        dialog = TemplateDialog(TemplateWorkspace(load('standard.docx', standard)), w)
        self.addCleanup(dialog.close)
        with (patch('desktop.template_dialog.QFileDialog.getSaveFileName', return_value=(str(self.source), '')),
              patch('desktop.template_dialog.QMessageBox.question', return_value=QMessageBox.StandardButton.Yes)):
            dialog.save_template()
        self.assertEqual(self.source.read_bytes(), original)
        self.assertIn('保留原稿', dialog.message.text())
        with (patch('desktop.window.QFileDialog.getSaveFileName', return_value=(str(self.template), '')),
              patch('desktop.window.QMessageBox.question', return_value=QMessageBox.StandardButton.Yes)):
            w.choose_save(); self.wait()
        self.assertEqual(self.template.read_bytes(), standard)
        self.assertIn('保留原稿', str(self.errors[-1]))

    def test_satisfied_template_still_installs_explicit_protection(self):
        from desktop.template_dialog import TemplateApplicationDialog
        from template_workflow import TemplateWorkspace
        w = self.window
        workspace = TemplateWorkspace(load('protected.docx', plain_template(15)))
        workspace.set_protections([workspace.model().scope('all')])
        w.template_baseline = workspace.baseline()
        def resolve(dialog):
            dialog.resolutions = {'protection:0': w.model.scope('all')}
            dialog.refresh()
            self.assertTrue(dialog.buttons.button(QDialogButtonBox.StandardButton.Ok).isEnabled())
            return QDialog.DialogCode.Accepted
        with patch.object(TemplateApplicationDialog, 'exec', resolve):
            w.apply_template(); self.wait()
        self.assertTrue(w.state.protected_scopes())
        self.assertEqual(w.state.session.current, 'original')
        w.format_panel.controls['bold'].setCurrentIndex(w.format_panel.controls['bold'].findData(True))
        w.apply(); self.wait()
        self.assertEqual(w.state.session.current, 'original')
        self.assertEqual(current_values(w.model, w.model.scope('all'), 'bold'), [False])


if __name__ == '__main__':
    unittest.main()
