"""Native UI lifecycle, stale ranges, and durable save regression tests."""
from io import BytesIO
from pathlib import Path
import os
import tempfile
import time
import unittest
from unittest.mock import patch

from PIL import Image

from desktop.controller import DocumentState, scope_choices
from docx_formatting import FormatError
from tests.test_v06 import ENV, sample
from v06_model import build_model
from v06_patch import current_values
from v06_plan import build_plan, intent
from v06_session import EditingSession

try:
    from PySide6.QtCore import QPoint, Qt
    from PySide6.QtGui import QColor, QPalette
    from PySide6.QtTest import QTest
    from PySide6.QtWidgets import QApplication
    from desktop.window import MainWindow
except ImportError:
    QApplication = None


def render(data, extension="docx"):
    out = BytesIO()
    Image.new("RGB", (100, 140), "white").save(out, format="PNG")
    return {"pages": 1, "images": [out.getvalue()], "pdf": b"test", "text": "synthetic"}


def load(name, data):
    session = EditingSession.load(name, data, renderer=render)
    session.environment = ENV
    return session


class SaveTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source = self.root / "原稿.docx"
        self.data = sample(False)
        self.source.write_bytes(self.data)
        self.state = DocumentState.open(self.source, loader=load)

    def edit(self):
        session = self.state.session
        model = session.model()
        plan = build_plan(model, [intent(model.scope("all"), {"font_family": "黑体"})]).plan
        version, _ = session.execute(model, plan, environment=ENV)
        session.select(version.id)
        return version

    def test_copy_is_byte_identical_and_source_cannot_be_overwritten(self):
        target = self.root / "副本.docx"
        self.state.save(target)
        self.assertEqual(target.read_bytes(), self.data)
        with self.assertRaisesRegex(FormatError, "保留原稿"):
            self.state.save(self.source)
        self.assertEqual(self.source.read_bytes(), self.data)

    def test_hardlink_alias_cannot_overwrite_source(self):
        alias = self.root / "链接.docx"
        os.link(self.source, alias)
        with self.assertRaisesRegex(FormatError, "保留原稿"):
            self.state.save(alias)
        self.assertEqual(self.source.read_bytes(), self.data)

    def test_modified_version_requires_review_and_keeps_extension(self):
        version = self.edit()
        self.assertTrue(self.state.unsaved)
        with self.assertRaisesRegex(FormatError, "核对"):
            self.state.save(self.root / "结果.docx")
        self.state.reviewed.add(version.id)
        with self.assertRaisesRegex(FormatError, "原文件格式"):
            self.state.save(self.root / "结果.doc")
        self.state.save(self.root / "结果.docx")
        self.assertFalse(self.state.unsaved)

    def test_failed_replace_preserves_existing_file_and_unsaved_version(self):
        version = self.edit()
        self.state.reviewed.add(version.id)
        target = self.root / "结果.docx"
        target.write_bytes(b"previous output")
        with patch("desktop.controller.os.replace", side_effect=PermissionError("occupied")):
            with self.assertRaises(PermissionError):
                self.state.save(target, overwrite=True)
        self.assertEqual(target.read_bytes(), b"previous output")
        self.assertEqual(self.source.read_bytes(), self.data)
        self.assertTrue(self.state.unsaved)
        self.assertEqual(list(self.root.glob(".wenxu-*")), [])

    def test_existing_output_requires_explicit_overwrite(self):
        target = self.root / "已存在.docx"
        target.write_bytes(b"existing content")
        with self.assertRaisesRegex(FormatError, "已存在"):
            self.state.save(target)
        self.assertEqual(target.read_bytes(), b"existing content")

    def test_scope_choices_follow_document_not_fixed_menu(self):
        model = build_model(self.data)
        labels = [label for label, _ in scope_choices(model)]
        self.assertIn("一级标题", labels)
        self.assertNotIn("二级标题", labels)
        self.assertIn("全部表格文字", labels)


@unittest.skipUnless(QApplication, "PySide6 is an optional desktop dependency")
class ThemeTests(unittest.TestCase):
    def test_dark_palette_keeps_window_toolbar_and_page_title_readable(self):
        app = QApplication.instance() or QApplication([])
        original = app.palette()
        dark = QPalette(original)
        for role, color in ((QPalette.ColorRole.Window, "#202020"),
                            (QPalette.ColorRole.WindowText, "#f4f4f4"),
                            (QPalette.ColorRole.Base, "#303030"),
                            (QPalette.ColorRole.Text, "#f4f4f4"),
                            (QPalette.ColorRole.Button, "#303030"),
                            (QPalette.ColorRole.ButtonText, "#f4f4f4")):
            dark.setColor(role, QColor(color))
        app.setPalette(dark)
        window = MainWindow()
        try:
            window.show()
            app.processEvents()
            for surface in (window, window.toolbar, window.current_view.scroll):
                self.assertLess(surface.palette().color(QPalette.ColorRole.Window).lightness(), 128)
                self.assertGreater(surface.palette().color(QPalette.ColorRole.WindowText).lightness(), 128)
            self.assertGreater(window.current_view.title.palette().color(QPalette.ColorRole.WindowText).lightness(), 128)
        finally:
            window.close()
            app.setPalette(original)

    def test_existing_window_recolors_when_system_palette_changes(self):
        app = QApplication.instance() or QApplication([])
        original = app.palette()
        dark = QPalette(original)
        light = QPalette(original)
        for role, dark_color, light_color in (
            (QPalette.ColorRole.Window, "#202020", "#f0f0f0"),
            (QPalette.ColorRole.WindowText, "#f4f4f4", "#000000"),
            (QPalette.ColorRole.Base, "#303030", "#ffffff"),
            (QPalette.ColorRole.Text, "#f4f4f4", "#000000"),
            (QPalette.ColorRole.Button, "#303030", "#f0f0f0"),
            (QPalette.ColorRole.ButtonText, "#f4f4f4", "#000000"),
        ):
            dark.setColor(role, QColor(dark_color))
            light.setColor(role, QColor(light_color))
        app.setPalette(dark)
        window = MainWindow()
        try:
            window.show()
            app.processEvents()
            app.setPalette(light)
            app.processEvents()
            for widget in (window.heading, window.notice, window.toolbar, window.current_view.title):
                self.assertEqual(widget.palette().color(QPalette.ColorRole.WindowText), QColor("#000000"))
            self.assertEqual(window.current_view.scroll.palette().color(QPalette.ColorRole.Window), QColor("#f0f0f0"))
            app.setPalette(dark)
            app.processEvents()
            self.assertEqual(window.heading.palette().color(QPalette.ColorRole.WindowText), QColor("#f4f4f4"))
            self.assertEqual(window.current_view.scroll.palette().color(QPalette.ColorRole.Window), QColor("#202020"))
        finally:
            window.close()
            app.setPalette(original)


@unittest.skipUnless(QApplication, "PySide6 is an optional desktop dependency")
class WindowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        SaveTests.setUp(self)
        self.window = MainWindow(loader=lambda path: DocumentState.open(path, loader=load), office_check=lambda: ENV)
        self.errors = []
        self.window.error = self.errors.append
        self.window.confirm_discard = lambda: True
        self.window.show()
        self.window.open_path(self.source)
        self.wait_job()
        self.addCleanup(self.window.close)

    def wait_job(self):
        deadline = time.monotonic() + 10
        while self.window.job is not None and time.monotonic() < deadline:
            self.app.processEvents()
            QTest.qWait(5)
        self.assertIsNone(self.window.job)
        self.app.processEvents()

    def apply_font(self):
        window = self.window
        window.scope.setCurrentIndex(next(i for i in range(window.scope.count()) if window.scope.itemText(i) == "全文文字"))
        window.format_panel.font_combo.setCurrentIndex(window.format_panel.font_combo.findData("黑体"))
        QTest.mouseClick(window.apply_button, Qt.MouseButton.LeftButton)
        self.wait_job()

    def test_apply_undo_select_and_review_are_bound_to_current_version(self):
        self.apply_font()
        window = self.window
        self.assertEqual(self.errors, [])
        self.assertEqual(window.state.session.current, "v1")
        self.assertEqual(current_values(window.model, window.model.scope("all"), "font_east_asia"), ["黑体"])
        self.assertFalse(window.save_action.isEnabled())
        QTest.mouseClick(window.reviewed, Qt.MouseButton.LeftButton, pos=QPoint(10, window.reviewed.height() // 2))
        self.assertTrue(window.save_action.isEnabled())
        self.assertEqual(window.format_panel.values(), {})
        window.undo()
        self.assertEqual(window.state.session.current, "original")
        self.assertEqual(window.model.revision, "original")
        window.choose_version(window.versions.findData("v1"))
        self.assertTrue(window.reviewed.isChecked())
        self.assertEqual(window.chosen_scope()["revision"], "v1")

    def test_repeated_font_edit_is_explicit_noop(self):
        self.apply_font()
        self.apply_font()
        self.assertEqual(len(self.window.state.session.versions), 2)
        self.assertIn("无需修改", self.window.statusBar().currentMessage())

    def test_search_requires_visible_selection_and_clears_old_results(self):
        window = self.window
        window.scope.setCurrentIndex(window.scope.findData("search"))
        window.search.setText("重复短语")
        self.assertEqual(window.targets.count(), 2)
        with self.assertRaisesRegex(FormatError, "选择"):
            window.chosen_scope()
        window.targets.item(1).setSelected(True)
        chosen = window.chosen_scope()
        self.assertEqual(chosen["spans"][0]["pid"], 3)
        window.search.setText("不存在的词")
        self.assertEqual(window.targets.count(), 0)
        with self.assertRaises(FormatError):
            window.chosen_scope()

    def test_unchanged_comparison_uses_the_same_visual_scale(self):
        window = self.window
        window.compare.setChecked(True)
        QTest.qWait(100)
        left, right = window.original_view.page.pixmap(), window.current_view.page.pixmap()
        self.assertLessEqual(abs(left.width() - right.width()), 1)
        window.resize(1100, 720)
        QTest.qWait(100)
        left, right = window.original_view.page.pixmap(), window.current_view.page.pixmap()
        self.assertLessEqual(abs(left.width() - right.width()), 1)

    def test_failed_open_keeps_existing_session(self):
        window = self.window
        old = window.state
        window.open_path(self.root / "missing.docx")
        self.wait_job()
        self.assertIs(window.state, old)
        self.assertIsInstance(self.errors[-1], FileNotFoundError)

    def test_failure_and_busy_close_do_not_destroy_session(self):
        window = self.window
        old = window.state
        window.start_job("working", lambda: time.sleep(0.1), lambda _: None)
        self.assertFalse(window.open_action.isEnabled())
        self.assertFalse(window.save_action.isEnabled())
        window.close()
        self.assertTrue(window.isVisible())
        self.wait_job()
        self.assertIs(window.state, old)
        with patch.object(window.state.session, "execute", side_effect=FormatError("render failed")):
            self.apply_font()
        self.assertEqual(window.state.session.current, "original")
        self.assertEqual(len(window.state.session.versions), 1)
        self.assertIn("render failed", str(self.errors[-1]))


if __name__ == "__main__":
    unittest.main()
