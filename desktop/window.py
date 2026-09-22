"""Native, asynchronous document workspace. The shared engine owns all edits."""
from __future__ import annotations

import json
from pathlib import Path

from PySide6.QtCore import QEvent, Qt, QThread, QTimer
from PySide6.QtGui import QAction, QKeySequence, QPixmap
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QCheckBox, QComboBox, QDoubleSpinBox,
    QFileDialog, QFormLayout, QFrame, QGroupBox, QHBoxLayout, QLabel, QLineEdit,
    QListWidget, QListWidgetItem, QMainWindow, QMessageBox, QProgressBar,
    QPushButton, QScrollArea, QSpinBox, QSplitter, QTabWidget, QToolBar,
    QVBoxLayout, QWidget,
)

from desktop import VERSION
from desktop.controller import DocumentState, require_office, scope_choices
from docx_formatting import FormatError
from v06_patch import current_values
from v06_plan import build_plan, intent, resolve_scope
from v06_session import source_font_warnings


class Job(QThread):
    def __init__(self, task, parent):
        super().__init__(parent)
        self.task, self.value, self.error = task, None, None

    def run(self):
        try:
            self.value = self.task()
        except Exception as exc:
            self.error = exc


class PageView(QWidget):
    def __init__(self, title):
        super().__init__()
        layout = QVBoxLayout(self)
        self.title = QLabel(title)
        self.title.setObjectName("pageTitle")
        layout.addWidget(self.title)
        self.scroll = QScrollArea()
        self.scroll.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop)
        self.scroll.setWidgetResizable(False)
        self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.page = QLabel("打开一份 Word 文档，开始整理格式")
        self.page.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.page.setMargin(16)
        self.scroll.setWidget(self.page)
        layout.addWidget(self.scroll)
        self.pixmap = None
        self.zoom = 1.0
        self.resize_timer = QTimer(self)
        self.resize_timer.setSingleShot(True)
        self.resize_timer.timeout.connect(self.update_size)
        self.scroll.viewport().installEventFilter(self)

    def eventFilter(self, watched, event):
        if watched is self.scroll.viewport() and event.type() == QEvent.Type.Resize:
            self.resize_timer.start(0)
        return super().eventFilter(watched, event)

    def show_page(self, image, zoom):
        self.zoom = zoom
        self.pixmap = QPixmap()
        if image:
            self.pixmap.loadFromData(image)
        self.update_size()

    def update_size(self):
        if self.pixmap and not self.pixmap.isNull():
            width = max(100, int((self.scroll.viewport().width() - 36) * self.zoom))
            rendered = self.pixmap.scaledToWidth(width, Qt.TransformationMode.SmoothTransformation)
            self.page.setPixmap(rendered)
            self.page.resize(rendered.width() + 32, rendered.height() + 32)
        else:
            self.page.clear()
            self.page.setText("此版本没有这一页")
            self.page.resize(260, 100)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.resize_timer.start(0)


class FormatPanel(QWidget):
    """Every value is explicitly opted in; a reset can never imply an edit."""
    def __init__(self, *, page=False):
        super().__init__()
        self.controls = {}
        layout = QVBoxLayout(self)
        main = QFormLayout()
        layout.addLayout(main)
        if page:
            for key, label in (("top_margin", "上边距"), ("bottom_margin", "下边距"),
                               ("left_margin", "左边距"), ("right_margin", "右边距")):
                self.number(main, key, label, 0, 720, 72)
            note = QLabel("单位：磅，28.35 磅约为 1 厘米。\n页面设置作用于整份文档。")
            note.setWordWrap(True)
            layout.addWidget(note)
        else:
            self.font_combo = self.choice(main, "font_family", "字体", [])
            self.number(main, "font_size", "字号（磅）", 6, 72, 12)
            self.choice(main, "bold", "加粗", [("开启", True), ("关闭", False)])
            self.choice(main, "alignment", "对齐", [("左对齐", "left"), ("居中", "center"),
                                                       ("右对齐", "right"), ("两端对齐", "justify")])
            toggle = QPushButton("更多格式 ▸")
            toggle.setCheckable(True)
            layout.addWidget(toggle)
            extra = QWidget()
            form = QFormLayout(extra)
            form.setContentsMargins(0, 4, 0, 0)
            self.choice(form, "italic", "斜体", [("开启", True), ("关闭", False)])
            self.choice(form, "underline", "下划线", [("开启", True), ("关闭", False)])
            color = QLineEdit()
            color.setPlaceholderText("保持原样；例如 #333333")
            self.controls["color"] = color
            form.addRow("文字颜色", color)
            for key, label in (("first_line_indent", "首行缩进"), ("left_indent", "左缩进"),
                               ("right_indent", "右缩进"), ("space_before", "段前"), ("space_after", "段后")):
                self.number(form, key, label + "（磅）", 0, 720, 0)
            self.spacing = self.choice(form, "spacing_mode", "行距", [("固定值（磅）", "exact"), ("倍数", "multiple")])
            self.spacing_value = QDoubleSpinBox()
            self.spacing_value.setRange(6, 144)
            self.spacing_value.setValue(28)
            self.spacing_value.setSingleStep(0.5)
            self.spacing_value.setEnabled(False)
            form.addRow("行距值", self.spacing_value)
            self.spacing.currentIndexChanged.connect(self.spacing_changed)
            self.choice(form, "keep_with_next", "与下段同页", [("开启", True), ("关闭", False)])
            layout.addWidget(extra)
            extra.hide()
            toggle.toggled.connect(extra.setVisible)
            toggle.toggled.connect(lambda checked: toggle.setText("更多格式 ▾" if checked else "更多格式 ▸"))
        layout.addStretch()

    def choice(self, form, key, label, options):
        widget = QComboBox()
        widget.addItem("保持原样", None)
        for title, value in options:
            widget.addItem(title, value)
        widget.setMinimumWidth(150)
        self.controls[key] = widget
        form.addRow(label, widget)
        return widget

    def number(self, form, key, label, low, high, initial):
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        check, value = QCheckBox(), QDoubleSpinBox()
        check.setAccessibleName("设置" + label)
        value.setAccessibleName(label)
        value.setRange(low, high)
        value.setValue(initial)
        value.setSingleStep(0.5)
        value.setEnabled(False)
        check.toggled.connect(value.setEnabled)
        layout.addWidget(check)
        layout.addWidget(value, 1)
        form.addRow(label, row)
        self.controls[key] = (check, value)

    def spacing_changed(self):
        mode = self.spacing.currentData()
        self.spacing_value.setEnabled(mode is not None)
        self.spacing_value.setRange(*( (0.5, 5) if mode == "multiple" else (6, 144)))
        self.spacing_value.setValue(1.5 if mode == "multiple" else 28)

    def reset(self, fonts=()):
        for widget in self.controls.values():
            if isinstance(widget, tuple):
                widget[0].setChecked(False)
            elif isinstance(widget, QComboBox):
                widget.setCurrentIndex(0)
            else:
                widget.clear()
        if "font_family" in self.controls:
            self.font_combo.clear()
            self.font_combo.addItem("保持原样", None)
            for name in sorted(fonts, key=str.casefold):
                self.font_combo.addItem(name, name)

    def values(self):
        result = {}
        for key, widget in self.controls.items():
            if isinstance(widget, tuple):
                if widget[0].isChecked():
                    result[key] = widget[1].value()
            elif isinstance(widget, QComboBox):
                if widget.currentData() is not None:
                    result[key] = widget.currentData()
            elif widget.text().strip():
                result[key] = widget.text().strip()
        mode = result.pop("spacing_mode", None)
        if mode:
            result["line_spacing"] = {"mode": mode, "value": self.spacing_value.value()}
        return result


class MainWindow(QMainWindow):
    def __init__(self, *, loader=DocumentState.open, office_check=require_office):
        super().__init__()
        self.loader, self.office_check = loader, office_check
        self.state = self.model = self.job = None
        self.after_job = None
        self.previews = {}
        self.setWindowTitle("文序 · Windows 桌面版")
        self.resize(1280, 850)
        self.setMinimumSize(960, 650)
        self.toolbar = QToolBar("文档")
        self.toolbar.setMovable(False)
        self.addToolBar(self.toolbar)
        self.open_action = self.action("打开文档", self.choose_file, QKeySequence.StandardKey.Open)
        self.undo_action = self.action("撤销", self.undo, QKeySequence.StandardKey.Undo)
        self.toolbar.addSeparator()
        self.versions = QComboBox()
        self.versions.setMinimumWidth(145)
        self.versions.setAccessibleName("当前版本")
        self.versions.activated.connect(self.choose_version)
        self.toolbar.addWidget(self.versions)
        self.toolbar.addSeparator()
        self.save_action = self.action("另存副本", self.choose_save, QKeySequence.StandardKey.SaveAs)
        self.toolbar.addSeparator()
        self.action("关于", self.about)
        root = QWidget()
        layout = QVBoxLayout(root)
        layout.setContentsMargins(16, 12, 16, 10)
        self.heading = QLabel("让格式整理回到文档本身")
        self.heading.setObjectName("heading")
        layout.addWidget(self.heading)
        self.notice = QLabel("选择本机 DOC / DOCX 文件。原稿始终保留，处理结果另存为副本。")
        self.notice.setWordWrap(True)
        self.notice.setTextFormat(Qt.TextFormat.PlainText)
        layout.addWidget(self.notice)
        split = QSplitter()
        layout.addWidget(split, 1)
        workspace = QWidget()
        center = QVBoxLayout(workspace)
        center.setContentsMargins(0, 0, 8, 0)
        navigation = QHBoxLayout()
        self.compare = QCheckBox("对照原稿")
        self.compare.toggled.connect(self.show_pages)
        navigation.addWidget(self.compare)
        navigation.addStretch()
        navigation.addWidget(QLabel("页码"))
        self.page_number = QSpinBox()
        self.page_number.setRange(1, 1)
        self.page_number.valueChanged.connect(self.show_pages)
        navigation.addWidget(self.page_number)
        self.page_count = QLabel("/ 0")
        navigation.addWidget(self.page_count)
        self.zoom = QComboBox()
        for label, value in (("适合宽度", 1.0), ("放大 125%", 1.25), ("放大 150%", 1.5), ("放大 200%", 2.0)):
            self.zoom.addItem(label, value)
        self.zoom.currentIndexChanged.connect(self.show_pages)
        navigation.addWidget(self.zoom)
        center.addLayout(navigation)
        panes = QHBoxLayout()
        self.original_view = PageView("原稿")
        self.current_view = PageView("当前版本")
        panes.addWidget(self.original_view, 1)
        panes.addWidget(self.current_view, 1)
        self.original_view.hide()
        center.addLayout(panes, 1)
        self.reviewed = QCheckBox("已核对当前版本的预览")
        self.reviewed.toggled.connect(self.mark_reviewed)
        center.addWidget(self.reviewed)
        split.addWidget(workspace)
        self.sidebar = QWidget()
        side = QVBoxLayout(self.sidebar)
        side.setContentsMargins(12, 0, 0, 0)
        side.addWidget(QLabel("修改范围"))
        self.scope = QComboBox()
        self.scope.currentIndexChanged.connect(self.scope_changed)
        side.addWidget(self.scope)
        self.local_scope = QWidget()
        local = QVBoxLayout(self.local_scope)
        local.setContentsMargins(0, 0, 0, 0)
        self.search = QLineEdit()
        self.search.setPlaceholderText("输入要修改的文字")
        self.search.textChanged.connect(self.populate_targets)
        local.addWidget(self.search)
        self.targets = QListWidget()
        self.targets.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.targets.setMaximumHeight(170)
        self.targets.itemSelectionChanged.connect(self.scope_summary)
        local.addWidget(self.targets)
        side.addWidget(self.local_scope)
        self.current_format = QLabel()
        self.current_format.setWordWrap(True)
        self.current_format.setTextFormat(Qt.TextFormat.PlainText)
        side.addWidget(self.current_format)
        self.tabs = QTabWidget()
        self.format_panel = FormatPanel()
        self.page_panel = FormatPanel(page=True)
        for title, panel in (("文字与段落", self.format_panel), ("页面", self.page_panel)):
            scroll = QScrollArea()
            scroll.setFrameShape(QFrame.Shape.NoFrame)
            scroll.setWidgetResizable(True)
            scroll.setWidget(panel)
            self.tabs.addTab(scroll, title)
        self.tabs.currentChanged.connect(self.scope_summary)
        side.addWidget(self.tabs, 1)
        self.apply_button = QPushButton("应用修改")
        self.apply_button.setObjectName("primary")
        self.apply_button.clicked.connect(self.apply)
        side.addWidget(self.apply_button)
        split.addWidget(self.sidebar)
        split.setSizes([880, 330])
        split.setStretchFactor(0, 1)
        split.setStretchFactor(1, 0)
        self.sidebar.setMinimumWidth(285)
        self.sidebar.setMaximumWidth(430)
        self.setCentralWidget(root)
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setMaximumWidth(160)
        self.progress.hide()
        self.statusBar().addPermanentWidget(self.progress)
        self.statusBar().showMessage("本地处理 · 原格式保存")
        self.setStyleSheet("""
            QMainWindow, QWidget { font-family: 'Microsoft YaHei UI'; font-size: 13px; }
            QToolBar { padding: 8px; spacing: 9px; border-bottom: 1px solid palette(mid); }
            #heading { font-size: 21px; font-weight: 600; padding: 4px 0; }
            #pageTitle { padding: 4px; }
            QScrollArea { border: 1px solid palette(mid); border-radius: 5px; }
            QComboBox, QLineEdit, QDoubleSpinBox, QSpinBox { min-height: 28px; }
            QPushButton { min-height: 30px; padding: 3px 12px; }
            QPushButton#primary { background: #2563eb; color: white; border: 0; border-radius: 5px; min-height: 38px; }
            QPushButton#primary:disabled { background: palette(button); color: palette(button-text); }
            QTabWidget::pane { border: 0; }
        """)
        self.refresh_enabled()

    def action(self, text, callback, shortcut=None):
        action = QAction(text, self)
        action.triggered.connect(callback)
        if shortcut:
            action.setShortcut(shortcut)
        self.toolbar.addAction(action)
        return action

    def error(self, exc):
        QMessageBox.warning(self, "未完成此操作", str(exc))

    def confirm_discard(self):
        if self.state and self.state.unsaved:
            return QMessageBox.question(self, "还有未保存的版本", "继续将丢弃本次会话中未保存的版本。是否继续？",
                                        QMessageBox.StandardButton.Discard | QMessageBox.StandardButton.Cancel,
                                        QMessageBox.StandardButton.Cancel) == QMessageBox.StandardButton.Discard
        return True

    def start_job(self, text, task, callback):
        if self.job:
            return
        self.job = Job(task, self)
        self.after_job = callback
        self.job.finished.connect(self.job_done)
        self.statusBar().showMessage(text)
        self.progress.show()
        self.refresh_enabled()
        self.job.start()

    def job_done(self):
        job, callback = self.job, self.after_job
        self.job = self.after_job = None
        self.progress.hide()
        self.refresh_enabled()
        try:
            if job.error:
                self.statusBar().showMessage("操作未完成；当前文件和版本已保留")
                self.error(job.error)
            else:
                callback(job.value)
        except Exception as exc:
            self.error(exc)
        finally:
            job.deleteLater()

    def refresh_enabled(self):
        ready = self.state is not None and self.job is None
        self.open_action.setEnabled(self.job is None)
        self.sidebar.setEnabled(ready)
        self.versions.setEnabled(ready)
        self.undo_action.setEnabled(ready and self.state.session.current != "original")
        self.save_action.setEnabled(ready and (self.state.session.current == "original" or self.reviewed.isChecked()))
        self.reviewed.setEnabled(ready and self.state.session.current != "original")

    def choose_file(self):
        if not self.confirm_discard():
            return
        path, _ = QFileDialog.getOpenFileName(self, "打开 Word 文档", "", "Word 文档 (*.doc *.docx)")
        if path:
            self.open_path(path)

    def open_path(self, path):
        def load():
            self.office_check()
            state = self.loader(path)
            return state, state.session.model(), state.session.preview()
        self.start_job("正在打开并核对文档，请稍候…", load, self.loaded)

    def loaded(self, value):
        self.state, self.model, preview = value
        self.previews = {"original": preview}
        self.compare.setChecked(False)
        self.page_number.setValue(1)
        self.heading.setText(self.state.source.name)
        self.heading.setTextFormat(Qt.TextFormat.PlainText)
        self.setWindowTitle(self.state.source.name + " — 文序")
        warnings = source_font_warnings(self.model, self.state.session.environment)
        self.notice.setText("原稿已保留。" + (" ".join(warnings) if warnings else "选择范围和格式后应用修改。"))
        self.sync_version()
        self.statusBar().showMessage("已打开 · " + str(preview["pages"]) + " 页 · 保留 " + self.state.source.suffix.upper() + " 格式")

    def sync_version(self):
        session = self.state.session
        self.versions.blockSignals(True)
        self.versions.clear()
        for ident in session.versions:
            self.versions.addItem("导入原稿" if ident == "original" else "修改版本 " + ident, ident)
        self.versions.setCurrentIndex(self.versions.findData(session.current))
        self.versions.blockSignals(False)
        self.scope.blockSignals(True)
        self.scope.clear()
        for label, value in scope_choices(self.model):
            self.scope.addItem(label, value)
        self.scope.addItem("选取具体段落…", "paragraphs")
        self.scope.addItem("查找指定文字…", "search")
        self.scope.blockSignals(False)
        self.search.clear()
        self.format_panel.reset(session.environment.get("font_families", []))
        self.page_panel.reset()
        self.reviewed.blockSignals(True)
        self.reviewed.setChecked(session.current in self.state.reviewed)
        self.reviewed.blockSignals(False)
        self.scope_changed()
        self.show_pages()
        self.refresh_enabled()

    def scope_changed(self):
        value = self.scope.currentData()
        local = isinstance(value, str)
        self.local_scope.setVisible(local)
        self.search.setVisible(value == "search")
        self.populate_targets()

    def populate_targets(self):
        self.targets.blockSignals(True)
        self.targets.clear()
        if self.model and self.scope.currentData() == "paragraphs":
            for row in self.model.paragraphs:
                if row["text"].strip():
                    item = QListWidgetItem(row["location"] + " · " + row["text"][:100])
                    item.setToolTip(row["text"])
                    item.setData(Qt.ItemDataRole.UserRole, row["pid"])
                    self.targets.addItem(item)
        elif self.model and self.scope.currentData() == "search":
            for hit in self.model.search(self.search.text()):
                item = QListWidgetItem(hit["context"])
                item.setToolTip(hit["context"])
                item.setData(Qt.ItemDataRole.UserRole, hit["spans"])
                self.targets.addItem(item)
        self.targets.blockSignals(False)
        self.scope_summary()

    def chosen_scope(self):
        if self.tabs.currentIndex() == 1:
            return self.model.scope("page")
        value = self.scope.currentData()
        if isinstance(value, dict):
            return value
        selected = [i.data(Qt.ItemDataRole.UserRole) for i in self.targets.selectedItems()]
        if not selected:
            raise FormatError("请在列表中选择要修改的内容；可按 Ctrl 多选。")
        if value == "paragraphs":
            return self.model.scope("paragraphs", pids=selected)
        return self.model.scope("selection", spans=[s for group in selected for s in group])

    def scope_summary(self):
        if self.model is None:
            return
        is_page = self.tabs.currentIndex() == 1
        self.scope.setEnabled(not is_page)
        self.local_scope.setEnabled(not is_page)
        try:
            scope = self.chosen_scope()
            if is_page:
                text = "勾选需要调整的页边距；未勾选项保持原样。"
            else:
                spans, _ = resolve_scope(self.model, scope)
                fonts = current_values(self.model, scope, "font_east_asia")
                summary = "、".join(str(v) if v is not None else "默认字体" for v in fonts[:4]) or "无可编辑文字"
                text = f"已选 {len({s['pid'] for s in spans})} 个段落 · 当前中文字体：{summary}" + ("等" if len(fonts) > 4 else "")
            self.current_format.setText(text)
        except FormatError as exc:
            self.current_format.setText(str(exc))

    def apply(self):
        if not self.state or self.job:
            return
        try:
            panel = self.page_panel if self.tabs.currentIndex() == 1 else self.format_panel
            review = build_plan(self.model, [intent(self.chosen_scope(), panel.values(), source="桌面控件")])
            if review.blockers:
                raise FormatError("\n".join(review.blockers))
            if review.notices and QMessageBox.question(self, "确认修改范围", "\n".join(review.notices) + "\n继续应用？") != QMessageBox.StandardButton.Yes:
                return
            model, session = self.model, self.state.session
            def execute():
                version, created = session.execute(model, review.plan, environment=session.environment)
                if version is None:
                    return None
                # Version selection stays on the UI thread after the job succeeds.
                return version, session.preview(version.id)
            self.start_job("正在应用格式并检查结果，请稍候…", execute, self.applied)
        except Exception as exc:
            self.error(exc)

    def applied(self, result):
        if result is None:
            self.statusBar().showMessage("所选内容已经符合要求，无需修改；没有新增版本。")
            return
        version, preview = result
        self.state.session.select(version.id)
        self.model = self.state.session.model()
        self.previews[version.id] = preview
        self.sync_version()
        report = json.loads(version.report_json)
        self.notice.setText(f"已应用修改 · {report['pages_before']} → {report['pages_after']} 页。请核对预览后另存副本。")
        self.statusBar().showMessage("修改成功 · " + version.name)

    def choose_version(self, index):
        if not self.state or self.job:
            return
        self.state.session.select(self.versions.itemData(index))
        self.model = self.state.session.model()
        self.sync_version()
        self.notice.setText("已切换版本。后续修改基于当前版本，原稿始终保留。")

    def undo(self):
        if self.state and not self.job:
            self.state.session.undo()
            self.model = self.state.session.model()
            self.sync_version()
            self.notice.setText("已撤销到上一版本；可在版本列表中返回后续版本。")

    def show_pages(self):
        self.original_view.setVisible(self.compare.isChecked())
        if not self.state:
            return
        current = self.previews[self.state.session.current]
        original = self.previews["original"]
        count = max(current["pages"], original["pages"]) if self.compare.isChecked() else current["pages"]
        self.page_number.blockSignals(True)
        self.page_number.setMaximum(count)
        self.page_number.blockSignals(False)
        self.page_count.setText("/ " + str(count))
        index, zoom = self.page_number.value() - 1, self.zoom.currentData()
        for view, preview in ((self.original_view, original), (self.current_view, current)):
            view.show_page(preview["images"][index] if index < len(preview["images"]) else None, zoom)
        self.current_view.title.setText("当前版本 · " + self.state.session.current + f" · {current['pages']} 页")
        self.original_view.resize_timer.start(0)
        self.current_view.resize_timer.start(0)

    def mark_reviewed(self, checked):
        if self.state:
            if checked:
                self.state.reviewed.add(self.state.session.current)
            else:
                self.state.reviewed.discard(self.state.session.current)
        self.refresh_enabled()

    def choose_save(self):
        if not self.state or self.job:
            return
        version = self.state.session.versions[self.state.session.current]
        name = version.name if version.id != "original" else self.state.source.stem + "_副本." + version.extension
        path, _ = QFileDialog.getSaveFileName(self, "另存副本", str(self.state.source.with_name(name)),
                                            f"Word 文档 (*.{version.extension})",
                                            options=QFileDialog.Option.DontConfirmOverwrite)
        if path:
            if not Path(path).suffix:
                path += "." + version.extension
            overwrite = Path(path).exists()
            if overwrite and QMessageBox.question(self, "替换已有副本", "该文件已存在，是否替换？\n" + str(path),
                                                  QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                                  QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes:
                return
            state = self.state
            self.start_job("正在保存副本…", lambda: state.save(path, overwrite=overwrite),
                           lambda saved: self.statusBar().showMessage("已保存：" + str(saved)))

    def about(self):
        QMessageBox.information(self, "关于文序", "文序 Windows 桌面版 " + VERSION +
                                "\nLTS 开发预览版 · 本地处理 · 原格式保存\n需要兼容的 WPS 与文档所用字体。\n会话版本仅在本次运行保留，请及时另存副本。")

    def closeEvent(self, event):
        if self.job:
            self.statusBar().showMessage("正在处理文档，请等待当前操作结束后关闭。")
            event.ignore()
        elif self.confirm_discard():
            event.accept()
        else:
            event.ignore()
