"""Small native dialogs for explicit template questions; no document rewriting."""
from pathlib import Path
from uuid import uuid4

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView, QCheckBox, QComboBox, QDialog, QDialogButtonBox,
    QFormLayout, QHBoxLayout, QInputDialog, QLabel, QLineEdit, QListWidget,
    QListWidgetItem, QMessageBox, QPlainTextEdit, QPushButton, QScrollArea,
    QSpinBox, QTabWidget, QVBoxLayout, QWidget, QFileDialog,
)

from desktop.classification import ClassificationPanel
from desktop.controller import save_bytes, paragraph_page
from desktop.window import FormatPanel, Job, PageView
from docx_formatting import FormatError
from v06_model import ROLE_LABELS
from v06_plan import resolve_scope, validate_values


class RuleDialog(QDialog):
    def __init__(self, model, rule=None, fonts=(), *, scope_only=False, parent=None):
        super().__init__(parent)
        self.setWindowTitle('明确作用范围' if scope_only else '核对范围与格式要求')
        self.resize(620, 700)
        self.model, self.scope_only = model, scope_only
        self.result_scope = self.result_values = None
        layout = QVBoxLayout(self)
        self.scope = QComboBox()
        initial = (rule or {}).get('scope')
        if initial:
            self.scope.addItem('当前范围（保留精确选区）', initial)
        for key, label in ROLE_LABELS.items():
            if key != 'unknown':
                self.scope.addItem('全部' + label, model.scope('role', role=key))
        for column in model.columns:
            self.scope.addItem('栏目：' + column['label'], model.scope('named_column', column_id=column['id']))
        self.scope.addItem('指定段落（下方多选）', model.scope('paragraphs', pids=[]))
        self.scope.addItem('全文普通文字', model.scope('all'))
        layout.addWidget(self.scope)
        self.paragraphs = QListWidget()
        self.paragraphs.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.paragraphs.setMaximumHeight(180)
        for row in model.paragraphs:
            item = QListWidgetItem(row['location'] + ' · ' + (row['text'][:70] or '（空白填写区）'))
            item.setToolTip(row['text']); item.setData(Qt.ItemDataRole.UserRole, row['pid'])
            self.paragraphs.addItem(item)
        layout.addWidget(self.paragraphs)
        self.scope.currentIndexChanged.connect(self.scope_changed)
        self.scope_changed()
        self.exception = QCheckBox('这是明确的局部例外（仅覆盖本次属性）')
        self.exception.setChecked(bool((rule or {}).get('exception')))
        self.panel = FormatPanel()
        self.panel.reset(fonts)
        values = (rule or {}).get('set', {})
        for key, value in values.items():
            control = self.panel.controls.get(key)
            if isinstance(control, tuple):
                control[0].setChecked(True); control[1].setValue(value)
            elif isinstance(control, QComboBox):
                index = control.findData(value)
                if index < 0:
                    control.addItem(str(value), value); index = control.count() - 1
                control.setCurrentIndex(index)
            elif isinstance(control, QLineEdit):
                control.setText(str(value))
        if isinstance(values.get('line_spacing'), dict):
            spacing = values['line_spacing']
            self.panel.spacing.setCurrentIndex(self.panel.spacing.findData(spacing['mode']))
            self.panel.spacing_value.setValue(spacing['value'])
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        field_box = QWidget(); fields = QVBoxLayout(field_box)
        script_fonts = QFormLayout()
        self.chinese, self.western = QLineEdit(), QLineEdit()
        self.chinese.setText(values.get('font_east_asia', ''))
        self.western.setText(values.get('font_western', ''))
        self.chinese.setPlaceholderText('留空不单独指定中文字体')
        self.western.setPlaceholderText('留空不单独指定西文字体')
        script_fonts.addRow('中文字体', self.chinese); script_fonts.addRow('西文字体', self.western)
        fields.addLayout(script_fonts); fields.addWidget(self.panel)
        scroll.setWidget(field_box)
        if not scope_only:
            layout.addWidget(scroll, 1); layout.addWidget(self.exception)
            note = QLabel('选择“字体”将同时设置中西文；仅指定中文或西文时使用上方输入。未设置的属性保持原样。')
            note.setWordWrap(True); layout.addWidget(note)
        else:
            scroll.hide()
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept_rule); buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def scope_changed(self):
        scope = self.scope.currentData()
        self.paragraphs.setVisible(scope['kind'] == 'paragraphs' and not scope.get('pids'))

    def accept_rule(self):
        scope = dict(self.scope.currentData())
        if scope['kind'] == 'paragraphs' and not scope.get('pids'):
            scope['pids'] = [i.data(Qt.ItemDataRole.UserRole) for i in self.paragraphs.selectedItems()]
            if not scope['pids']:
                QMessageBox.warning(self, '请选择范围', '请选择至少一个实际段落。'); return
        values = self.panel.values() if not self.scope_only else {}
        if 'font_family' not in values and not self.scope_only:
            for key, widget in (('font_east_asia', self.chinese), ('font_western', self.western)):
                if widget.text().strip(): values[key] = widget.text().strip()
        if not self.scope_only and not values:
            QMessageBox.warning(self, '请选择格式', '请至少指定一个支持的格式属性。'); return
        self.result_scope, self.result_values = scope, values
        self.accept()


class TemplateDialog(QDialog):
    def __init__(self, workspace, parent=None):
        super().__init__(parent)
        self.workspace, self.baseline, self.changed, self.job = workspace, None, False, None
        self.setWindowTitle('核对教案模板'); self.resize(980, 760)
        layout = QVBoxLayout(self)
        self.status = QLabel(); self.status.setWordWrap(True); self.status.setTextFormat(Qt.TextFormat.PlainText)
        layout.addWidget(self.status)
        self.tabs = QTabWidget(); layout.addWidget(self.tabs, 1)
        rules_page = QWidget(); rules_layout = QVBoxLayout(rules_page)
        self.items = QListWidget(); self.items.currentRowChanged.connect(self.selected)
        rules_layout.addWidget(self.items, 1)
        self.detail = QPlainTextEdit(); self.detail.setReadOnly(True); self.detail.setMaximumHeight(130)
        rules_layout.addWidget(self.detail)
        actions = QHBoxLayout()
        for text, slot in (('设置格式与范围', self.edit_rule), ('保留为人工事项', self.manual),
                           ('标记不适用', self.not_applicable), ('设置保护范围', self.protect)):
            button = QPushButton(text); button.clicked.connect(slot); actions.addWidget(button)
        rules_layout.addLayout(actions)
        extra = QHBoxLayout()
        add = QPushButton('新增局部例外'); add.clicked.connect(self.add_rule)
        reset = QPushButton('撤销所选处理'); reset.clicked.connect(self.reset_rule)
        extra.addWidget(add); extra.addWidget(reset); extra.addStretch()
        rules_layout.addLayout(extra)
        self.tabs.addTab(rules_page, '旁批与样式')
        self.roles = ClassificationPanel()
        self.roles.roles_changed.connect(self.correct_roles)
        self.roles.locate.connect(self.locate_paragraph)
        self.tabs.addTab(self.roles, '正文与标题分类')
        preview_page = QWidget(); pv = QVBoxLayout(preview_page)
        self.page = QSpinBox(); self.page.setPrefix('第 '); self.page.setSuffix(' 页')
        self.page.valueChanged.connect(self.show_preview)
        pv.addWidget(self.page)
        self.page_view = PageView('模板实际文件预览'); pv.addWidget(self.page_view, 1)
        self.tabs.addTab(preview_page, '模板预览')
        self.reviewed = QCheckBox('已核对当前模板的预览及未自动验证事项')
        self.reviewed.toggled.connect(self.enable_actions)
        layout.addWidget(self.reviewed)
        self.message = QLabel(); self.message.setWordWrap(True); self.message.setTextFormat(Qt.TextFormat.PlainText)
        layout.addWidget(self.message)
        buttons = QHBoxLayout()
        self.calibrate_button = QPushButton('校验并修正模板'); self.calibrate_button.clicked.connect(self.calibrate)
        self.save_button = QPushButton('另存模板文件'); self.save_button.clicked.connect(self.save_template)
        self.adopt_button = QPushButton('采用此金标准'); self.adopt_button.clicked.connect(self.adopt)
        self.cancel_button = QPushButton('返回'); self.cancel_button.clicked.connect(self.reject)
        for button in (self.calibrate_button, self.save_button, self.adopt_button, self.cancel_button): buttons.addWidget(button)
        layout.addLayout(buttons)
        self.refresh()

    def refresh(self):
        self.review = self.workspace.review()
        self.items.clear()
        known = set()
        for c in self.review.comments:
            item = QListWidgetItem('旁批 · ' + c['text'].replace('\n', ' ')[:100])
            item.setData(Qt.ItemDataRole.UserRole, ('comment', c['id']))
            self.items.addItem(item); known.add(c['key'])
        for rule in self.review.rules:
            if rule['key'] in known: continue
            role_label = ROLE_LABELS.get(rule['scope'].get('role'), '')
            item = QListWidgetItem('样式 · ' + role_label + ' · ' + str(rule.get('source', rule['key'])))
            item.setData(Qt.ItemDataRole.UserRole, ('rule', rule['key'])); self.items.addItem(item)
        self.roles.set_model(self.workspace.model())
        pids = {a['pid'] for a in self.review.plan.read()['assignments'] if a['pid']} if self.review.plan else set()
        self.roles.set_included(pids)
        self.status.setText(('按模板现有样式提取' if self.review.basis_kind == 'sample_styles' else '按旁批校准模板')
                            + f' · {len(self.review.differences)} 项属性差异 · {len(self.review.blockers)} 项待处理'
                            + f' · {len(self.review.manual_items)} 项未自动验证事项')
        self.message.setText('\n'.join(self.review.blockers) if self.review.blockers else '明确的格式要求已就绪；人工事项不会被计为自动校验通过。')
        if self.items.count(): self.items.setCurrentRow(0)
        self.preview = self.workspace.session.preview()
        self.page.setRange(1, self.preview['pages']); self.show_preview()
        self.enable_actions()

    def selected(self, index=0):
        item = self.items.currentItem()
        if not item: return
        kind, key = item.data(Qt.ItemDataRole.UserRole)
        if kind == 'comment':
            c = next(c for c in self.review.comments if c['id'] == key)
            context = c.get('context', '')
            status = {'pending': '待核对', 'parsed': '已提取', 'explicit': '已明确', 'manual': '人工事项',
                      'not_applicable': '不适用', 'protection': '保护范围'}.get(c.get('status'), '待核对')
            self.detail.setPlainText(c['text'] + '\n\n对应内容：' + str(context) + '\n状态：' + status)
            rule_key = c['key']
        else:
            rule = next(r for r in self.review.rules if r['key'] == key)
            self.detail.setPlainText('来源：' + str(rule.get('source', '模板样式')) + '\n目标格式：' + self.values_label(rule.get('set', {})))
            rule_key = key
        rule = next((r for r in self.review.rules if r['key'] == rule_key), None)
        if rule:
            try:
                from template_workflow import _value_reader
                model = self.workspace.model(); scope = rule['scope']
                spans, _ = resolve_scope(model, scope)
                pids = list(dict.fromkeys(s['pid'] for s in spans))
                scope_label = ROLE_LABELS.get(scope.get('role'), '指定范围')
                positions = '；'.join(model.row(pid)['location'] for pid in pids[:4])
                read = _value_reader(model)
                actual = {prop: read(scope, prop) for prop in validate_values(rule['set'])}
                self.detail.appendPlainText(f'\n实际作用范围：{scope_label} · {len(pids)} 段 · {positions}'
                                            + '\n当前有效格式：' + self.values_label(actual)
                                            + '\n要求格式：' + self.values_label(rule['set']))
            except FormatError as exc:
                self.detail.appendPlainText('\n范围待核对：' + str(exc))
        else:
            self.detail.appendPlainText('\n尚未形成完整格式目标；请明确要求类型与实际作用范围。')

    @staticmethod
    def values_label(values):
        labels = {'font_east_asia': '中文字体', 'font_western': '西文字体', 'font_family': '字体',
                  'font_size': '字号（磅）', 'bold': '加粗', 'italic': '斜体', 'underline': '下划线',
                  'alignment': '对齐', 'line_spacing': '行距', 'first_line_indent': '首行缩进',
                  'left_indent': '左缩进', 'right_indent': '右缩进', 'space_before': '段前',
                  'space_after': '段后', 'keep_with_next': '与下段同页', 'color': '文字颜色'}
        def display(value):
            if isinstance(value, list): return ' / '.join(display(v) for v in value)
            if isinstance(value, dict) and 'mode' in value:
                return ('固定 ' if value['mode'] == 'exact' else '倍数 ') + str(value['value'])
            if value is None: return '继承或未指定'
            return '是' if value is True else '否' if value is False else str(value)
        return '；'.join(labels.get(k, k) + '：' + display(v) for k, v in values.items())

    def current_key(self):
        item = self.items.currentItem()
        return item.data(Qt.ItemDataRole.UserRole) if item else (None, None)

    def edited(self):
        self.changed = True; self.reviewed.setChecked(False); self.refresh()

    def edit_rule(self):
        kind, key = self.current_key()
        if not key: return
        rule_key = 'comment:' + str(key) if kind == 'comment' else key
        rule = next((r for r in self.review.rules if r['key'] == rule_key), None)
        if rule is None and kind == 'comment':
            c = next(c for c in self.review.comments if c['id'] == key)
            if c.get('spans') and not c.get('point'):
                rule = {'scope': self.workspace.model().scope('selection', spans=c['spans']), 'set': {}}
            elif c.get('point'):
                rule = {'scope': self.workspace.model().scope('paragraphs', pids=[c['spans'][0]['pid']]), 'set': {}}
        dialog = RuleDialog(self.workspace.model(), rule, self.workspace.session.environment.get('font_families', []), parent=self)
        if dialog.exec():
            try:
                self.workspace.set_rule(rule_key, dialog.result_scope, dialog.result_values, exception=dialog.exception.isChecked())
                self.edited()
            except Exception as exc: self.message.setText(str(exc))

    def add_rule(self):
        dialog = RuleDialog(self.workspace.model(), fonts=self.workspace.session.environment.get('font_families', []), parent=self)
        dialog.scope.setCurrentIndex(dialog.scope.count() - 2)
        dialog.exception.setChecked(True)
        if dialog.exec():
            try:
                self.workspace.set_rule('local:' + uuid4().hex[:10], dialog.result_scope, dialog.result_values,
                                        exception=dialog.exception.isChecked())
                self.edited()
            except Exception as exc: self.message.setText(str(exc))

    def reset_rule(self):
        kind, key = self.current_key()
        if not key: return
        try:
            self.workspace.reset_rule('comment:' + str(key) if kind == 'comment' else key)
            self.edited()
        except Exception as exc: self.message.setText(str(exc))

    def classify(self, classification):
        kind, key = self.current_key()
        if kind != 'comment':
            self.message.setText('请选择具体旁批；模板样式冲突须通过格式或分类纠正。'); return
        reason, ok = QInputDialog.getText(self, '说明处理依据', '请填写处理理由（保留原旁批，不会标记自动通过）：')
        if not ok or not reason.strip(): return
        choices = {'正式教案填写时': 'manuscript', '授课结束后': 'after_class', '打印时': 'printing', '模板本身（仍需处理）': 'template'}
        label, ok = QInputDialog.getItem(self, '适用阶段', '这条要求属于哪个阶段？', list(choices), editable=False)
        if not ok: return
        try:
            self.workspace.classify_comment(key, classification, reason=reason.strip(), stage=choices[label])
            self.edited()
        except Exception as exc: self.message.setText(str(exc))

    def manual(self): self.classify('manual')
    def not_applicable(self): self.classify('not_applicable')

    def protect(self):
        kind, key = self.current_key()
        if kind != 'comment': return
        dialog = RuleDialog(self.workspace.model(), scope_only=True, parent=self)
        if not dialog.exec(): return
        reason, ok = QInputDialog.getText(self, '保护范围依据', '说明此范围如何对应旁批要求：')
        if not ok or not reason.strip(): return
        try:
            self.workspace.classify_comment(key, 'protection', reason=reason.strip(), stage='template', scope=dialog.result_scope)
            self.edited()
        except Exception as exc: self.message.setText(str(exc))

    def correct_roles(self, pids, role):
        if not pids or self.job: return
        try:
            self.workspace.set_roles({pid: role for pid in pids}); self.edited()
        except Exception as exc: self.message.setText(str(exc))

    def enable_actions(self):
        busy = self.job is not None
        self.tabs.setEnabled(not busy); self.reviewed.setEnabled(not busy)
        self.cancel_button.setEnabled(not busy)
        self.calibrate_button.setEnabled(not busy and not self.review.blockers and bool(self.review.differences))
        self.save_button.setEnabled(not busy and (self.workspace.session.current == 'original' or self.reviewed.isChecked()))
        self.adopt_button.setEnabled(not busy and self.review.ready and not self.review.differences
                                     and (self.review.basis_kind == 'sample_styles' or self.reviewed.isChecked()))

    def calibrate(self):
        if self.job: return
        self.job = Job(lambda: self.workspace.calibrate(environment=self.workspace.session.environment), self)
        self.job.finished.connect(self.calibrated)
        self.message.setText('正在保存并重新核验模板…'); self.enable_actions(); self.job.start()

    def calibrated(self):
        job, self.job = self.job, None
        try:
            if job.error:
                self.message.setText(str(job.error)); self.enable_actions()
            else:
                self.edited(); self.tabs.setCurrentIndex(2)
        except Exception as exc:
            self.message.setText(str(exc)); self.enable_actions()
        finally:
            job.deleteLater()

    def show_preview(self):
        if not hasattr(self, 'preview'): return
        images = self.preview['images']; index = self.page.value() - 1
        self.page_view.show_page(images[index] if index < len(images) else None, 1.0)

    def locate_paragraph(self, pid):
        row = self.workspace.model().row(pid)
        page = paragraph_page(self.preview, row['text'])
        if page:
            self.page.setValue(page); self.tabs.setCurrentIndex(2)
            self.message.setText('已定位模板预览：' + row['location'])
        else:
            self.message.setText(row['location'] + '；页面文字无法唯一定位，请结合栏目与原文核对。')

    def adopt(self):
        try:
            self.baseline = self.workspace.baseline(reviewed=self.reviewed.isChecked()); self.accept()
        except Exception as exc: self.message.setText(str(exc))

    def save_template(self):
        version = self.workspace.session.versions[self.workspace.session.current]
        path, _ = QFileDialog.getSaveFileName(self, '保存模板文件（完整来源另用模板包导出）', version.name,
                                             f'Word 模板文件 (*.{version.extension})', options=QFileDialog.Option.DontConfirmOverwrite)
        if not path: return
        if not Path(path).suffix: path += '.' + version.extension
        try:
            if Path(path).suffix.lower() != '.' + version.extension: raise FormatError('请保留模板原格式。')
            overwrite = Path(path).exists()
            if overwrite and QMessageBox.question(self, '替换模板文件', '目标文件已存在，是否替换？') != QMessageBox.StandardButton.Yes: return
            source = getattr(self.parent(), 'template_source', None)
            state = getattr(self.parent(), 'state', None)
            protected = [p for p in (source, state.source if state else None) if p]
            save_bytes(path, version.data, overwrite=overwrite, protected=protected)
            self.message.setText('已保存模板文件：' + path)
        except Exception as exc: self.message.setText(str(exc))

    def reject(self):
        if self.job:
            self.message.setText('请等待当前保存与核验结束。'); return
        super().reject()

    def closeEvent(self, event):
        if self.job: event.ignore()
        else: super().closeEvent(event)


class TemplateApplicationDialog(QDialog):
    def __init__(self, baseline, model, parent=None, *, protections=None):
        super().__init__(parent)
        self.baseline, self.model, self.resolutions = baseline, model, {}
        self.protections = protections or []
        self.setWindowTitle('核对模板在当前教案中的范围'); self.resize(750, 550)
        layout = QVBoxLayout(self)
        note = QLabel('模板中的局部位置不能直接套到另一份教案。请为对应要求选择当前文档范围；分类不明时返回主界面更正分类。')
        note.setWordWrap(True); layout.addWidget(note)
        self.items = QListWidget(); layout.addWidget(self.items)
        self.message = QPlainTextEdit(); self.message.setReadOnly(True); layout.addWidget(self.message)
        edit = QPushButton('指定当前教案范围'); edit.clicked.connect(self.resolve); layout.addWidget(edit)
        self.buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.buttons.accepted.connect(self.accept); self.buttons.rejected.connect(self.reject); layout.addWidget(self.buttons)
        self.refresh()

    def refresh(self):
        review = self.baseline.application_review(self.model, resolutions=self.resolutions, protections=self.protections)
        self.items.clear()
        for rule in review.rules:
            role = ROLE_LABELS.get(rule.get('scope', {}).get('role'), '')
            item = QListWidgetItem(role + ' · ' + str(rule.get('source', rule['key'])))
            item.setData(Qt.ItemDataRole.UserRole, rule['key']); self.items.addItem(item)
        self.message.setPlainText('\n'.join(review.blockers) or '作用范围已明确，可以按模板整理。')
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setEnabled(not review.blockers and (review.ready or review.plan is not None))

    def resolve(self):
        item = self.items.currentItem()
        if not item: return
        key = item.data(Qt.ItemDataRole.UserRole)
        dialog = RuleDialog(self.model, scope_only=True, parent=self)
        if dialog.exec():
            self.resolutions[key] = dialog.result_scope; self.refresh()
