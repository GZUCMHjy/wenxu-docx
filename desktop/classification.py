"""Expandable document roles using native Qt selection and keyboard behavior."""
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QAbstractItemView, QComboBox, QHBoxLayout, QLabel, QPlainTextEdit,
    QPushButton, QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget,
)

from v06_model import ROLE_LABELS
from v06_patch import current_values


class ClassificationPanel(QWidget):
    roles_changed = Signal(list, str)
    locate = Signal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(['内容分类', '位置', '本次范围'])
        self.tree.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.tree.setAccessibleName('按正文和标题分类的具体内容')
        self.tree.setMinimumHeight(130)
        self.tree.setMaximumHeight(240)
        self.tree.setColumnWidth(0, 230)
        self.tree.setColumnWidth(2, 75)
        layout.addWidget(self.tree)
        self.detail = QPlainTextEdit()
        self.detail.setReadOnly(True)
        self.detail.setPlaceholderText('选择内容查看原文与分类依据；双击尝试定位预览。')
        self.detail.setMaximumHeight(95)
        layout.addWidget(self.detail)
        row = QHBoxLayout()
        self.role = QComboBox()
        self.role.setAccessibleName('更正后的内容分类')
        for key, label in ROLE_LABELS.items():
            self.role.addItem(label, key)
        self.correct = QPushButton('更正分类')
        row.addWidget(self.role, 1); row.addWidget(self.correct)
        layout.addLayout(row)
        note = QLabel('分类用于选择修改范围；更正分类不会自动改写 Word 标题样式。')
        note.setWordWrap(True)
        layout.addWidget(note)
        self.tree.itemSelectionChanged.connect(self.selection_changed)
        self.tree.itemDoubleClicked.connect(self.double_clicked)
        self.correct.clicked.connect(lambda: self.roles_changed.emit(self.selected_pids(), self.role.currentData()))
        self.model = None
        self.correct.setEnabled(False)

    def set_model(self, model):
        expanded = {self.tree.topLevelItem(i).data(0, Qt.ItemDataRole.UserRole)
                    for i in range(self.tree.topLevelItemCount()) if self.tree.topLevelItem(i).isExpanded()}
        self.model = model
        self.tree.clear()
        self.detail.clear()
        for role, label in ROLE_LABELS.items():
            rows = [p for p in model.paragraphs if p['role'] == role and p['text'].strip()]
            if not rows:
                continue
            group = QTreeWidgetItem([f'{label}（{len(rows)}）', ''])
            group.setData(0, Qt.ItemDataRole.UserRole, role)
            group.setFlags(group.flags() & ~Qt.ItemFlag.ItemIsSelectable)
            self.tree.addTopLevelItem(group)
            for p in rows:
                item = QTreeWidgetItem([p['text'][:90].replace('\n', ' '), p['location']])
                item.setData(0, Qt.ItemDataRole.UserRole, p['pid'])
                item.setToolTip(0, p['text'] + '\n分类依据：' + p['evidence'])
                group.addChild(item)
            group.setExpanded(role in expanded or role == 'unknown')

    def selected_pids(self):
        return [item.data(0, Qt.ItemDataRole.UserRole) for item in self.tree.selectedItems()
                if type(item.data(0, Qt.ItemDataRole.UserRole)) is int]

    def set_included(self, pids, protected=()):
        for i in range(self.tree.topLevelItemCount()):
            group = self.tree.topLevelItem(i)
            for j in range(group.childCount()):
                item = group.child(j)
                pid = item.data(0, Qt.ItemDataRole.UserRole)
                item.setText(2, ('含保护' if pid in protected else '包含') if pid in pids else '—')

    def selection_changed(self):
        pids = self.selected_pids()
        self.correct.setEnabled(bool(pids))
        if pids and self.model:
            row = self.model.row(pids[0])
            scope = self.model.scope('paragraphs', pids=[pids[0]])
            parts = []
            for prop, label in (('font_east_asia', '中文字体'), ('font_size', '字号（磅）'), ('bold', '加粗')):
                values = current_values(self.model, scope, prop)
                display = ' / '.join('是' if v is True else '否' if v is False else '继承或未指定' if v is None else str(v) for v in values)
                parts.append(label + '：' + display)
            self.detail.setPlainText(row['location'] + '\n依据：' + row['evidence'] + '\n当前格式：'
                                     + '；'.join(parts) + '\n\n' + row['text'])
        else:
            self.detail.clear()

    def double_clicked(self, item, column):
        pid = item.data(0, Qt.ItemDataRole.UserRole)
        if type(pid) is int:
            self.locate.emit(pid)
