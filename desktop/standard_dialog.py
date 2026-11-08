"""The four user-facing teaching-plan rules."""
from PySide6.QtWidgets import (
    QComboBox, QDialog, QDialogButtonBox, QGridLayout, QLabel, QVBoxLayout,
)
from lesson_regions import GROUP_LABELS, validate_settings


class StandardDialog(QDialog):
    def __init__(self, settings, heading_policy, fonts=(), parent=None):
        super().__init__(parent)
        self.setWindowTitle('调整金标准')
        self.setMinimumWidth(560)
        layout = QVBoxLayout(self)
        note = QLabel('按教案区域统一字体与字号；文字、表格结构及其他格式保持原样。')
        note.setWordWrap(True)
        layout.addWidget(note)
        grid = QGridLayout()
        layout.addLayout(grid)
        self.controls = {}
        for row, (key, label) in enumerate(GROUP_LABELS.items()):
            grid.addWidget(QLabel(label), row, 0)
            font, size = QComboBox(), QComboBox()
            font.setAccessibleName(label + '字体')
            size.setAccessibleName(label + '字号')
            font.addItems(sorted(set(fonts) | {settings[key]['font_family']}))
            font.setCurrentText(settings[key]['font_family'])
            for text, value in [('小五 · 9 磅', 9), ('五号 · 10.5 磅', 10.5), ('小四 · 12 磅', 12), ('四号 · 14 磅', 14), ('小三 · 15 磅', 15), ('三号 · 16 磅', 16)]:
                size.addItem(text, value)
            if size.findData(settings[key]['font_size']) < 0:
                size.addItem(str(settings[key]['font_size']) + ' 磅', settings[key]['font_size'])
            size.setCurrentIndex(size.findData(settings[key]['font_size']))
            grid.addWidget(font, row, 1); grid.addWidget(size, row, 2)
            self.controls[key] = (font, size)
        self.headings = QComboBox()
        self.headings.addItem('内容里的标题：保留原有字号（默认）', 'preserve')
        self.headings.addItem('内容里的标题：同教学内容字号', 'same_as_content')
        self.headings.setCurrentIndex(self.headings.findData(heading_policy))
        layout.addWidget(self.headings)
        explanation = QLabel('小结、作业、板书三行的填写内容优先采用最后一组；行头仍采用第一组。\n标题默认作为例外保留，整理后可查看具体位置。')
        explanation.setWordWrap(True)
        layout.addWidget(explanation)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText('应用并预览')
        buttons.accepted.connect(self.accept); buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def values(self):
        settings = {key: {'font_family': font.currentText(), 'font_size': size.currentData()}
                    for key, (font, size) in self.controls.items()}
        policy = self.headings.currentData()
        return validate_settings(settings, policy), policy
