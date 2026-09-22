"""File lifecycle and scope choices; no GUI or document rewriting here."""
from __future__ import annotations

from dataclasses import dataclass, field
import os
from pathlib import Path
import tempfile

from docx_formatting import FormatError, MAX_FILE_BYTES
from v06_model import ROLE_LABELS
from v06_office import native_office
from v06_session import EditingSession


def require_office():
    # The Windows edition uses the same engine for every stage of a session.
    if os.name != "nt":
        raise FormatError("此版本面向 Windows，请使用 Windows 桌面程序。")
    os.environ["WENXU_RENDERER"] = "wps"
    native_office.cache_clear()
    try:
        office = native_office()
        if office is None:
            raise FormatError("WPS 接口不可用。")
        return office.environment()
    except FormatError as exc:
        raise FormatError("文序需要本机 WPS 的文档处理接口。请安装或修复 WPS 后重试。\n" + str(exc)) from exc


def scope_choices(model):
    """Only show ranges present in this version, paired with visible context."""
    roles = {p["role"] for p in model.paragraphs if p["text"].strip()}
    choices = [(ROLE_LABELS[r], model.scope("role", role=r))
               for r in ("body", "title", "heading_1", "heading_2", "heading_3") if r in roles]
    if any(c["text"].strip() for c in model.cells):
        choices.append(("全部表格文字", model.scope("table", table="all")))
        for table in sorted({c["table"] for c in model.cells}):
            choices.append((f"表 {table}", model.scope("table", table=table)))
    choices.append(("全文文字", model.scope("all")))
    for column in model.columns:
        choices.append(("栏目 · " + column["label"], model.scope("named_column", column_id=column["id"])))
    return choices


@dataclass
class DocumentState:
    source: Path
    session: EditingSession
    saved: set[str] = field(default_factory=set)
    reviewed: set[str] = field(default_factory=set)

    @classmethod
    def open(cls, path, *, loader=EditingSession.load):
        source = Path(path).resolve(strict=True)
        if source.suffix.lower() not in {".doc", ".docx"}:
            raise FormatError("请选择 DOC 或 DOCX 文件。")
        # Bound the actual read as well as the size check (a file may change).
        with source.open("rb") as handle:
            data = handle.read(MAX_FILE_BYTES + 1)
        if not data or len(data) > MAX_FILE_BYTES:
            raise FormatError("文件为空或超过 10 MB。")
        session = loader(source.name, data)
        if session.conversion.get("blocking_issues"):
            raise FormatError("；".join(session.conversion["blocking_issues"]))
        session.preview()  # A failed preview must not replace the open session.
        return cls(source, session)

    @property
    def unsaved(self):
        return bool(set(self.session.versions) - {"original"} - self.saved)

    def save(self, path, *, overwrite=False):
        version = self.session.versions[self.session.current]
        if version.id != "original" and version.id not in self.reviewed:
            raise FormatError("请先核对当前版本的预览。")
        target = Path(path).absolute()
        if target.suffix.lower() != "." + version.extension:
            raise FormatError("请保留原文件格式：." + version.extension)
        if (target.resolve() == self.source or
                (target.exists() and self.source.exists() and os.path.samefile(target, self.source))):
            raise FormatError("请使用其他文件名；文序始终保留原稿。")
        # Do not follow a destination link during replace or hide its identity.
        if target.is_symlink():
            raise FormatError("保存位置是文件链接，请选择普通文件路径。")
        if target.exists() and not overwrite:
            raise FormatError("目标文件已存在，请确认覆盖或使用其他文件名。")
        temporary = None
        try:
            with tempfile.NamedTemporaryFile(prefix=".wenxu-", suffix=".tmp", dir=target.parent, delete=False) as handle:
                temporary = Path(handle.name)
                handle.write(version.data)
                handle.flush()
                os.fsync(handle.fileno())
            # Windows rename fails if a new target appeared since the dialog.
            (os.replace if overwrite else os.rename)(temporary, target)
        finally:
            if temporary is not None:
                temporary.unlink(missing_ok=True)
        self.saved.add(version.id)
        return target
