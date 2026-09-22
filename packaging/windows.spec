# Build on Windows using the isolated, pinned desktop environment.
from pathlib import Path
import os
import sys

root = Path(SPECPATH).parent
a = Analysis(
    [str(root / "desktop_app.py")],
    pathex=[str(root)],
    binaries=[],
    datas=[(str(root / "office_bridge.ps1"), ".")],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=["streamlit", "tkinter", "PySide6.QtQml", "PySide6.QtQuick", "PySide6.QtWebEngineCore"],
    noarchive=False,
)
allowed = [Path(sys.prefix).resolve(), Path(sys.base_prefix).resolve(),
           Path(os.environ.get("SystemRoot", r"C:\Windows")).resolve()]
for destination, source, kind in a.binaries:
    if not any(Path(source).resolve().is_relative_to(base) for base in allowed):
        raise RuntimeError("Unexpected binary outside the Python/Windows runtime: " + source)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name="Wenxu",
    debug=False,
    strip=False,
    upx=False,
    console=False,
    version=str(root / "packaging/windows-version.txt"),
    icon=str(root / "build/desktop-assets/wenxu.ico"),
)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name="Wenxu")
