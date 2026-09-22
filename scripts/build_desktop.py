"""Build an explicit allowlist of runtime files; never package local documents."""
from hashlib import sha256
from importlib.metadata import distribution, distributions
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys

from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from desktop import VERSION


def main():
    if os.name != "nt":
        raise SystemExit("Build the Windows package on Windows.")
    assets = ROOT / "build/desktop-assets"
    assets.mkdir(parents=True, exist_ok=True)
    icon = Image.new("RGBA", (256, 256))
    painter = ImageDraw.Draw(icon)
    painter.rounded_rectangle((8, 8, 248, 248), radius=52, fill="#2563eb")
    painter.line([(48, 75), (82, 184), (128, 108), (174, 184), (208, 75)], fill="white", width=18, joint="curve")
    icon.save(assets / "wenxu.ico", sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    # DLL resolution must not inherit unrelated toolchains (e.g. Poppler's ICU).
    # Qt uses the Windows ICU ABI; a same-named PATH DLL may be incompatible.
    environment = os.environ.copy()
    windows = Path(environment.get("SystemRoot", r"C:\Windows"))
    environment["PATH"] = os.pathsep.join(map(str, (Path(sys.executable).parent, Path(sys.base_prefix), windows / "System32", windows)))
    environment["PYTHONNOUSERSITE"] = "1"
    for key in ("PYTHONPATH", "PYTHONHOME", "QT_PLUGIN_PATH", "QT_QPA_PLATFORM_PLUGIN_PATH", "QML2_IMPORT_PATH"):
        environment.pop(key, None)
    # Force dependency discovery on each build, including after PATH changes.
    (ROOT / "build/windows/Analysis-00.toc").unlink(missing_ok=True)
    subprocess.run([sys.executable, "-m", "PyInstaller", "--noconfirm",
                    "--workpath", str(ROOT / "build"), "--distpath", str(ROOT / "dist"),
                    str(ROOT / "packaging/windows.spec")], cwd=ROOT, env=environment, check=True)
    target = ROOT / "dist/Wenxu"
    shutil.copyfile(ROOT / "docs/desktop-quickstart.md", target / "使用说明.md")
    licenses = target / "THIRD-PARTY"
    licenses.mkdir(exist_ok=True)
    packages = ("python-docx", "lxml", "PyMuPDF", "olefile", "Pillow", "PySide6-Essentials", "shiboken6", "typing_extensions")
    manifest = {}
    for name in packages:
        dist = distribution(name)
        manifest[name] = dist.version
        dest = licenses / name
        dest.mkdir(exist_ok=True)
        (dest / "METADATA.txt").write_text(dist.read_text("METADATA") or "", encoding="utf-8")
        for entry in dist.files or []:
            if "license" in str(entry).lower() or "copying" in entry.name.lower():
                source = Path(dist.locate_file(entry))
                if source.is_file():
                    relative = Path(*[p for p in entry.parts if p not in {".", ".."}])
                    output = dest / relative
                    output.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copyfile(source, output)
    python_license = Path(sys.base_prefix) / "LICENSE.txt"
    if python_license.exists():
        shutil.copyfile(python_license, licenses / "Python-LICENSE.txt")
    (target / "build-info.json").write_text(json.dumps({
        "version": VERSION, "python": sys.version.split()[0], "dependencies": manifest,
        "build_tools": {d.metadata['Name']: d.version for d in distributions()
                        if d.metadata['Name'].lower() in {"pyinstaller", "pyinstaller-hooks-contrib", "setuptools", "pefile", "altgraph", "pywin32-ctypes", "packaging"}},
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    archive = shutil.make_archive(str(ROOT / f"dist/Wenxu-{VERSION}-windows-x64"), "zip", target.parent, target.name)
    checksum = sha256(Path(archive).read_bytes()).hexdigest()
    Path(archive + ".sha256").write_text(checksum + "  " + Path(archive).name + "\n", encoding="ascii")
    print(json.dumps({"executable": str(target / "Wenxu.exe"), "archive": archive, "sha256": checksum}, indent=2))


if __name__ == "__main__":
    main()
