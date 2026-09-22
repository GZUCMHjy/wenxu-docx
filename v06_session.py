"""Local-only import, independent rendering, atomic versions and delivery."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from html import escape
from io import BytesIO
import json
import os
from pathlib import Path
import platform
import shutil
import subprocess
import tempfile

from docx.oxml.ns import qn

from docx_formatting import FormatError, MAX_FILE_BYTES, NS, _xml
from prd_formatting import safe_stem
from prd_workflow import installed_font_families
from v06_model import build_model, digest
from v06_import import IMPORT_VERSION, compare_rendered_pages, compare_rendered_text
from v06_office import native_office
from v06_native import NATIVE_VERSION, edit_native_doc, inspect_native_doc
from v06_patch import ENGINE_VERSION, patch_document, verify_result
from v06_plan import encode


def _office():
    executable = os.environ.get("WENXU_LIBREOFFICE") or shutil.which("libreoffice") or shutil.which("soffice")
    if platform.system() != "Linux" or not executable or not Path(executable).is_file():
        raise FormatError("请在 WSL/Linux 中运行并设置 WENXU_LIBREOFFICE。")
    return executable


def render_environment():
    native = native_office()
    if native:
        return native.environment()
    executable = _office()
    result = subprocess.run([executable, "--version"], capture_output=True, text=True, timeout=20)
    if result.returncode:
        raise FormatError("无法读取渲染器版本。")
    fonts = sorted(installed_font_families())
    return {"platform": platform.platform(), "renderer": result.stdout.strip(), "font_families": fonts, "parameters": "120dpi/max50/macro-disabled/no-link-update", "python": platform.python_version()}


def _convert(data, extension, target):
    native = native_office()
    if native:
        return native.convert(data, extension, target)
    return _libreoffice_convert(data, extension, target)


def _libreoffice_convert(data, extension, target):
    office = _office()
    with tempfile.TemporaryDirectory(prefix="wenxu-v06-") as tmp:
        root = Path(tmp)
        profile = root / "profile"
        (profile / "user").mkdir(parents=True)
        # Prevent document macros and automatic link updates in the disposable
        # rendering profile. No user desktop profile is read or changed.
        (profile / "user" / "registrymodifications.xcu").write_text('''<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry">
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop></item>
<item oor:path="/org.openoffice.Office.Writer/Content/Update"><prop oor:name="Link" oor:op="fuse"><value>2</value></prop></item>
<item oor:path="/org.openoffice.Office.Writer/Content/Update"><prop oor:name="Field" oor:op="fuse"><value>false</value></prop><prop oor:name="Chart" oor:op="fuse"><value>false</value></prop></item>
</oor:items>''', encoding="utf-8")
        source = root / ("source." + extension)
        source.write_bytes(data)
        try:
            process = subprocess.run([office, "--headless", "--nologo", "--nodefault", "--norestore", f"-env:UserInstallation={profile.as_uri()}", "--convert-to", target, "--outdir", str(root), str(source)], capture_output=True, timeout=120)
        except (subprocess.TimeoutExpired, OSError) as exc:
            raise FormatError("文件转换超时或渲染器启动失败；原稿保留。") from exc
        output = root / ("source." + target.split(":", 1)[0])
        if process.returncode or not output.is_file() or output.stat().st_size == 0:
            raise FormatError("转换或渲染失败，未产生可验证结果。")
        return output.read_bytes()


def render_document(data, extension="docx"):
    import pymupdf
    pdf = _convert(data, extension, "pdf")
    try:
        with pymupdf.open(stream=pdf, filetype="pdf") as doc:
            if not 1 <= len(doc) <= 50:
                raise FormatError("仅支持渲染后 1–50 页的稿件。")
            return {"pdf": pdf, "pages": len(doc), "images": [p.get_pixmap(matrix=pymupdf.Matrix(120/72, 120/72), alpha=False).tobytes("png") for p in doc], "text": "\n".join(p.get_text() for p in doc)}
    except FormatError:
        raise
    except Exception as exc:
        raise FormatError("无法读取 PDF 或生成页面预览。") from exc


def validate_doc(data):
    if not data or len(data) > MAX_FILE_BYTES or not data.startswith(bytes.fromhex("D0CF11E0A1B11AE1")):
        raise FormatError("文件不是支持的二进制 DOC，或超过 10 MB。")
    import olefile
    try:
        with olefile.OleFileIO(BytesIO(data)) as ole:
            names = ["/".join(p) for p in ole.listdir()]
            if not ole.exists("WordDocument"):
                raise FormatError("未发现 WordDocument 流；不能按 DOC 读取。")
            if any(any(x in n.lower().split("/") for x in ("vba", "macros", "encryptedpackage", "encryptioninfo")) for n in names):
                raise FormatError("不处理带宏或加密的 DOC。")
            header = ole.openstream("WordDocument").read(32)
            if len(header) < 12 or int.from_bytes(header[10:12], "little") & 0x8100:
                raise FormatError("DOC 加密或混淆保护，无法读取。")
            return {"ole_streams": len(names), "embedded_streams": sum(n.startswith("ObjectPool/") for n in names)}
    except FormatError:
        raise
    except Exception as exc:
        raise FormatError("DOC 容器损坏，无法可靠读取。") from exc


def source_font_warnings(model, environment):
    declared = set()
    for name, raw in model.parts.items():
        if name.endswith(".xml") and name.startswith("word/"):
            for fonts in _xml(raw).findall(".//w:rFonts", NS):
                for attr in ("eastAsia", "ascii", "hAnsi", "cs"):
                    if fonts.get(qn("w:" + attr)):
                        declared.add(fonts.get(qn("w:" + attr)))
    installed = {f.casefold() for f in environment.get("font_families", [])}
    missing = sorted(f for f in declared if f.casefold() not in installed)
    return ["文件声明了当前环境未安装的字体，预览可能使用替代字体（含未使用样式声明）：" + "、".join(missing)] if missing else []


@dataclass(frozen=True)
class Version:
    id: str
    name: str
    data: bytes
    parent: str | None
    plan_hash: str | None = None
    report_json: str = "{}"
    protections_json: str = "[]"
    roles_json: str = "{}"
    analysis_data: bytes | None = None
    native_map_json: str = '[]'

    @property
    def model_data(self):
        return self.analysis_data if self.analysis_data is not None else self.data

    @property
    def extension(self):
        return Path(self.name).suffix.lstrip('.').lower()


@dataclass
class EditingSession:
    original_name: str
    original: bytes
    versions: dict
    current: str = "original"
    conversion: dict = field(default_factory=dict)
    renders: dict = field(default_factory=dict)
    jobs: dict = field(default_factory=dict)
    environment: dict = field(default_factory=dict)
    next_number: int = 1
    renderer: object = render_document

    @classmethod
    def load(cls, name, data, *, renderer=render_document, inspector=inspect_native_doc):
        extension = Path(name).suffix.lower()
        conversion = {"status": "not_needed", "visual_review": "not_needed"}
        renders = {}
        environment = render_environment() if renderer is render_document else {}
        baseline = data
        analysis_data = None
        mapping = []
        if extension == ".doc":
            inventory = validate_doc(data)
            analysis_data, mapping = inspector(data)
            indexed = build_model(analysis_data)
            # The editable artifact is the exact original DOC bytes. Structural
            # snapshots are indexes only; neither preview nor export uses them.
            old, new = renderer(data, "doc"), renderer(baseline, "doc")
            text_check = compare_rendered_text(old, new)
            page_check = compare_rendered_pages(old, new)
            blockers = [] if page_check['pixels_equal'] and text_check['sequence_equal'] else ['同一原格式文件的两次独立渲染不一致，已停止编辑；原稿保留。']
            conversion = {"status": "original_format", "visual_review": "not_needed", "import_version": IMPORT_VERSION,
                          "original_hash": digest(data), "working_hash": digest(baseline), "byte_identical": data == baseline,
                          "source_pages": old['pages'], "working_pages": new['pages'], "source_inventory": inventory,
                          "working_inventory": indexed.inventory, "rendered_text_equal": text_check['sequence_equal'],
                          "text_check": text_check, "page_check": page_check, "blocking_issues": blockers,
                          "renderer": environment.get('renderer', 'test renderer'), "format": 'doc'}
            renders["conversion_source"], renders["conversion_working"] = old, new
        elif extension != ".docx":
            raise FormatError("只支持 DOCX 和 DOC 原格式稿件。")
        build_model(analysis_data if analysis_data is not None else baseline)
        stem, number = safe_stem(name)
        version = Version("original", stem + ('_工作副本.doc' if extension == '.doc' else '.docx'), baseline, None,
                          analysis_data=analysis_data, native_map_json=encode(mapping))
        return cls(name, data, {"original": version}, conversion=conversion, renders=renders, next_number=number, renderer=renderer, environment=environment)

    def model(self):
        v = self.versions[self.current]
        return build_model(v.model_data, v.id, {int(k): value for k, value in json.loads(v.roles_json).items()})

    def select(self, ident):
        if ident not in self.versions:
            raise FormatError("所选版本不存在。")
        self.current = ident

    def undo(self):
        self.current = self.versions[self.current].parent or self.current

    def preview(self, version_id=None):
        v = self.versions[version_id or self.current]
        env = self.environment or render_environment()
        key = digest((digest(v.data) + encode(env)).encode())
        if key not in self.renders:
            self.renders[key] = self.renderer(v.data, v.extension)
        return self.renders[key]

    def execute(self, model, plan, *, environment=None):
        source = self.versions[self.current]
        if model.revision != self.current or model.data != source.model_data:
            raise FormatError("当前输入已切换，旧计划已失效。")
        if self.conversion.get("blocking_issues"):
            raise FormatError("；".join(self.conversion["blocking_issues"]))
        env = environment or render_environment()
        frozen = plan.read()
        installed = {f.casefold() for f in env.get("font_families", [])}
        requested = {a["value"] for a in frozen["assignments"] if a["property"].startswith("font_") and a["property"] != "font_size"}
        missing = sorted(f for f in requested if f.casefold() not in installed)
        if missing:
            raise FormatError("缺少要求字体：" + "、".join(missing) + "。未创建输出版本。")
        key = digest((plan.plan_hash + ENGINE_VERSION + NATIVE_VERSION + digest(source.data) + encode(env)).encode())
        if key in self.jobs:
            return self.versions[self.jobs[key]], False
        output, operations = patch_document(model, plan)
        checks = verify_result(model, output, plan)
        # For DOC, model_data is a read-only snapshot of these exact native
        # bytes. Already-satisfied attributes require no save: preserving the
        # original bytes is intentional, and baseline download remains enabled.
        if not operations:
            return None, False
        analysis_data = None
        mapping = []
        if source.extension == 'doc':
            output, analysis_data, mapping, checks = edit_native_doc(source.data, model, plan, json.loads(source.native_map_json))
            validate_doc(output)
        # Structural verification always precedes expensive rendering and commit.
        old_key = digest((digest(source.data) + encode(env)).encode())
        new_key = digest((digest(output) + encode(env)).encode())
        before = self.renders.get(old_key) or self.renderer(source.data, source.extension)
        after = self.renders.get(new_key) or self.renderer(output, source.extension)
        checks["render"] = "passed"
        checks["delivery_status"] = "pending_visual_review"
        v_id = "v" + str(self.next_number)
        stem, _ = safe_stem(self.original_name)
        report = {"version": v_id, "parent": self.current, "input_hash": digest(source.data), "index_hash": model.source_hash, "output_hash": digest(output), "source_format": source.extension, "output_format": source.extension, "plan_hash": plan.plan_hash, "engine_version": NATIVE_VERSION if source.extension == 'doc' else ENGINE_VERSION, "created_at": datetime.now(timezone.utc).isoformat(), "plan": frozen, "operations": operations, "checks": checks, "environment": env, "font_warnings": source_font_warnings(model, env), "pages_before": before["pages"], "pages_after": after["pages"], "conversion": self.conversion.copy()}
        protections = [{**p, "revision": v_id} for p in frozen["protections"]]
        version = Version(v_id, f"{stem}_{v_id}.{source.extension}", output, self.current, plan.plan_hash, encode(report), encode(protections), encode(model.role_overrides), analysis_data, encode(mapping))
        self.versions[v_id] = version
        self.jobs[key] = v_id
        self.renders[old_key], self.renders[new_key] = before, after
        self.environment = env
        self.next_number += 1
        return version, True


def report_html(version, *, visual_review="pending"):
    report = json.loads(version.report_json)
    report["checks"]["visual_review"] = visual_review
    report["checks"]["delivery_status"] = "user_reviewed" if visual_review == "confirmed" else "pending_visual_review"
    return ("<!doctype html><html lang='zh-CN'><meta charset='utf-8'><title>文序格式检查报告</title><style>body{font:15px/1.65 system-ui;max-width:1000px;margin:40px auto;padding:0 24px}pre{white-space:pre-wrap;overflow-wrap:anywhere;background:#f6f8fa;padding:20px}</style><h1>文序格式检查报告</h1><p>结构、目标属性、对象保全和渲染分别记录。人工版式复核状态：" + escape(visual_review) + "。客户端兼容性未运行。</p><pre>" + escape(json.dumps(report, ensure_ascii=False, indent=2)) + "</pre></html>").encode("utf-8")
