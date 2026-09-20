"""Requirement extraction, rule review, rendering and delivery report for PRD v0.5."""

from __future__ import annotations

from datetime import datetime, timezone
from hashlib import sha256
from html import escape
from io import BytesIO
import json
import os
from pathlib import Path
import platform
import re
import shutil
import subprocess
import tempfile
import uuid

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from docx_formatting import FormatError, _paragraphs, _runs
from prd_formatting import (
    CHARACTER_PROPERTIES,
    PAGE_PROPERTIES,
    PARAGRAPH_PROPERTIES,
    PROPERTY_LABELS,
    ROLE_LABELS,
    _effective_paragraph_value,
    _effective_run_value,
    current_values,
)


FONT_NAMES = (
    "宋体",
    "黑体",
    "仿宋",
    "楷体",
    "微软雅黑",
    "方正小标宋简体",
    "Noto Serif CJK SC",
    "Noto Sans CJK SC",
    "Liberation Serif",
    "DejaVu Serif",
)

CHINESE_SIZES = {
    "初号": 42,
    "小初": 36,
    "一号": 26,
    "小一": 24,
    "二号": 22,
    "小二": 18,
    "三号": 16,
    "小三": 15,
    "四号": 14,
    "小四": 12,
    "五号": 10.5,
    "小五": 9,
    "六号": 7.5,
    "小六": 6.5,
}

PROPERTY_UNITS = {
    "font_family": "字体名",
    "font_size": "磅",
    "bold": "布尔",
    "italic": "布尔",
    "color": "RGB",
    "alignment": "方式",
    "first_line_indent": "磅",
    "left_indent": "磅",
    "right_indent": "磅",
    "space_before": "磅",
    "space_after": "磅",
    "line_spacing": "固定磅/倍数",
    "keep_with_next": "布尔",
    "top_margin": "磅",
    "bottom_margin": "磅",
    "left_margin": "磅",
    "right_margin": "磅",
}

ROLE_PATTERNS = [
    (r"文档标题|主标题", "title"),
    (r"一级标题|标题一", "heading_1"),
    (r"二级标题|标题二", "heading_2"),
    (r"三级标题|标题三", "heading_3"),
    (r"副标题", "title"),
    (r"标题", "heading_1"),
    (r"正文", "body"),
    (r"落款|署名|日期", "signature"),
    (r"页面|页边距", "page"),
]


def _new_rule(scope, prop, target, unit, source_type, source_detail, *, apply=True, status="待核对", reason=""):
    return {
        "id": uuid.uuid4().hex[:10],
        "apply": apply,
        "scope": scope,
        "property": prop,
        "target": target,
        "unit": unit,
        "source_type": source_type,
        "source_detail": source_detail.strip(),
        "status": status,
        "reason": reason,
    }


def _scope_in(text: str):
    for pattern, scope in ROLE_PATTERNS:
        if re.search(pattern, text):
            return scope
    return ""


def _parse_clause(clause: str, scope: str, source_type: str):
    rows = []
    detail = clause.strip()
    for font in FONT_NAMES:
        if font in clause:
            rows.append(_new_rule(scope, "font_family", font, "字体名", source_type, detail))
            break

    size_match = re.search(r"(?<!行距)(?<!固定值)(?<!缩进)(初号|小初|一号|小一|二号|小二|三号|小三|四号|小四|五号|小五|六号|小六)(?:字)?", clause)
    if size_match:
        rows.append(_new_rule(scope, "font_size", str(CHINESE_SIZES[size_match.group(1)]), "磅", source_type, detail))
    else:
        size_match = re.search(r"(?:字号|字体大小|文字大小)\s*(?:为|改为|设为|设置为)?\s*(\d+(?:\.\d+)?)\s*(磅|pt)", clause, re.I)
        if size_match:
            rows.append(_new_rule(scope, "font_size", size_match.group(1), "磅", source_type, detail))

    exact = re.search(r"(?:固定(?:值)?(?:行距)?|行距(?:为|改为|设为|设置为)?固定(?:值)?)\s*(\d+(?:\.\d+)?)\s*(磅|pt)?", clause, re.I)
    multiple = re.search(r"(?:行距(?:为|改为|设为|设置为)?\s*)?(\d+(?:\.\d+)?)\s*倍(?:行距)?", clause)
    if exact:
        unit = "固定磅" if exact.group(2) else ""
        rows.append(_new_rule(scope, "line_spacing", exact.group(1), unit, source_type, detail))
    elif multiple:
        rows.append(_new_rule(scope, "line_spacing", multiple.group(1), "倍数", source_type, detail))
    elif re.search(r"行距\s*(?:参照|按照|见)\s*(?:附件|实例|范文)", clause):
        rows.append(_new_rule(scope, "line_spacing", "", "", source_type, detail, status="待补充", reason="需要从参考实例确定行距"))
    else:
        incomplete = re.search(r"行距\s*(?:为|改为|设为|设置为)?\s*(\d+(?:\.\d+)?)\s*$", clause)
        if incomplete:
            rows.append(_new_rule(scope, "line_spacing", incomplete.group(1), "", source_type, detail, status="待补充", reason="缺少固定磅或倍数单位"))

    alignment = next(((label, key) for label, key in (("两端对齐", "justify"), ("居中", "center"), ("右对齐", "right"), ("左对齐", "left")) if label in clause), None)
    if alignment:
        rows.append(_new_rule(scope, "alignment", alignment[1], "方式", source_type, detail))

    if "不加粗" in clause or "取消加粗" in clause:
        rows.append(_new_rule(scope, "bold", "否", "布尔", source_type, detail))
    elif "加粗" in clause:
        rows.append(_new_rule(scope, "bold", "是", "布尔", source_type, detail))
    if "取消斜体" in clause or "不斜体" in clause:
        rows.append(_new_rule(scope, "italic", "否", "布尔", source_type, detail))
    elif "斜体" in clause:
        rows.append(_new_rule(scope, "italic", "是", "布尔", source_type, detail))

    color = re.search(r"#([0-9a-fA-F]{6})", clause)
    if color:
        rows.append(_new_rule(scope, "color", "#" + color.group(1).upper(), "RGB", source_type, detail))
    elif "黑色" in clause:
        rows.append(_new_rule(scope, "color", "#000000", "RGB", source_type, detail))
    elif "红色" in clause:
        rows.append(_new_rule(scope, "color", "#FF0000", "RGB", source_type, detail))

    length_specs = (
        (r"首行缩进", "first_line_indent"),
        (r"左缩进", "left_indent"),
        (r"右缩进", "right_indent"),
        (r"段前", "space_before"),
        (r"段后", "space_after"),
        (r"上页边距", "top_margin"),
        (r"下页边距", "bottom_margin"),
        (r"左页边距", "left_margin"),
        (r"右页边距", "right_margin"),
    )
    for pattern, prop in length_specs:
        match = re.search(pattern + r"\s*(?:为|改为|设为|设置为)?\s*(\d+(?:\.\d+)?)\s*(磅|pt|厘米|cm|字符)?", clause, re.I)
        if match:
            unit = (match.group(2) or "").lower()
            rows.append(_new_rule(scope or ("page" if prop in PAGE_PROPERTIES else ""), prop, match.group(1), unit, source_type, detail))

    if "与下段同页" in clause:
        rows.append(_new_rule(scope, "keep_with_next", "是", "布尔", source_type, detail))
    if re.search(r"(?:改写|润色|扩写|缩写|重写)(?:正文|内容)|(?:正文|内容)(?:改写|润色|扩写|缩写|重写)", clause):
        rows.append(_new_rule(scope, "unsupported", detail, "", source_type, detail, status="待补充", reason="正文改写超出格式纠正范围"))
    return rows


def parse_text_requirements(text: str) -> list[dict]:
    """Extract every explicit supported property and retain unparsed clauses as pending rows."""
    rows = []
    for sentence in [part.strip() for part in re.split(r"[\n。；;]+", text or "") if part.strip()]:
        current_scope = _scope_in(sentence)
        clauses = [part.strip() for part in re.split(r"[，,]+", sentence) if part.strip()]
        for clause in clauses:
            current_scope = _scope_in(clause) or current_scope
            parsed = _parse_clause(clause, current_scope, "文字要求")
            if parsed:
                rows.extend(parsed)
            elif not re.fullmatch(r"(?:请|需要|要求|格式|设置|调整|修改)+", clause):
                rows.append(_new_rule(current_scope, "unparsed", clause, "", "文字要求", clause, status="待补充", reason="未能识别该要求，请人工选择属性和值"))
    return rows


def _display_target(prop, value):
    if value is None:
        return "未指定", ""
    if prop == "line_spacing":
        units = {"exact": "固定磅", "multiple": "倍数", "at_least": "至少磅"}
        return str(value["value"]), units.get(value["mode"], "")
    if prop in {"bold", "italic", "keep_with_next"}:
        return "是" if value else "否", "布尔"
    if prop == "font_size":
        return f"{value:g}", "磅"
    if prop in {"first_line_indent", "left_indent", "right_indent", "space_before", "space_after", *PAGE_PROPERTIES}:
        return f"{value:g}", "磅"
    if prop == "color":
        return str(value).upper(), "RGB"
    if prop == "alignment":
        return value, "方式"
    return str(value), PROPERTY_UNITS.get(prop, "")


def extract_reference_requirements(data: bytes) -> list[dict]:
    """Extract observed styles from an example; inconsistent values remain separate candidates."""
    from prd_formatting import build_structure_mapping, preflight_document

    preflight_document(data)
    mapping = build_structure_mapping(data)
    document = Document(BytesIO(data))
    rows = []
    observed = {}
    roles = {int(row["index"]): row["scope"] for row in mapping}
    props = list(CHARACTER_PROPERTIES | PARAGRAPH_PROPERTIES)
    for index, element in enumerate(_paragraphs(document.element), 1):
        role = roles.get(index, "body")
        if role not in {"title", "heading_1", "heading_2", "heading_3", "body", "signature"}:
            continue
        paragraph = Paragraph(element, document)
        text = "".join(element.itertext(tag=qn("w:t"))).strip()
        for prop in props:
            values = []
            if prop in CHARACTER_PROPERTIES:
                for run_element in _runs(element):
                    value = _effective_run_value(Run(run_element, document), paragraph, prop)
                    if value is not None and value not in values:
                        values.append(value)
            else:
                value = _effective_paragraph_value(paragraph, prop)
                if value is not None:
                    values.append(value)
            for value in values:
                key = (role, prop, json.dumps(value, ensure_ascii=False, sort_keys=True))
                observed.setdefault(key, {"value": value, "locations": []})["locations"].append((index, text[:40]))
    for (role, prop, _), item in observed.items():
        target, unit = _display_target(prop, item["value"])
        locations = "；".join(f"第 {index} 段 {text}" for index, text in item["locations"][:4])
        rows.append(_new_rule(role, prop, target, unit, "参考实例", locations))
    return rows


def _to_bool(value):
    if isinstance(value, bool):
        return value
    value = str(value).strip().casefold()
    if value in {"是", "true", "1", "yes", "加粗", "斜体"}:
        return True
    if value in {"否", "false", "0", "no", "不加粗", "不斜体"}:
        return False
    raise ValueError("请输入是或否")


def _normalize_length(value, unit):
    number = float(value)
    if not math_is_finite(number) or number < 0 or number > 720:
        raise ValueError("数值须在 0–720 之间")
    unit = str(unit).strip().casefold()
    if unit in {"磅", "pt"}:
        return number
    if unit in {"厘米", "cm"}:
        return number * 72 / 2.54
    if unit == "字符":
        raise ValueError("字符单位需结合字号换算；请改用磅或厘米")
    raise ValueError("缺少磅或厘米单位")


def math_is_finite(number):
    return number == number and number not in {float("inf"), float("-inf")}


def normalize_rules(rows: list[dict], manuscript: bytes | None = None, mapping: list[dict] | None = None):
    """Derive current values/status, merge identical rules and expose every blocker."""
    normalized = []
    for position, original in enumerate(rows):
        row = dict(original)
        row.setdefault("id", uuid.uuid4().hex[:10])
        row["apply"] = bool(row.get("apply", True))
        scope, prop = str(row.get("scope", "")).strip(), str(row.get("property", "")).strip()
        target, unit = row.get("target", ""), str(row.get("unit", "")).strip()
        blockers = []
        if not row["apply"]:
            row["status"] = "本次不应用"
            row["normalized_target"] = None
            normalized.append(row)
            continue
        if scope not in ROLE_LABELS or scope == "preserve":
            blockers.append("请选择有效作用对象")
        if prop not in PROPERTY_LABELS:
            blockers.append(row.get("reason") or "请选择当前支持的格式属性")
        elif scope == "page" and prop not in PAGE_PROPERTIES:
            blockers.append("页面对象只能设置上下左右页边距")
        elif prop in PAGE_PROPERTIES and scope != "page":
            blockers.append("页边距属性只能作用于页面")
        value = None
        if not blockers:
            try:
                if prop == "font_family":
                    value = str(target).strip()
                    if not value:
                        raise ValueError("字体名称不能为空")
                elif prop == "font_size":
                    value = float(target)
                    if not 6 <= value <= 72 or value * 2 != int(value * 2):
                        raise ValueError("字号须为 6–72 磅，步长 0.5")
                    if unit.casefold() not in {"磅", "pt"}:
                        raise ValueError("字号单位须为磅")
                elif prop == "color":
                    value = str(target).strip().upper()
                    if not re.fullmatch(r"#[0-9A-F]{6}", value):
                        raise ValueError("颜色须为 #RRGGBB")
                elif prop in {"bold", "italic", "keep_with_next"}:
                    value = _to_bool(target)
                elif prop == "alignment":
                    value = str(target).strip().casefold()
                    mapping_names = {"左对齐": "left", "居中": "center", "右对齐": "right", "两端对齐": "justify"}
                    value = mapping_names.get(value, value)
                    if value not in {"left", "center", "right", "justify"}:
                        raise ValueError("对齐值须为左对齐、居中、右对齐或两端对齐")
                elif prop == "line_spacing":
                    number = float(target)
                    if unit in {"固定磅", "磅", "pt"}:
                        if not 6 <= number <= 144:
                            raise ValueError("固定行距须为 6–144 磅")
                        value = {"mode": "exact", "value": number}
                    elif unit in {"倍数", "倍"}:
                        if not 0.5 <= number <= 5:
                            raise ValueError("倍数行距须为 0.5–5 倍")
                        value = {"mode": "multiple", "value": number}
                    else:
                        raise ValueError("行距须注明固定磅或倍数")
                elif prop in {"first_line_indent", "left_indent", "right_indent", "space_before", "space_after", *PAGE_PROPERTIES}:
                    value = _normalize_length(target, unit)
            except (TypeError, ValueError) as exc:
                blockers.append(str(exc))
        row["normalized_target"] = value
        row["reason"] = "；".join(blockers)
        row["status"] = "待补充" if blockers else "待核对"
        if manuscript is not None and mapping is not None and not blockers:
            values = current_values(manuscript, mapping, scope, prop)
            row["current"] = "存在多种格式：" + " / ".join(format_value(prop, item) for item in values) if len(values) > 1 else (format_value(prop, values[0]) if values else "未指定")
        else:
            row["current"] = row.get("current", "待导入稿件后读取")
        normalized.append(row)

    groups = {}
    for row in normalized:
        if row["apply"] and not row["reason"] and row["property"] in PROPERTY_LABELS:
            groups.setdefault((row["scope"], row["property"]), []).append(row)
    for group in groups.values():
        values = {json.dumps(row["normalized_target"], ensure_ascii=False, sort_keys=True) for row in group}
        if len(values) > 1:
            for row in group:
                row["status"] = "冲突"
                row["reason"] = "同一对象和属性存在不同目标值；请统一目标值或取消不应用项"
        elif len(group) > 1:
            first = group[0]
            first["source_detail"] = "；".join(dict.fromkeys(row.get("source_detail", "") for row in group if row.get("source_detail")))
            for duplicate in group[1:]:
                duplicate["apply"] = False
                duplicate["status"] = "本次不应用"
                duplicate["reason"] = "与上一条相同，来源已合并"
    return normalized


def executable_rules(rows: list[dict]):
    selected = [row for row in rows if row.get("apply")]
    blockers = [row for row in selected if row.get("status") in {"待补充", "冲突"} or row.get("reason")]
    return selected, blockers


def format_value(prop, value):
    if value is None:
        return "未指定"
    if prop == "line_spacing":
        if value["mode"] == "exact":
            return f"固定 {value['value']:g} 磅"
        if value["mode"] == "at_least":
            return f"至少 {value['value']:g} 磅"
        return f"{value['value']:g} 倍"
    if prop in {"bold", "italic", "keep_with_next"}:
        return "是" if value else "否"
    if prop == "font_size":
        name = next((label for label, points in CHINESE_SIZES.items() if points == value), "")
        return f"{name}（{value:g} 磅）" if name else f"{value:g} 磅"
    if prop in {"first_line_indent", "left_indent", "right_indent", "space_before", "space_after", *PAGE_PROPERTIES}:
        return f"{value:g} 磅"
    return str(value)


def rule_fingerprint(input_hash: str, rules: list[dict], mapping: list[dict], engine_version="prd-poc-1"):
    payload = {
        "input": input_hash,
        "rules": [{key: row.get(key) for key in ("id", "apply", "scope", "property", "normalized_target", "source_type", "source_detail", "reason")} for row in rules],
        "mapping": [{key: row.get(key) for key in ("index", "scope", "excluded_reason")} for row in mapping],
        "engine": engine_version,
    }
    return sha256(json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode()).hexdigest()


def installed_font_families():
    command = shutil.which("fc-list")
    if not command:
        return set()
    result = subprocess.run([command, ":", "family"], capture_output=True, text=True, timeout=20, check=False)
    families = set()
    for line in result.stdout.splitlines():
        for family in line.split(","):
            if family.strip():
                families.add(family.strip().casefold())
    return families


def missing_required_fonts(rules: list[dict]):
    installed = installed_font_families()
    if not installed:
        return sorted({str(row["normalized_target"]) for row in rules if row.get("apply") and row.get("property") == "font_family"})
    return sorted({str(row["normalized_target"]) for row in rules if row.get("apply") and row.get("property") == "font_family" and str(row["normalized_target"]).casefold() not in installed})


def render_docx(data: bytes, label="document", *, max_pages=50):
    """Render DOCX with Linux LibreOffice and rasterize its PDF in-process."""
    configured_office = os.environ.get("WENXU_LIBREOFFICE", "").strip()
    office = configured_office or shutil.which("libreoffice") or shutil.which("soffice")
    if not office or not Path(office).is_file():
        raise FormatError("缺少 Linux LibreOffice；请安装它或设置 WENXU_LIBREOFFICE。")
    try:
        import pymupdf
    except ImportError as exc:
        raise FormatError("缺少 PyMuPDF，无法生成逐页预览。") from exc
    safe = re.sub(r"[^A-Za-z0-9_-]", "_", label)[:40] or "document"
    with tempfile.TemporaryDirectory(prefix="wenxu-render-") as temp_name:
        temp = Path(temp_name)
        source = temp / f"{safe}.docx"
        source.write_bytes(data)
        profile = temp / "profile"
        profile.mkdir()
        process = subprocess.run(
            [office, "--headless", f"-env:UserInstallation=file://{profile}", "--convert-to", "pdf", "--outdir", str(temp), str(source)],
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
        pdf = temp / f"{safe}.pdf"
        if process.returncode or not pdf.exists() or not pdf.stat().st_size:
            raise FormatError("LibreOffice 渲染失败，未生成可检查的 PDF。")
        pdf_bytes = pdf.read_bytes()
        try:
            rendered = pymupdf.open(stream=pdf_bytes, filetype="pdf")
            pages = rendered.page_count
            if not 1 <= pages <= max_pages:
                raise FormatError(f"渲染页数为 {pages}，POC 仅支持 1–{max_pages} 页。")
            matrix = pymupdf.Matrix(120 / 72, 120 / 72)
            images = [page.get_pixmap(matrix=matrix, alpha=False).tobytes("png") for page in rendered]
            rendered.close()
        except FormatError:
            raise
        except Exception as exc:
            raise FormatError("PDF 页面图生成失败，无法完成视觉预览检查。") from exc
        return {
            "pdf": pdf_bytes,
            "images": images,
            "pages": pages,
            "renderer": process.stdout.strip() or f"LibreOffice headless ({Path(office).name})",
        }


def build_html_report(*, task_id, input_name, input_hash, output_name, parent_version, rules, mapping, summary, checks, environment, warnings=None):
    warnings = warnings or []
    now = datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")
    applied = [row for row in rules if row.get("apply")]
    excluded = [row for row in rules if not row.get("apply")]
    preserved = [row for row in mapping if row.get("scope") == "preserve"]
    rows = "".join(
        "<tr>"
        + "".join(
            f"<td>{escape(str(value))}</td>"
            for value in (
                ROLE_LABELS.get(rule["scope"], rule["scope"]),
                PROPERTY_LABELS.get(rule["property"], rule["property"]),
                rule.get("current", ""),
                format_value(rule["property"], rule.get("normalized_target")),
                rule.get("source_type", ""),
                rule.get("source_detail", ""),
            )
        )
        + "</tr>"
        for rule in applied
    )
    check_rows = "".join(f"<li class='{escape(item['level'])}'><strong>{escape(item['name'])}</strong>：{escape(item['detail'])}</li>" for item in checks)
    warning_rows = "".join(f"<li>{escape(item)}</li>" for item in warnings) or "<li>无</li>"
    excluded_rows = "".join(f"<li>{escape(ROLE_LABELS.get(item.get('scope'), str(item.get('scope'))))} · {escape(PROPERTY_LABELS.get(item.get('property'), str(item.get('property'))))}：{escape(item.get('reason') or '用户选择本次不应用')}</li>" for item in excluded) or "<li>无</li>"
    preserved_rows = "".join(f"<li>第 {item['index']} 段：{escape(item.get('text', '')[:120])}；{escape(item.get('excluded_reason') or '用户指定原样保留')}</li>" for item in preserved) or "<li>无</li>"
    return f"""<!doctype html><html lang='zh-CN'><head><meta charset='utf-8'><title>{escape(output_name)} 格式检查报告</title>
<style>body{{font:15px/1.65 system-ui,sans-serif;max-width:1120px;margin:36px auto;padding:0 24px;color:#17202a}}h1,h2{{color:#111}}table{{width:100%;border-collapse:collapse}}th,td{{border:1px solid #d8dee4;padding:8px;vertical-align:top}}th{{background:#f3f6f8}}.pass{{color:#08783e}}.warning{{color:#8a5a00}}.block{{color:#b42318}}code{{word-break:break-all}}.meta{{display:grid;grid-template-columns:180px 1fr;gap:6px}}</style></head><body>
<h1>文序格式检查报告</h1><div class='meta'><strong>任务 ID</strong><code>{escape(task_id)}</code><strong>处理时间</strong><span>{escape(now)}</span><strong>输入文件</strong><span>{escape(input_name)}</span><strong>输入指纹</strong><code>{escape(input_hash)}</code><strong>来源版本</strong><span>{escape(parent_version)}</span><strong>输出文件</strong><span>{escape(output_name)}</span><strong>处理环境</strong><span>{escape(environment)}</span></div>
<h2>已应用要求</h2><table><thead><tr><th>对象</th><th>属性</th><th>输入当前值</th><th>目标值</th><th>来源</th><th>依据</th></tr></thead><tbody>{rows}</tbody></table>
<h2>检查结果</h2><ul>{check_rows}</ul><h2>已确认警告</h2><ul>{warning_rows}</ul><h2>本次不应用要求</h2><ul>{excluded_rows}</ul><h2>未应用规范区域</h2><ul>{preserved_rows}</ul>
<h2>变更摘要</h2><p>变更文字片段 {summary.get('changed_runs', 0)} 个，变更段落 {summary.get('changed_paragraphs', 0)} 个。正文未插入系统水印或检查摘要。</p></body></html>""".encode("utf-8")


def runtime_label(render_info=None):
    renderer = render_info.get("renderer", "未渲染") if render_info else "未渲染"
    return f"{platform.system()} {platform.release()} · Python {platform.python_version()} · {renderer}"
