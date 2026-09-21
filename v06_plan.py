"""Deterministic input adapters and immutable object/property editing plans."""
from __future__ import annotations

from dataclasses import dataclass
import json
import math
import re

from docx_formatting import FormatError
from prd_formatting import PARAGRAPH_PROPERTIES, PAGE_PROPERTIES
from prd_workflow import CHINESE_SIZES, FONT_NAMES
from v06_model import DocumentModel, MODEL_VERSION, digest

CHAR_PROPS = {"font_family", "font_east_asia", "font_western", "font_size", "bold", "italic", "underline", "color"}
PROPERTIES = CHAR_PROPS | PARAGRAPH_PROPERTIES | PAGE_PROPERTIES


def encode(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def validate_values(values):
    if not isinstance(values, dict) or not values or set(values) - PROPERTIES:
        raise FormatError("请选择至少一个支持的属性；未指定属性保持原样。")
    result = {}
    for key, value in values.items():
        if value is None:
            raise FormatError("保持原样请移除该属性，不使用空值。")
        if key in {"font_family", "font_east_asia", "font_western"}:
            if not isinstance(value, str) or not value.strip() or len(value) > 100:
                raise FormatError("字体名称无效。")
            result[key] = value.strip()
        elif key in {"bold", "italic", "underline", "keep_with_next"}:
            if type(value) is not bool:
                raise FormatError("开关属性必须明确开启或关闭。")
            result[key] = value
        elif key == "color":
            if not isinstance(value, str) or not re.fullmatch(r"#[0-9a-fA-F]{6}", value):
                raise FormatError("颜色必须为 #RRGGBB。")
            result[key] = value.upper()
        elif key == "alignment":
            if value not in {"left", "right", "center", "justify"}:
                raise FormatError("对齐方式无效。")
            result[key] = value
        elif key == "line_spacing":
            if not isinstance(value, dict) or set(value) != {"mode", "value"} or value["mode"] not in {"exact", "multiple"}:
                raise FormatError("行距必须明确固定磅值或倍数。")
            n = value["value"]
            bounds = (6, 144) if value["mode"] == "exact" else (0.5, 5)
            if type(n) not in {int, float} or not math.isfinite(n) or not bounds[0] <= n <= bounds[1]:
                raise FormatError("行距超出支持范围。")
            result[key] = dict(value)
        else:
            bounds = (6, 72) if key == "font_size" else (0, 720)
            if type(value) not in {int, float} or not math.isfinite(value) or not bounds[0] <= value <= bounds[1]:
                raise FormatError("字号或距离超出支持范围。")
            if key == "font_size" and value * 2 != round(value * 2):
                raise FormatError("字号步长为 0.5 磅。")
            result[key] = float(value)
    # Use script-specific atomic properties so conflicting font controls overlap.
    if "font_family" in result:
        family = result.pop("font_family")
        for script in ("font_east_asia", "font_western"):
            if script in result and result[script] != family:
                raise FormatError("同时设置全部字体和不同的分文字字体，请拆成明确例外。")
            result[script] = family
    return result


def intent(scope, values, *, source="控件", exception=False, active=True):
    return {"scope": scope, "set": values, "source": source, "exception": exception, "active": active}


def resolve_scope(model, scope):
    if scope.get("revision") != model.revision:
        raise FormatError("选区属于其他输入版本，请重新选择。")
    kind = scope.get("kind")
    if kind == "page":
        return [], []
    warnings = []
    if kind == "selection":
        spans = scope.get("spans", [])
        for span in spans:
            p = model.row(span.get("pid"))
            a, b = span.get("start"), span.get("end")
            if type(a) is not int or type(b) is not int or not 0 <= a < b <= len(p["text"]):
                raise FormatError("字符选区已失效或为空。")
        return spans, warnings
    if kind == "paragraphs":
        pids = scope.get("pids", [])
        for pid in pids:
            model.row(pid)
    elif kind == "role":
        role = scope.get("role")
        include = scope.get("include_tables", True)
        pids = [p["pid"] for p in model.paragraphs if p["role"] == role and (include or not p["table"])]
        pending = [p["pid"] for p in model.paragraphs if p["role"] == "unknown" and p["text"].strip() and (include or not p["table"])]
        if role == "body" and pending:
            warnings.append("正文覆盖范围存在待归类段落：" + "、".join(map(str, pending)))
    elif kind == "all":
        pids = [p["pid"] for p in model.paragraphs]
    elif kind in {"table", "cell", "row", "column"}:
        cells = model.cells
        if kind == "cell":
            cells = [c for c in cells if c["id"] in scope.get("cell_ids", [])]
        else:
            cells = [c for c in cells if scope.get("table") == "all" or c["table"] == scope.get("table")]
            if kind == "row":
                cells = [c for c in cells if c["row"] <= scope.get("row", 0) <= c["end_row"]]
            if kind == "column":
                cells = [c for c in cells if c["col"] <= scope.get("column", 0) <= c["end_col"]]
        pids = sorted({pid for cell in cells for pid in cell["pids"]})
    elif kind == "named_column":
        columns = [c for c in model.columns if c["id"] == scope.get("column_id")] if scope.get("column_id") else [c for c in model.columns if c["label"] == scope.get("label")]
        if len(columns) != 1:
            raise FormatError("栏目未定位或同名栏目有多处，请选择具体填写区。")
        pids = columns[0]["pids"]
    else:
        raise FormatError("范围未理解或不受支持。")
    if not pids:
        raise FormatError("当前范围没有命中对象。")
    return [{"pid": pid, "start": 0, "end": len(model.row(pid)["text"])} for pid in sorted(set(pids))], warnings


@dataclass(frozen=True)
class EditPlan:
    payload: str
    plan_hash: str

    def read(self):
        if digest(self.payload.encode()) != self.plan_hash:
            raise FormatError("计划快照校验失败。")
        return json.loads(self.payload)


@dataclass
class PlanReview:
    plan: EditPlan | None
    blockers: list[str]
    notices: list[str]
    preview: list[dict]
    excluded: list[dict]


def build_plan(model: DocumentModel, intents: list[dict], protections=None):
    protections = protections or []
    blocked, notices, candidates, excluded = [], [], [], []
    protected = []
    for scope in protections:
        spans, _ = resolve_scope(model, scope)
        protected.extend(spans)
    for number, request in enumerate(intents, 1):
        if not request.get("active", True):
            excluded.append({"request": number, "reason": "用户取消", "source": request.get("source")})
            continue
        try:
            if request.get("unparsed"):
                raise FormatError("未理解内容：" + request["unparsed"])
            scope = request["scope"]
            spans, warnings = resolve_scope(model, scope)
            notices.extend(warnings)
            values = validate_values(request["set"])
            is_page = scope["kind"] == "page"
            if any((p in PAGE_PROPERTIES) != is_page for p in values):
                raise FormatError("页面设置与字符/段落属性需分别添加。")
            priority = 1 if request.get("exception") and scope["kind"] not in {"role", "all", "page"} else 0
            if request.get("exception") and not priority:
                raise FormatError("局部例外需绑定明确的局部对象。")
            for prop, value in values.items():
                if is_page:
                    candidates.append({"pid": 0, "start": 0, "end": 0, "property": prop, "value": value, "priority": 0, "request": number})
                    continue
                for span in spans:
                    pid, start, end = span["pid"], span["start"], span["end"]
                    row = model.row(pid)
                    own_protection = [s for s in protected if s["pid"] == pid]
                    if prop in PARAGRAPH_PROPERTIES:
                        if own_protection or row["field"]:
                            excluded.append({"request": number, "pid": pid, "property": prop, "reason": "段内保护区/字段排除整个段落属性"})
                            continue
                        if start != 0 or end != len(row["text"]):
                            notices.append(f"第 {pid} 段的 {prop} 作用于整个段落，超出所选字符。")
                        candidates.append({"pid": pid, "start": 0, "end": 0, "property": prop, "value": value, "priority": priority, "request": number})
                    else:
                        ranges = [(start, end)]
                        for a, b in [(s["start"], s["end"]) for s in own_protection] + row["protected_spans"]:
                            new = []
                            for x, y in ranges:
                                if y <= a or x >= b:
                                    new.append((x, y))
                                else:
                                    excluded.append({"request": number, "pid": pid, "start": max(a, x), "end": min(b, y), "property": prop, "reason": "保护区或字段"})
                                    if x < a:
                                        new.append((x, a))
                                    if y > b:
                                        new.append((b, y))
                            ranges = new
                        for a, b in ranges:
                            if a < b:
                                candidates.append({"pid": pid, "start": a, "end": b, "property": prop, "value": value, "priority": priority, "request": number})
        except (FormatError, KeyError, TypeError) as exc:
            blocked.append(f"第 {number} 项：{exc}")
    grouped = {}
    for c in candidates:
        grouped.setdefault((c["pid"], c["property"]), []).append(c)
    resolved = []
    for (pid, prop), group in sorted(grouped.items()):
        bounds = sorted({c[k] for c in group for k in ("start", "end")})
        intervals = list(zip(bounds, bounds[1:])) if prop in CHAR_PROPS else [(0, 0)]
        for a, b in intervals:
            hits = [c for c in group if prop not in CHAR_PROPS or c["start"] <= a and c["end"] >= b]
            if not hits:
                continue
            high = max(c["priority"] for c in hits)
            chosen = [c for c in hits if c["priority"] == high]
            if len({encode(c["value"]) for c in chosen}) != 1:
                blocked.append(f"冲突：段落 {pid} 字符 {a}–{b} 的 {prop}，请修改或取消互斥项。")
                continue
            c = chosen[0]
            item = {"pid": pid, "start": a, "end": b, "property": prop, "value": c["value"], "sources": sorted({h["request"] for h in chosen})}
            if resolved and all(resolved[-1][k] == item[k] for k in ("pid", "property", "value", "sources")) and resolved[-1]["end"] == a and prop in CHAR_PROPS:
                resolved[-1]["end"] = b
            else:
                resolved.append(item)
    if not resolved and not blocked:
        blocked.append("没有可执行的修改，请选择属性及未被保护的范围。")
    payload = encode({"input_revision": model.revision, "input_hash": model.source_hash, "model_version": MODEL_VERSION, "roles": model.role_overrides, "assignments": resolved, "intents": intents, "protections": protections, "excluded": excluded, "notices": sorted(set(notices))})
    return PlanReview(None if blocked else EditPlan(payload, digest(payload.encode())), sorted(set(blocked)), sorted(set(notices)), resolved, excluded)


def _number_zh(text):
    if text.isdigit():
        return int(text)
    digits = dict(zip("一二三四五六七八九", range(1, 10)))
    if text == "十":
        return 10
    if "十" in text:
        a, b = text.split("十")
        return digits.get(a, 1) * 10 + digits.get(b, 0)
    return digits.get(text, 0)


def parse_text(model, text):
    """Small exhaustive grammar: any unconsumed substantive text blocks its item."""
    results, last_scope = [], None
    for clause in filter(None, (s.strip() for s in re.split(r"[；;。\n，,]+", text or ""))):
        remaining, scope = clause, None
        exception = bool(re.match(r"^(?:但|其中|例外)", remaining))
        remaining = re.sub(r"^(?:但|其中|例外)", "", remaining)
        match = re.search(r"第([一二三四五六七八九十\d]+)段", remaining)
        named = re.search(r"[“\"]([^”\"]+)[”\"](?:这一栏|栏|填写区|栏目)", remaining)
        if match:
            scope = model.scope("paragraphs", pids=[_number_zh(match.group(1))])
            remaining = remaining.replace(match.group(), "", 1)
            exception = True  # explicitly named local target in a compound instruction
        elif named:
            scope = model.scope("named_column", label=named.group(1))
            remaining = remaining.replace(named.group(), "", 1)
            exception = True
        else:
            for column in sorted(model.columns, key=lambda c: len(c["label"]), reverse=True):
                match_column = re.search(re.escape(column["label"]) + r"(?:这一栏|栏|填写区|栏目)", remaining)
                if match_column:
                    scope = model.scope("named_column", label=column["label"])
                    remaining = remaining.replace(match_column.group(), "", 1)
                    exception = True
                    break
            for label, role in (("一级标题", "heading_1"), ("二级标题", "heading_2"), ("三级标题", "heading_3"), ("文档标题", "title"), ("正文", "body")):
                if scope is None and label in remaining:
                    scope = model.scope("role", role=role, include_tables=True)
                    remaining = remaining.replace(label, "", 1)
                    break
            if scope is None and "全文" in remaining:
                scope = model.scope("all")
                remaining = remaining.replace("全文", "", 1)
        scope = scope or last_scope
        if scope:
            last_scope = scope
        values = {}
        for font in sorted(FONT_NAMES, key=len, reverse=True):
            if font in remaining:
                values["font_family"] = font
                remaining = remaining.replace(font, "", 1)
                break
        for size in sorted(CHINESE_SIZES, key=len, reverse=True):
            if size in remaining:
                values["font_size"] = CHINESE_SIZES[size]
                remaining = remaining.replace(size, "", 1)
                break
        patterns = [
            (r"(?:字号|字体大小)(\d+(?:\.\d+)?)(?:磅|pt)", "font_size", lambda m: float(m[1])),
            (r"(?:固定(?:值)?(?:行距)?|行距固定(?:值)?)(\d+(?:\.\d+)?)(?:磅|pt)", "line_spacing", lambda m: {"mode": "exact", "value": float(m[1])}),
            (r"(?:行距)?(\d+(?:\.\d+)?)倍(?:行距)?", "line_spacing", lambda m: {"mode": "multiple", "value": float(m[1])}),
        ]
        remaining = re.sub(r"设置为|改成|改为|设为|统一为|调整为", "", remaining)
        remaining = re.sub(r"\s+", "", remaining)
        for pattern, prop, convert in patterns:
            m = re.search(pattern, remaining, re.I)
            if m:
                values[prop] = convert(m)
                remaining = remaining.replace(m.group(), "", 1)
        for key, label in (("bold", "加粗"), ("italic", "斜体"), ("underline", "下划线"), ("keep_with_next", "与下段同页")):
            m = re.search(r"(不|取消|关闭)?" + label, remaining)
            if m:
                values[key] = not bool(m[1])
                remaining = remaining.replace(m.group(), "", 1)
        for label, value in (("两端对齐", "justify"), ("居中", "center"), ("左对齐", "left"), ("右对齐", "right")):
            if label in remaining:
                values["alignment"] = value
                remaining = remaining.replace(label, "", 1)
        for label, value in (("红色", "#FF0000"), ("黑色", "#000000")):
            if label in remaining:
                values["color"] = value
                remaining = remaining.replace(label, "", 1)
        m = re.search(r"#[0-9a-fA-F]{6}", remaining)
        if m:
            values["color"] = m.group().upper()
            remaining = remaining.replace(m.group(), "", 1)
        for label, prop in (("首行缩进", "first_line_indent"), ("左缩进", "left_indent"), ("右缩进", "right_indent"), ("段前", "space_before"), ("段后", "space_after")):
            m = re.search(label + r"(\d+(?:\.\d+)?)(磅|pt|厘米|cm)", remaining, re.I)
            if m:
                values[prop] = float(m[1]) * (72/2.54 if m[2].lower() in {"厘米", "cm"} else 1)
                remaining = remaining.replace(m.group(), "", 1)
        remaining = re.sub(r"请|所有|全部|统一|同时|字体|字号|文字|内容|改|为|设|成|仅|只|的|、|号|字|这一栏|填写", "", remaining)
        item = intent(scope or model.scope("unknown"), values, source=clause, exception=exception)
        if remaining or not scope or not values:
            item["unparsed"] = remaining or clause
        results.append(item)
    return results
