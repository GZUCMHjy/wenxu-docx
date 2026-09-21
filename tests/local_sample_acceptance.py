"""Optional private-file acceptance. All generated content must stay outside Git."""
import argparse
from pathlib import Path
import json
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from v06_model import build_model, digest
from v06_plan import build_plan, intent
from v06_patch import patch_document, verify_result
from v06_session import EditingSession, render_environment, render_document, report_html


def save_render(root, name, rendered):
    (root / (name+".pdf")).write_bytes(rendered["pdf"])
    for n, image in enumerate(rendered["images"], 1):
        (root / f"{name}-{n:02}.png").write_bytes(image)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    args = parser.parse_args()
    dest = args.output_dir.resolve()
    if dest.is_relative_to(ROOT.resolve()):
        parser.error("真实稿件验证产物必须保存到仓库外。")
    dest.mkdir(parents=True, exist_ok=True)
    original = args.source.read_bytes()
    session = EditingSession.load(args.source.name, original)
    baseline = session.versions["original"].data
    (dest / "working.docx").write_bytes(baseline)
    env = render_environment()
    if session.conversion["status"] == "converted":
        save_render(dest, "original", session.renders["conversion_source"])
        save_render(dest, "working", session.renders["conversion_working"])
    else:
        save_render(dest, "working", render_document(baseline))
    m = session.model()
    # This diagnostic path intentionally does not register a deliverable version
    # before conversion/visual review. Files are candidates for local inspection.
    operations = []
    cases = []
    available = [c for c in m.columns if c["pids"] and any(m.row(pid)["text"].strip() for pid in c["pids"])]
    if available:
        column = available[0]
        requests = [intent(m.scope("named_column", column_id=column["id"]), {"font_size": 10.5})]
        plan = build_plan(m, requests).plan
        candidate, operations = patch_document(m, plan)
        checks = verify_result(m, candidate, plan)
        cases.append({"case": "column-fill", "checks": checks, "operations": len(operations)})
    else:
        candidate = baseline
    # Whole-document size reaches text adjacent to all declared preserved types.
    all_plan = build_plan(m, [intent(m.scope("all"), {"font_size": 12})]).plan
    output, all_operations = patch_document(m, all_plan)
    checks = verify_result(m, output, all_plan)
    (dest / "candidate.docx").write_bytes(output)
    save_render(dest, "candidate", render_document(output))
    m2 = build_model(output, "candidate")
    same_plan = build_plan(m2, [intent(m2.scope("all"), {"font_size": 12})]).plan
    repeated, repeated_operations = patch_document(m2, same_plan)
    assert repeated == output and not repeated_operations
    # Independent local character selection, never derived from comments.
    p = next(p for p in m.paragraphs if len(p["text"]) >= 8 and not p["protected_spans"])
    local_plan = build_plan(m, [intent(m.scope("selection", spans=[{"pid": p["pid"], "start": 2, "end": 6}]), {"color": "#FF0000"})]).plan
    local, local_operations = patch_document(m, local_plan)
    local_checks = verify_result(m, local, local_plan)
    cases.extend([{"case": "all-text-size", "checks": checks, "operations": len(all_operations)}, {"case": "four-characters", "checks": local_checks, "operations": len(local_operations)}, {"case": "idempotency", "passed": True}])
    report = {"source_hash": digest(original), "working_hash": digest(baseline), "output_hash": digest(output), "inventory": m.inventory, "conversion": session.conversion, "environment": env, "cases": cases, "visual_review": "pending", "client_compatibility": "not_run"}
    (dest / "acceptance.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    assert args.source.read_bytes() == original
    print(json.dumps({"inventory": m.inventory, "cases": len(cases), "original_preserved": True, "visual_review": "pending"}, ensure_ascii=False))


if __name__ == "__main__":
    main()
