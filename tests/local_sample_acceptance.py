"""Optional private-file acceptance. All generated content must stay outside Git."""
import argparse
from pathlib import Path
import json
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from v06_model import digest
from v06_plan import build_plan, intent
from v06_session import EditingSession, render_environment


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
    version = session.versions['original']
    baseline = version.data
    extension = version.extension
    assert baseline == original
    (dest / ('working.' + extension)).write_bytes(baseline)
    env = render_environment()
    if session.conversion['status'] == 'original_format':
        assert session.conversion['page_check']['pixels_equal']
        save_render(dest, 'original', session.renders['conversion_source'])
        save_render(dest, 'working', session.renders['conversion_working'])
    else:
        save_render(dest, 'working', session.preview())
    model = session.model()
    cases = []
    requests = []
    columns = [c for c in model.columns if c['pids'] and any(model.row(pid)['text'].strip() for pid in c['pids'])]
    if columns:
        requests.append(('column-fill', intent(model.scope('named_column', column_id=columns[0]['id']), {'font_size': 10.5})))
    requests.append(('all-text-size', intent(model.scope('all'), {'font_size': 12})))
    paragraph = next(p for p in model.paragraphs if len(p['text']) >= 8 and not p['protected_spans'])
    requests.append(('four-characters', intent(model.scope('selection', spans=[{'pid': paragraph['pid'], 'start': 2, 'end': 6}]), {'color': '#FF0000'})))
    for label, request in requests:
        session.select('original')
        plan = build_plan(model, [request]).plan
        result, created = session.execute(model, plan, environment=env)
        if result is None:
            cases.append({'case': label, 'already_satisfied': True}); continue
        assert created and result.extension == extension
        (dest / (label + '.' + extension)).write_bytes(result.data)
        save_render(dest, label, session.preview(result.id))
        report = json.loads(result.report_json)
        cases.append({'case': label, 'checks': report['checks'], 'operations': len(report['operations'])})
        session.select(result.id)
        updated = session.model()
        new_scope = {**request['scope'], 'revision': result.id}
        same = build_plan(updated, [intent(new_scope, request['set'])]).plan
        assert session.execute(updated, same, environment=env) == (None, False)
    assert session.versions['original'].data == original
    report = {'source_hash': digest(original), 'working_hash': digest(baseline), 'inventory': model.inventory,
              'conversion': session.conversion, 'environment': env, 'cases': cases, 'visual_review': 'pending'}
    (dest / 'acceptance.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    assert args.source.read_bytes() == original
    print(json.dumps({"inventory": model.inventory, "cases": len(cases), "original_preserved": True, "visual_review": "pending"}, ensure_ascii=False))


if __name__ == "__main__":
    main()
