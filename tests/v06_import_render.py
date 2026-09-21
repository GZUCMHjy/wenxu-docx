"""Local LibreOffice regression: grouped vector formulas survive import and edits."""
from collections import Counter
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from test_v06_import import FORMULAS, grouped_sample
from v06_import import repair_imported_picture_groups
from v06_model import build_model
from v06_plan import build_plan, intent
from v06_patch import patch_document, verify_result
from v06_session import render_document


def main():
    dest = ROOT / 'verification' / 'v06-import-render'
    dest.mkdir(parents=True, exist_ok=True)
    fixed, repairs = repair_imported_picture_groups(grouped_sample())
    assert repairs[0]['pictures'] == 6
    model = build_model(fixed)
    plan = build_plan(model, [intent(model.scope('all'), {'font_size': 12})]).plan
    edited, _ = patch_document(model, plan)
    verify_result(model, edited, plan)
    for label, data in (('imported', fixed), ('edited', edited)):
        render = render_document(data)
        text = ''.join(render['text'].split())
        for formula, count in Counter(FORMULAS).items():
            assert text.count(formula) == count, (label, formula, text)
        (dest / (label+'.docx')).write_bytes(data)
        (dest / (label+'.pdf')).write_bytes(render['pdf'])
        (dest / (label+'.png')).write_bytes(render['images'][0])
    print('Six synthetic vector formulas remain visible after import and text formatting: PASS')


if __name__ == '__main__':
    main()
