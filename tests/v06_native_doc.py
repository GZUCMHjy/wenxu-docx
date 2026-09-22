"""Native DOC format properties, re-open verification and undo; synthetic only."""
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT)); sys.path.insert(0, str(ROOT / 'tests'))
from test_v06 import sample
from v06_model import digest
from v06_office import native_office
from v06_plan import build_plan, intent
from v06_session import EditingSession, _libreoffice_convert


def main():
    office = native_office()
    assert office is not None, 'Requires local WPS.'
    data = _libreoffice_convert(sample(False), 'docx', 'doc:MS Word 97')
    # Generate a native fixture. This setup is not the application's import.
    data, _, _ = office.inspect_doc(data, {'input_hash': digest(data), 'assignments': []})
    session = EditingSession.load('synthetic.doc', data)
    assert session.versions['original'].data == data
    assert session.conversion['page_check']['pixels_equal']
    model = session.model()
    changes = {'font_east_asia': '宋体', 'font_western': 'Arial', 'font_size': 14,
               'bold': True, 'italic': False, 'underline': True, 'color': '#13579B',
               'alignment': 'center', 'first_line_indent': 24, 'left_indent': 12,
               'right_indent': 6, 'space_before': 6, 'space_after': 8,
               'keep_with_next': True, 'line_spacing': {'mode': 'exact', 'value': 24}}
    plan = build_plan(model, [intent(model.scope('selection', spans=[{'pid': 2, 'start': 2, 'end': 6}]), changes),
                             intent(model.scope('page'), {'top_margin': 40, 'bottom_margin': 45, 'left_margin': 50, 'right_margin': 55})]).plan
    version, created = session.execute(model, plan)
    assert created and version.extension == 'doc'
    assert json.loads(version.report_json)['checks']['preservation'] == 'passed'
    session.select(version.id)
    model = session.model()
    scope = model.scope('paragraphs', pids=[2])
    plan = build_plan(model, [intent(scope, {'bold': False, 'italic': True, 'underline': False,
                                          'keep_with_next': False, 'line_spacing': {'mode': 'multiple', 'value': 1.5}})]).plan
    second, created = session.execute(model, plan)
    assert created and second.parent == version.id
    session.select(second.id); session.undo(); assert session.current == version.id
    session.undo(); assert session.current == 'original'
    assert session.versions['original'].data == data
    output = ROOT / 'verification/v06-native-doc'; output.mkdir(parents=True, exist_ok=True)
    (output / 'edited.doc').write_bytes(second.data)
    (output / 'check.json').write_text(second.report_json, encoding='utf-8')
    print('Native DOC character/paragraph/page properties, two saved versions and undo: PASS')


if __name__ == '__main__': main()
