"""Synthetic DOC edit and cross-upload review isolation in Chromium."""
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT)); sys.path.insert(0, str(ROOT / 'tests'))
from playwright.sync_api import expect, sync_playwright
from editor_browser import EditorBrowser
from test_v06 import sample, character_values
from v06_session import _libreoffice_convert
from v06_office import native_office
from v06_model import digest
from v06_native import inspect_native_doc


def main():
    dest = ROOT / 'verification/v06-doc-browser'; dest.mkdir(parents=True, exist_ok=True)
    data = sample(complex_objects=False)
    doc = _libreoffice_convert(data, 'docx', 'doc:MS Word 97')
    doc, _, _ = native_office().inspect_doc(doc, {'input_hash': digest(doc), 'assignments': []})
    (dest / 'synthetic.doc').write_bytes(doc)
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={'width': 1512, 'height': 1080}, accept_downloads=True)
        ui = EditorBrowser(page); ui.upload(dest / 'synthetic.doc')
        ui.tab('预览与下载')
        assert ui.download('下载未修改的工作副本', dest / 'baseline.doc') == doc
        ui.tab('编辑')
        ui.choose('字体', '黑体')
        ui.applied('synthetic_v1.doc')
        expect(page.get_by_role('button', name='下载修改后的 DOC', exact=True)).to_be_disabled()
        ui.confirm()
        edited = ui.download('下载修改后的 DOC', dest / 'edited.doc')
        assert edited.startswith(bytes.fromhex('D0CF11E0A1B11AE1'))
        ui.tab('历史与高级')
        assert ui.download('下载保留的原始上传文件', dest / 'retained.doc') == doc
        ui.choose('本次输入版本', 'original · synthetic_工作副本.doc')
        expect(page.get_by_role('tab', name='预览与下载', exact=True)).to_have_attribute('aria-selected', 'true')
        assert ui.download('下载未修改的工作副本', dest / 'baseline.doc') == doc
        # A different document with v1 must not inherit settings or download confirmation.
        (dest / 'second.docx').write_bytes(data)
        ui.upload(dest / 'second.docx')
        expect(page.get_by_role('button', name='应用并预览', exact=True)).to_be_disabled()
        ui.expand('用一句话设置')
        ui.text('一句话修改', '正文四号')
        ui.click('解析并加入清单')
        ui.applied('second_v1.docx')
        expect(page.get_by_role('button', name='下载修改后的 DOCX', exact=True)).to_be_disabled()
        page.screenshot(path=str(dest / 'second-upload.png'))
        browser.close()
    snapshot, _ = inspect_native_doc(edited)
    assert set(character_values(snapshot, 2, 'font_east_asia')) == {'黑体'}
    assert set(character_values(snapshot, 2, 'font_western')) == {'黑体'}
    print('PASS: DOC font edit and saved properties, retained bytes, history selection, cross-upload settings and review isolation')


if __name__ == '__main__': main()
