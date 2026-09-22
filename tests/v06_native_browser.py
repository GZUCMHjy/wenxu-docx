"""Native-engine font selection and real downloaded DOCX, synthetic data only."""
import os
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT)); sys.path.insert(0, str(ROOT / 'tests'))
from playwright.sync_api import expect, sync_playwright
from streamlit.proto.ForwardMsg_pb2 import ForwardMsg
from test_v06 import sample
from v06_model import build_model
from v06_office import native_office


def main():
    office = native_office()
    assert office is not None, 'This integration test requires the local native renderer.'
    installed = {font.casefold() for font in office.environment()['font_families']}
    output = ROOT / 'verification/v06-native-browser'; output.mkdir(parents=True, exist_ok=True)
    source = sample(complex_objects=False); (output / 'fonts.docx').write_bytes(source)
    with sync_playwright() as p:
        browser = p.chromium.launch(); page = browser.new_page(accept_downloads=True)
        page.set_default_timeout(180_000); expect.set_options(timeout=180_000)
        sockets = []; page.on('websocket', lambda socket: sockets.append(socket))
        page.goto(os.environ.get('WENXU_BASE_URL', 'http://127.0.0.1:8501'))
        expect(page.get_by_role('heading', name='文序 · 精确格式调整')).to_be_visible()
        def finished(payload):
            msg = ForwardMsg(); msg.ParseFromString(payload)
            return msg.WhichOneof('type') == 'script_finished' and msg.script_finished == 0
        def rerun(action):
            with sockets[0].expect_event('framereceived', predicate=finished): action()
        rerun(lambda: page.locator('input[type="file"]').first.set_input_files(str(output / 'fonts.docx')))
        props = page.get_by_test_id('stMultiSelect').filter(has=page.get_by_text('本次设置的属性（其他保持原样）', exact=True))
        props.locator('input').click()
        rerun(lambda: page.get_by_role('option', name='字体（中文与西文）', exact=True).click())
        page.keyboard.press('Escape')
        selector = page.get_by_test_id('stSelectbox').filter(has=page.get_by_text('字体（中文与西文）', exact=True))
        selected = selector.locator('[data-baseweb="select"]').inner_text().strip()
        assert selected.casefold() in installed, selected
        rerun(lambda: page.get_by_role('button', name='加入待应用清单', exact=True).click())
        rerun(lambda: page.get_by_role('button', name='应用修改并检查', exact=True).click())
        expect(page.get_by_text('fonts_v1.docx 已完成结构、目标属性、对象保全和渲染检查。', exact=True)).to_be_visible()
        rerun(lambda: page.get_by_text('已查看修改前后页面，接受当前版式并下载', exact=True).click())
        with page.expect_download() as download: page.get_by_role('button', name='下载修改后的 DOCX', exact=True).click()
        download.value.save_as(str(output / 'downloaded.docx'))
        before, after = build_model(source), build_model((output / 'downloaded.docx').read_bytes())
        assert [p['text'] for p in before.paragraphs] == [p['text'] for p in after.paragraphs]
        assert page.get_by_test_id('stException').count() == 0
        browser.close()
    print('Native font default is installed; real format/edit/download preserves text: PASS')


if __name__ == '__main__': main()
