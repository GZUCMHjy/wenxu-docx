"""Direct font changes, tab state, saved properties and actual rendered output."""
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT)); sys.path.insert(0, str(ROOT / 'tests'))
from playwright.sync_api import expect, sync_playwright
from editor_browser import EditorBrowser
from test_v06 import sample, character_values
from v06_import import compare_rendered_pages
from v06_model import build_model
from v06_office import native_office
from v06_session import render_document


def main():
    office = native_office()
    assert office is not None, 'This test requires the local native renderer.'
    assert '黑体' in office.environment()['font_families']
    output = ROOT / 'verification/v06-native-browser'; output.mkdir(parents=True, exist_ok=True)
    source = sample(complex_objects=False); (output / 'fonts.docx').write_bytes(source)
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={'width': 1512, 'height': 1080}, accept_downloads=True)
        ui = EditorBrowser(page)
        ui.upload(output / 'fonts.docx')
        expect(page.get_by_role('button', name='应用并预览', exact=True)).to_be_disabled()
        expect(page.get_by_role('heading', name='历史版本', exact=True)).not_to_be_visible()
        expect(page.get_by_test_id('stImage')).to_have_count(0)
        expect(page.get_by_text('当前中文字体：', exact=False)).to_be_visible()
        ui.choose('字体', '黑体')
        ui.tab('历史与高级'); ui.tab('编辑')
        expect(page.get_by_text('本次将应用：字体 → 黑体', exact=True)).to_be_visible()
        page.screenshot(path=str(output / 'editor.png'))
        ui.applied('fonts_v1.docx')
        expect(page.get_by_text('已应用：字体 → 黑体', exact=True)).to_be_visible()
        expect(page.get_by_test_id('stImage')).to_have_count(1)
        ui.rerun(lambda: page.get_by_text('对比修改前', exact=True).click())
        expect(page.get_by_test_id('stImage')).to_have_count(2)
        expect(page.get_by_role('button', name='下载修改后的 DOCX', exact=True)).to_be_disabled()
        ui.confirm()
        data = ui.download('下载修改后的 DOCX', output / 'downloaded.docx')
        before, after = build_model(source), build_model(data)
        assert [p['text'] for p in before.paragraphs] == [p['text'] for p in after.paragraphs]
        for prop in ['font_east_asia', 'font_western']:
            assert set(character_values(data, 2, prop)) == {'黑体'}
        assert character_values(data, 1, 'font_east_asia') == character_values(source, 1, 'font_east_asia')
        assert character_values(data, 2, 'font_size') == character_values(source, 2, 'font_size')
        page.screenshot(path=str(output / 'font-result.png'))
        ui.click('继续修改')
        expect(page.get_by_text('当前中文字体：黑体', exact=False)).to_be_visible()
        ui.choose('字体', '黑体')
        ui.click('应用并预览')
        expect(page.get_by_text('所选范围已经是这些设置，文件没有变化。', exact=True)).to_be_visible()
        ui.tab('历史与高级')
        box = page.get_by_test_id('stSelectbox').filter(has=page.get_by_text('本次输入版本', exact=True))
        box.locator('input').click()
        expect(page.get_by_role('option')).to_have_count(2)
        page.keyboard.press('Escape')
        browser.close()
    actual = compare_rendered_pages(render_document(source), render_document(data))
    assert actual['differing_pages'], actual
    print('PASS: one-click native font edit, tab state, saved fonts, real pixels, unchanged properties, explicit no-op')


if __name__ == '__main__': main()
