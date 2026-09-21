"""Synthetic DOC conversion gate and cross-upload review isolation in Chromium."""
import os
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT)); sys.path.insert(0, str(ROOT / "tests"))
from playwright.sync_api import expect, sync_playwright
from test_v06 import sample
from v06_session import _convert


def main():
    dest = ROOT / "verification" / "v06-doc-browser"
    dest.mkdir(parents=True, exist_ok=True)
    data = sample(complex_objects=False)
    doc = _convert(data, "docx", "doc:MS Word 97")
    (dest / "synthetic.doc").write_bytes(doc)
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 1512, "height": 1080}, accept_downloads=True)
        page.set_default_timeout(180_000); expect.set_options(timeout=180_000)
        page.goto(os.environ.get("WENXU_BASE_URL", "http://127.0.0.1:8501"))
        page.locator('input[type="file"]').first.set_input_files(str(dest / "synthetic.doc"))
        expect(page.get_by_role("heading", name="设置本次修改")).to_be_visible()
        text = page.get_by_role("textbox", name="一句话修改", exact=True)
        text.fill("正文五号"); text.press("Tab")
        page.get_by_role("button", name="解析并加入清单", exact=True).click()
        apply = page.get_by_role("button", name="应用修改并检查", exact=True)
        expect(apply).to_be_disabled()
        page.get_by_text("已核对转换前后页面，接受此工作副本", exact=True).click()
        expect(apply).to_be_enabled(); apply.click()
        expect(page.get_by_text("synthetic_v1.docx 已完成结构、目标属性、对象保全和渲染检查。", exact=True)).to_be_visible()
        expect(page.get_by_role("button", name="下载修改后的 DOCX", exact=True)).to_be_disabled()
        page.get_by_text("已查看修改前后页面，接受当前版式并下载", exact=True).click()
        expect(page.get_by_role("button", name="下载修改后的 DOCX", exact=True)).to_be_enabled()
        with page.expect_download() as original:
            page.get_by_role("button", name="下载保留的原始上传文件", exact=True).click()
        original.value.save_as(str(dest / "retained.doc"))
        assert (dest / "retained.doc").read_bytes() == doc
        # A different document with v1 must not inherit the preceding review.
        (dest / "second.docx").write_bytes(sample(complex_objects=False))
        page.locator('input[type="file"]').first.set_input_files(str(dest / "second.docx"))
        expect(page.get_by_text("original · second.docx", exact=True)).to_be_visible()
        text = page.get_by_role("textbox", name="一句话修改", exact=True)
        text.fill("正文四号"); text.press("Tab")
        page.get_by_role("button", name="解析并加入清单", exact=True).click()
        page.get_by_role("button", name="应用修改并检查", exact=True).click()
        expect(page.get_by_text("second_v1.docx 已完成结构、目标属性、对象保全和渲染检查。", exact=True)).to_be_visible()
        expect(page.get_by_role("button", name="下载修改后的 DOCX", exact=True)).to_be_disabled()
        page.get_by_role("heading", name="文序 · 精确格式调整").scroll_into_view_if_needed()
        page.screenshot(path=str(dest / "second-upload.png"))
        assert page.get_by_test_id("stException").count() == 0
        browser.close()
    print('DOC conversion gate, retained original and cross-upload review isolation: PASS')


if __name__ == "__main__":
    main()
