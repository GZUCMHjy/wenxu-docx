"""Real local Chromium workflow; generated synthetic data only."""
from io import BytesIO
import json
import os
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT)); sys.path.insert(0, str(ROOT / "tests"))
from playwright.sync_api import expect, sync_playwright
from test_v06 import sample, character_values


def main():
    output = ROOT / "verification" / "v06-browser"
    output.mkdir(parents=True, exist_ok=True)
    source = sample(complex_objects=False)
    (output / "synthetic.docx").write_bytes(source)
    evidence = []
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 1512, "height": 1080}, accept_downloads=True)
        page.set_default_timeout(150_000)
        expect.set_options(timeout=150_000)
        page.goto(os.environ.get("WENXU_BASE_URL", "http://127.0.0.1:8501"))
        expect(page.get_by_role("heading", name="文序 · 精确格式调整")).to_be_visible()
        page.locator('input[type="file"]').first.set_input_files(str(output / "synthetic.docx"))
        expect(page.get_by_role("heading", name="设置本次修改")).to_be_visible()

        def choose(label, option, multi=False):
            container = page.get_by_test_id("stMultiSelect" if multi else "stSelectbox").filter(has=page.get_by_text(label, exact=True)).first
            container.locator('input').click()
            page.get_by_role("option", name=option, exact=True).click()
            if multi:
                page.keyboard.press("Escape")

        choose("本次设置的属性（其他保持原样）", "字号", True)
        page.get_by_role("spinbutton", name="字号（磅）", exact=True).first.fill("10.5")
        page.get_by_role("spinbutton", name="字号（磅）", exact=True).first.press("Tab")
        page.get_by_role("button", name="加入待应用清单", exact=True).click()
        expect(page.get_by_role("button", name="应用修改并检查", exact=True)).to_be_enabled()
        page.screenshot(path=str(output / "plan.png"), full_page=True)
        page.get_by_role("button", name="应用修改并检查", exact=True).click()
        # Capture actionable UI errors as well as the expected success result.
        page.wait_for_function("""() => document.body.innerText.includes('synthetic_v1.docx 已完成') || [...document.querySelectorAll('[data-testid=stAlert]')].some(e => e.innerText.includes('失败') || e.innerText.includes('无法') || e.innerText.includes('失效'))""")
        print(page.get_by_test_id("stAlert").all_text_contents(), flush=True)
        page.screenshot(path=str(output / "after-apply.png"), full_page=True)
        expect(page.get_by_text("synthetic_v1.docx 已完成结构、目标属性、对象保全和渲染检查。", exact=True)).to_be_visible()
        expect(page.get_by_role("button", name="下载修改后的 DOCX", exact=True)).to_be_disabled()
        page.get_by_text("已查看修改前后页面，接受当前版式并下载", exact=True).click()
        with page.expect_download() as download:
            page.get_by_role("button", name="下载修改后的 DOCX", exact=True).click()
        download.value.save_as(str(output / "v1.docx"))
        v1 = (output / "v1.docx").read_bytes()
        assert set(character_values(v1, 2, "font_size")) == {10.5}
        assert character_values(v1, 2, "bold") == character_values(source, 2, "bold")
        evidence.append("A02 UI only changes explicit font size; render and gated download")
        page.screenshot(path=str(output / "v1.png"), full_page=True)

        page.get_by_text("局部调整", exact=True).click()
        search = page.get_by_role("textbox", name="搜索原文", exact=True)
        search.fill("重复短语"); search.press("Tab")
        expect(page.get_by_text("找到 2 处；请明确选择一个或多个命中。", exact=True)).to_be_visible()
        box = page.get_by_test_id("stMultiSelect").filter(has=page.get_by_text("选择命中位置", exact=True))
        box.locator('input').click()
        page.get_by_role("option", name="第 2 处", exact=False).click(); page.keyboard.press("Escape")
        choose("本次设置的属性（其他保持原样）", "字体颜色", True)
        page.get_by_role("button", name="加入待应用清单", exact=True).click()
        page.get_by_role("button", name="应用修改并检查", exact=True).click()
        expect(page.get_by_text("synthetic_v2.docx 已完成结构、目标属性、对象保全和渲染检查。", exact=True)).to_be_visible()
        page.get_by_text("已查看修改前后页面，接受当前版式并下载", exact=True).click()
        with page.expect_download() as download:
            page.get_by_role("button", name="下载修改后的 DOCX", exact=True).click()
        download.value.save_as(str(output / "v2.docx"))
        v2 = (output / "v2.docx").read_bytes()
        assert character_values(v2, 2, "color") == character_values(v1, 2, "color")
        assert character_values(v2, 3, "color")[:4] == ["#FF0000"] * 4
        assert character_values(v2, 2, "font_size") == character_values(v1, 2, "font_size")
        with page.expect_download() as report:
            page.get_by_role("button", name="下载 HTML 检查报告", exact=True).click()
        report.value.save_as(str(output / "report.html"))
        assert "confirmed" in (output / "report.html").read_text()
        evidence.append("A03/A09 only second match changes; v1 font size persists in v2; report downloaded")
        page.get_by_role("button", name="撤销本次", exact=True).click()
        expect(page.get_by_text("synthetic_v1.docx 已完成结构、目标属性、对象保全和渲染检查。", exact=True)).to_be_visible()
        evidence.append("A13 undo returns v1 and resets editing draft")
        box = page.get_by_role("textbox", name="一句话修改", exact=True)
        box.fill("正文小四；正文五号"); box.press("Tab")
        page.get_by_role("button", name="解析并加入清单", exact=True).click()
        expect(page.get_by_role("button", name="应用修改并检查", exact=True)).to_be_disabled()
        evidence.append("A11 conflicting explicit sizes block apply in actual browser")
        page.screenshot(path=str(output / "conflict.png"), full_page=True)
        assert page.get_by_test_id("stException").count() == 0
        browser.close()
    (output / "evidence.json").write_text(json.dumps({"checks": evidence}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"browser_checks": len(evidence), "result": "passed"}))


if __name__ == "__main__":
    main()
