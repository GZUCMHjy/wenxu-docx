"""Real browser upload -> widget changes -> DOCX download -> independent inspection.

Run in WSL after starting Streamlit: python tests/browser_smoke.py
"""

import hashlib
from io import BytesIO
import json
from pathlib import Path
import platform
import sys
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from docx import Document
from docx.oxml.ns import qn
from playwright.sync_api import sync_playwright, expect
from test_docx_formatting import heading_sample_docx


def inspect_output(source, output, *, size, color, line, rule):
    before, after = Document(BytesIO(source)), Document(BytesIO(output))
    assert before.element.xpath(".//w:t/text()") == after.element.xpath(".//w:t/text()")
    p = after.paragraphs[1]
    assert p.runs[0].font.size.pt == size
    assert str(p.runs[0].font.color.rgb) == color
    assert p.runs[1].bold and p.runs[2].italic
    assert p.paragraph_format.space_after.pt == 9
    assert p.paragraph_format.first_line_indent.pt == 24
    for para in after.element.xpath("./w:body//w:p"):
        spacing = para.find("w:pPr/w:spacing", para.nsmap)
        assert spacing.get(qn("w:line")) == line
        assert spacing.get(qn("w:lineRule"), "auto") == rule
    with ZipFile(BytesIO(source)) as a, ZipFile(BytesIO(output)) as b:
        assert a.namelist() == b.namelist()
        for name in a.namelist():
            if name != "word/document.xml":
                assert a.read(name) == b.read(name), name


def main():
    source = heading_sample_docx()
    sample = ROOT / "examples/sample.docx"
    sample.parent.mkdir(exist_ok=True)
    sample.write_bytes(source)
    artifacts = ROOT / "verification"
    artifacts.mkdir(exist_ok=True)
    report = {"runtime": platform.platform(), "python": sys.version.split()[0], "source_sha256": hashlib.sha256(source).hexdigest(), "checks": []}
    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page(viewport={"width": 1280, "height": 960}, accept_downloads=True)
        page.set_default_timeout(30000)
        page.goto("http://127.0.0.1:8501")
        execute = page.get_by_role("button", name="应用格式并生成 DOCX")
        expect(execute).to_be_disabled()
        page.locator('input[type="file"]').set_input_files(str(sample))
        expect(execute).to_be_disabled()
        page.get_by_text("修改字号", exact=True).click()
        size_input = page.get_by_role("spinbutton", name="字号（磅）")
        size_input.fill("18")
        size_input.press("Tab")
        page.get_by_text("修改字体颜色", exact=True).click()
        page.get_by_text("修改行距", exact=True).click()
        line_input = page.get_by_role("spinbutton", name="固定行距（磅）")
        line_input.fill("30")
        line_input.press("Tab")
        execute.click()
        expect(page.get_by_text("修改完成，内容与未选属性校验通过。", exact=True)).to_be_visible()
        page.screenshot(path=str(artifacts / "browser-success.png"), full_page=True)
        with page.expect_download() as download_info:
            page.get_by_role("button", name="下载修改后的 DOCX").click()
        downloaded = download_info.value
        assert downloaded.suggested_filename == "sample_v1.docx"
        path = artifacts / downloaded.suggested_filename
        downloaded.save_as(str(path))
        output = path.read_bytes()
        inspect_output(source, output, size=18, color="000000", line="600", rule="exact")
        report["checks"].append("browser: upload -> size 18pt / black / exact 30pt -> download; XML and untouched parts verified")

        # Change a setting after success: the old download must disappear.
        size_input.fill("16")
        size_input.press("Tab")
        expect(page.get_by_role("button", name="下载修改后的 DOCX")).to_have_count(0)
        report["checks"].append("settings change invalidates old download")
        page.get_by_text("修改字号", exact=True).click()
        page.get_by_text("修改字体颜色", exact=True).click()
        page.get_by_role("combobox", name="行距模式").click()
        page.get_by_role("option", name="倍数", exact=True).click()
        page.get_by_role("combobox", name="行距倍数").click()
        page.get_by_role("option", name="2.0", exact=True).click()
        execute.click()
        expect(page.get_by_text("修改完成，内容与未选属性校验通过。", exact=True)).to_be_visible()
        with page.expect_download() as download_info:
            page.get_by_role("button", name="下载修改后的 DOCX").click()
        path = artifacts / "spacing-only.docx"
        download_info.value.save_as(str(path))
        inspect_output(source, path.read_bytes(), size=12, color="222222", line="480", rule="auto")
        report["checks"].append("line-spacing-only download: 2x, original font size/color retained")

        # Scope plus font is checked independently against actual built-in heading styles.
        page.get_by_text("修改行距", exact=True).click()
        page.get_by_text("修改字体", exact=True).click()
        for scope_name, style_name, family in [("一级标题", "Heading 1", "黑体"), ("二级标题", "Heading 2", "宋体"), ("三级标题", "Heading 3", "黑体")]:
            page.get_by_role("combobox", name="修改范围").click()
            page.get_by_role("option", name=scope_name, exact=True).click()
            expect(page.get_by_role("button", name="下载修改后的 DOCX")).to_have_count(0)
            page.get_by_role("combobox", name="字体").click()
            page.get_by_role("option", name=family, exact=True).click()
            execute.click()
            expect(page.get_by_text("修改完成，内容与未选属性校验通过。", exact=True)).to_be_visible()
            with page.expect_download() as download_info:
                page.get_by_role("button", name="下载修改后的 DOCX").click()
            path = artifacts / f"{style_name.replace(' ', '-')}.docx"
            download_info.value.save_as(str(path))
            before, after = Document(BytesIO(source)), Document(path)
            matched = 0
            for old, new in zip(before.paragraphs, after.paragraphs, strict=True):
                if old.style.name == style_name:
                    matched += 1
                    for run in new.runs:
                        fonts = run._r.find("w:rPr/w:rFonts", run._r.nsmap)
                        assert fonts.get(qn("w:eastAsia")) == family
                        assert run.font.name == family
                        assert run.font.size == old.runs[0].font.size
                else:
                    assert old._p.xml == new._p.xml
            assert matched == 1
            with ZipFile(BytesIO(source)) as a, ZipFile(path) as b:
                assert a.namelist() == b.namelist()
                for name in a.namelist():
                    if name != "word/document.xml":
                        assert a.read(name) == b.read(name)
            report["checks"].append(f"{scope_name}: {family}, other paragraphs and package parts unchanged")
        page.get_by_role("button", name="下载修改后的 DOCX").scroll_into_view_if_needed()
        page.screenshot(path=str(artifacts / "browser-heading-result.png"), full_page=True)
        page.get_by_role("combobox", name="修改范围").click()
        page.get_by_role("option", name="九级标题", exact=True).click()
        expect(execute).to_be_disabled()
        expect(page.get_by_role("button", name="下载修改后的 DOCX")).to_have_count(0)
        report["checks"].append("empty heading scope blocked and previous result invalidated")

        # Replacing the input invalidates results, and malformed data cannot be downloaded.
        page.locator('input[type="file"]').set_input_files({"name": "broken.docx", "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "buffer": b"not a zip"})
        expect(page.get_by_role("button", name="下载修改后的 DOCX")).to_have_count(0)
        expect(page.get_by_text("文档已损坏或不是有效 DOCX，请重新选择文件。", exact=True)).to_be_visible()
        expect(execute).to_be_disabled()
        expect(page.get_by_role("button", name="下载修改后的 DOCX")).to_have_count(0)
        page.screenshot(path=str(artifacts / "browser-invalid.png"), full_page=True)
        report["checks"].append("bad DOCX blocked without stale download")
        page.set_viewport_size({"width": 390, "height": 844})
        page.screenshot(path=str(artifacts / "browser-mobile.png"), full_page=True)
        browser.close()
    assert sample.read_bytes() == source
    report["checks"].append("uploaded sample bytes unchanged")
    report["result"] = "PASS"
    (artifacts / "browser-report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
