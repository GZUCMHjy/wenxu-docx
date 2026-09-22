"""Shared real-browser actions that wait for each Streamlit run to finish."""
import os

from playwright.sync_api import expect
from streamlit.proto.ForwardMsg_pb2 import ForwardMsg


class EditorBrowser:
    def __init__(self, page):
        self.page = page
        self.sockets = []
        page.set_default_timeout(30_000)
        expect.set_options(timeout=30_000)
        page.on('websocket', lambda socket: self.sockets.append(socket))
        page.goto(os.environ.get('WENXU_BASE_URL', 'http://127.0.0.1:8501'))
        expect(page.get_by_role('heading', name='文序 · 精确格式调整')).to_be_visible()

    def rerun(self, action):
        def finished(payload):
            msg = ForwardMsg()
            msg.ParseFromString(payload)
            return msg.WhichOneof('type') == 'script_finished' and msg.script_finished == 0
        with self.sockets[0].expect_event('framereceived', predicate=finished, timeout=180_000):
            action()
        assert self.page.get_by_test_id('stException').count() == 0, self.page.get_by_test_id('stException').all_text_contents()

    def upload(self, path):
        self.rerun(lambda: self.page.locator('input[type="file"]').first.set_input_files(str(path)))
        expect(self.page.get_by_role('tab', name='编辑', exact=True)).to_have_attribute('aria-selected', 'true')

    def choose(self, label, option, multi=False, exact=True):
        box = self.page.get_by_test_id('stMultiSelect' if multi else 'stSelectbox').filter(has=self.page.get_by_text(label, exact=True)).first
        box.locator('input').click()
        self.rerun(lambda: self.page.get_by_role('option', name=option, exact=exact).click())
        if multi:
            self.page.keyboard.press('Escape')

    def tab(self, name):
        self.rerun(lambda: self.page.get_by_role('tab', name=name, exact=True).click())

    def click(self, label):
        self.rerun(lambda: self.page.get_by_role('button', name=label, exact=True).click())

    def text(self, label, value):
        field = self.page.get_by_role('textbox', name=label, exact=True)
        field.fill(value)
        self.rerun(lambda: field.press('Tab'))

    def expand(self, label):
        summary = self.page.get_by_text(label, exact=True).locator('xpath=ancestor::summary')
        if not summary.evaluate('(el) => el.parentElement.open'):
            summary.click()

    def download(self, label, path):
        with self.page.expect_download() as download:
            self.click(label)
        download.value.save_as(str(path))
        return path.read_bytes()

    def applied(self, name):
        self.click('应用并预览')
        expect(self.page.get_by_role('tab', name='预览与下载', exact=True),
               str(self.page.get_by_test_id('stAlert').all_text_contents())).to_have_attribute('aria-selected', 'true', timeout=10_000)
        expect(self.page.get_by_text('当前结果：' + name, exact=True)).to_be_visible()

    def confirm(self):
        self.rerun(lambda: self.page.get_by_text('已查看预览，确认下载', exact=True).click())
