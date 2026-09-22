"""Full-page import evidence and the local native adapter's failure boundary."""
from io import BytesIO
import json
import os
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import patch

from PIL import Image

from docx_formatting import FormatError
from test_v06_import import grouped_sample, rendered
from v06_import import compare_rendered_pages
from v06_office import LocalOffice, native_office
from v06_plan import build_plan, intent
from v06_session import EditingSession, _convert
from test_v06 import sample, ENV


class FidelityTests(unittest.TestCase):
    def test_edit_report_uses_actual_page_changes_including_later_pages(self):
        original = rendered('same')
        original['pages'] = 3
        original['images'] *= 3
        altered = {**original, 'images': list(original['images'])}
        image = BytesIO()
        Image.new('RGB', (10, 10), 'black').save(image, format='PNG')
        altered['images'][2] = image.getvalue()
        for after, expected in [(altered, [3]), (original, [])]:
            with self.subTest(expected=expected):
                views = iter([original, after])
                session = EditingSession.load('synthetic.docx', sample(False), renderer=lambda *args: next(views))
                model = session.model()
                plan = build_plan(model, [intent(model.scope('all'), {'font_family': '黑体'})]).plan
                result, _ = session.execute(model, plan, environment=ENV)
                evidence = json.loads(result.report_json)['page_check']
                self.assertEqual(evidence['differing_pages'], expected)
                self.assertEqual(evidence['pixels_equal'], not expected)

    def load(self, old, new):
        views = iter([old, new])
        with patch('v06_session.validate_doc', return_value={}):
            return EditingSession.load('synthetic.doc', b'original', inspector=lambda data: (grouped_sample(), []),
                                       renderer=lambda *args: next(views))

    def assert_blocked(self, session):
        session.conversion['visual_review'] = 'confirmed'
        model = session.model()
        plan = build_plan(model, [intent(model.scope('all'), {'font_size': 12})]).plan
        with self.assertRaises(FormatError):
            session.execute(model, plan, environment={'font_families': []})
        self.assertEqual(list(session.versions), ['original'])

    def test_added_numbering_is_blocked_even_when_no_character_is_missing(self):
        session = self.load(rendered('first second'), rendered('first、second'))
        self.assertEqual(session.conversion['text_check']['missing_count'], 0)
        self.assert_blocked(session)

    def test_repagination_is_blocked_even_with_identical_text(self):
        new = rendered('same'); new['pages'] = 2; new['images'] *= 2
        session = self.load(rendered('same'), new)
        self.assertTrue(session.conversion['rendered_text_equal'])
        self.assertEqual(session.conversion['page_check']['differing_pages'], [2])
        self.assert_blocked(session)

    def test_missing_picture_is_not_hidden_by_identical_pdf_text(self):
        old, new = rendered('same'), rendered('same')
        image = Image.new('RGB', (4, 4), 'white'); image.putpixel((2, 1), (0, 0, 0))
        buf = BytesIO(); image.save(buf, format='PNG'); old['images'] = [buf.getvalue()]
        session = self.load(old, new)
        self.assertTrue(session.conversion['rendered_text_equal'])
        check = session.conversion['page_check']
        self.assertFalse(check['pixels_equal'])
        self.assertEqual(check['differing_pages'], [1])
        self.assertEqual(check['pages'][0]['difference_bbox'], [2, 1, 3, 2])
        self.assertEqual(check['pages'][0]['changed_pixels'], 1)
        self.assert_blocked(session)

    def test_all_pages_and_image_dimensions_are_checked(self):
        old, new = rendered('same'), rendered('same')
        old['pages'] = new['pages'] = 3; old['images'] *= 3; new['images'] *= 3
        self.assertTrue(compare_rendered_pages(old, new)['pixels_equal'])
        buf = BytesIO(); Image.new('RGB', (5, 4), 'white').save(buf, format='PNG')
        new['images'][2] = buf.getvalue()
        check = compare_rendered_pages(old, new)
        self.assertEqual(check['identical_pages'], [1, 2])
        self.assertEqual(check['differing_pages'], [3])
        self.assertFalse(check['pages'][2]['same_size'])

    def test_incomplete_page_evidence_cannot_pass(self):
        new = rendered('same'); new['images'] = []
        self.assert_blocked(self.load(rendered('same'), new))

    def test_reordered_words_with_different_page_cannot_be_accepted(self):
        old, new = rendered('first second'), rendered('second first')
        buf = BytesIO(); Image.new('RGB', (4, 4), 'black').save(buf, format='PNG')
        new['images'] = [buf.getvalue()]
        self.assert_blocked(self.load(old, new))

    def test_unstable_extraction_from_identical_bytes_is_blocked(self):
        session = self.load(rendered('left right'), rendered('right left'))
        self.assertFalse(session.conversion['rendered_text_equal'])
        self.assertTrue(session.conversion['page_check']['pixels_equal'])
        self.assert_blocked(session)


class NativeAdapterTests(unittest.TestCase):
    def tearDown(self):
        native_office.cache_clear()

    def test_explicit_libreoffice_does_not_probe_native_office(self):
        native_office.cache_clear()
        with patch.dict(os.environ, {'WENXU_RENDERER': 'libreoffice'}), patch('v06_office._powershell') as probe:
            self.assertIsNone(native_office()); probe.assert_not_called()

    def test_explicit_native_mode_fails_closed_if_unavailable(self):
        native_office.cache_clear()
        with patch.dict(os.environ, {'WENXU_RENDERER': 'wps'}), patch('v06_office._powershell', return_value=None):
            with self.assertRaises(FormatError): native_office()

    def test_native_conversion_failure_never_falls_back_to_another_engine(self):
        with patch('v06_session.native_office') as native, patch('v06_session._libreoffice_convert') as fallback:
            native.return_value.convert.side_effect = FormatError('native failure')
            with self.assertRaises(FormatError): _convert(b'data', 'doc', 'pdf')
            fallback.assert_not_called()

    def test_doc_without_native_editor_is_not_converted(self):
        from v06_native import inspect_native_doc
        with patch('v06_native.native_office', return_value=None):
            with self.assertRaisesRegex(FormatError, '保留 DOC 原格式'):
                inspect_native_doc(b'original')

    def test_office_busy_error_is_actionable_and_does_not_leak_stderr(self):
        with patch('v06_office._path', side_effect=lambda path, **kw: str(path)): office = LocalOffice('powershell.exe', '/tmp')
        with patch('v06_office.subprocess.run', return_value=subprocess.CompletedProcess([], 1, b'', b'OFFICE_IN_USE private-name')):
            with self.assertRaisesRegex(FormatError, '正在处理其他文档') as error: office.call('metadata')
            self.assertNotIn('private-name', str(error.exception))

    def test_output_path_need_not_exist_and_temporary_input_is_cleaned(self):
        with tempfile.TemporaryDirectory() as temp:
            with patch('v06_office._path', side_effect=lambda path, **kw: str(path)): office = LocalOffice('powershell.exe', temp)
            seen = []
            def call(action, *args):
                directory = next(Path(temp).iterdir())
                self.assertEqual((directory / 'input.doc').read_bytes(), b'private')
                self.assertFalse((directory / 'output.docx').exists())
                self.assertEqual(action, 'docx'); seen.append(args)
                (directory / 'output.docx').write_bytes(b'converted')
            with patch.object(office, 'call', side_effect=call), patch('v06_office._path', side_effect=str):
                self.assertEqual(office.convert(b'private', 'doc', 'docx:filter'), b'converted')
            self.assertEqual(list(Path(temp).iterdir()), [])
            self.assertTrue(seen)


if __name__ == '__main__': unittest.main()
