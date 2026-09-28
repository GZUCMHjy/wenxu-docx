"""Template provenance follows document versions, never mutable UI selection."""
import json
import unittest

from docx_formatting import FormatError
from tests.test_v06 import ENV, fake_render, sample
from v06_plan import build_plan, intent
from v06_session import EditingSession


class TemplateVersionTests(unittest.TestCase):
    def setUp(self):
        self.session = EditingSession.load('synthetic.docx', sample(False), renderer=fake_render)
        self.binding = {'fingerprint': 'a' * 64, 'file_hash': 'b' * 64,
                        'name': '合成规范', 'basis_kind': 'sample_styles'}

    def edit(self, values, **kwargs):
        model = self.session.model()
        plan = build_plan(model, [intent(model.scope('all'), values)]).plan
        return self.session.execute(model, plan, environment=ENV, **kwargs)[0]

    def test_binding_is_frozen_and_inherited_only_from_selected_parent(self):
        v1 = self.edit({'font_size': 12}, template_binding=self.binding)
        self.binding['name'] = '外部对象后来变化'
        self.assertEqual(json.loads(v1.template_json)['name'], '合成规范')
        self.session.select(v1.id)
        v2 = self.edit({'font_size': 14})
        self.assertEqual(v2.template_json, v1.template_json)
        self.assertEqual(json.loads(v2.report_json)['template']['name'], '合成规范')
        self.session.undo()
        self.session.select('original')
        independent = self.edit({'font_size': 16})
        self.assertEqual(json.loads(independent.template_json), {})

    def test_same_plan_with_different_template_has_distinct_provenance(self):
        v1 = self.edit({'font_size': 12}, template_binding=self.binding)
        other = {**self.binding, 'fingerprint': 'c' * 64}
        v2 = self.edit({'font_size': 12}, template_binding=other)
        self.assertNotEqual(v1.id, v2.id)
        self.assertEqual(json.loads(v2.template_json)['fingerprint'], 'c' * 64)
        self.assertEqual(json.loads(v1.template_json)['fingerprint'], 'a' * 64)
        self.session.select(v2.id)
        self.assertIsNone(self.edit({'font_size': 12}, template_binding=self.binding))
        self.assertEqual(json.loads(v2.template_json)['fingerprint'], 'c' * 64)

    def test_invalid_binding_fails_before_new_version(self):
        with self.assertRaises(FormatError):
            self.edit({'font_size': 12}, template_binding={'fingerprint': 'bad'})
        self.assertEqual(list(self.session.versions), ['original'])


if __name__ == '__main__':
    unittest.main()
