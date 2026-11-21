import tempfile
import unittest
from pathlib import Path

import markdown
from docx import Document

from converter import AppConfig, MarkdownToDocxConverter


class EscapedEmailTests(unittest.TestCase):
    def setUp(self):
        self.converter = MarkdownToDocxConverter()
        self.output_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.output_dir.cleanup)

    def export(self, source):
        output = self.converter.convert(source, AppConfig(self.output_dir.name, str(Path.cwd())))
        self.assertEqual(self.converter.last_warnings, [])
        return Document(output)

    def text(self, document):
        return ''.join(document.element.xpath('.//w:t/text()'))

    def test_exact_email_sample(self):
        source = (Path(__file__).parent / 'fixtures/escaped_emails.md').read_text(encoding='utf-8')
        self.assertEqual(
            self.text(self.export(source)),
            '{hpark233,xruan9,djia7,tangw}@uic.edu, peixi.xiong@intel.com, pei.yu@microsoft.com',
        )

    def test_paired_backslashes_keep_their_literal_meaning(self):
        for count in range(1, 5):
            with self.subTest(backslashes=count):
                source = 'user' + '\\' * count + '@example.com'
                self.assertEqual(self.text(self.export(source)), 'user' + '\\' * (count // 2) + '@example.com')

    def test_inline_fenced_and_indented_code_stay_literal(self):
        for source in (r'`user\@example.com`', '```text\nuser\\@example.com\n```', '    user\\@example.com'):
            with self.subTest(source=source):
                self.assertEqual(self.text(self.export(source)), r'user\@example.com')

    def test_existing_escapes_and_math_stay_intact(self):
        document = self.export(r'路径 C:\Temp，邮箱 user\@example.com，\*星号\*，\_下划线\_，公式 $x_i+\alpha$。')
        text = self.text(document)
        self.assertIn(r'C:\Temp', text)
        self.assertIn('user@example.com', text)
        self.assertIn('*星号*', text)
        self.assertIn('_下划线_', text)
        self.assertEqual(len(document.element.xpath('.//m:oMath')), 1)
        self.assertEqual(self.converter._math_tokens['MATHINLINETOK0END']['latex'], r'x_i+\alpha')

    def test_email_link_text_and_target_are_preserved(self):
        document = self.export(r'[peixi.xiong\@intel.com](mailto:peixi.xiong@intel.com)')
        self.assertEqual(self.text(document), 'peixi.xiong@intel.com')
        targets = [relationship.target_ref for relationship in document.part.rels.values() if relationship.is_external]
        self.assertEqual(targets, ['mailto:peixi.xiong@intel.com'])

    def test_conversion_does_not_change_other_markdown_instances(self):
        before = list(markdown.Markdown().ESCAPED_CHARS)
        rendered_before = markdown.markdown(r'user\@example.com')
        self.export(r'user\@example.com')
        self.assertEqual(markdown.Markdown().ESCAPED_CHARS, before)
        self.assertEqual(markdown.markdown(r'user\@example.com'), rendered_before)


if __name__ == '__main__':
    unittest.main()
