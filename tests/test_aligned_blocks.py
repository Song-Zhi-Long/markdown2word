import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from converter import AppConfig, MarkdownToDocxConverter


class AlignedBlockTests(unittest.TestCase):
    def setUp(self):
        self.converter = MarkdownToDocxConverter()
        self.output_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.output_dir.cleanup)

    def export(self, source, indent):
        output = self.converter.convert(source, AppConfig(
            self.output_dir.name, str(Path.cwd()), body_first_line_indent=indent
        ))
        self.assertEqual(self.converter.last_warnings, [])
        return Document(output)

    def assert_format(self, paragraph, alignment, indent_pt=0):
        effective_alignment = paragraph.alignment
        if effective_alignment is None:
            effective_alignment = paragraph.style.paragraph_format.alignment
        self.assertEqual(effective_alignment, alignment)
        self.assertEqual(paragraph.paragraph_format.first_line_indent.pt, indent_pt)

    def test_exact_rare_sample_in_both_indent_modes(self):
        source = (Path(__file__).parent / 'fixtures/rare_aligned_escaped.md').read_text(encoding='utf-8')
        for indent in (True, False):
            with self.subTest(indent=indent):
                document = self.export(source, indent)
                self.assertEqual(len(document.paragraphs), 8)
                for paragraph in document.paragraphs[:5]:
                    self.assert_format(paragraph, WD_ALIGN_PARAGRAPH.CENTER)
                research = document.paragraphs[5]
                self.assertTrue(research.text.startswith('研究方向：'))
                self.assertTrue(any(run.bold and '研究方向：' in run.text for run in research.runs))
                self.assert_format(research, WD_ALIGN_PARAGRAPH.JUSTIFY)
                self.assert_format(document.paragraphs[-1], WD_ALIGN_PARAGRAPH.JUSTIFY, 24 if indent else 0)
                text = ''.join(document.element.xpath('.//w:t/text()'))
                self.assertNotIn('div', text)
                self.assertNotIn('\\', text)
                self.assertIn('peixi.xiong@intel.com', text)
                self.assertIn('pei.yu@microsoft.com', text)

    def test_native_divs_parse_markdown_and_override_global_indent(self):
        source = '<div align="center">\n\n**Authors**\n\nInstitutions\n\n</div>\n\n<div align="left">\n\n**Research:** text\n\n</div>\n\nBody'
        for indent in (True, False):
            with self.subTest(indent=indent):
                paragraphs = self.export(source, indent).paragraphs
                self.assertEqual([p.text for p in paragraphs], ['Authors', 'Institutions', 'Research: text', 'Body'])
                for paragraph in paragraphs[:2]:
                    self.assert_format(paragraph, WD_ALIGN_PARAGRAPH.CENTER)
                self.assertTrue(paragraphs[0].runs[0].bold)
                self.assert_format(paragraphs[2], WD_ALIGN_PARAGRAPH.JUSTIFY)
                self.assert_format(paragraphs[3], WD_ALIGN_PARAGRAPH.JUSTIFY, 24 if indent else 0)

    def test_nested_alignment_uses_nearest_container_and_restores_parent(self):
        source = '<div align="center">\n\nOuter\n\n<div align="left">\n\nInner\n\n</div>\n\nOuter again\n\n</div>'
        paragraphs = self.export(source, True).paragraphs
        self.assertEqual([p.text for p in paragraphs], ['Outer', 'Inner', 'Outer again'])
        for paragraph, alignment in zip(paragraphs, (WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.JUSTIFY, WD_ALIGN_PARAGRAPH.CENTER)):
            self.assert_format(paragraph, alignment)

    def test_simple_multiline_blocks_and_line_breaks_inherit_alignment(self):
        for source in ('<p align="center">First\nSecond</p>', '<div align="center">First<br>Second</div>'):
            with self.subTest(source=source):
                paragraphs = self.export(source, True).paragraphs
                self.assertEqual([p.text for p in paragraphs], ['First', 'Second'])
                for paragraph in paragraphs:
                    self.assert_format(paragraph, WD_ALIGN_PARAGRAPH.CENTER)

    def test_mixed_html_container_preserves_inline_text_and_blocks(self):
        document = self.export('<div align="center">First <b>bold</b><p>Second</p>Third</div>', True)
        self.assertEqual([p.text for p in document.paragraphs], ['First bold', 'Second', 'Third'])
        for paragraph in document.paragraphs:
            self.assert_format(paragraph, WD_ALIGN_PARAGRAPH.CENTER)

    def test_code_with_escaped_divs_stays_literal(self):
        for source in ('`\\<div align="center">`', '```html\n\\<div align="center">\\\n\\*\\*literal\\*\\*\\\n\\</div>\n```'):
            with self.subTest(source=source):
                document = self.export(source, True)
                text = ''.join(document.element.xpath('.//w:t/text()'))
                self.assertIn(r'\<div align="center">', text)
                if '```' in source:
                    self.assertIn(r'\*\*literal\*\*' + '\\', text)

    def test_math_inside_aligned_blocks_stays_editable_and_display_math_centered(self):
        document = self.export('<div align="left">\n\nText $x_i$\n\n$$\\sum_{i=1}^n -x_i\\tag{1}$$\n\n</div>', True)
        self.assertEqual(len(document.element.xpath('.//m:oMath')), 2)
        self.assert_format(document.paragraphs[0], WD_ALIGN_PARAGRAPH.JUSTIFY)
        self.assert_format(document.paragraphs[1], WD_ALIGN_PARAGRAPH.CENTER)


if __name__ == '__main__':
    unittest.main()
