import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from converter import AppConfig, MarkdownToDocxConverter


class UnderbraceWidthTests(unittest.TestCase):
    def test_depth_units_remain_centered_under_editable_braces(self):
        sample = Path(__file__).parent / 'fixtures/depth_units_underbraces.md'
        converter = MarkdownToDocxConverter()
        with tempfile.TemporaryDirectory() as directory:
            path = converter.convert(sample.read_text(encoding='utf-8'), AppConfig(directory, directory))
            doc = Document(path)
        self.assertEqual(converter.last_warnings, [])
        limits = doc.element.xpath('.//m:f/m:num/m:limLow') + doc.element.xpath('.//m:f/m:den/m:limLow')
        self.assertEqual(len(limits), 3)
        for limit, symbol, unit in zip(limits, ('fy', 'h', 'h2d'), ('pixel', 'm', 'pixel')):
            brace = limit.find(qn('m:e')).find(qn('m:groupChr'))
            self.assertEqual(brace.find(qn('m:groupChrPr')).find(qn('m:chr')).get(qn('m:val')), '⏟')
            base = brace.find(qn('m:e'))
            self.assertEqual(len(base), 3)
            left, content, right = base
            self.assertTrue(''.join(left.itertext()).isspace())
            self.assertEqual(''.join(left.itertext()), ''.join(right.itertext()))
            self.assertEqual(''.join(content.itertext()), symbol)
            self.assertEqual(''.join(limit.find(qn('m:lim')).itertext()), unit)

    def test_long_compound_nested_and_upper_braces_keep_natural_width(self):
        converter = MarkdownToDocxConverter()
        for source in (r'\underbrace{abcdef}', r'\underbrace{a+b}', r'\underbrace{\frac{x}{y}}',
                       r'\underbrace{x_{\frac{a}{b}}}', r'\overbrace{h}'):
            with self.subTest(source=source):
                math = converter.latex_to_omml(source, False)
                brace = math.find('.//' + qn('m:groupChr'))
                self.assertIsNotNone(brace)
                self.assertFalse(any(text.text.isspace() for text in brace.iter(qn('m:t'))))
