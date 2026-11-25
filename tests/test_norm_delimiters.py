import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from converter import AppConfig, MarkdownToDocxConverter


BODY = r'L_{offset}=\sum_p\left\|\hat{\delta}_p-\delta_p^{GT}\right\|_1'


class NormDelimiterTests(unittest.TestCase):
    def setUp(self):
        self.converter = MarkdownToDocxConverter()

    def extract(self, source):
        self.converter._prepare_markdown(source)
        return next(iter(self.converter._math_tokens.values()))['latex']

    def assert_norm(self, omml, subscript='1'):
        norms = [
            node for node in omml.xpath('.//*[local-name()="d"]')
            if node.find(qn('m:dPr') + '/' + qn('m:begChr')) is not None
            and node.find(qn('m:dPr') + '/' + qn('m:begChr')).get(qn('m:val')) == '‖'
        ]
        self.assertEqual(len(norms), 1)
        norm = norms[0]
        self.assertEqual(norm.find(qn('m:dPr') + '/' + qn('m:endChr')).get(qn('m:val')), '‖')
        script = norm.getparent().getparent()
        self.assertEqual(script.tag, qn('m:sSub'))
        self.assertEqual(''.join(script.find(qn('m:sub')).itertext()), subscript)
        self.assertFalse(any('\\' in value for node in omml.iter() for value in node.attrib.values()))
        return norm

    def test_exact_offset_loss_has_native_norm_and_complete_summand(self):
        source = (Path(__file__).parent / 'fixtures/offset_norm_escaped.md').read_text(encoding='utf-8')
        with tempfile.TemporaryDirectory() as output_dir:
            output = self.converter.convert(source, AppConfig(output_dir, str(Path.cwd())))
            document = Document(output)
        self.assertEqual(self.converter.last_warnings, [])
        norm = self.assert_norm(document.element)
        self.assertEqual(len(norm.xpath('.//*[local-name()="acc"]')), 1)
        self.assertIn('GT', ''.join(norm.itertext()))
        nary = document.element.xpath('.//m:nary')[0]
        self.assertEqual(norm.getparent().getparent().getparent(), nary.find(qn('m:e')))

    def test_mixed_and_fully_escaped_norms_in_supported_delimiters(self):
        for width in (2, 4, 8):
            for command_width in (1, 2, 4):
                for bar_width in (1, 2, 4):
                    with self.subTest(width=width, command_width=command_width, bar_width=bar_width):
                        body = BODY.replace('\\', '\\' * command_width)
                        body = body.replace('\\' * command_width + '|', '\\' * (command_width * bar_width) + '|')
                        latex = self.extract('\\' * width + '[' + body + '\\' * width + ']')
                        self.assertEqual(latex, BODY)
                        self.assert_norm(self.converter.latex_to_omml(latex, True))

    def test_native_norm_forms_stay_compatible(self):
        for opening, closing in (('$', '$'), ('$$', '$$'), (r'\(', r'\)'), (r'\[', r'\]')):
            with self.subTest(opening=opening):
                self.assertEqual(self.extract(opening + BODY + closing), BODY)
                self.assert_norm(self.converter.latex_to_omml(BODY, opening not in ('$', r'\(')))

    def test_matrix_row_break_before_plain_bar_is_not_collapsed(self):
        body = r'\begin{bmatrix}a\\|x|\end{bmatrix}'
        self.assertEqual(self.extract(r'\\[' + body + r'\\]'), body)
        omml = self.converter.latex_to_omml(body, True)
        array = omml.xpath('.//*[local-name()="eqArr"]')[0]
        self.assertEqual([''.join(row.itertext()) for row in array], ['a', '|x|'])

    def test_norm_around_matrix_keeps_rows_and_outer_subscript(self):
        source = r'\\[\left\\|\begin{bmatrix}a\\b\end{bmatrix}\right\\|_2\\]'
        latex = self.extract(source)
        norm = self.assert_norm(self.converter.latex_to_omml(latex, True), '2')
        self.assertEqual(len(norm.xpath('.//*[local-name()="eqArr"]/*[local-name()="e"]')), 2)

    def test_literal_text_and_code_are_preserved(self):
        source = r'\\[\text{literal \left\\|}\left\\|x\right\\|_1\\]'
        latex = self.extract(source)
        self.assertEqual(latex, r'\text{literal \left\\|}\left\|x\right\|_1')
        code = chr(96) + r'\left\\|x\right\\|' + chr(96)
        self.assertEqual(self.converter._prepare_markdown(code), code)
        self.assertEqual(self.converter._math_tokens, {})


if __name__ == '__main__':
    unittest.main()
