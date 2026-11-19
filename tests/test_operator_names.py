import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from converter import AppConfig, MarkdownToDocxConverter


def math_text(node):
    return ''.join(node.xpath('.//*[local-name()="t"]/text()'))


class OperatorNameTests(unittest.TestCase):
    def setUp(self):
        self.converter = MarkdownToDocxConverter()

    def test_exact_smooth_l1_sample_has_complete_summand(self):
        source = (Path(__file__).parent / 'fixtures/smooth_l1_escaped.md').read_text(encoding='utf-8')
        with tempfile.TemporaryDirectory() as output_dir:
            output = self.converter.convert(source, AppConfig(output_dir, str(Path.cwd())))
            document = Document(output)
        self.assertEqual(self.converter.last_warnings, [])
        sums = document.element.xpath('.//m:borderBox//m:nary')
        self.assertEqual(len(sums), 1)
        nary = sums[0]
        self.assertEqual(math_text(nary.xpath('./*[local-name()="sub"]')[0]), 'b∈{x,y,z,w,l,h,θ}')
        operand = nary.xpath('./*[local-name()="e"]')[0]
        self.assertEqual(math_text(operand), 'SmoothL1(Δb)')
        self.assertIsNone(nary.getnext())
        self.assert_upright_name(operand, 'SmoothL1')

    def assert_upright_name(self, omml, name):
        runs = omml.xpath('.//*[local-name()="r"]')
        matching_runs = [run for run in runs if math_text(run) == name]
        self.assertTrue(matching_runs, name)
        for run in matching_runs:
            styles = run.xpath('./*[local-name()="rPr"]/*[local-name()="sty"]')
            self.assertEqual([style.get(qn('m:val')) for style in styles], ['p'])

    def test_letter_and_digit_names_remain_upright_in_inline_and_display_math(self):
        for name in ('SmoothL1', 'L2Loss', 'Huber', 'Loss123', 'L1 L2'):
            for display in (False, True):
                with self.subTest(name=name, display=display):
                    omml = self.converter.latex_to_omml(r'\operatorname{' + name + '}(x)', display)
                    self.assertEqual(math_text(omml), name + '(x)')
                    self.assert_upright_name(omml, name)
                    self.assertNotIn('OPERATORNAMETOKEN', math_text(omml))

    def test_multiple_names_and_limit_variants(self):
        for command in (r'\operatorname', r'\operatorname*', r'\operatornamewithlimits'):
            with self.subTest(command=command):
                omml = self.converter.latex_to_omml(
                    command + r'{Loss1}(x)+\operatorname{Loss2}(y)', True
                )
                self.assertEqual(math_text(omml), 'Loss1(x)+Loss2(y)')
                self.assert_upright_name(omml, 'Loss1')
                self.assert_upright_name(omml, 'Loss2')

    def test_negative_named_summand_and_external_addition(self):
        omml = self.converter.latex_to_omml(r'\sum_{i=1}^n -\operatorname{Loss1}(x_i)+c', True)
        nary = omml.xpath('.//*[local-name()="nary"]')[0]
        self.assertEqual(math_text(nary.xpath('./*[local-name()="e"]')[0]), '−Loss1(xi)')
        self.assertEqual(math_text(nary.getnext()), '+c')

    def test_numbered_nested_sum_owns_function_argument(self):
        with tempfile.TemporaryDirectory() as output_dir:
            output = self.converter.convert(
                r'$$\sum_i\sum_j\operatorname{Loss2}(x_{ij})\tag{2}$$',
                AppConfig(output_dir, str(Path.cwd())),
            )
            document = Document(output)
        self.assertEqual(self.converter.last_warnings, [])
        sums = document.element.xpath('.//m:oMathPara//m:nary')
        self.assertEqual(len(sums), 2)
        outer = sums[0].xpath('./*[local-name()="e"]')[0]
        inner = sums[1].xpath('./*[local-name()="e"]')[0]
        self.assertIs(sums[1].getparent(), outer)
        self.assertEqual(math_text(inner), 'Loss2(xij)')

    def test_literal_text_and_marker_like_names_are_preserved(self):
        for text in (r'literal \operatorname{Loss1}', 'OPERATORNAMETOKENAEND'):
            with self.subTest(text=text):
                omml = self.converter.latex_to_omml(
                    r'\text{' + text + r'}+\operatorname{Loss2}(x)', True
                )
                self.assertEqual(math_text(omml), text + '+Loss2(x)')
                self.assert_upright_name(omml, 'Loss2')


if __name__ == '__main__':
    unittest.main()
