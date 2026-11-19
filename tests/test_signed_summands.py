import tempfile
import unittest
from pathlib import Path

from docx import Document
from latex2mathml.converter import convert
from lxml import etree

from converter import AppConfig, MarkdownToDocxConverter


def math_text(node):
    return ''.join(node.xpath('.//*[local-name()="t"]/text()'))


class SignedSummandTests(unittest.TestCase):
    def setUp(self):
        self.converter = MarkdownToDocxConverter()

    def test_exact_boxed_loss_has_complete_nested_summand(self):
        source = (Path(__file__).parent / 'fixtures/focal_loss_expanded_escaped.md').read_text(encoding='utf-8')
        with tempfile.TemporaryDirectory() as output_dir:
            output = self.converter.convert(source, AppConfig(output_dir, str(Path.cwd())))
            document = Document(output)
        self.assertEqual(self.converter.last_warnings, [])
        boxes = document.element.xpath('.//m:borderBox')
        self.assertEqual(len(boxes), 1)
        sums = boxes[0].xpath('.//*[local-name()="nary"]')
        self.assertEqual(len(sums), 2)
        outer_operand = sums[0].find('{http://schemas.openxmlformats.org/officeDocument/2006/math}e')
        inner_operand = sums[1].find('{http://schemas.openxmlformats.org/officeDocument/2006/math}e')
        self.assertIn(sums[1], list(outer_operand))
        self.assertTrue(len(inner_operand))
        self.assertEqual(math_text(inner_operand), '−αu,v(1−pt(u,v))2logpt(u,v)')
        self.assertEqual(sums[1].getparent(), outer_operand)
        # The entire signed loss must belong to the inner sum, not sit beside it.
        self.assertEqual(len(outer_operand), 1)
        self.assertEqual(sums[0].getnext(), None)

    def test_leading_sign_is_part_of_operand_but_later_binary_sign_is_not(self):
        for sign, rendered_sign in (('-', '−'), ('+', '+'), (r'\pm', '±'), (r'\mp', '∓')):
            for display in (True, False):
                with self.subTest(sign=sign, display=display):
                    # Inspect n-ary ownership before inline conversion flattens it.
                    latex = self.converter._preprocess_latex_for_math(r'\sum_{i=1}^n ' + sign + r' x_i-y', display)
                    root = etree.fromstring(convert(latex).encode())
                    self.converter._normalize_mathml_tree(root)
                    omml = self.converter._math_transform(root).getroot()
                    self.converter._repair_nary_operand(omml)
                    operand = omml.xpath('.//*[local-name()="nary"]/*[local-name()="e"]')[0]
                    self.assertEqual(math_text(operand), rendered_sign + 'xi')
                    nary = operand.getparent()
                    self.assertEqual(math_text(nary.getnext()), '−y')

    def test_parenthesized_signs_stay_inside_and_later_addition_stays_outside(self):
        for body, expected in (
            (r'-(a_i+b_i)+c', '−(ai+bi)'),
            (r'-\alpha_i(1-p_i)^2\log p_i+c', '−αi(1−pi)2logpi'),
            (r'-\left(a_i-b_i\right)+c', '−ai−bi'),
            (r'\frac{-a_i}{1-p_i}+c', '−ai1−pi'),
        ):
            with self.subTest(body=body):
                omml = self.converter.latex_to_omml(r'\sum_{i=1}^n ' + body, True)
                nary = omml.xpath('.//*[local-name()="nary"]')[0]
                operand = nary.xpath('./*[local-name()="e"]')[0]
                self.assertEqual(math_text(operand), expected)
                self.assertEqual(math_text(nary.getnext()), '+c')
                if r'\left' in body:
                    self.assertTrue(operand.xpath('.//*[local-name()="d"]'))

    def test_adjacent_sums_and_external_subtraction_remain_separate(self):
        omml = self.converter.latex_to_omml(r'\sum_{i=1}^n x_i-\sum_{j=1}^m -y_j', True)
        sums = omml.xpath('.//*[local-name()="nary"]')
        self.assertEqual(len(sums), 2)
        self.assertEqual(math_text(sums[0].xpath('./*[local-name()="e"]')[0]), 'xi')
        self.assertEqual(math_text(sums[0].getnext()), '−')
        self.assertEqual(math_text(sums[1].xpath('./*[local-name()="e"]')[0]), '−yj')
        self.assertIs(sums[0].getparent(), sums[1].getparent())


if __name__ == '__main__':
    unittest.main()
