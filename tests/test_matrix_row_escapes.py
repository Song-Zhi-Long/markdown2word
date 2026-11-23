import tempfile
import unittest
from pathlib import Path

from docx import Document

from converter import AppConfig, MarkdownToDocxConverter


def math_text(node):
    return ''.join(node.xpath('.//*[local-name()="t"]/text()'))


class MatrixRowEscapeTests(unittest.TestCase):
    def setUp(self):
        self.converter = MarkdownToDocxConverter()

    def vector_rows(self, omml):
        arrays = omml.xpath('.//*[local-name()="eqArr"]')
        self.assertEqual(len(arrays), 1)
        return [math_text(row).strip() for row in arrays[0].xpath('./*[local-name()="e"]')]

    def test_exact_projection_sample(self):
        source = (Path(__file__).parent / 'fixtures/projection_matrices_escaped.md').read_text(encoding='utf-8')
        with tempfile.TemporaryDirectory() as output_dir:
            output = self.converter.convert(source, AppConfig(output_dir, str(Path.cwd())))
            document = Document(output)
        self.assertEqual(self.converter.last_warnings, [])
        equations = document.element.xpath('.//m:oMathPara')
        self.assertEqual(len(equations), 3)
        self.assertNotIn('\\', math_text(document.element))
        self.assertEqual(self.vector_rows(equations[0]), ['xi', 'yi', 'zi'])
        self.assertEqual(self.vector_rows(equations[1]), ['2', '1', '10'])
        arrays = equations[2].xpath('.//*[local-name()="eqArr"]')
        self.assertEqual(len(arrays), 2)
        self.assertEqual([
            [math_text(row).strip() for row in array.xpath('./*[local-name()="e"]')]
            for array in arrays
        ], [['2', '1', '10'], ['3400', '2500', '10']])
        self.assertEqual(len(equations[2].xpath('.//*[local-name()="d"]')), 3)

    def test_exact_camera_projection_treats_v_as_a_variable(self):
        source = (Path(__file__).parent / 'fixtures/camera_projection_escaped.md').read_text(encoding='utf-8')
        with tempfile.TemporaryDirectory() as output_dir:
            output = self.converter.convert(source, AppConfig(output_dir, str(Path.cwd())))
            document = Document(output)
        self.assertEqual(self.converter.last_warnings, [])
        self.assertNotIn('\\', math_text(document.element))
        arrays = document.element.xpath('.//m:eqArr')
        self.assertEqual(len(arrays), 2)
        self.assertEqual([
            [math_text(row).strip() for row in array.xpath('./*[local-name()="e"]')]
            for array in arrays
        ], [['u', 'v', '1'], ['X', 'Y', 'Z', '1']])

    def test_accent_letters_without_arguments_are_plain_matrix_variables(self):
        for variable in ('v', 'u', 'b', 'c', 'd', 'k', 'r', 't', 'H'):
            for ending in ('', ' ', '_i', '+1'):
                with self.subTest(variable=variable, ending=ending):
                    latex = r'\begin{bmatrix}1\\' + '\\' + variable + ending + r'\end{bmatrix}'
                    normalized = latex.replace('\\\\\\' + variable, '\\\\' + variable)
                    self.assertEqual(self.converter._preprocess_latex_for_math(latex, True), normalized)
                    self.assertNotIn('\\', math_text(self.converter.latex_to_omml(latex, True)))

    def test_accent_commands_with_arguments_are_preserved(self):
        for argument in ('{a}', ' a', r'\alpha'):
            for accent in ('v', 'u', 'c', 'H'):
                with self.subTest(accent=accent, argument=argument):
                    latex = r'\begin{bmatrix}1\\' + '\\' + accent + argument + r'\\2\end{bmatrix}'
                    self.assertEqual(self.converter._preprocess_latex_for_math(latex, True), latex)

    def test_copied_rows_before_variables_numbers_and_signs(self):
        for body, rows in (
            (r'x_i\\\y_i\\\z_i', ['xi', 'yi', 'zi']),
            (r'2\\\1\\\10', ['2', '1', '10']),
            (r'3400\\\2500\\\10', ['3400', '2500', '10']),
            (r'2\\\-1\\\+10', ['2', '−1', '+10']),
        ):
            for display in (False, True):
                with self.subTest(body=body, display=display):
                    latex = r'\begin{bmatrix}' + body + r'\end{bmatrix}'
                    self.assertEqual(self.vector_rows(self.converter.latex_to_omml(latex, display)), rows)

    def test_legitimate_commands_after_row_break_are_preserved(self):
        for body in (
            r'1\\\vdots\\1',
            r'1\\\alpha\\\frac{1}{2}',
            r'1\\\mathbf{x}\\\text{row}',
        ):
            with self.subTest(body=body):
                latex = r'\begin{bmatrix}' + body + r'\end{bmatrix}'
                self.assertEqual(self.converter._preprocess_latex_for_math(latex, True), latex)
                self.assertNotIn('\\', math_text(self.converter.latex_to_omml(latex, True)))

    def test_native_rows_and_intentional_empty_rows_are_preserved(self):
        for body, rows in ((r'1\\2\\3', ['1', '2', '3']), (r'1\\\\2', ['1', '', '2'])):
            with self.subTest(body=body):
                latex = r'\begin{bmatrix}' + body + r'\end{bmatrix}'
                self.assertEqual(self.converter._preprocess_latex_for_math(latex, True), latex)
                self.assertEqual(self.vector_rows(self.converter.latex_to_omml(latex, True)), rows)

    def test_literal_text_and_unknown_multiletter_commands_are_not_rewritten(self):
        for body in (
            r'\text{literal\\\x}\\\y_i',
            r'\text{nested {literal\\\x}}\\\y_i',
            r'1\\\text{literal\\\x}\\\y_i',
            r'1\\\custommacro',
            r'1\\\i',
        ):
            with self.subTest(body=body):
                latex = r'\begin{bmatrix}' + body + r'\end{bmatrix}'
                prepared = self.converter._preprocess_latex_for_math(latex, True)
                self.assertEqual(prepared, latex.replace(r'\\\y_i', r'\\y_i'))
        self.assertEqual(self.converter._preprocess_latex_for_math(r'x\\\y', True), r'x\\\y')


if __name__ == '__main__':
    unittest.main()
