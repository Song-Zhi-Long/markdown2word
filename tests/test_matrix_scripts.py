import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from converter import AppConfig, MarkdownToDocxConverter


def child(node, name):
    return node.find(qn('m:' + name))


def math_text(node):
    return ''.join(node.xpath('.//*[local-name()="t"]/text()'))


class MatrixScriptTests(unittest.TestCase):
    def setUp(self):
        self.converter = MarkdownToDocxConverter()

    def assert_scripted_matrix(self, omml, opening='[', closing=']', sub=None, sup=None, rows=2):
        delimiters = omml.xpath('.//*[local-name()="d"]')
        self.assertEqual(len(delimiters), 1)
        delimiter = delimiters[0]
        properties = child(delimiter, 'dPr')
        begin = child(properties, 'begChr')
        end = child(properties, 'endChr')
        self.assertEqual(begin.get(qn('m:val')) if begin is not None else '(', opening)
        self.assertEqual(end.get(qn('m:val')) if end is not None else ')', closing)
        grow = child(properties, 'grow')
        if grow is not None:
            self.assertIn(grow.get(qn('m:val')), ('1', 'true', 'on'))
        base = delimiter.getparent()
        self.assertEqual(base.tag, qn('m:e'))
        script = base.getparent()
        self.assertEqual(script.tag, qn('m:sSubSup' if sub is not None and sup is not None else 'm:sSub' if sub is not None else 'm:sSup'))
        if sub is not None:
            self.assertEqual(math_text(child(script, 'sub')), sub)
        if sup is not None:
            self.assertEqual(math_text(child(script, 'sup')), sup)
        matrix_rows = delimiter.xpath('.//*[local-name()="mr"]')
        if not matrix_rows:
            arrays = delimiter.xpath('.//*[local-name()="eqArr"]')
            self.assertEqual(len(arrays), 1)
            matrix_rows = arrays[0].xpath('./*[local-name()="e"]')
        self.assertEqual(len(matrix_rows), rows)
        return matrix_rows

    def test_exact_column_vector_preserves_brackets_rows_and_size(self):
        source = (Path(__file__).parent / 'fixtures/column_vector_escaped.md').read_text(encoding='utf-8')
        with tempfile.TemporaryDirectory() as output_dir:
            output = self.converter.convert(source, AppConfig(output_dir, str(Path.cwd())))
            document = Document(output)
        self.assertEqual(self.converter.last_warnings, [])
        self.assertEqual(len(document.element.xpath('.//m:oMathPara')), 1)
        rows = self.assert_scripted_matrix(document.element, sub='N×1', rows=4)
        self.assertEqual([math_text(row).strip() for row in rows], ['1', '1', '⋮', '1'])
        self.assertNotIn('\\', math_text(document.element))
        bold = document.element.xpath('.//m:sty[@m:val="b"]')
        self.assertTrue(bold)
        self.assertIn('然后将公式（1）改写为', ''.join(document.element.xpath('.//w:t/text()')))

    def test_matrix_environment_scripts_in_inline_and_display_modes(self):
        for environment, opening, closing in (
            ('bmatrix', '[', ']'), ('pmatrix', '(', ')'),
            ('Bmatrix', '{', '}'), ('vmatrix', '|', '|'), ('Vmatrix', '‖', '‖'),
        ):
            for suffix, sub, sup in ((r'_{N\times2}', 'N×2', None), ('^T', None, 'T'), (r'_n^{-1}', 'n', '−1')):
                for display in (False, True):
                    with self.subTest(environment=environment, suffix=suffix, display=display):
                        body = r'\begin{' + environment + r'}a&b\\c&d\end{' + environment + '}' + suffix
                        omml = self.converter.latex_to_omml(body, display)
                        self.assert_scripted_matrix(omml, opening, closing, sub, sup)

    def test_grouped_and_stretchy_matrix_bases_stay_compatible(self):
        for body in (
            r'{\begin{bmatrix}1\\2\end{bmatrix}}_n',
            r'\left[\begin{matrix}1\\2\end{matrix}\right]_n',
        ):
            with self.subTest(body=body):
                self.assert_scripted_matrix(self.converter.latex_to_omml(body, True), sub='n')

    def test_nested_scripted_matrices_keep_both_delimiters(self):
        body = r'\begin{pmatrix}\begin{bmatrix}1\\2\end{bmatrix}_n&0\\0&1\end{pmatrix}^T'
        omml = self.converter.latex_to_omml(body, True)
        delimiters = omml.xpath('.//*[local-name()="d"]')
        self.assertEqual(len(delimiters), 2)
        scripts = omml.xpath('.//*[local-name()="sSub" or local-name()="sSup"]')
        self.assertEqual(len(scripts), 2)
        for script in scripts:
            self.assertTrue(child(script, 'e').xpath('./*[local-name()="d"]'))
        self.assertEqual(sorted(math_text(child(script, 'sub') if child(script, 'sub') is not None else child(script, 'sup')) for script in scripts), ['T', 'n'])

    def test_unfenced_matrices_and_ordinary_scripts_are_unchanged(self):
        omml = self.converter.latex_to_omml(r'\begin{matrix}1\\2\end{matrix}_n+x_i^2', True)
        self.assertEqual(omml.xpath('.//*[local-name()="d"]'), [])
        sub = omml.xpath('.//*[local-name()="sSub"]')[0]
        self.assertEqual(math_text(child(sub, 'sub')), 'n')
        self.assertTrue(child(sub, 'e').xpath('./*[local-name()="eqArr"]'))
        self.assertEqual(len(omml.xpath('.//*[local-name()="sSubSup"]')), 1)


if __name__ == '__main__':
    unittest.main()
