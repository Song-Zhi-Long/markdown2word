import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from converter import AppConfig, MarkdownToDocxConverter


def xpath(node, expression):
    return etree.XPath(expression, namespaces={
        'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    })(node)


class NumberedMathTests(unittest.TestCase):
    def setUp(self):
        self.converter = MarkdownToDocxConverter()
        self.output_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.output_dir.cleanup)

    def export(self, source):
        path = self.converter.convert(
            source, AppConfig(self.output_dir.name, str(Path.cwd()))
        )
        self.assertEqual(self.converter.last_warnings, [])
        return Document(path)

    def assert_display_number(self, document, tag):
        displays = document.element.xpath('.//m:oMathPara')
        self.assertEqual(len(displays), 1)
        display = displays[0]
        paragraph = display.getparent()
        # No regular text or tabs may share this paragraph with the equation:
        # Word otherwise switches it to inline and overrides above/below limits.
        self.assertEqual(paragraph.xpath('./w:r'), [])
        arrays = xpath(display, './m:oMath/m:eqArr')
        self.assertEqual(len(arrays), 1)
        self.assertEqual(
            xpath(arrays[0], './m:eqArrPr/m:maxDist')[0].get(qn('m:val')), '1'
        )
        runs = xpath(arrays[0], './m:e/m:r')
        self.assertEqual(xpath(runs[-2], './m:t')[0].text, '#')
        self.assertEqual(xpath(runs[-1], './m:t')[0].text, tag)
        self.assertEqual(len(xpath(runs[-1], './m:rPr/m:nor')), 1)
        return display

    def test_exact_focal_loss_sample(self):
        source = (Path(__file__).parent / 'fixtures/focal_loss_escaped.md').read_text(
            encoding='utf-8'
        )
        document = self.export(source)
        display = self.assert_display_number(document, '(3)')
        sums = xpath(display, './/m:nary')
        self.assertEqual(len(sums), 2)
        for nary, lower, upper in zip(sums, ('u=1', 'v=1'), ('WF', 'HF')):
            self.assertEqual(xpath(nary, './m:naryPr/m:chr')[0].get(qn('m:val')), '∑')
            self.assertEqual(xpath(nary, './m:naryPr/m:limLoc')[0].get(qn('m:val')), 'undOvr')
            self.assertEqual(''.join(xpath(nary, './m:sub//m:t/text()')), lower)
            self.assertEqual(''.join(xpath(nary, './m:sup//m:t/text()')), upper)
            self.assertTrue(xpath(nary, './m:e/*'))
        self.assertEqual(len(xpath(display, './/m:f')), 1)
        self.assertEqual(len(xpath(display, './/m:acc')), 1)
        self.assertEqual(len(document.element.xpath('.//m:oMath')), 2)

    def test_numbered_delimiters_and_labels(self):
        for opening, closing in (('$$', '$$'), (r'\[', r'\]'), (r'\\[', r'\\]')):
            for suffix, tag in ((r'\tag{A.3}', '(A.3)'), (r'\tag*{(3)}', '(3)'), (r'\quad (3)', '(3)')):
                with self.subTest(opening=opening, suffix=suffix):
                    document = self.export(opening + r'\sum_{i=1}^n i ' + suffix + closing)
                    self.assert_display_number(document, tag)

    def test_unnumbered_display_is_not_wrapped_in_numbering_array(self):
        document = self.export(r'$$\sum_{i=1}^n i$$')
        self.assertEqual(len(document.element.xpath('.//m:oMathPara')), 1)
        self.assertEqual(document.element.xpath('.//m:eqArr'), [])

    def test_inline_sum_keeps_side_scripts(self):
        document = self.export(r'行内 $\sum_{i=1}^n i$ 继续正文。')
        self.assertEqual(document.element.xpath('.//m:oMathPara'), [])
        self.assertEqual(document.element.xpath('.//m:nary'), [])
        self.assertEqual(len(document.element.xpath('.//m:sSubSup')), 1)

    def test_explicit_nolimits_is_preserved(self):
        document = self.export(r'$$\sum\nolimits_{i=1}^n i\tag{1}$$')
        display = self.assert_display_number(document, '(1)')
        self.assertEqual(xpath(display, './/m:naryPr/m:limLoc')[0].get(qn('m:val')), 'subSup')

    def test_existing_numbered_cases_and_binomials_remain_editable(self):
        for body, element in (
            (r'\boxed{\begin{cases}x=0\\y=1\end{cases}}', 'm:borderBox'),
            (r'\binom n2=\frac{n(n-1)}{2}', 'm:d'),
            (r'\begin{aligned}a&=b\\c&=d\end{aligned}', 'm:m'),
        ):
            with self.subTest(body=body):
                document = self.export('$$' + body + r'\tag{4}$$')
                display = self.assert_display_number(document, '(4)')
                self.assertTrue(xpath(display, './/' + element))


if __name__ == '__main__':
    unittest.main()
