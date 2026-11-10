import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from converter import AppConfig, MarkdownToDocxConverter


CASES_BODY = (
    r"\boxed{ \begin{cases} m_X(t)=0\\[4pt] "
    r"R_X(t_1,t_2)=\dfrac{E(V^2)}{2} \cos\omega(t_1-t_2) \end{cases} }"
)
USER_SAMPLE = (
    r"只要振幅 \\(V\\) 与均匀随机相位 \\(\Theta\\) 相互独立，就有："
    "\n\n" + r"\\[ " + CASES_BODY.replace(r"\\[4pt]", r"\\\\[4pt]") + r" \\]"
    "\n\n因此，这类随机相位正弦过程都具有宽平稳性。"
)


class MathConversionTests(unittest.TestCase):
    def setUp(self):
        self.converter = MarkdownToDocxConverter()

    def extract(self, source):
        prepared = self.converter._prepare_markdown(source)
        return prepared, list(self.converter._math_tokens.values())

    def assert_cases_layout(self, omml):
        boxes = omml.xpath(".//*[local-name()='borderBox']")
        self.assertEqual(len(boxes), 1)
        delimiters = boxes[0].xpath("./*[local-name()='e']/*[local-name()='d']")
        self.assertEqual(len(delimiters), 1)
        delimiter = delimiters[0]
        properties = delimiter.xpath("./*[local-name()='dPr']")[0]
        values = {etree.QName(child).localname: child.get(qn("m:val")) for child in properties}
        self.assertEqual(values, {"begChr": "{", "endChr": "", "grow": "1"})
        matrices = delimiter.xpath("./*[local-name()='e']/*[local-name()='m']")
        self.assertEqual(len(matrices), 1)
        rows = matrices[0].xpath("./*[local-name()='mr']")
        self.assertEqual(len(rows), 2)
        for row in rows:
            self.assertTrue("".join(row.itertext()).strip())
        alignment = matrices[0].xpath(".//*[local-name()='mcJc']")[0]
        self.assertEqual(alignment.get(qn("m:val")), "left")
        self.assertEqual(len(omml.xpath(".//*[local-name()='f']")), 1)

    def test_exact_user_sample(self):
        prepared, formulas = self.extract(USER_SAMPLE)
        self.assertNotIn("\\", prepared)
        self.assertEqual(len(formulas), 3)
        body = next(item["latex"] for item in formulas if item["display"])
        self.assertEqual(body, CASES_BODY)
        self.assert_cases_layout(self.converter.latex_to_omml(body, True))

    def test_standard_and_escaped_cases_variants(self):
        for opening, closing in (("$$", "$$"), (r"\[", r"\]"), (r"\\[", r"\\]")):
            for spacing in (r"\\", r"\\[4pt]", r"\\ [4pt]", r"\\[-2pt]"):
                with self.subTest(opening=opening, spacing=spacing):
                    body = CASES_BODY.replace(r"\\[4pt]", spacing)
                    _, formulas = self.extract(opening + body + closing)
                    self.assert_cases_layout(self.converter.latex_to_omml(formulas[0]["latex"], True))
        _, formulas = self.extract(r"\\[" + CASES_BODY.replace("\\", "\\\\") + r"\\]")
        self.assert_cases_layout(self.converter.latex_to_omml(formulas[0]["latex"], True))

    def test_export_contains_editable_equations_without_fallback(self):
        with tempfile.TemporaryDirectory() as output_dir:
            output = self.converter.convert(
                USER_SAMPLE, AppConfig(output_dir=output_dir, asset_root=str(Path.cwd()))
            )
            document = Document(output)
            self.assertEqual(self.converter.last_warnings, [])
            self.assertEqual(len(document.element.xpath(".//m:oMath")), 3)
            display = document.element.xpath(".//m:oMathPara")[0]
            self.assert_cases_layout(display)
            text = "".join(document.element.xpath(".//w:t/text()"))
            self.assertNotIn("\\", text)
            self.assertNotIn("$$", text)
            self.assertIn("宽平稳性", text)

    def test_previous_covariance_sample(self):
        source = r"\\[ C_X=E\left\\{(U-\mu_U)^2\right\\} \\]"
        prepared, formulas = self.extract(source)
        self.assertNotIn("\\", prepared)
        self.assertEqual(formulas[0]["latex"], r"C_X=E\left\{(U-\mu_U)^2\right\}")
        self.converter.latex_to_omml(formulas[0]["latex"], True)

    def test_original_delimiters(self):
        body = r"\frac{a}{b}+\mu_U\cos 2t"
        for opening, closing, display in (
            ("$", "$", False), ("$$", "$$", True),
            (r"\(", r"\)", False), (r"\[", r"\]", True),
            (r"\\(", r"\\)", False), (r"\\[", r"\\]", True),
        ):
            with self.subTest(opening=opening):
                _, formulas = self.extract(opening + body + closing)
                self.assertEqual(formulas, [{"latex": body, "display": display}])

    def test_native_matrix_rows_are_preserved(self):
        for environment in ("pmatrix", "aligned", "cases", "array"):
            columns = "{rl}" if environment == "array" else ""
            body = rf"\begin{{{environment}}}{columns}a&=b\\c&=d\end{{{environment}}}"
            for opening, closing in (("$$", "$$"), (r"\[", r"\]"), (r"\\[", r"\\]")):
                with self.subTest(environment=environment, opening=opening):
                    _, formulas = self.extract(opening + body + closing)
                    self.assertEqual(formulas[0]["latex"], body)
                    omml = self.converter.latex_to_omml(body, True)
                    matrices = omml.xpath(".//*[local-name()='m']")
                    self.assertEqual(len(matrices), 1)
                    self.assertEqual(len(matrices[0].xpath("./*[local-name()='mr']")), 2)

    def test_intentional_empty_rows_without_spacing_are_preserved(self):
        body = r"\begin{pmatrix}a\\\\b\end{pmatrix}"
        _, formulas = self.extract(r"\\[" + body + r"\\]")
        self.assertEqual(formulas[0]["latex"], body)

    def test_nested_matrix_delimiters(self):
        for body in (
            r"\boxed{\begin{pmatrix}a&b\\c&d\end{pmatrix}}",
            r"\frac{\begin{pmatrix}a&b\\c&d\end{pmatrix}}{2}",
            r"\boxed{\boxed{\begin{cases}x\\y\end{cases}}}",
        ):
            with self.subTest(body=body):
                omml = self.converter.latex_to_omml(body, True)
                self.assertEqual(len(omml.xpath(".//*[local-name()='d']")), 1)
                self.assertNotIn("{", "".join(omml.xpath(".//*[local-name()='t']/text()")))

    def test_complex_formula_features(self):
        for body, tag in (
            (r"\sum_{i=1}^n\sum_{j=1}^m a_{ij}", "nary"),
            (r"\widehat{X}", "acc"),
            (r"\underbrace{a+b}_{n}+\overbrace{c+d}^{m}", "groupChr"),
        ):
            with self.subTest(body=body):
                omml = self.converter.latex_to_omml(body, True)
                self.assertTrue(omml.xpath(f".//*[local-name()='{tag}']"))

    def test_blockquote_and_code(self):
        for slashes in ("\\", "\\\\"):
            prepared, formulas = self.extract(f"> {slashes}[\n> x+y\n> {slashes}]")
            self.assertEqual(formulas, [{"latex": "x+y", "display": True}])
            self.assertIn("> MATHBLOCKTOK0END", prepared)
        source = r"`\\(x\\)`" + "\n\n```latex\n" + r"\\[x+y\\]" + "\n```"
        prepared, formulas = self.extract(source)
        self.assertEqual(prepared, source)
        self.assertEqual(formulas, [])


if __name__ == "__main__":
    unittest.main()
