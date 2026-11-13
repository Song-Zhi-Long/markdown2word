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

CANONICAL_GAUSSIAN_FORMULAS = (
    r"G=\{g_i\}_{i=1}^{N}",
    r"\boldsymbol{\sigma}(\tilde{t}) = \boldsymbol{\sigma} + \boldsymbol{\Sigma}_{1:3,4} "
    r"\boldsymbol{\Sigma}_{4,4}^{-1} (\tilde{t}-\mu_t). \tag{1}",
    r"o(\tilde{t})=\omega(\tilde{t})*o, \tag{2}",
    r"\omega(\tilde{t}) = \exp\left( -\frac{1}{2} "
    r"\frac{(\tilde{t}-\mu_t)^2}{\Sigma_{4,4}} \right). \tag{3}",
)
ESCAPED_GAUSSIAN_FORMULAS = (
    r"G=\\\\{g_i\\\\}\\\_{i=1}^{N}",
    CANONICAL_GAUSSIAN_FORMULAS[1].replace("_", r"\\\_"),
    CANONICAL_GAUSSIAN_FORMULAS[2].replace("*", r"\\\*"),
    CANONICAL_GAUSSIAN_FORMULAS[3].replace(r"\Sigma_", r"\Sigma\\\_"),
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

    def test_multilayer_gaussian_formulas(self):
        for width in (2, 4):
            slashes = "\\" * width
            source = "\n".join(
                slashes + "[ " + body + " " + slashes + "]" + ("\\" if index == 2 else "")
                for index, body in enumerate(ESCAPED_GAUSSIAN_FORMULAS)
            )
            with self.subTest(width=width):
                prepared, formulas = self.extract(source)
                self.assertNotIn("\\", prepared)
                self.assertEqual([formula["latex"] for formula in formulas], list(CANONICAL_GAUSSIAN_FORMULAS))
                for formula, canonical in zip(formulas, CANONICAL_GAUSSIAN_FORMULAS):
                    actual_body, actual_tag = self.converter._extract_equation_tag(formula["latex"])
                    expected_body, expected_tag = self.converter._extract_equation_tag(canonical)
                    self.assertEqual(actual_tag, expected_tag)
                    actual = self.converter.latex_to_omml(actual_body, True)
                    expected = self.converter.latex_to_omml(expected_body, True)
                    self.assertEqual(etree.tostring(actual), etree.tostring(expected))
                    math_text = "".join(actual.xpath(".//*[local-name()='t']/text()"))
                    self.assertNotIn("_", math_text)
                    self.assertNotIn("\\", math_text)

    def test_gaussian_word_subscripts_multiplication_and_tags(self):
        source = "\n".join(r"\\\\[" + body + r"\\\\]" for body in ESCAPED_GAUSSIAN_FORMULAS)
        with tempfile.TemporaryDirectory() as output_dir:
            output = self.converter.convert(source, AppConfig(output_dir=output_dir, asset_root=str(Path.cwd())))
            document = Document(output)
            self.assertEqual(self.converter.last_warnings, [])
            self.assertEqual(len(document.element.xpath(".//m:oMath")), 4)
            scripts = document.element.xpath(".//m:sub")
            script_text = ["".join(script.itertext()) for script in scripts]
            for expected in ("i=1", "1:3,4", "4,4"):
                self.assertIn(expected, script_text)
            ordinary_text = "".join(document.element.xpath(".//w:t/text()"))
            for tag in ("(1)", "(2)", "(3)"):
                self.assertIn(tag, ordinary_text)
            math_text = "".join(document.element.xpath(".//m:t/text()"))
            self.assertIn("*", math_text)
            self.assertNotIn("_", math_text)
            self.assertNotIn("\\", math_text)
            self.assertNotIn("\\", ordinary_text)

    def test_four_slash_inline_blockquote_and_code(self):
        prepared, formulas = self.extract(r"文字 \\\\(x_i\\\\) 结束")
        self.assertNotIn("\\", prepared)
        self.assertEqual(formulas, [{"latex": "x_i", "display": False}])
        prepared, formulas = self.extract("> " + r"\\\\[" + "\n> x_i\n> " + r"\\\\]")
        self.assertIn("> MATHBLOCKTOK0END", prepared)
        self.assertEqual(formulas, [{"latex": "x_i", "display": True}])
        source = "```latex\n" + r"\\\\[x\\\_i\\\\]" + "\n```"
        self.assertEqual(self.extract(source), (source, []))

    def test_fully_escaped_matrix_commands_and_rows(self):
        body = r"\begin{pmatrix}a&b\\c&d\end{pmatrix}"
        for width in (2, 4):
            slashes = "\\" * width
            source = slashes + "[" + body.replace("\\", slashes) + slashes + "]"
            _, formulas = self.extract(source)
            self.assertEqual(formulas[0]["latex"], body)
            omml = self.converter.latex_to_omml(formulas[0]["latex"], True)
            self.assertEqual(len(omml.xpath(".//*[local-name()='mr']")), 2)

    def test_literal_underscore_is_preserved(self):
        for opening, closing in (("$$", "$$"), (r"\[", r"\]"), (r"\\[", r"\\]"), (r"\\\\[", r"\\\\]")):
            body = r"\text{file\_name}+x_i"
            _, formulas = self.extract(opening + body + closing)
            self.assertEqual(formulas[0]["latex"], body)

    def test_multilayer_mismatched_delimiters_are_not_consumed(self):
        for source in (r"\\\\[x\\]", r"\\[x\\\\]", r"\\\[x\\\]"):
            self.assertEqual(self.extract(source), (source, []))

    def test_tilde_and_hat_use_native_accents(self):
        for command, expected in (("tilde", "\u0303"), ("widetilde", "\u0303"), ("widehat", "\u0302")):
            for display in (False, True):
                with self.subTest(command=command, display=display):
                    omml = self.converter.latex_to_omml("\\" + command + "{t}", display)
                    accents = omml.xpath(".//*[local-name()='acc']")
                    self.assertEqual(len(accents), 1)
                    symbol = accents[0].xpath("./*[local-name()='accPr']/*[local-name()='chr']")[0]
                    self.assertEqual(symbol.get(qn("m:val")), expected)
                    self.assertEqual("".join(accents[0].xpath("./*[local-name()='e']//*[local-name()='t']/text()")), "t")
                    self.assertFalse(omml.xpath(".//*[local-name()='limUpp']"))


if __name__ == "__main__":
    unittest.main()
