import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from converter import AppConfig, MarkdownToDocxConverter


GAUSSIAN_SAMPLE = (Path(__file__).parent / "fixtures" / "gaussian_escaped.md").read_text(encoding="utf-8")
EXPECTED_BLOCKS = [
    r"G=\{g_i\}_{i=1}^{N}",
    r"\boldsymbol{\Sigma}=RSS^TR^T",
    r"\boldsymbol{\sigma}(\tilde{t}) = \boldsymbol{\sigma} + \boldsymbol{\Sigma}_{1:3,4} "
    r"\boldsymbol{\Sigma}_{4,4}^{-1} (\tilde{t}-\mu_t). \tag{1}",
    r"o(\tilde{t})=\omega(\tilde{t})*o, \tag{2}",
    r"\omega(\tilde{t}) = \exp\left( -\frac{1}{2} \frac{(\tilde{t}-\mu_t)^2}{\Sigma_{4,4}} \right). \tag{3}",
    r"c(t,\theta,\phi) = \cos\left(\frac{2\pi n}{T}t\right) Y_{lm}(\theta,\phi). \tag{4}",
]


class EscapedMathTests(unittest.TestCase):
    def setUp(self):
        self.converter = MarkdownToDocxConverter()

    def extract(self, source):
        prepared = self.converter._prepare_markdown(source)
        return prepared, list(self.converter._math_tokens.values())

    def test_full_gaussian_sample_is_normalized(self):
        prepared, formulas = self.extract(GAUSSIAN_SAMPLE)
        self.assertTrue(prepared.startswith("### **3.1."))
        self.assertNotIn("\\", prepared)
        self.assertEqual(len(formulas), 27)
        self.assertEqual([item["latex"] for item in formulas if item["display"]], EXPECTED_BLOCKS)
        self.assertIn({"latex": r"Y_{lm}(\theta,\phi)", "display": False}, formulas)

    def test_multilayer_delimiters_and_latex(self):
        for width in (1, 2, 4, 8):
            for opening, closing, display in (("(", ")", False), ("[", "]", True)):
                for body in (r"\frac{a_i}{b^2}+\mu_t", r"\text{file\_name}"):
                    for body_width in (1, width):
                        with self.subTest(width=width, opening=opening, body=body, body_width=body_width):
                            source_body = body.replace("\\", "\\" * body_width)
                            _, formulas = self.extract("\\" * width + opening + source_body + "\\" * width + closing)
                            self.assertEqual(formulas, [{"latex": body, "display": display}])

    def test_multilayer_matrices_preserve_rows(self):
        body = r"\begin{pmatrix}a&b\\c&d\end{pmatrix}"
        for width in (2, 4, 8):
            for body_width in (1, width):
                with self.subTest(width=width, body_width=body_width):
                    source_body = body.replace("\\", "\\" * body_width)
                    _, formulas = self.extract("\\" * width + "[" + source_body + "\\" * width + "]")
                    self.assertEqual(formulas[0]["latex"], body)

    def test_markdown_escaped_subscripts_and_star(self):
        for width in (2, 4, 8):
            for underscores in (1, 2, 3, 4):
                with self.subTest(width=width, underscores=underscores):
                    body = "Y" + "\\" * underscores + "_{lm}+x" + "\\" * underscores + "*o"
                    _, formulas = self.extract("\\" * width + "(" + body + "\\" * width + ")")
                    self.assertEqual(formulas[0]["latex"], "Y_{lm}+x*o")

    def test_literal_text_and_native_escaped_symbols_are_preserved(self):
        body = r"\text{file\_{id}\_name\*}+\operatorname{some\_name}"
        for width in (1, 2, 4, 8):
            with self.subTest(width=width):
                _, formulas = self.extract("\\" * width + "(" + body + "\\" * width + ")")
                self.assertEqual(formulas[0]["latex"], body)
        _, formulas = self.extract(r"$Y\_{lm}+x\*o$")
        self.assertEqual(formulas[0]["latex"], r"Y\_{lm}+x\*o")
        for source in (r"$a\_b$", r"\(a\_b\)"):
            _, formulas = self.extract(source)
            self.assertEqual(formulas[0]["latex"], r"a\_b")

    def test_copied_unbraced_subscripts(self):
        for width in (2, 4, 8):
            for escape_width in (1, 2, 3, 4):
                for body, expected in (
                    (r"[x]\_n", "[x]_n"), (r"x\_0", "x_0"),
                    (r"a\_b", "a_b"), (r"x\_\alpha", r"x_\alpha"),
                    (r"[x]\_n^2", "[x]_n^2"), (r"\sum\_k x^k", r"\sum_k x^k"),
                ):
                    with self.subTest(width=width, escape_width=escape_width, body=body):
                        copied = body.replace(r"\_", "\\" * escape_width + "_")
                        _, formulas = self.extract("\\" * width + "(" + copied + "\\" * width + ")")
                        self.assertEqual(formulas[0]["latex"], expected)

    def test_falling_factorial_exports_real_subscripts(self):
        with tempfile.TemporaryDirectory() as output_dir:
            for width, escape_width, command_width in ((2, 1, 1), (4, 3, 1), (4, 3, 4)):
                with self.subTest(width=width, escape_width=escape_width, command_width=command_width):
                    slashes = "\\" * width
                    underscore = "\\" * escape_width + "_"
                    source = (
                        "PPT 中定义：\n\n" + slashes + "[ [x]" + underscore + "n="
                        + "\\" * command_width + "sum" + underscore + "{k=0}^{n}S_1(n,k)x^k " + slashes + "]"
                        + "\n\n其中，" + slashes + "([x]" + underscore + "n" + slashes + ") 表示下降阶乘："
                    )
                    output = self.converter.convert(
                        source, AppConfig(output_dir=output_dir, asset_root=str(Path.cwd()))
                    )
                    document = Document(output)
                    self.assertEqual(self.converter.last_warnings, [])
                    self.assertEqual(len(document.element.xpath(".//m:oMath")), 2)
                    math_text = "".join(document.element.xpath(".//m:t/text()"))
                    self.assertNotIn("_", math_text)
                    self.assertNotIn("\\", math_text)
                    subscripts = ["".join(node.itertext()) for node in document.element.xpath(".//m:sub")]
                    self.assertEqual(subscripts.count("n"), 2)
                    self.assertIn("k=0", subscripts)
                    self.assertIn("1", subscripts)

    def test_tilde_and_hat_are_word_accents(self):
        for command, character in (("tilde", "\u0303"), ("widetilde", "\u0303"), ("hat", "\u0302"), ("widehat", "\u0302")):
            for display in (False, True):
                with self.subTest(command=command, display=display):
                    omml = self.converter.latex_to_omml("\\" + command + "{t}", display)
                    accents = omml.xpath(".//*[local-name()='acc']")
                    self.assertEqual(len(accents), 1)
                    mark = accents[0].xpath("./*[local-name()='accPr']/*[local-name()='chr']")[0]
                    self.assertEqual(mark.get(qn("m:val")), character)
                    self.assertEqual(omml.xpath(".//*[local-name()='limUpp']"), [])

    def test_export_preserves_scripts_accents_and_equation_tags(self):
        with tempfile.TemporaryDirectory() as output_dir:
            output = self.converter.convert(
                GAUSSIAN_SAMPLE, AppConfig(output_dir=output_dir, asset_root=str(Path.cwd()))
            )
            document = Document(output)
            self.assertEqual(self.converter.last_warnings, [])
            self.assertEqual(len(document.element.xpath(".//m:oMath")), 27)
            math_text = "".join(document.element.xpath(".//m:t/text()"))
            self.assertNotIn("\\", math_text)
            self.assertNotIn("_", math_text)
            self.assertIn("*", math_text)
            subscripts = ["".join(node.itertext()) for node in document.element.xpath(".//m:sub")]
            self.assertIn("1:3,4", subscripts)
            self.assertEqual(subscripts.count("4,4"), 2)
            self.assertEqual(subscripts.count("lm"), 2)
            scripted_set = document.element.xpath(".//m:sSubSup")[0]
            self.assertEqual("".join(scripted_set.xpath("./*[local-name()='sub']//*[local-name()='t']/text()")), "i=1")
            self.assertEqual("".join(scripted_set.xpath("./*[local-name()='sup']//*[local-name()='t']/text()")), "N")
            self.assertTrue(document.element.xpath(".//m:acc"))
            self.assertFalse(document.element.xpath(".//m:limUpp"))
            text = "".join(document.element.xpath(".//w:t/text()"))
            self.assertNotIn("\\", text)
            for number in range(1, 5):
                self.assertIn(f"({number})", text)
            self.assertEqual(document.paragraphs[0].style.name, "Heading 2")

    def test_code_and_intentionally_escaped_markdown_stay_literal(self):
        for source in (
            r"\### literal heading" + "\n\n" + r"\*\*literal bold\*\*",
            "```latex\n" + GAUSSIAN_SAMPLE + "\n```",
            r"`\\\\(Y\\\_{lm}\\\\)`",
        ):
            with self.subTest(source=source[:30]):
                prepared, formulas = self.extract(source)
                self.assertEqual(prepared, source.rstrip("\n"))
                self.assertEqual(formulas, [])

    def test_mismatched_or_nonstandard_delimiters_are_not_partially_consumed(self):
        for source in (r"\\(x\)", r"\\\\[x\\]", r"\\\(x\\\)", r"\\\\[x"):
            prepared, formulas = self.extract(source)
            self.assertEqual(prepared, source)
            self.assertEqual(formulas, [])

    def test_multilayer_blockquote(self):
        for width in (1, 2, 4, 8):
            slashes = "\\" * width
            prepared, formulas = self.extract(f"> {slashes}[\n> Y_{{lm}}\n> {slashes}]")
            self.assertEqual(formulas, [{"latex": "Y_{lm}", "display": True}])
            self.assertIn("> MATHBLOCKTOK0END", prepared)


if __name__ == "__main__":
    unittest.main()
