import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from converter import AppConfig, MarkdownToDocxConverter


COMBINATION_FORMULA = r"C(n,2)=\binom n2=\frac{n!}{2!(n-2)!}=\frac{n(n-1)}{2}"


class BinomialTests(unittest.TestCase):
    def setUp(self):
        self.converter = MarkdownToDocxConverter()

    def assert_growing_binomials(self, omml, count=1):
        delimiters = omml.xpath(".//*[local-name()='d']")
        self.assertEqual(len(delimiters), count)
        for delimiter in delimiters:
            properties = delimiter.find(qn("m:dPr"))
            self.assertIsNotNone(properties)
            for name, value in (("begChr", "("), ("endChr", ")"), ("grow", "1")):
                prop = properties.find(qn("m:" + name))
                self.assertIsNotNone(prop, name)
                self.assertEqual(prop.get(qn("m:val")), value)
            fraction = delimiter.find(f"{qn('m:e')}/{qn('m:f')}")
            self.assertIsNotNone(fraction)
            self.assertEqual(fraction.find(f"{qn('m:fPr')}/{qn('m:type')}").get(qn("m:val")), "noBar")

    def test_short_braced_and_style_variants(self):
        for body in (r"\binom n2", r"\binom{n}{2}", r"\dbinom{n+1}{k}", r"\tbinom{n}{k}"):
            for display in (False, True):
                with self.subTest(body=body, display=display):
                    self.assert_growing_binomials(self.converter.latex_to_omml(body, display))

    def test_full_combination_formula_preserves_other_fractions_and_parentheses(self):
        omml = self.converter.latex_to_omml(COMBINATION_FORMULA, True)
        self.assert_growing_binomials(omml)
        fraction_types = [node.get(qn("m:val")) for node in omml.xpath(".//*[local-name()='fPr']/*[local-name()='type']")]
        self.assertEqual(fraction_types, ["noBar", "bar", "bar"])
        ordinary = "".join(omml.xpath("./*[local-name()='r']/*[local-name()='t']/text()"))
        self.assertEqual(ordinary, "C(n,2)===")
        denominators = ["".join(node.itertext()) for node in omml.xpath(".//*[local-name()='den']")]
        self.assertIn("2!(n−2)!", denominators)

    def test_nested_binomials_inside_boxes_and_fractions(self):
        for body, count in (
            (r"\boxed{\binom{n}{k}}", 1),
            (r"\frac{\binom{n}{k}}{2}", 1),
            (r"\binom{\binom{n}{k}}{2}", 2),
        ):
            with self.subTest(body=body):
                self.assert_growing_binomials(self.converter.latex_to_omml(body, True), count)

    def test_plain_fractions_and_unfenced_stacks_are_unchanged(self):
        for body in (r"(\frac{n}{2})", r"\frac{n}{2}", r"{n\atop k}", "C(n,2)"):
            with self.subTest(body=body):
                omml = self.converter.latex_to_omml(body, True)
                self.assertEqual(omml.xpath(".//*[local-name()='d']"), [])

    def test_native_and_escaped_markdown_delimiters(self):
        for width in (1, 2, 4, 8):
            for opening, closing, display in (("(", ")", False), ("[", "]", True)):
                with self.subTest(width=width, opening=opening):
                    slashes = "\\" * width
                    self.converter._prepare_markdown(slashes + opening + COMBINATION_FORMULA + slashes + closing)
                    payload = next(iter(self.converter._math_tokens.values()))
                    self.assertEqual(payload["latex"], COMBINATION_FORMULA)
                    self.assertEqual(payload["display"], display)
                    self.assert_growing_binomials(self.converter.latex_to_omml(payload["latex"], display))

    def test_user_sample_exports_editable_word_math(self):
        source = "又因为组合数：\n\n" + r"\\[ " + COMBINATION_FORMULA + r" \\]" + "\n\n最终得到："
        with tempfile.TemporaryDirectory() as output_dir:
            output = self.converter.convert(source, AppConfig(output_dir=output_dir, asset_root=str(Path.cwd())))
            document = Document(output)
            self.assertEqual(self.converter.last_warnings, [])
            self.assertEqual(len(document.element.xpath(".//m:oMath")), 1)
            self.assertEqual(len(document.element.xpath(".//m:oMathPara")), 1)
            self.assert_growing_binomials(document.element)
            text = "".join(document.element.xpath(".//w:t/text()"))
            self.assertNotIn("\\", text)
            self.assertNotIn("$$", text)
            self.assertIn("最终得到", text)


if __name__ == "__main__":
    unittest.main()
