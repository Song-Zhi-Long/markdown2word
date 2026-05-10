from __future__ import annotations

import re
from copy import deepcopy
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import markdown
from docx import Document
from docx.document import Document as _DocumentType
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from latex2mathml.converter import convert as latex_to_mathml
from lxml import etree, html as lxml_html
from PIL import Image

TOKEN_PATTERN = re.compile(r"(MATH(?:BLOCK|INLINE)TOK\d+END)")
INLINE_MATH_TOKEN = "MATHINLINE"
BLOCK_MATH_TOKEN = "MATHBLOCK"
MATHML_NS = "http://www.w3.org/1998/Math/MathML"
MATHML_TAG = f"{{{MATHML_NS}}}"
SCRIPT_MATH_CHARS = {
    "𝒜": "A",
    "ℬ": "B",
    "𝒞": "C",
    "𝒟": "D",
    "ℰ": "E",
    "ℱ": "F",
    "𝒢": "G",
    "ℋ": "H",
    "ℐ": "I",
    "𝒥": "J",
    "𝒦": "K",
    "ℒ": "L",
    "ℳ": "M",
    "𝒩": "N",
    "𝒪": "O",
    "𝒫": "P",
    "𝒬": "Q",
    "ℛ": "R",
    "𝒮": "S",
    "𝒯": "T",
    "𝒰": "U",
    "𝒱": "V",
    "𝒲": "W",
    "𝒳": "X",
    "𝒴": "Y",
    "𝒵": "Z",
}

LATIN_FONT = "Times New Roman"
CJK_FONT = "宋体"
CODE_FONT = "Consolas"

H1_SIZE_PT = 16.0
H2_SIZE_PT = 14.0
H3_SIZE_PT = 12.0
BODY_SIZE_PT = 12.0
CAPTION_SIZE_PT = 10.5
CODE_SIZE_PT = 10.5
BODY_FIRST_LINE_INDENT_PT = 24.0
LIST_INDENT_PT = 18.0
LIST_HANGING_INDENT_PT = 18.0


@dataclass
class AppConfig:
    output_dir: str
    asset_root: str
    title_chars: int = 12
    auto_timestamp: bool = True


@dataclass
class TextStyle:
    bold: bool = False
    italic: bool = False
    underline: bool = False
    code: bool = False
    size_pt: float = BODY_SIZE_PT

    def copy(self) -> "TextStyle":
        return TextStyle(
            bold=self.bold,
            italic=self.italic,
            underline=self.underline,
            code=self.code,
            size_pt=self.size_pt,
        )


class MarkdownToDocxConverter:
    def __init__(self, xsl_path: Optional[str] = None) -> None:
        root = Path(__file__).resolve().parent
        self.xsl_path = Path(xsl_path) if xsl_path else root / "mml2omml.xsl"
        if not self.xsl_path.exists():
            raise FileNotFoundError(f"Cannot find XSLT file: {self.xsl_path}")

        self._math_transform = etree.XSLT(etree.parse(str(self.xsl_path)))
        self._math_tokens: Dict[str, Dict[str, object]] = {}
        self.last_warnings: List[str] = []
        self._document: Optional[_DocumentType] = None

    def convert(self, markdown_text: str, config: AppConfig) -> str:
        if not markdown_text or not markdown_text.strip():
            raise ValueError("Markdown text is empty.")

        output_dir = Path(config.output_dir).expanduser().resolve()
        output_dir.mkdir(parents=True, exist_ok=True)

        self.last_warnings = []
        prepared_markdown = self._prepare_markdown(markdown_text)
        html_text = markdown.markdown(
            prepared_markdown,
            extensions=[
                "tables",
                "fenced_code",
                "sane_lists",
                "md_in_html",
            ],
            output_format="html5",
        )

        document = Document()
        self._document = document
        self._configure_document_styles(document)

        root = lxml_html.fragment_fromstring(html_text, create_parent="div")
        for child in root:
            self._render_block(child, document, config)

        output_path = output_dir / self._build_output_filename(markdown_text, config)
        document.save(str(output_path))
        return str(output_path)

    def latex_to_omml(self, latex: str, display: bool) -> etree._Element:
        normalized_latex = self._preprocess_latex_for_math(latex, display)
        mathml_text = latex_to_mathml(normalized_latex)
        mathml_root = etree.fromstring(mathml_text.encode("utf-8"))
        mathml_root.set("display", "block" if display else "inline")
        self._normalize_mathml_tree(mathml_root)

        transformed = self._math_transform(mathml_root)
        if transformed.getroot() is None:
            raise ValueError("Math transform returned empty result.")
        omml_root = deepcopy(transformed.getroot())
        self._normalize_omml_math(omml_root)
        return omml_root

    def _preprocess_latex_for_math(self, latex: str, display: bool) -> str:
        normalized = latex

        # `\pmb` should preserve the original italic math style for variables.
        normalized = re.sub(r"\\pmb(?=\s*\{)", r"\\boldsymbol", normalized)

        if display:
            large_ops = r"sum|prod|coprod|bigcup|bigcap|bigoplus|bigotimes|bigvee|bigwedge|bigsqcup"
            normalized = re.sub(
                rf"\\({large_ops})(?!\s*\\(?:limits|nolimits))(?=\s*(?:_|\^))",
                lambda match: f"\\{match.group(1)}\\limits",
                normalized,
            )

        return normalized

    def _normalize_mathml_tree(self, root: etree._Element) -> None:
        self._normalize_mathml_script_chars(root)
        self._convert_mathml_spaces(root)
        self._convert_mathml_stretchy_fences(root)

    def _normalize_mathml_script_chars(self, root: etree._Element) -> None:
        for node in root.xpath(".//*[local-name()='mi']"):
            text = (node.text or "").strip()
            if text in SCRIPT_MATH_CHARS:
                node.text = SCRIPT_MATH_CHARS[text]
                node.set("mathvariant", "script")

    def _convert_mathml_spaces(self, root: etree._Element) -> None:
        for node in root.xpath(".//*[local-name()='mspace']"):
            width = (node.get("width") or "").strip().lower()
            replacement = etree.Element(f"{MATHML_TAG}mtext")
            replacement.text = self._mathml_space_text(width)

            parent = node.getparent()
            if parent is not None:
                parent.replace(node, replacement)

    def _convert_mathml_stretchy_fences(self, node: etree._Element) -> None:
        for child in list(node):
            self._convert_mathml_stretchy_fences(child)

        if self._mathml_local_name(node) != "mrow":
            return

        children = list(node)
        if len(children) < 2:
            return

        first = children[0]
        last = children[-1]
        if not self._is_mathml_stretchy_fence(first, "prefix"):
            return
        if not self._is_mathml_stretchy_fence(last, "postfix"):
            return

        mfenced = etree.Element(f"{MATHML_TAG}mfenced")
        mfenced.set("open", self._mathml_text(first))
        mfenced.set("close", self._mathml_text(last))
        mfenced.set("separators", "")

        inner_row = etree.Element(f"{MATHML_TAG}mrow")
        for child in children[1:-1]:
            inner_row.append(child)
        mfenced.append(inner_row)

        parent = node.getparent()
        if parent is not None:
            parent.replace(node, mfenced)

    def _is_mathml_stretchy_fence(self, node: etree._Element, expected_form: str) -> bool:
        return (
            self._mathml_local_name(node) == "mo"
            and (node.get("stretchy") or "").lower() == "true"
            and (node.get("fence") or "").lower() == "true"
            and (node.get("form") or "").lower() == expected_form
        )

    def _mathml_local_name(self, node: etree._Element) -> str:
        return etree.QName(node).localname.lower()

    def _mathml_text(self, node: etree._Element) -> str:
        return "".join(node.itertext())

    def _mathml_space_text(self, width: str) -> str:
        if width.endswith("em"):
            try:
                amount = float(width[:-2])
            except ValueError:
                amount = 1.0
        else:
            amount = 1.0

        if amount >= 1.5:
            return "\u2003\u2003"
        if amount >= 0.75:
            return "\u2003"
        if amount >= 0.4:
            return "\u2002"
        return " "

    def _normalize_omml_math(self, omml_root: etree._Element) -> None:
        for math_node in self._iter_omath_nodes(omml_root):
            self._repair_nary_operand(math_node)
            self._repair_matrix_delimiter(math_node)

    def _iter_omath_nodes(self, root: etree._Element):
        local = self._tag_name(root)
        if local == "omath":
            yield root
            return
        if local == "omathpara":
            for node in root.xpath("./*[local-name()='oMath']"):
                yield node

    def _repair_nary_operand(self, omath: etree._Element) -> None:
        children = list(omath)
        for idx, child in enumerate(children):
            if self._tag_name(child) != "nary":
                continue
            e_nodes = child.xpath("./*[local-name()='e']")
            if not e_nodes:
                continue
            e_node = e_nodes[0]
            if len(e_node) != 0:
                continue
            if idx + 1 >= len(children):
                continue

            next_expr = children[idx + 1]
            if self._tag_name(next_expr) in {"dPr", "ctrlPr"}:
                continue
            if self._tag_name(next_expr) == "r" and self._is_operator_run(next_expr):
                continue
            e_node.append(deepcopy(next_expr))
            omath.remove(next_expr)
            children = list(omath)

    def _repair_matrix_delimiter(self, omath: etree._Element) -> None:
        bracket_pairs = {"(": ")", "[": "]", "{": "}", "|": "|"}
        target_tags = {"m", "eqarr"}
        changed = True

        while changed:
            changed = False
            children = list(omath)
            for idx, expr_node in enumerate(children):
                if self._tag_name(expr_node) not in target_tags:
                    continue
                if idx == 0 or idx >= len(children) - 1:
                    continue

                left_node = children[idx - 1]
                right_node = children[idx + 1]
                left = self._peek_open_bracket_from_run(left_node)
                right = self._peek_close_bracket_from_run(right_node)

                if left is None or right is None:
                    continue
                left_bracket, left_remain = left
                right_bracket, right_remain = right
                if bracket_pairs.get(left_bracket) != right_bracket:
                    continue

                delim = self._build_delimiter_omml(left_bracket, right_bracket, expr_node)
                omath.replace(expr_node, delim)

                self._set_run_text(left_node, left_remain)
                self._set_run_text(right_node, right_remain)
                if not left_remain:
                    omath.remove(left_node)
                if not right_remain:
                    omath.remove(right_node)

                changed = True
                break

    def _peek_open_bracket_from_run(self, node: etree._Element) -> Optional[Tuple[str, str]]:
        if self._tag_name(node) != "r":
            return None
        text = self._get_run_text(node)
        if not text:
            return None

        stripped = text.rstrip()
        if not stripped:
            return None
        bracket = stripped[-1]
        if bracket not in {"(", "[", "{", "|"}:
            return None
        remain = stripped[:-1] + text[len(stripped) :]
        return bracket, remain

    def _peek_close_bracket_from_run(self, node: etree._Element) -> Optional[Tuple[str, str]]:
        if self._tag_name(node) != "r":
            return None
        text = self._get_run_text(node)
        if not text:
            return None

        leading_trimmed = text.lstrip()
        if not leading_trimmed:
            return None
        bracket = leading_trimmed[0]
        if bracket not in {")", "]", "}", "|"}:
            return None
        remain = text[: len(text) - len(leading_trimmed)] + leading_trimmed[1:]
        return bracket, remain

    def _get_run_text(self, node: etree._Element) -> str:
        return "".join((t.text or "") for t in node.xpath(".//*[local-name()='t']"))

    def _set_run_text(self, node: etree._Element, text: str) -> None:
        text_nodes = node.xpath("./*[local-name()='t']")
        if text_nodes:
            text_nodes[0].text = text
            for extra in text_nodes[1:]:
                node.remove(extra)
            return

        if text:
            t = OxmlElement("m:t")
            t.text = text
            node.append(t)

    def _is_operator_run(self, node: etree._Element) -> bool:
        text = "".join((t.text or "") for t in node.xpath(".//*[local-name()='t']")).strip()
        return text in {"+", "-", "*", "/", "=", ",", ";", ":"}

    def _build_delimiter_omml(self, left: str, right: str, expr_node: etree._Element) -> etree._Element:
        delim = OxmlElement("m:d")
        d_pr = OxmlElement("m:dPr")

        beg = OxmlElement("m:begChr")
        beg.set(qn("m:val"), left)
        d_pr.append(beg)

        end = OxmlElement("m:endChr")
        end.set(qn("m:val"), right)
        d_pr.append(end)

        grow = OxmlElement("m:grow")
        grow.set(qn("m:val"), "1")
        d_pr.append(grow)

        delim.append(d_pr)
        e = OxmlElement("m:e")
        e.append(deepcopy(expr_node))
        delim.append(e)
        return delim

    def resolve_image_path(self, src: str, asset_root: str) -> Optional[Path]:
        if not src:
            return None

        src = src.strip()
        if src.startswith("http://") or src.startswith("https://") or src.startswith("data:"):
            return None

        image_path = Path(src)
        if image_path.is_absolute() and image_path.exists():
            return image_path

        base_dir = Path(asset_root).expanduser().resolve() if asset_root else Path.cwd()
        candidate = (base_dir / image_path).resolve()
        if candidate.exists():
            return candidate

        local_candidate = (Path.cwd() / image_path).resolve()
        if local_candidate.exists():
            return local_candidate

        return None

    def _configure_document_styles(self, document: _DocumentType) -> None:
        self._configure_style(document, "Normal", LATIN_FONT, CJK_FONT, BODY_SIZE_PT)
        self._configure_style(document, "Heading 1", LATIN_FONT, CJK_FONT, H1_SIZE_PT)
        self._configure_style(document, "Heading 2", LATIN_FONT, CJK_FONT, H2_SIZE_PT)
        self._configure_style(document, "Heading 3", LATIN_FONT, CJK_FONT, H3_SIZE_PT)
        self._configure_style(document, "Caption", LATIN_FONT, CJK_FONT, CAPTION_SIZE_PT)
        self._configure_hyperlink_style(document, "Hyperlink", RGBColor(5, 99, 193))
        self._configure_hyperlink_style(document, "FollowedHyperlink", RGBColor(149, 79, 114))

    def _configure_style(self, document: _DocumentType, style_name: str, latin_font: str, east_asia_font: str, size_pt: float) -> None:
        try:
            style = document.styles[style_name]
        except KeyError:
            return

        style.font.name = latin_font
        style.font.size = Pt(size_pt)

        r_pr = style._element.get_or_add_rPr()
        r_fonts = r_pr.find(qn("w:rFonts"))
        if r_fonts is None:
            r_fonts = OxmlElement("w:rFonts")
            r_pr.append(r_fonts)
        r_fonts.set(qn("w:ascii"), latin_font)
        r_fonts.set(qn("w:hAnsi"), latin_font)
        r_fonts.set(qn("w:eastAsia"), east_asia_font)
        r_fonts.set(qn("w:cs"), latin_font)

    def _configure_hyperlink_style(self, document: _DocumentType, style_name: str, rgb_color: RGBColor) -> None:
        try:
            style = document.styles[style_name]
        except KeyError:
            style = document.styles.add_style(style_name, WD_STYLE_TYPE.CHARACTER)

        style.font.name = LATIN_FONT
        style.font.underline = True
        style.font.color.rgb = rgb_color

        r_pr = style._element.get_or_add_rPr()
        r_fonts = r_pr.find(qn("w:rFonts"))
        if r_fonts is None:
            r_fonts = OxmlElement("w:rFonts")
            r_pr.append(r_fonts)
        r_fonts.set(qn("w:ascii"), LATIN_FONT)
        r_fonts.set(qn("w:hAnsi"), LATIN_FONT)
        r_fonts.set(qn("w:eastAsia"), CJK_FONT)
        r_fonts.set(qn("w:cs"), LATIN_FONT)

    def _prepare_markdown(self, text: str) -> str:
        masked_text, code_tokens = self._mask_code_regions(text)
        list_ready_text = self._split_reset_ordered_lists(masked_text)
        math_ready_text = self._extract_math_tokens(list_ready_text)
        return self._restore_tokens(math_ready_text, code_tokens)

    def _mask_code_regions(self, text: str) -> Tuple[str, Dict[str, str]]:
        tokens: Dict[str, str] = {}

        def fenced_replacer(match: re.Match[str]) -> str:
            token = f"CODEBLOCKTOK{len(tokens)}END"
            tokens[token] = match.group(0)
            return token

        text = re.sub(r"```[\s\S]*?```|~~~[\s\S]*?~~~", fenced_replacer, text)

        def inline_replacer(match: re.Match[str]) -> str:
            token = f"CODEINLINETOK{len(tokens)}END"
            tokens[token] = match.group(0)
            return token

        text = re.sub(r"`[^`\n]+`", inline_replacer, text)
        return text, tokens

    def _restore_tokens(self, text: str, tokens: Dict[str, str]) -> str:
        restored = text
        for token, value in tokens.items():
            restored = restored.replace(token, value)
        return restored

    def _extract_math_tokens(self, text: str) -> str:
        self._math_tokens = {}

        def create_token(latex: str, display: bool) -> str:
            token_type = BLOCK_MATH_TOKEN if display else INLINE_MATH_TOKEN
            token = f"{token_type}TOK{len(self._math_tokens)}END"
            self._math_tokens[token] = {"latex": latex.strip(), "display": display}
            if display:
                return f"\n\n{token}\n\n"
            return token

        text = re.sub(r"\$\$(.+?)\$\$", lambda m: create_token(m.group(1), True), text, flags=re.DOTALL)
        text = re.sub(r"\\\[(.+?)\\\]", lambda m: create_token(m.group(1), True), text, flags=re.DOTALL)
        text = re.sub(r"\\\((.+?)\\\)", lambda m: create_token(m.group(1), False), text)
        text = re.sub(r"(?<!\\)\$(?!\$)([^\n]+?)(?<!\\)\$", lambda m: create_token(m.group(1), False), text)
        return text

    def _split_reset_ordered_lists(self, text: str) -> str:
        ordered_item_pattern = re.compile(r"^(\s*)(\d+)([.)])\s+")
        separator = '<div class="md2docx-list-split"></div>'
        lines = text.splitlines()
        result: List[str] = []
        index = 0

        while index < len(lines):
            current = lines[index]
            if current.strip():
                result.append(current)
                index += 1
                continue

            blank_start = index
            while index < len(lines) and not lines[index].strip():
                result.append(lines[index])
                index += 1

            prev_index = blank_start - 1
            while prev_index >= 0 and not lines[prev_index].strip():
                prev_index -= 1

            if prev_index < 0 or index >= len(lines):
                continue

            prev_match = ordered_item_pattern.match(lines[prev_index])
            next_match = ordered_item_pattern.match(lines[index])
            if not prev_match or not next_match:
                continue

            same_indent = len(prev_match.group(1)) == len(next_match.group(1))
            restarts_at_one = next_match.group(2) == "1"
            if same_indent and restarts_at_one:
                result.append(separator)

        return "\n".join(result)

    def _render_block(self, node: etree._Element, container, config: AppConfig) -> None:
        tag = self._tag_name(node)
        if not tag:
            return

        if tag in {"h1", "h2", "h3", "h4", "h5", "h6"}:
            level = int(tag[1])
            para = container.add_paragraph(style=f"Heading {level}")
            self._apply_alignment(para, node)
            size = {1: H1_SIZE_PT, 2: H2_SIZE_PT, 3: H3_SIZE_PT}.get(level, BODY_SIZE_PT)
            self._render_inline_node(node, para, container, config, TextStyle(size_pt=size))
            return

        if tag in {"p", "div", "center"}:
            if not node.text_content().strip() and not node.xpath(".//*[local-name()='img']"):
                return

            single_block = self._single_block_math_token(node)
            if single_block:
                para = container.add_paragraph()
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para.paragraph_format.first_line_indent = Pt(0)
                self._append_math(paragraph=para, token=single_block, inline=False)
                return

            para = container.add_paragraph()
            if tag == "center":
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para.paragraph_format.first_line_indent = Pt(0)
            else:
                self._apply_alignment(para, node)

            style = TextStyle(size_pt=BODY_SIZE_PT)
            if self._is_caption_text(node.text_content()):
                try:
                    para.style = "Caption"
                except KeyError:
                    pass
                style.size_pt = CAPTION_SIZE_PT
                para.paragraph_format.first_line_indent = Pt(0)
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif tag != "center":
                para.paragraph_format.first_line_indent = Pt(BODY_FIRST_LINE_INDENT_PT)

            self._render_inline_node(node, para, container, config, style)
            return

        if tag in {"ul", "ol"}:
            self._render_list(node, container, config, ordered=(tag == "ol"), level=1)
            return

        if tag == "pre":
            self._render_code_block(node, container)
            return

        if tag == "blockquote":
            for child in node:
                self._render_block(child, container, config)
            return

        if tag == "table":
            self._render_table(node, container, config)
            return

        if tag == "img":
            para = container.add_paragraph()
            self._insert_image(para, node, config)
            return

        if tag == "hr":
            container.add_paragraph("-" * 40)
            return

        for child in node:
            self._render_block(child, container, config)

    def _render_code_block(self, node: etree._Element, container) -> None:
        code_text = (node.text_content() or "").replace("\r\n", "\n").replace("\r", "\n")
        lines = code_text.split("\n")
        if lines and lines[-1] == "":
            lines = lines[:-1]
        if not lines:
            lines = [""]

        for idx, line in enumerate(lines):
            para = container.add_paragraph()
            para.paragraph_format.first_line_indent = Pt(0)
            para.paragraph_format.space_before = Pt(6 if idx == 0 else 0)
            para.paragraph_format.space_after = Pt(6 if idx == len(lines) - 1 else 0)
            self._set_paragraph_shading(para, "EDEDED")

            run = para.add_run(line if line else " ")
            self._apply_run_style(run, TextStyle(code=True, size_pt=CODE_SIZE_PT))

    def _render_list(self, node: etree._Element, container, config: AppConfig, ordered: bool, level: int) -> None:
        next_number = self._extract_list_start(node) if ordered else 1

        for li in node.xpath("./li"):
            para = container.add_paragraph()
            self._configure_list_paragraph(para, level)
            prefix = f"{next_number}. " if ordered else self._bullet_prefix(level)
            prefix_run = para.add_run(prefix)
            self._apply_run_style(prefix_run, TextStyle(size_pt=BODY_SIZE_PT))
            self._render_list_item_inline(li, para, container, config)
            if ordered:
                next_number += 1

            for nested in li:
                nested_tag = self._tag_name(nested)
                if nested_tag in {"ul", "ol"}:
                    self._render_list(nested, container, config, ordered=(nested_tag == "ol"), level=min(level + 1, 3))

    def _render_list_item_inline(self, li: etree._Element, paragraph, container, config: AppConfig) -> None:
        base_style = TextStyle(size_pt=BODY_SIZE_PT)
        if li.text and li.text.strip():
            paragraph = self._append_text(paragraph, container, li.text, base_style)

        for child in li:
            tag = self._tag_name(child)
            if tag in {"ul", "ol"}:
                continue
            paragraph = self._render_inline_node(child, paragraph, container, config, base_style)
            if child.tail and child.tail.strip():
                paragraph = self._append_text(paragraph, container, child.tail, base_style)

    def _render_table(self, table_node: etree._Element, container, config: AppConfig) -> None:
        rows = table_node.xpath("./thead/tr|./tbody/tr|./tfoot/tr|./tr")
        if not rows:
            return

        occupancy: Dict[Tuple[int, int], bool] = {}
        anchors: Dict[Tuple[int, int], Tuple[etree._Element, int, int]] = {}
        max_col = 0
        max_row = 0

        for r_idx, row in enumerate(rows):
            c_idx = 0
            cells = row.xpath("./th|./td")
            for cell in cells:
                while occupancy.get((r_idx, c_idx), False):
                    c_idx += 1

                rowspan = max(1, int(cell.get("rowspan", "1") or "1"))
                colspan = max(1, int(cell.get("colspan", "1") or "1"))
                anchors[(r_idx, c_idx)] = (cell, rowspan, colspan)

                for rr in range(r_idx, r_idx + rowspan):
                    for cc in range(c_idx, c_idx + colspan):
                        occupancy[(rr, cc)] = True
                        max_row = max(max_row, rr)
                        max_col = max(max_col, cc)

                c_idx += colspan

        table = container.add_table(rows=max_row + 1, cols=max_col + 1)
        table.style = "Table Grid"

        for (r_idx, c_idx), (_, rowspan, colspan) in anchors.items():
            if rowspan > 1 or colspan > 1:
                table.cell(r_idx, c_idx).merge(table.cell(r_idx + rowspan - 1, c_idx + colspan - 1))

        for (r_idx, c_idx), (cell_node, _, _) in anchors.items():
            cell = table.cell(r_idx, c_idx)
            self._clear_cell(cell)
            para = cell.add_paragraph()
            style = TextStyle(bold=(self._tag_name(cell_node) == "th"), size_pt=BODY_SIZE_PT)
            self._render_inline_node(cell_node, para, cell, config, style)

    def _render_inline_node(self, node: etree._Element, paragraph, container, config: AppConfig, style: TextStyle):
        if node.text and (node.text.strip() or len(node) == 0):
            paragraph = self._append_text(paragraph, container, node.text, style)

        for child in node:
            child_tag = self._tag_name(child)
            child_style = style.copy()

            if child_tag in {"strong", "b"}:
                child_style.bold = True
                paragraph = self._render_inline_node(child, paragraph, container, config, child_style)
            elif child_tag in {"em", "i"}:
                child_style.italic = True
                paragraph = self._render_inline_node(child, paragraph, container, config, child_style)
            elif child_tag == "u":
                child_style.underline = True
                paragraph = self._render_inline_node(child, paragraph, container, config, child_style)
            elif child_tag == "code":
                child_style.code = True
                child_style.size_pt = CODE_SIZE_PT
                paragraph = self._render_inline_node(child, paragraph, container, config, child_style)
            elif child_tag == "a":
                paragraph = self._append_hyperlink_from_node(paragraph, container, child, child_style, config)
            elif child_tag == "br":
                paragraph = self._start_new_paragraph_after(paragraph, container)
            elif child_tag == "img":
                self._insert_image(paragraph, child, config)
            elif child_tag in {"sub", "sup"}:
                paragraph = self._append_script_text(
                    paragraph,
                    container,
                    child.text_content(),
                    child_style,
                    is_subscript=(child_tag == "sub"),
                    is_superscript=(child_tag == "sup"),
                )
            else:
                paragraph = self._render_inline_node(child, paragraph, container, config, child_style)

            tail_text = child.tail or ""
            if child_tag == "br" and tail_text.startswith("\n"):
                tail_text = tail_text[1:]
            if tail_text and (tail_text.strip() or self._tag_name(node) in {"code", "pre"}):
                paragraph = self._append_text(paragraph, container, tail_text, style)

        return paragraph

    def _append_hyperlink_from_node(self, paragraph, container, anchor_node: etree._Element, style: TextStyle, config: AppConfig):
        href = (anchor_node.get("href") or "").strip()
        text = self._normalize_inline_text(anchor_node.text_content()).strip() or href

        if not href:
            return self._append_text(paragraph, container, text, style)
        resolved_href = self._resolve_hyperlink_target(href, config.asset_root)

        for part in self._split_text_with_paragraphs(text):
            if part == "\n":
                paragraph = self._start_new_paragraph_after(paragraph, container)
                continue
            if not part:
                continue
            if part in self._math_tokens:
                self._append_math(paragraph=paragraph, token=part, inline=not self._math_tokens[part]["display"])
                continue
            self._append_hyperlink(paragraph, resolved_href, part, style)
        return paragraph

    def _resolve_hyperlink_target(self, href: str, asset_root: str) -> str:
        href = href.strip()
        if not href:
            return href
        if href.startswith("#"):
            return href

        path_like = Path(href)
        if path_like.is_absolute():
            try:
                return path_like.resolve().as_uri()
            except ValueError:
                return str(path_like)

        if re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*:", href):
            return href

        base_dir = Path(asset_root).expanduser().resolve() if asset_root else Path.cwd()
        resolved = (base_dir / path_like).resolve()
        try:
            return resolved.as_uri()
        except ValueError:
            return str(resolved)

    def _append_hyperlink(self, paragraph, url: str, text: str, style: TextStyle) -> None:
        part = paragraph.part
        r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)

        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("r:id"), r_id)
        hyperlink.set(qn("w:history"), "1")

        run = OxmlElement("w:r")
        r_pr = OxmlElement("w:rPr")

        r_style = OxmlElement("w:rStyle")
        r_style.set(qn("w:val"), "Hyperlink")
        r_pr.append(r_style)

        r_fonts = OxmlElement("w:rFonts")
        if style.code:
            r_fonts.set(qn("w:ascii"), CODE_FONT)
            r_fonts.set(qn("w:hAnsi"), CODE_FONT)
            r_fonts.set(qn("w:eastAsia"), CODE_FONT)
            r_fonts.set(qn("w:cs"), CODE_FONT)
        else:
            r_fonts.set(qn("w:ascii"), LATIN_FONT)
            r_fonts.set(qn("w:hAnsi"), LATIN_FONT)
            r_fonts.set(qn("w:eastAsia"), CJK_FONT)
            r_fonts.set(qn("w:cs"), LATIN_FONT)
        r_pr.append(r_fonts)

        size_half = str(int(round(style.size_pt * 2)))
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), size_half)
        r_pr.append(sz)
        sz_cs = OxmlElement("w:szCs")
        sz_cs.set(qn("w:val"), size_half)
        r_pr.append(sz_cs)

        if style.bold:
            r_pr.append(OxmlElement("w:b"))
        if style.italic:
            r_pr.append(OxmlElement("w:i"))
        if style.underline:
            u = OxmlElement("w:u")
            u.set(qn("w:val"), "single")
            r_pr.append(u)
        if style.code:
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"), "EDEDED")
            r_pr.append(shd)

        run.append(r_pr)
        text_element = OxmlElement("w:t")
        text_element.text = text
        run.append(text_element)
        hyperlink.append(run)
        paragraph._p.append(hyperlink)

    def _append_text(self, paragraph, container, text: str, style: TextStyle):
        if not text:
            return paragraph

        normalized_text = self._normalize_inline_text(text)
        if not normalized_text:
            return paragraph

        for part in self._split_text_with_paragraphs(normalized_text):
            if part == "\n":
                paragraph = self._start_new_paragraph_after(paragraph, container)
                continue
            if not part:
                continue
            if part in self._math_tokens:
                self._append_math(paragraph=paragraph, token=part, inline=not self._math_tokens[part]["display"])
                continue

            run = paragraph.add_run(part)
            self._apply_run_style(run, style)
        return paragraph

    def _append_script_text(self, paragraph, container, text: str, style: TextStyle, is_subscript: bool, is_superscript: bool):
        normalized_text = self._normalize_inline_text(text)
        if not normalized_text:
            return paragraph

        for part in self._split_text_with_paragraphs(normalized_text):
            if part == "\n":
                paragraph = self._start_new_paragraph_after(paragraph, container)
                continue
            if not part:
                continue
            run = paragraph.add_run(part)
            self._apply_run_style(run, style)
            run.font.subscript = is_subscript
            run.font.superscript = is_superscript
        return paragraph

    def _normalize_inline_text(self, text: str) -> str:
        return text.replace("\r\n", "\n").replace("\r", "\n")

    def _split_text_with_paragraphs(self, text: str) -> List[str]:
        parts: List[str] = []
        for line_index, line in enumerate(text.split("\n")):
            if line_index > 0:
                parts.append("\n")
            parts.extend(TOKEN_PATTERN.split(line))
        return parts

    def _start_new_paragraph_after(self, paragraph, container):
        new_paragraph = container.add_paragraph()
        self._copy_paragraph_format(paragraph, new_paragraph)
        return new_paragraph

    def _copy_paragraph_format(self, source, target) -> None:
        source_ppr = source._p.pPr
        if source_ppr is None:
            return

        target_ppr = target._p.pPr
        if target_ppr is not None:
            target._p.remove(target_ppr)
        target._p.insert(0, deepcopy(source_ppr))

    def _apply_run_style(self, run, style: TextStyle) -> None:
        if style.bold:
            run.bold = True
        if style.italic:
            run.italic = True
        if style.underline:
            run.underline = True

        if style.code:
            run.font.name = CODE_FONT
            run.font.size = Pt(style.size_pt or CODE_SIZE_PT)
            self._set_run_fonts(run, CODE_FONT, CODE_FONT)
            self._set_run_shading(run, "EDEDED")
        else:
            run.font.name = LATIN_FONT
            run.font.size = Pt(style.size_pt or BODY_SIZE_PT)
            self._set_run_fonts(run, LATIN_FONT, CJK_FONT)

    def _set_run_fonts(self, run, latin_font: str, east_asia_font: str) -> None:
        r_pr = run._element.get_or_add_rPr()
        r_fonts = r_pr.find(qn("w:rFonts"))
        if r_fonts is None:
            r_fonts = OxmlElement("w:rFonts")
            r_pr.append(r_fonts)
        r_fonts.set(qn("w:ascii"), latin_font)
        r_fonts.set(qn("w:hAnsi"), latin_font)
        r_fonts.set(qn("w:eastAsia"), east_asia_font)
        r_fonts.set(qn("w:cs"), latin_font)

    def _set_run_shading(self, run, fill: str) -> None:
        r_pr = run._element.get_or_add_rPr()
        existing = r_pr.find(qn("w:shd"))
        if existing is not None:
            r_pr.remove(existing)

        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill)
        r_pr.append(shd)

    def _set_paragraph_shading(self, paragraph, fill: str) -> None:
        p_pr = paragraph._p.get_or_add_pPr()
        existing = p_pr.find(qn("w:shd"))
        if existing is not None:
            p_pr.remove(existing)

        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill)
        p_pr.append(shd)

    def _append_math(self, paragraph, token: str, inline: bool) -> None:
        payload = self._math_tokens.get(token)
        if payload is None:
            return

        latex = str(payload.get("latex", "")).strip()
        display = bool(payload.get("display", False))

        if inline:
            detached_scripts = self._extract_detached_script_tokens(latex)
            if detached_scripts:
                for marker, text in detached_scripts:
                    run = paragraph.add_run(text)
                    self._apply_run_style(run, TextStyle(size_pt=BODY_SIZE_PT))
                    run.font.superscript = marker == "^"
                    run.font.subscript = marker == "_"
                return

        try:
            omml = self.latex_to_omml(latex, display=display)
            local_tag = self._tag_name(omml)

            if inline:
                run = paragraph.add_run()
                self._set_run_fonts(run, LATIN_FONT, CJK_FONT)
                if local_tag == "omathpara":
                    math_nodes = omml.xpath(".//*[local-name()='oMath']")
                    if math_nodes:
                        run._r.append(deepcopy(math_nodes[0]))
                    else:
                        run.text = f"${latex}$"
                elif local_tag == "omath":
                        run._r.append(omml)
                else:
                    run.text = f"${latex}$"
            else:
                self._apply_display_math_paragraph_format(paragraph)
                if local_tag == "omathpara":
                    self._set_omathpara_center(omml)
                    paragraph._p.append(omml)
                elif local_tag == "omath":
                    wrapper = OxmlElement("m:oMathPara")
                    wrapper.append(omml)
                    self._set_omathpara_center(wrapper)
                    paragraph._p.append(wrapper)
                else:
                    run = paragraph.add_run()
                    self._set_run_fonts(run, LATIN_FONT, CJK_FONT)
                    run.text = f"$$ {latex} $$"
        except Exception as exc:  # noqa: BLE001
            self.last_warnings.append(f"Formula fallback: {latex[:60]} ({exc})")
            if inline:
                run = paragraph.add_run(f"${latex}$")
                self._apply_run_style(run, TextStyle(size_pt=BODY_SIZE_PT))
            else:
                self._apply_display_math_paragraph_format(paragraph)
                run = paragraph.add_run(f"$$ {latex} $$")
                self._apply_run_style(run, TextStyle(size_pt=BODY_SIZE_PT))

    def _extract_detached_script_tokens(self, latex: str) -> Optional[List[Tuple[str, str]]]:
        stripped = latex.strip()
        if not stripped:
            return None

        pattern = re.compile(r"([_^])(?:\{([^{}]+)\}|([^\s_^{}]+))")
        tokens: List[Tuple[str, str]] = []
        position = 0

        for match in pattern.finditer(stripped):
            if match.start() != position:
                return None
            value = match.group(2) or match.group(3) or ""
            if not value:
                return None
            tokens.append((match.group(1), value))
            position = match.end()

        if position != len(stripped):
            return None
        return tokens or None

    def _apply_display_math_paragraph_format(self, paragraph) -> None:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.first_line_indent = Pt(0)

    def _set_omathpara_center(self, omathpara_node: etree._Element) -> None:
        if self._tag_name(omathpara_node) != "omathpara":
            return

        para_pr = None
        for child in list(omathpara_node):
            if self._tag_name(child) == "omathparapr":
                para_pr = child
                break

        if para_pr is None:
            para_pr = OxmlElement("m:oMathParaPr")
            omathpara_node.insert(0, para_pr)

        jc = None
        for child in list(para_pr):
            if self._tag_name(child) == "jc":
                jc = child
                break
        if jc is None:
            jc = OxmlElement("m:jc")
            para_pr.append(jc)
        jc.set(qn("m:val"), "center")

    def _insert_image(self, paragraph, image_node: etree._Element, config: AppConfig) -> None:
        src = image_node.get("src", "").strip()
        resolved = self.resolve_image_path(src, config.asset_root)

        if resolved is None or not resolved.exists():
            if src:
                self.last_warnings.append(f"Image not found: {src}")
                run = paragraph.add_run(f"[Image not found: {src}]")
                self._apply_run_style(run, TextStyle(size_pt=BODY_SIZE_PT))
            return

        width_inches = self._parse_image_width(image_node.get("width"), resolved)
        run = paragraph.add_run()
        if width_inches:
            run.add_picture(str(resolved), width=Inches(width_inches))
        else:
            run.add_picture(str(resolved))

    def _parse_image_width(self, width_attr: Optional[str], image_path: Path) -> Optional[float]:
        max_width = self._page_content_width_inches()
        if not width_attr:
            return self._natural_or_max_width(image_path, max_width)

        raw = width_attr.strip().lower()
        try:
            if raw.endswith("%"):
                return max(0.2, min(max_width * max(min(float(raw.rstrip("%")), 100.0), 1.0) / 100.0, max_width))
            if raw.endswith("px"):
                return max(0.2, min(float(raw[:-2]) / 96.0, max_width))
            if raw.endswith("cm"):
                return max(0.2, min(float(raw[:-2]) / 2.54, max_width))
            if raw.endswith("mm"):
                return max(0.2, min(float(raw[:-2]) / 25.4, max_width))
            if raw.endswith("in"):
                return max(0.2, min(float(raw[:-2]), max_width))
            if raw.endswith("pt"):
                return max(0.2, min(float(raw[:-2]) / 72.0, max_width))
            return max(0.2, min(float(raw) / 96.0, max_width))
        except ValueError:
            return self._natural_or_max_width(image_path, max_width)

    def _natural_or_max_width(self, image_path: Path, max_width: float) -> float:
        try:
            with Image.open(image_path) as img:
                dpi = img.info.get("dpi", (96, 96))[0] or 96
                natural_width = img.width / float(dpi)
                return max(0.2, min(natural_width, max_width))
        except Exception:
            return max_width

    def _page_content_width_inches(self) -> float:
        if self._document is None:
            return 6.0
        section = self._document.sections[0]
        return float(section.page_width - section.left_margin - section.right_margin) / 914400.0

    def _extract_list_start(self, node: etree._Element) -> int:
        raw_start = (node.get("start") or "1").strip()
        try:
            return max(1, int(raw_start))
        except ValueError:
            return 1

    def _configure_list_paragraph(self, paragraph, level: int) -> None:
        left_indent = Pt(LIST_INDENT_PT * level)
        paragraph.paragraph_format.left_indent = left_indent
        paragraph.paragraph_format.first_line_indent = Pt(-LIST_HANGING_INDENT_PT)

    def _bullet_prefix(self, level: int) -> str:
        bullets = ["• ", "◦ ", "▪ "]
        return bullets[min(max(level - 1, 0), len(bullets) - 1)]

    def _add_paragraph_with_style(self, container, style_names: Iterable[str]):
        for style_name in style_names:
            try:
                return container.add_paragraph(style=style_name)
            except KeyError:
                continue
        return container.add_paragraph()

    def _apply_alignment(self, paragraph, node: etree._Element) -> None:
        align = self._extract_alignment(node)
        if align is not None:
            paragraph.alignment = align

    def _extract_alignment(self, node: etree._Element) -> Optional[WD_ALIGN_PARAGRAPH]:
        align_attr = (node.get("align") or "").strip().lower()
        style_attr = (node.get("style") or "").lower()

        if "text-align" in style_attr:
            match = re.search(r"text-align\s*:\s*(left|center|right|justify)", style_attr)
            if match:
                align_attr = match.group(1)

        mapping = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }
        return mapping.get(align_attr)

    def _build_output_filename(self, markdown_text: str, config: AppConfig) -> str:
        cleaned = re.sub(r"```[\s\S]*?```|~~~[\s\S]*?~~~", " ", markdown_text)
        cleaned = re.sub(r"\$\$.+?\$\$", " ", cleaned, flags=re.DOTALL)
        cleaned = re.sub(r"\\\[.+?\\\]", " ", cleaned, flags=re.DOTALL)
        cleaned = re.sub(r"\\\(.+?\\\)", " ", cleaned)
        cleaned = re.sub(r"`[^`\n]+`", " ", cleaned)
        cleaned = re.sub(r"<[^>]+>", " ", cleaned)
        cleaned = re.sub(r"[\r\n\t]+", " ", cleaned)
        cleaned = re.sub(r"\s+", " ", cleaned).strip()

        normalized = re.sub(r"[^\w\u4e00-\u9fff]+", "", cleaned, flags=re.UNICODE)
        base = normalized[: max(1, config.title_chars)] or "document"

        if config.auto_timestamp:
            return f"{base}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        return f"{base}.docx"

    def _single_block_math_token(self, node: etree._Element) -> Optional[str]:
        if len(node) != 0:
            return None

        text = (node.text or "").strip()
        if text in self._math_tokens and bool(self._math_tokens[text]["display"]):
            return text
        return None

    def _is_caption_text(self, text: str) -> bool:
        normalized = re.sub(r"\s+", " ", text or "").strip().lower()
        return bool(re.match(r"^(表|图)\s*\d+\s*[:：]|^(table|figure)\s*\d+\s*:", normalized))

    def _tag_name(self, node) -> str:
        if node is None:
            return ""
        if isinstance(node, str):
            return node
        if not hasattr(node, "tag"):
            return ""
        if isinstance(node.tag, str):
            return node.tag.split("}")[-1].lower()
        return ""

    def _clear_cell(self, cell) -> None:
        tc = cell._tc
        for paragraph in cell.paragraphs:
            tc.remove(paragraph._p)


def build_converter() -> MarkdownToDocxConverter:
    return MarkdownToDocxConverter()
