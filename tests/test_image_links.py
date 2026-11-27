import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from urllib.error import HTTPError

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from PIL import Image

from converter import AppConfig, MarkdownToDocxConverter


class ImageLinkTests(unittest.TestCase):
    def setUp(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.root = Path(folder.name)
        self.converter = MarkdownToDocxConverter()
        buffer = io.BytesIO()
        Image.new('RGB', (24, 24)).save(buffer, format='PNG')
        self.png = buffer.getvalue()

    def convert(self, source):
        cfg = AppConfig(str(self.root), str(self.root))
        return Document(self.converter.convert(source, cfg, 'test.docx'))

    def test_local_images_use_existing_image_layout(self):
        (self.root / 'diagram one.png').write_bytes(self.png)
        img = '[diagram](diagram%20one.png)'
        doc = self.convert(img + '\n\nBefore ' + img + ' after\n\n![explicit](diagram one.png)')
        self.assertEqual(len(doc.inline_shapes), 3)
        self.assertEqual(doc.element.xpath('.//w:hyperlink'), [])
        self.assertEqual(doc.paragraphs[0].alignment, WD_ALIGN_PARAGRAPH.CENTER)
        self.assertEqual(doc.paragraphs[1].text, 'Before  after')

    def test_local_non_images_missing_links_and_code_are_preserved(self):
        (self.root / 'not-an-image.png').write_text('plain text', encoding='utf-8')
        doc = self.convert('[file](not-an-image.png) [missing](missing.jpg) [anchor](#section)\n\n'
                           '`[example](diagram.png)`')
        self.assertEqual(len(doc.inline_shapes), 0)
        self.assertEqual(len(doc.element.xpath('.//w:hyperlink')), 3)
        self.assertEqual(doc.paragraphs[-1].text, '[example](diagram.png)')
        self.assertEqual(self.converter.last_warnings, [])

    def test_remote_image_without_extension_and_webpage_with_image_suffix(self):
        image_url = 'https://example.test/resource?purpose=inline'
        page_url = 'https://example.test/page.png'
        payload = io.BytesIO(self.png)
        payload.headers = {'Content-Type': 'image/png; charset=binary'}
        page = io.BytesIO(b'<html>Web page</html>')
        page.headers = {'Content-Type': 'text/html'}
        with patch('converter.urllib.request.urlopen', side_effect=lambda request, **_: {
            image_url: payload, page_url: page,
        }[request.full_url]) as opened, patch.object(page, 'read', wraps=page.read) as page_read:
            doc = self.convert(f'[image]({image_url})\n\n[source]({page_url})\n\n[again]({image_url})')
        self.assertEqual(opened.call_count, 2)
        page_read.assert_not_called()
        self.assertEqual(len(doc.inline_shapes), 2)
        links = doc.element.xpath('.//w:hyperlink')
        self.assertEqual(len(links), 1)
        self.assertEqual(doc.part.rels[links[0].get(qn('r:id'))].target_ref, page_url)

    def test_failed_download_and_invalid_image_payload_fall_back_to_links(self):
        invalid = io.BytesIO(b'not a PNG')
        invalid.headers = {'Content-Type': 'image/png'}

        def open_url(request, **kwargs):
            if request.full_url.endswith('/timeout'):
                raise TimeoutError
            return invalid

        with patch('converter.urllib.request.urlopen', side_effect=open_url):
            doc = self.convert('[timeout](https://example.test/timeout) [invalid](https://example.test/invalid)')
        self.assertEqual(len(doc.inline_shapes), 0)
        self.assertEqual(len(doc.element.xpath('.//w:hyperlink')), 2)
        self.assertEqual(self.converter.last_warnings, [])

    def test_escaped_brackets_in_link_labels_stay_literal(self):
        (self.root / 'diagram.png').write_bytes(self.png)
        doc = self.convert(
            r'[Paper \[Quick Review\]](diagram.png)' + '\n\n'
            + r'[Source \[Review\]](#source)' + '\n\n'
            + r'\[x^2\]'
        )
        self.assertEqual(len(doc.inline_shapes), 1)
        self.assertEqual(doc.element.xpath('.//w:hyperlink//w:t/text()'), ['Source [Review]'])
        self.assertEqual(len(doc.element.xpath('.//m:oMathPara')), 1)

    def test_http_404_png_body_is_embedded_and_html_body_remains_a_link(self):
        favicon_url = 'https://example.test/favicon?domain=site.example&sz=32'
        page_url = 'https://example.test/missing'

        def open_url(request, timeout):
            if request.full_url == favicon_url:
                raise HTTPError(favicon_url, 404, 'Not Found', {'Content-Type': 'image/png'}, io.BytesIO(self.png))
            self.assertEqual(request.full_url, page_url)
            raise HTTPError(page_url, 404, 'Not Found', {'Content-Type': 'text/html'}, io.BytesIO(b'<html>Missing</html>'))

        with patch('converter.urllib.request.urlopen', side_effect=open_url) as opened:
            doc = self.convert(
                r'[image](https://example.test/favicon?domain=site.example\&sz=32)'
                + f'\n\n[source]({page_url})'
            )
        self.assertCountEqual([call.args[0].full_url for call in opened.call_args_list], [favicon_url, page_url])
        self.assertEqual(len(doc.inline_shapes), 1)
        link, = doc.element.xpath('.//w:hyperlink')
        self.assertEqual(doc.part.rels[link.get(qn('r:id'))].target_ref, page_url)

    def test_consecutive_images_form_rows_and_mosaics_without_crossing_text(self):
        photos = []
        for index, size in enumerate(((600, 540), (610, 550), (630, 520), (600, 520), (880, 400), (710, 520))):
            name = f'photo{index}.png'
            Image.new('RGB', size, (index * 30, 0, 0)).save(self.root / name)
            photos.append(f'[photo{index}]({name})')
        doc = self.convert('\n\n'.join(photos[:3] + ['Between groups'] + photos[3:]))
        row, mosaic = doc.tables
        for table, count in ((row, 3), (mosaic, 2)):
            self.assertEqual(len(table.rows), 1)
            self.assertEqual(len(table.columns), count)
            self.assertFalse(table.autofit)
        self.assertEqual([len(cell._tc.xpath('.//wp:inline')) for cell in row.rows[0].cells], [1, 1, 1])
        self.assertEqual([len(cell._tc.xpath('.//wp:inline')) for cell in mosaic.rows[0].cells], [1, 2])
        self.assertEqual([paragraph.text for paragraph in doc.paragraphs], ['Between groups'])
        self.assertEqual(len(doc.inline_shapes), 6)
        self.assertAlmostEqual(doc.inline_shapes[3].width / doc.inline_shapes[3].height, 600 / 520, places=4)

    def test_four_explicit_images_form_two_rows_and_caption_separates_single_image(self):
        for index in range(5):
            Image.new('RGB', (640, 480)).save(self.root / f'{index}.png')
        images = '\n'.join(f'![photo]({index}.png)' for index in range(4))
        doc = self.convert(images + '\n\nFigure 1: A caption\n\n![single](4.png)')
        self.assertEqual([len(table.columns) for table in doc.tables], [2, 2])
        self.assertEqual(len(doc.paragraphs[-1]._p.xpath('.//wp:inline')), 1)
        self.assertEqual(len(doc.inline_shapes), 5)

    def test_paper_card_keeps_thumbnail_title_metadata_and_link_together(self):
        Image.new('RGB', (400, 640)).save(self.root / 'paper.png')
        doc = self.convert(
            '## Section\n\n[cover](paper.png)\n\n**Paper title**\n\n'
            '- 作者：Example Author\n- 会议：ICCV 2021\n- [论文链接](#paper)\n\nPaper summary.'
        )
        table, = doc.tables
        preview, details = table.rows[0].cells
        self.assertEqual(len(preview._tc.xpath('.//wp:inline')), 1)
        self.assertEqual([p.text for p in details.paragraphs][:3],
                         ['Paper title', '• \t作者：Example Author', '• \t会议：ICCV 2021'])
        link, = details._tc.xpath('.//w:hyperlink')
        self.assertEqual(doc.part.rels[link.get(qn('r:id'))].target_ref, '#paper')
        self.assertEqual(doc.paragraphs[-1].text, 'Paper summary.')

    def test_generic_image_and_bold_heading_with_list_stay_as_separate_blocks(self):
        (self.root / 'diagram.png').write_bytes(self.png)
        doc = self.convert('![diagram](diagram.png)\n\n**Parts of the diagram**\n\n- Front\n- Back')
        self.assertFalse(doc.tables)
        self.assertEqual(len(doc.inline_shapes), 1)
        self.assertEqual(doc.paragraphs[1].text, 'Parts of the diagram')

    def test_favicons_and_adjacent_source_lines_are_removed_before_requests(self):
        source = (
            r'[image](https://www.google.com/s2/favicons?domain=https://doi.org\&sz=32)'
            '\n\nDOI+2\n\n'
            '[image](https://www.google.com/s2/favicons?domain=https://openaccess.thecvf.com&sz=32)'
            '\n开放访问计算机视觉基金会+3\nKeep this sentence.\n\n'
            '![icon](https://example.test/favicon-32x32.png)  \nCVlibs+1\n\n'
            '[image](https://example.test/favicon.ico)\n\nActual paragraph.\n\nElsewhere+1'
        )
        with patch('converter.urllib.request.urlopen') as opened:
            doc = self.convert(source)
        opened.assert_not_called()
        self.assertFalse(doc.inline_shapes)
        self.assertEqual([p.text for p in doc.paragraphs],
                         ['Keep this sentence.', 'Actual paragraph.', 'Elsewhere+1'])
