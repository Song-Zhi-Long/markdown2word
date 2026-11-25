import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

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
