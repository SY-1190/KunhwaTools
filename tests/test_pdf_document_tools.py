import base64
import io
import unittest
from pathlib import Path
from zipfile import ZipFile

import pymupdf as fitz
from PIL import Image

from pdf_document_tools import analyze_pdf, convert_word, edit_pdf


class PDFDocumentTests(unittest.TestCase):
    def setUp(self):
        doc = fitz.open()
        page = doc.new_page(width=595, height=842)
        page.insert_text((40, 60), "Original text", fontsize=14)
        page.insert_text((40, 100), "Secret to delete", fontsize=12)
        buffer = io.BytesIO()
        Image.new("RGB", (100, 80), "coral").save(buffer, format="PNG")
        self.image = buffer.getvalue()
        page.insert_image(fitz.Rect(40, 140, 140, 220), stream=self.image)
        for x in (40, 180, 320):
            page.draw_line((x, 280), (x, 370))
        for y in (280, 310, 340, 370):
            page.draw_line((40, y), (320, y))
        for y, values in ((300, ("Part", "Count")), (330, ("Bolt", "12")), (360, ("Nut", "24"))):
            for x, text in zip((50, 190), values):
                page.insert_text((x, y), text)
        doc.new_page().insert_text((40, 60), "Second page")
        self.data = doc.tobytes()
        doc.close()

    def test_actual_redaction_korean_and_image_replacement(self):
        objects = analyze_pdf(self.data)["pages"][0]["objects"]
        line = next(o for o in objects if o.get("text") == "Original text")
        secret = next(o for o in objects if o.get("text") == "Secret to delete")
        image = next(o for o in objects if o["type"] == "image")
        buffer = io.BytesIO()
        Image.new("RGB", (100, 80), "teal").save(buffer, format="PNG")
        output = edit_pdf(self.data, [
            dict(line, page=0, sourceId=line["id"], rect=[40, 40, 400, 80], text="수정한 한글", size=16),
            dict(secret, page=0, sourceId=secret["id"], deleted=True),
            dict(image, page=0, sourceId=image["id"], image="data:image/png;base64," + base64.b64encode(buffer.getvalue()).decode()),
        ])
        with fitz.open(stream=output, filetype="pdf") as doc:
            text = doc[0].get_text()
            self.assertNotIn("Original text", text)
            self.assertNotIn("Secret to delete", text)
            self.assertIn("수정한 한글", text)
            pix = doc[0].get_pixmap()
            self.assertEqual(pix.pixel(60, 160)[:3], (0, 128, 128))

    def test_word_is_editable_and_has_table_and_media(self):
        data, report = convert_word(self.data)
        with ZipFile(io.BytesIO(data)) as archive:
            xml = archive.read("word/document.xml").decode("utf-8")
            self.assertIn("Original text", xml)
            self.assertIn("<w:tbl>", xml)
            self.assertIn("Bolt", xml)
            self.assertTrue(any(n.startswith("word/media/") for n in archive.namelist()))
        self.assertEqual(report["tables"], 1)
        self.assertEqual(report["pages"], 2)

    def test_appearance_and_scanned_modes(self):
        _, report = convert_word(self.data, "appearance")
        self.assertEqual(report["images"], 2)
        scan = fitz.open()
        scan.new_page().insert_image(fitz.Rect(40, 40, 140, 120), stream=self.image)
        _, report = convert_word(scan.tobytes())
        self.assertEqual(report["scannedPages"], [1])
        self.assertEqual(report["paragraphs"], 0)

    def test_rotated_page_and_invalid_input(self):
        doc = fitz.open(stream=self.data, filetype="pdf")
        doc[0].set_rotation(90)
        self.assertAlmostEqual(analyze_pdf(doc.tobytes())["pages"][0]["width"], 842)
        for operation in ({"page": 99, "type": "text"}, {"page": 0, "type": "text", "rect": [-1, 0, 20, 20]}):
            with self.assertRaises(ValueError):
                edit_pdf(self.data, [operation])
        with self.assertRaises(ValueError):
            analyze_pdf(b"invalid file")

    def test_merged_word_cells(self):
        doc = fitz.open()
        page = doc.new_page()
        for y in (40, 80, 120):
            page.draw_line((40, y), (320, y))
        page.draw_line((40, 40), (40, 120))
        page.draw_line((320, 40), (320, 120))
        page.draw_line((180, 80), (180, 120))
        page.insert_text((50, 65), "Merged header")
        page.insert_text((50, 105), "Left")
        page.insert_text((190, 105), "Right")
        word, report = convert_word(doc.tobytes())
        self.assertEqual(report["tables"], 1)
        with ZipFile(io.BytesIO(word)) as archive:
            self.assertIn('w:gridSpan w:val="2"', archive.read("word/document.xml").decode())


if __name__ == "__main__":
    unittest.main()
