import unittest
import sys
import tempfile
import zipfile
from pathlib import Path

import numpy as np

sys.path.insert(0, str(Path(__file__).resolve().parent))
from compare_mathtype_standard_visuals import (
    extract_object_wmfs, extract_wmfs, registered_metrics, shift_on_canvas,
)


class RegisteredMetricsTest(unittest.TestCase):
    def test_extracts_wmfs_in_object_order_and_preserves_reused_previews(self):
        document = """<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"
          xmlns:v=\"urn:schemas-microsoft-com:vml\"
          xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>
          <w:object><v:imagedata r:id=\"rId2\"/></w:object>
          <w:object><v:imagedata r:id=\"rId1\"/></w:object>
          <w:object><v:imagedata r:id=\"rId2\"/></w:object>
          </w:body></w:document>"""
        relationships = """<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
          <Relationship Id=\"rId1\" Target=\"media/image1.wmf\" Type=\"image\"/>
          <Relationship Id=\"rId2\" Target=\"media/image2.wmf\" Type=\"image\"/>
          </Relationships>"""
        with tempfile.TemporaryDirectory() as directory:
            docx = Path(directory) / "sample.docx"
            with zipfile.ZipFile(docx, "w") as archive:
                archive.writestr("word/document.xml", document)
                archive.writestr("word/_rels/document.xml.rels", relationships)
                archive.writestr("word/media/image1.wmf", b"one")
                archive.writestr("word/media/image2.wmf", b"two")

            self.assertEqual(
                [("word/media/image2.wmf", b"two"),
                 ("word/media/image1.wmf", b"one"),
                 ("word/media/image2.wmf", b"two")],
                extract_wmfs(docx),
            )

    def test_preserves_object_slots_for_non_wmf_previews(self):
        document = """<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:v="urn:schemas-microsoft-com:vml"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
          <w:object><v:imagedata r:id="rId1"/></w:object>
          <w:object><v:imagedata r:id="rId2"/></w:object>
          </w:body></w:document>"""
        relationships = """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Target="media/image1.emf" Type="image"/>
          <Relationship Id="rId2" Target="media/image2.wmf" Type="image"/>
          </Relationships>"""
        with tempfile.TemporaryDirectory() as directory:
            docx = Path(directory) / "sample.docx"
            with zipfile.ZipFile(docx, "w") as archive:
                archive.writestr("word/document.xml", document)
                archive.writestr("word/_rels/document.xml.rels", relationships)
                archive.writestr("word/media/image1.emf", b"emf")
                archive.writestr("word/media/image2.wmf", b"wmf")

            self.assertEqual([None, ("word/media/image2.wmf", b"wmf")], extract_object_wmfs(docx))

    def test_recovers_one_pixel_rasterization_shift(self):
        standard = np.full((12, 12), 255.0)
        standard[3:9, 4:8] = 0.0
        generated = shift_on_canvas(standard, 1, -1)

        ssim, iou, missing, dy, dx = registered_metrics(standard, generated, 250, 1)

        self.assertAlmostEqual(1.0, ssim)
        self.assertAlmostEqual(1.0, iou)
        self.assertEqual(0, missing)
        self.assertEqual((-1, 1), (dy, dx))

    def test_does_not_hide_missing_major_component(self):
        standard = np.full((16, 16), 255.0)
        standard[2:7, 2:7] = 0.0
        standard[9:14, 9:14] = 0.0
        generated = standard.copy()
        generated[9:14, 9:14] = 255.0

        _, _, missing, _, _ = registered_metrics(standard, generated, 250, 1)

        self.assertEqual(1, missing)


if __name__ == "__main__":
    unittest.main()
