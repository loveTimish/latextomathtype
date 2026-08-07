import unittest

from restore_formatted_docx_geometry import (
    restore_document_xml,
    restore_object_geometry,
)


class RestoreFormattedGeometryTest(unittest.TestCase):
    def test_restores_object_box_without_replacing_formatted_ole(self):
        before = (
            '<w:object w:dxaOrig="200" w:dyaOrig="179">'
            '<v:shape style="width:10pt;height:9pt"><v:imagedata r:id="rId1"/></v:shape>'
            '<o:OLEObject r:id="rId2"/></w:object>'
        )
        formatted = (
            '<w:object w:dxaOrig="240" w:dyaOrig="220">'
            '<v:shape style="width:1in;height:11pt"><v:imagedata r:id="rId3"/></v:shape>'
            '<o:OLEObject r:id="formattedOle"/></w:object>'
        )

        restored = restore_object_geometry(before, formatted)

        self.assertIn('w:dxaOrig="200"', restored)
        self.assertIn('w:dyaOrig="179"', restored)
        self.assertIn('style="width:10pt;height:9pt"', restored)
        self.assertIn('r:id="formattedOle"', restored)

    def test_restores_position_by_formula_ordinal(self):
        before = (
            '<w:r><w:rPr><w:position w:val="-22"/></w:rPr>'
            '<w:object w:dxaOrig="200" w:dyaOrig="200"><v:shape style="width:10pt;height:10pt"/>'
            '</w:object></w:r>'
        )
        formatted = (
            '<w:r><w:rPr><w:position w:val="-24"/></w:rPr>'
            '<w:object w:dxaOrig="220" w:dyaOrig="220"><v:shape style="width:11pt;height:11pt"/>'
            '</w:object></w:r>'
        )

        restored, _ = restore_document_xml(before, formatted)

        self.assertIn('w:position w:val="-22"', restored)
        self.assertIn('style="width:10pt;height:10pt"', restored)


if __name__ == "__main__":
    unittest.main()
