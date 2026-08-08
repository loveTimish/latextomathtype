import unittest

from compare_docx_pair_metrics import parse_vml_dimension_pt


class ParseVmlDimensionTest(unittest.TestCase):
    def test_converts_word_vml_units_to_points(self):
        self.assertEqual(72.0, parse_vml_dimension_pt("width:1in;height:31.1pt", "width"))
        self.assertEqual(31.1, parse_vml_dimension_pt("width:1in;height:31.1pt", "height"))
        self.assertAlmostEqual(72.0, parse_vml_dimension_pt("width:2.54cm", "width"))
        self.assertAlmostEqual(72.0, parse_vml_dimension_pt("width:25.4mm", "width"))
        self.assertEqual(72.0, parse_vml_dimension_pt("width:96px", "width"))

    def test_requires_a_supported_explicit_unit(self):
        self.assertIsNone(parse_vml_dimension_pt("width:auto", "width"))
        self.assertIsNone(parse_vml_dimension_pt("height:12em", "height"))


if __name__ == "__main__":
    unittest.main()
