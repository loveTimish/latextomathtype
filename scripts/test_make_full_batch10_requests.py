import unittest
import json
import tempfile
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
from make_full_batch10_requests import normalize_math_delimiters, report_equations, tex_to_plain_blocks


class MathDelimiterNormalizationTest(unittest.TestCase):
    def test_empty_source_equation_is_preserved_without_period_repair(self):
        with tempfile.TemporaryDirectory() as directory:
            report_dir = Path(directory) / "4"
            report_dir.mkdir()
            (report_dir / "4.report.json").write_text(json.dumps({"equations": [{
                "status": "converted",
                "output": ".",
                "sourceRepairReason": "empty-output",
            }]}), encoding="utf-8")

            equations = report_equations(4, Path(directory))

        self.assertEqual(1, len(equations))
        self.assertEqual("", equations[0]["output"])

    def test_unresolved_graphic_before_adjacent_formulas_does_not_break_delimiters(self):
        source = (
            r"\includegraphics[width=1\textwidth] 22.docx.tmp/word/media/image448.emf"
            r"$$\pwmetrics{307.400,16.000}\frac{x-26}{6k}$$"
            r"$\pwmetrics{150.800,16.000}=\frac{1}{6k}\left\{x+14\right\}$"
        )

        normalized = normalize_math_delimiters(source)
        blocks = tex_to_plain_blocks(source, preserve_source_ole_spans=True)

        self.assertNotIn(r"\includegraphics", normalized)
        self.assertIn(r"$$\pwmetrics{307.400,16.000}\frac{x-26}{6k}$$", normalized)
        self.assertIn(r"$\pwmetrics{150.800,16.000}=\frac{1}{6k}", normalized)
        self.assertEqual(1, len(blocks))
        self.assertIn(r"\frac{x-26}{6k}", blocks[0])
        self.assertIn(r"=\frac{1}{6k}", blocks[0])

    def test_corrupt_metric_fragment_inside_graphic_filename_is_removed(self):
        source = (
            r"before \includegraphics[width=1\textwidth] "
            r"73.docx.tmp/word/media/image4$\pwmetrics{46.400,13.000}97.png after"
        )

        normalized = normalize_math_delimiters(source)

        self.assertEqual("before  after", normalized)
        self.assertNotIn(r"\pwmetrics", normalized)

    def test_plain_text_tabularx_commands_are_removed_without_touching_formula_alignment(self):
        source = (
            r"\begin table \begin tabularx \textwidth |p \dimexpr 0.252\linewidth"
            r"-2\tabcolsep-2\arrayrulewidth | \arraybackslash 立体图形 & "
            r"\arraybackslash 体积公式 $\pwmetrics{80.000,16.000}"
            r"\begin{aligned}V&=abh\\V&=Sh\end{aligned}$ "
            r"\textsuperscript 正方体 \end tabularx \end table"
        )

        blocks = tex_to_plain_blocks(source, preserve_source_ole_spans=True)
        result = " ".join(blocks)

        self.assertIn("立体图形", result)
        self.assertIn("体积公式", result)
        self.assertIn(r"\begin{aligned}V&=abh", result)
        self.assertNotIn(r"\arraybackslash", result)
        self.assertNotIn(r"\end tabularx", result)
        self.assertNotIn(r"\textsuperscript", result)

    def test_malformed_plain_text_color_prefix_is_removed_before_formula(self):
        source = (
            r"请在右图 \textcolor color-"
            r"$\pwmetrics{69.600,13.000}993300 4 \times 4$ 表格中填数"
        )

        result = " ".join(tex_to_plain_blocks(source, preserve_source_ole_spans=True))

        self.assertNotIn(r"\textcolor", result)
        self.assertIn(r"$\pwmetrics{69.600,13.000}993300 4 \times 4$", result)

    def test_loose_box_wrapper_preserves_heading_text(self):
        result = " ".join(tex_to_plain_blocks(r"\fbox 5-2-1.数的整除", preserve_source_ole_spans=True))

        self.assertEqual("5-2-1.数的整除", result)
        self.assertNotIn(r"\fbox", result)


if __name__ == "__main__":
    unittest.main()
