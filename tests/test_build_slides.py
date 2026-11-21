"""Tests for scripts/build_slides.py Ch. 10 PowerPoint (.pptx) builder."""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from pptx import Presentation
import build_slides


class TestBuildSlides(unittest.TestCase):
    """Verify Ch. 10 16:9 geometry, 40/30/24 pt typography, 4x7 tables, notes & CUI."""

    def test_build_deck_geometry_cui_notes_and_continued(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "brief.pptx"
            spec = build_slides.get_slides_schema(dept="dow", cui=True)
            spec["slides"].append({
                "title": "Long Bullet Topic",
                "bullets": [f"Action item {i}" for i in range(1, 10)],
                "notes": "Two slides auto-paginated for 7x7 compliance.",
            })
            spec["slides"][3]["table"]["rows"][1][2] = "100.00"
            build_slides.build_deck(
                spec,
                out,
                palette_name="ussf",
                dept="dow",
                cui=True,
                cui_category="OPSEC",
                cui_ldc="FEDCON",
            )
            prs = Presentation(str(out))
            self.assertAlmostEqual(prs.slide_width.inches, 13.333, places=2)
            self.assertAlmostEqual(prs.slide_height.inches, 7.5, places=2)
            # 4 base slides + 2 paginated from Long Bullet Topic = 6 slides
            self.assertEqual(len(prs.slides), 6)

            slide0_text = "\n".join(s.text for s in prs.slides[0].shapes if s.has_text_frame)
            self.assertIn("CUI//OPSEC//FEDCON", slide0_text)
            self.assertIn("Controlled by: Department of War", slide0_text)

            titles = [
                s.text for slide in prs.slides for s in slide.shapes
                if s.has_text_frame and any(r.font.size and r.font.size.pt == 40 for p in s.text_frame.paragraphs for r in p.runs)
            ]
            self.assertTrue(any("Continued" in t for t in titles))

            tbl_slide = prs.slides[3]
            tables = [s.table for s in tbl_slide.shapes if s.has_table]
            self.assertEqual(len(tables), 1)
            self.assertEqual(tables[0].cell(1, 2).text, "100")

            self.assertIn(
                "Current average package",
                prs.slides[1].notes_slide.notes_text_frame.text,
            )

    def test_rejects_oversized_table(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "bad.pptx"
            spec = build_slides.get_slides_schema()
            spec["slides"][3]["table"]["rows"] = [["A", "B"]] * 5
            with self.assertRaises(ValueError):
                build_slides.build_deck(spec, out)


if __name__ == "__main__":
    unittest.main()
