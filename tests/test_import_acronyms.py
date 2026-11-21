"""Tests for scripts/import_acronyms.py."""
from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import docx
import import_acronyms


class TestImportAcronyms(unittest.TestCase):
    """Verify command writing guide extraction, split schemas, and collisions."""

    def test_import_txt_splits_acronyms_and_abbreviations_with_collisions(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            guide = Path(tmp) / "guide.txt"
            guide.write_text(
                "ACRONYMS\n"
                "SpOC (Space Operations Command)\n"
                "OPR - Operational Readiness Review\n"
                "AI (Air Interdiction)\n"
                "ABBREVIATIONS\n"
                "coord (coordinate)\n"
                "w/ (with)\n",
                encoding="utf-8",
            )
            payload = import_acronyms.import_command_guide(
                guide, source_name="SpOC Guide", effective_date="21 Nov 26", mode="strict"
            )
            self.assertEqual(payload["mode"], "strict")
            self.assertIn("SpOC", payload["acronyms"])
            self.assertIn("coord", payload["abbreviations"])
            self.assertTrue(payload["abbreviations"]["coord"]["verb_capable"])
            self.assertFalse(payload["abbreviations"]["w/"]["formal_allowed"])
            self.assertIn("OPR", payload["collisions"])
            self.assertIn("Operational Readiness Review", payload["collisions"]["OPR"])
            self.assertIn("Office of Primary Responsibility", payload["collisions"]["OPR"])
            self.assertIn("AI", payload["collisions"])
            self.assertIn("Artificial Intelligence", payload["collisions"]["AI"])

    def test_import_csv_respects_explicit_kind(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            csv_path = Path(tmp) / "list.csv"
            csv_path.write_text(
                "term,meaning,kind\n"
                "DEL,Delta,abbreviation\n"
                "STARCOM,Space Training and Readiness Command,acronym\n",
                encoding="utf-8",
            )
            payload = import_acronyms.import_command_guide(
                csv_path, source_name="USSF CSV", effective_date="21 Nov 26", mode="additive"
            )
            self.assertIn("STARCOM", payload["acronyms"])
            self.assertIn("DEL", payload["abbreviations"])
            self.assertNotIn("DEL", payload["acronyms"])

    def test_import_docx_table_and_cli_write(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            doc_path = Path(tmp) / "guide.docx"
            out_path = Path(tmp) / "out.json"
            doc = docx.Document()
            doc.add_paragraph("ACRONYMS")
            table = doc.add_table(rows=2, cols=2)
            table.cell(0, 0).text = "SSC"
            table.cell(0, 1).text = "Space Systems Command"
            table.cell(1, 0).text = "SBD"
            table.cell(1, 1).text = "Space Base Delta"
            doc.save(str(doc_path))

            rc = import_acronyms.main([str(doc_path), "--out", str(out_path), "--source", "SSC Guide"])
            self.assertEqual(rc, 0)
            written = json.loads(out_path.read_text(encoding="utf-8"))
            self.assertEqual(written["acronyms"]["SSC"], ["Space Systems Command"])
            self.assertEqual(written["acronyms"]["SBD"], ["Space Base Delta"])


if __name__ == "__main__":
    unittest.main()
