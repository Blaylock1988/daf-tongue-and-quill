"""Tests for scripts/import_acronyms.py."""
from __future__ import annotations

import io
import json
import sys
import tempfile
import unittest
import urllib.error
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import patch

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

    def test_check_afpc_from_html_reports_diff_and_clean_status(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            html_new = Path(tmp) / "afpc_new.html"
            html_new.write_text(
                "<html><body>"
                "<p>Updated: 15 Dec 2026</p>"
                "<table>"
                "<tr><th>Acronym</th><th>Definition</th></tr>"
                "<tr><td>ABM</td><td>Air Battle Manager</td></tr>"
                "<tr><td>ACE</td><td>Agile Combat Employment</td></tr>"
                "<tr><td>ACE</td><td>Airspace Control Element</td></tr>"
                "<tr><td>ZETA9</td><td>Zero Emission Tactical Aircraft</td></tr>"
                "</table></body></html>",
                encoding="utf-8",
            )
            buf = io.StringIO()
            with redirect_stdout(buf):
                rc = import_acronyms.main(
                    ["--check-afpc", "--from-file", str(html_new), "--json"]
                )
            self.assertEqual(rc, 2)
            diff = json.loads(buf.getvalue())
            self.assertTrue(diff["has_changes"])
            self.assertEqual(diff["detected_date"], "15 Dec 2026")
            self.assertIn("ZETA9", diff["added"])
            self.assertEqual(diff["new_meanings"]["ACE"], ["Airspace Control Element"])

            html_same = Path(tmp) / "afpc_same.html"
            html_same.write_text(
                "<html><body><table>"
                "<tr><td>ABM</td><td>Air Battle Manager</td></tr>"
                "<tr><td>ACE</td><td>Agile Combat Employment</td></tr>"
                "</table></body></html>",
                encoding="utf-8",
            )
            buf2 = io.StringIO()
            with redirect_stdout(buf2):
                rc2 = import_acronyms.main(
                    ["--check-afpc", "--from-file", str(html_same), "--json"]
                )
            self.assertEqual(rc2, 0)
            self.assertFalse(json.loads(buf2.getvalue())["has_changes"])

    def test_sync_afpc_from_html_preserves_metadata_and_sets_last_synced(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            html_path = Path(tmp) / "afpc_update.html"
            html_path.write_text(
                "<html><body>"
                "<p>Effective 20 Jan 2027</p>"
                "<table><tr><td>XRAY</td><td>Extended Range Array</td></tr></table>"
                "</body></html>",
                encoding="utf-8",
            )
            out_json = Path(tmp) / "afpc_synced.json"
            buf = io.StringIO()
            with redirect_stdout(buf):
                rc = import_acronyms.main(
                    ["--sync-afpc", "--from-file", str(html_path), "--out", str(out_json)]
                )
            self.assertEqual(rc, 0)
            synced = json.loads(out_json.read_text(encoding="utf-8"))
            self.assertIn("XRAY", synced["acronyms"])
            self.assertEqual(synced["detected_source_date"], "20 Jan 2027")
            self.assertTrue(synced["last_synced"])
            self.assertIn("principles", synced)
            self.assertIn("categories", synced)
            self.assertIn("abbreviations", synced)
            self.assertIn("multi_meaning_collisions", synced)

    def test_check_afpc_http_403_exits_3_with_from_file_guidance(self) -> None:
        err = urllib.error.HTTPError(
            import_acronyms.AFPC_URL, 403, "Forbidden", hdrs=None, fp=None
        )
        buf = io.StringIO()
        try:
            with patch("urllib.request.urlopen", side_effect=err), redirect_stdout(buf):
                rc = import_acronyms.main(["--check-afpc"])
        finally:
            err.close()
        self.assertEqual(rc, 3)
        self.assertIn("--from-file", buf.getvalue())


if __name__ == "__main__":
    unittest.main()
