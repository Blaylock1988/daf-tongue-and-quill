"""Tests for scripts/tq_lookup.py."""
from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import tq_lookup


class TestTqLookup(unittest.TestCase):
    """Verify rank, DoW/DoD address, acronym collision, abbrev, wordy, fallacy, and CUI lookups."""

    def test_rank_lookup_covers_ussf_and_usaf_warrant_officers(self) -> None:
        spc = tq_lookup.lookup_rank("Spc3")
        self.assertTrue(any(r["service"] == "USSF" and r["abbrev"] == "Spc3" for r in spc))

        wo = tq_lookup.lookup_rank("WO1")
        self.assertTrue(any(r["service"] == "USAF" and r["abbrev"] == "WO1" for r in wo))

    def test_address_lookup_pairs_dow_and_dod(self) -> None:
        secwar = tq_lookup.lookup_address("Secretary of War")
        secdef = tq_lookup.lookup_address("Secretary of Defense")
        self.assertTrue(any("Secretary of War" in e["addressee"] for e in secwar))
        self.assertTrue(any("Secretary of Defense" in e["addressee"] for e in secdef))

    def test_acronym_collision_and_command_override_lookup(self) -> None:
        opr = tq_lookup.lookup_acronym("OPR", None)
        exact = [m for m in opr["matches"] if m["term"] == "OPR"]
        self.assertEqual(len(exact), 1)
        self.assertTrue(exact[0]["ambiguous"])
        self.assertGreaterEqual(len(exact[0]["meanings"]), 2)

        with tempfile.TemporaryDirectory() as tmp:
            cmd_json = Path(tmp) / "cmd.json"
            cmd_json.write_text(
                json.dumps({
                    "source": "Local Wing",
                    "mode": "strict",
                    "acronyms": {"ACE": ["Agile Combat Element"]},
                    "abbreviations": {},
                }),
                encoding="utf-8",
            )
            style_json = Path(tmp) / "style.json"
            style_json.write_text(
                json.dumps({"acronym_list_path": str(cmd_json), "acronym_mode": "strict"}),
                encoding="utf-8",
            )
            ace = tq_lookup.lookup_acronym("ACE", str(style_json))
            exact_ace = [m for m in ace["matches"] if m["term"] == "ACE"][0]
            self.assertEqual(ace["mode"], "strict")
            self.assertTrue(exact_ace["ambiguous"])
            self.assertIn("Agile Combat Employment", exact_ace["meanings"])
            self.assertIn("Agile Combat Element", exact_ace["meanings"])

    def test_abbrev_wordy_fallacy_and_cui_lookups(self) -> None:
        abbrevs = tq_lookup.lookup_abbrev("coord", None)
        self.assertTrue(any(a["abbreviation"] == "coord" and a["verb_capable"] for a in abbrevs))

        wordy = tq_lookup.lookup_wordy("in order to")
        self.assertTrue(any(w["avoid"] == "in order to" and w["try"] == "to" for w in wordy))

        fallacy = tq_lookup.lookup_fallacy("post hoc")
        self.assertTrue(any("False Cause" in f["name"] for f in fallacy))

        cui = tq_lookup.lookup_cui("FEDCON")
        self.assertTrue(any(item["ldc"] == "FEDCON" for item in cui["ldcs"]))


if __name__ == "__main__":
    unittest.main()
