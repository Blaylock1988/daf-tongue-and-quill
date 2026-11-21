"""Tests for scripts/validate_doc.py and scripts/validate_memo.py."""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import build_doc
import build_slides
import validate_doc
import validate_memo


class TestValidateDoc(unittest.TestCase):
    """Verify multi-format validation, PII/CUI security, DoW/DoD, and XO dump."""

    def test_clean_generated_docx_and_pptx_pass(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            memo_path = Path(tmp) / "clean_memo.docx"
            spec = build_doc.get_schema("memo", dept="dow", cui=True)
            spec["cui_office"] = "1 FW/XP"
            spec["cui_poc"] = "Maj Smith, DSN 555-0100"
            spec["paragraphs"] = [
                {"level": 0, "text": "Implement updated administrative standards by 15 December 2026."},
                {"level": 0, "text": "Direct questions to Maj Jordan Smith at DSN 555-0100."},
            ]
            build_doc.build_document(spec, memo_path, dept="dow", cui=True, cui_category="OPSEC", cui_ldc="FEDCON")
            rep = validate_doc.validate_file(memo_path, doc_type="memo", dept_mode="dow", strict=True)
            self.assertTrue(rep["passed"], msg=str(rep["issues"]))
            self.assertEqual(validate_memo.main([str(memo_path), "--dept", "dow", "--strict"]), 0)

            deck_path = Path(tmp) / "clean_deck.pptx"
            deck_spec = build_slides.get_slides_schema(dept="dod", cui=True)
            build_slides.build_deck(deck_spec, deck_path, dept="dod", cui=True, cui_category="OPSEC", cui_ldc="FEDCON")
            s_rep = validate_doc.validate_file(deck_path, doc_type="slides", dept_mode="dod")
            self.assertTrue(s_rep["passed"], msg=str(s_rep["issues"]))

    def test_security_pii_cui_and_fouo_flags(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            bad = Path(tmp) / "bad_sec.txt"
            bad.write_text(
                "21 Nov 2026\n"
                "MEMORANDUM FOR:  1 FW/CC\n"
                "FROM:  1 FW/XP\n"
                "SUBJECT:  Security Errors\n\n"
                "UNCLASSIFIED//CUI\n"
                "FOUO SSN 123-45-6789\n"
                "(S//NOFORN) Classified test line\n",
                encoding="utf-8",
            )
            rep = validate_doc.validate_file(bad, doc_type="memo")
            self.assertFalse(rep["passed"])
            codes = {i["code"] for i in rep["issues"]}
            self.assertIn("SEC-SSN", codes)
            self.assertIn("SEC-CLASSIFIED", codes)
            self.assertIn("CUI-FOUO-OBSOLETE", codes)
            self.assertIn("CUI-INVALID-BANNER", codes)
            self.assertIn("MEMO-COLON", codes)
            self.assertIn("FMT-MIXED-DATE", codes)

    def test_attachments_acronyms_deai_and_dow_dod_mixing(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            draft = Path(tmp) / "flawed.txt"
            draft.write_text(
                "21 November 2026\n"
                "MEMORANDUM FOR  1 FW/CC\n"
                "FROM:  1 FW/XP\n"
                "SUBJECT:  Mixed Usage\n\n"
                "1.  We must delve into Atch 2 and coordinate with OPR under Department of War and Department of Defense rules.\n"
                "2.  In order to coord this faster, this single sentence intentionally runs far past thirty words with lots of filler phrases and repetitive redundancy like advance planning so we leverage holistic synergy.\n\n"
                "Jane Doe, Col, USAF\n"
                "Acting Commander\n\n"
                "2 Attachments:\n"
                "1.  Summary\n",
                encoding="utf-8",
            )
            rep = validate_doc.validate_file(draft, doc_type="memo")
            self.assertFalse(rep["passed"])
            codes = {i["code"] for i in rep["issues"]}
            self.assertIn("ATCH-COUNT-MISMATCH", codes)
            self.assertIn("ATCH-MISSING-REF", codes)
            self.assertIn("SIG-ACTING-CMD", codes)
            self.assertIn("SIG-NAME-CAPS", codes)
            self.assertIn("DEPT-MIXED", codes)
            self.assertIn("ACR-COLLISION", codes)
            self.assertIn("ABBR-FORMAL", codes)
            self.assertIn("STYLE-AI-BUZZWORD", codes)
            self.assertIn("STYLE-WORDY", codes)
            self.assertIn("STYLE-REDUNDANT", codes)

    def test_bio_letter_and_slides_errors_plus_dump_text(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            bio_txt = Path(tmp) / "bio.txt"
            bio_txt.write_text(
                "COLONEL JANE DOE\n"
                "She and her husband moved after DoW guidance and flew over 2,000 hours.\n"
                "ASSIGNMENTS\n"
                "1. Jan. 2010 - May 2012, Flight Commander.\n"
                "EFFECTIVE DATES OF PROMOTION\n"
                "Colonel July 1, 2024\n",
                encoding="utf-8",
            )
            b_rep = validate_doc.validate_file(bio_txt, doc_type="bio")
            self.assertFalse(b_rep["passed"])
            b_codes = {i["code"] for i in b_rep["issues"]}
            self.assertIn("BIO-PRIVACY-FAMILY", b_codes)
            self.assertIn("BIO-ASSIGN-MONTHS", b_codes)
            self.assertIn("BIO-DOW-CASE", b_codes)
            self.assertIn("BIO-MORE-THAN", b_codes)
            self.assertIn("BIO-CURRENT-AS-OF", b_codes)

            ltr_txt = Path(tmp) / "letter.txt"
            ltr_txt.write_text(
                "Dear Colonel Smith:\n\nThank you.\n\nSincerely,\n\nFOR THE COMMANDER\n",
                encoding="utf-8",
            )
            l_codes = {i["code"] for i in validate_doc.validate_file(ltr_txt, doc_type="letter")["issues"]}
            self.assertIn("LETTER-OPEN-PUNCT", l_codes)
            self.assertIn("LETTER-CLOSE-PUNCT", l_codes)
            self.assertIn("LETTER-AUTH-LINE", l_codes)

            deck_md = Path(tmp) / "slides.md"
            deck_md.write_text(
                "## Slide 1: Bad Briefing\n"
                "- This single bullet line has way more than seven words on it\n"
                "- Line two\n- Line three\n- Line four\n- Line five\n- Line six\n- Line seven\n- Line eight\n"
                "- VALUE IS 100.00\n"
                "Notes: Speak plainly.\n",
                encoding="utf-8",
            )
            s_codes = {i["code"] for i in validate_doc.validate_file(deck_md, doc_type="slides")["issues"]}
            self.assertIn("SLIDE-7X7-LINES", s_codes)
            self.assertIn("SLIDE-7X7-WORDS", s_codes)
            self.assertIn("SLIDE-DECIMAL-ZERO", s_codes)

            dumped = validate_doc.dump_extracted(deck_md)
            self.assertEqual(len(dumped["slides"]), 1)
            self.assertEqual(dumped["slides"][0]["notes"], "Speak plainly.")


if __name__ == "__main__":
    unittest.main()
