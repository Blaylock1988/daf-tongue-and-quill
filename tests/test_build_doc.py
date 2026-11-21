"""Tests for scripts/build_doc.py Word (.docx) & Staff Package builder."""
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
from docx.shared import Inches
import build_doc


class TestBuildDoc(unittest.TestCase):
    """Verify T&Q Word geometry, presets, CUI banners/DI, and package bundling."""

    def test_memo_geometry_cui_and_signature_indent(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "memo.docx"
            spec = build_doc.get_schema("memo", preset="etp", dept="dow", cui=True)
            spec["paragraphs"] = [{"level": 0, "text": ("Detailed paragraph text. " * 18).strip()} for _ in range(4)]
            build_doc.build_document(
                spec,
                out,
                dept="dow",
                cui=True,
                cui_category="EXPT, CTI",
                cui_ldc="FEDCON",
                cui_office="1 FW/XP",
                cui_poc="Maj Smith, DSN 555-0100",
            )
            doc = docx.Document(str(out))
            sec = doc.sections[0]
            self.assertAlmostEqual(sec.left_margin.inches, 1.0, places=2)
            self.assertTrue(sec.different_first_page_header_footer)

            hdr1_text = "\n".join(p.text for p in sec.first_page_header.paragraphs)
            self.assertIn("CUI//CTI/EXPT//FEDCON", hdr1_text)
            self.assertIn("DEPARTMENT OF THE AIR FORCE", hdr1_text)

            ftr_text = "\n".join(p.text for p in sec.footer.paragraphs)
            self.assertIn("CUI//CTI/EXPT//FEDCON", ftr_text)

            cont_xml = sec.header._element.xml
            self.assertIn("PAGE", cont_xml)

            sig_paras = [p for p in doc.paragraphs if "JANE A. DOE" in p.text]
            self.assertEqual(len(sig_paras), 1)
            self.assertTrue(sig_paras[0].paragraph_format.keep_together)
            self.assertAlmostEqual(
                sec.left_margin.inches + sig_paras[0].paragraph_format.left_indent.inches,
                4.5,
                places=2,
            )
            all_text = "\n".join(p.text for p in doc.paragraphs)
            self.assertIn("CUI Category: CTI, EXPT", all_text)

    def test_short_memo_balances_margins_keeping_4_5_inch_signature(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "short.docx"
            spec = build_doc.get_schema("memo", preset="assumption-of-command")
            build_doc.build_document(spec, out)
            doc = docx.Document(str(out))
            sec = doc.sections[0]
            self.assertAlmostEqual(sec.left_margin.inches, 1.5, places=2)
            sig = [p for p in doc.paragraphs if "ALEXANDER M. GRANT" in p.text][0]
            self.assertAlmostEqual(
                sec.left_margin.inches + sig.paragraph_format.left_indent.inches,
                4.5,
                places=2,
            )

    def test_bio_personal_letter_and_bbp_geometry(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            bio_out = Path(tmp) / "bio.docx"
            build_doc.build_document(build_doc.get_schema("bio"), bio_out)
            bio = docx.Document(str(bio_out))
            id_cell = bio.tables[0].cell(0, 0)
            self.assertEqual(id_cell.paragraphs[0].runs[0].font.name, "Arial")
            self.assertEqual(id_cell.paragraphs[0].runs[0].font.size.pt, 13.5)
            self.assertIn("OFFICIAL PHOTOGRAPH", bio.tables[0].cell(0, 1).text)

            ltr_out = Path(tmp) / "letter.docx"
            ltr_spec = build_doc.get_schema("letter", dept="dow")
            ltr_spec["salutation"] = "Dear Mr. Secretary:"
            build_doc.build_document(ltr_spec, ltr_out)
            ltr = docx.Document(str(ltr_out))
            self.assertAlmostEqual(ltr.sections[0].top_margin.inches, 1.75, places=2)
            sal = [p for p in ltr.paragraphs if "Dear Mr. Secretary" in p.text][0]
            self.assertFalse(sal.text.endswith(":"))

            bbp_out = Path(tmp) / "bbp.docx"
            build_doc.build_document(build_doc.get_schema("bbp"), bbp_out)
            bbp = docx.Document(str(bbp_out))
            self.assertIn("BULLET BACKGROUND PAPER", "\n".join(p.text for p in bbp.paragraphs))

    def test_staff_package_bundler_and_init_style_cli(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            pkg_dir = Path(tmp) / "pkg"
            pkg_spec = {
                "dept": "dow",
                "subject": "Unified Wing Routing Policy",
                "suspense": "30 Nov 26",
                "from": "1 FW/XP",
                "to": "1 FW/CC",
                "cui": True,
                "cui_category": "OPSEC",
                "cui_ldc": "FEDCON",
                "cui_poc": "Maj Smith, DSN 555-0100",
                "tabs": [
                    {"label": "Outgoing Memo", "filename": "Tab1_Memo.docx", **build_doc.get_schema("memo", preset="policy", dept="dow")},
                    {"label": "BBP", "filename": "Tab2_BBP.docx", **build_doc.get_schema("bbp", dept="dow")},
                ],
            }
            files = build_doc.build_staff_package(pkg_spec, pkg_dir)
            names = {p.name for p in files}
            self.assertEqual(names, {"00_eSSS.docx", "Tab1_Memo.docx", "Tab2_BBP.docx", "00_Cover_Email.txt"})
            email = (pkg_dir / "00_Cover_Email.txt").read_text(encoding="utf-8")
            self.assertIn("CUI//OPSEC//FEDCON", email)
            self.assertIn("Unified Wing Routing Policy", email)

            style_out = Path(tmp) / "unit_style.json"
            rc = build_doc.main(["init-style", str(style_out), "--dept", "dow", "--service", "USSF"])
            self.assertEqual(rc, 0)
            cfg = json.loads(style_out.read_text(encoding="utf-8"))
            self.assertEqual(cfg["department_name"], "dow")
            self.assertEqual(cfg["service"], "USSF")


if __name__ == "__main__":
    unittest.main()
