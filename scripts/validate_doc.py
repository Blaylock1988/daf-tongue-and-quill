"""Multi-format DAFH 33-337 validator and extractor (.txt, .md, .docx, .pptx)."""
from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from tq_common import (
    StyleConfig,
    TRANSMITTAL_NOTICE,
    build_effective_lexicon,
    configure_utf8_stdout,
    load_style_config,
    load_style_rules,
)

ALL_TYPES = (
    "auto",
    "memo",
    "mfr",
    "indorsement",
    "esss",
    "bbp",
    "position",
    "talking",
    "point",
    "staff-study",
    "bio",
    "letter",
    "resume",
    "minutes",
    "slides",
)

SSN_RE = re.compile(r"\b\d{3}-\d{2}-\d{4}\b")
EDIPI_RE = re.compile(r"\b(?:DOD\s*ID|EDIPI)\s*[:#]?\s*\d{9,10}\b", re.IGNORECASE)
CLASSIFIED_LINE_RE = re.compile(
    r"^\s*(?:TOP\s+SECRET|SECRET|CONFIDENTIAL)(?://[A-Z0-9/, \-]+)?\s*$"
)
CLASSIFIED_PORTION_RE = re.compile(r"\((?:TS|S)(?://[A-Z0-9/, \-]+)?\)")
FOUO_RE = re.compile(r"\b(?:FOUO|FOR\s+OFFICIAL\s+USE\s+ONLY)\b", re.IGNORECASE)
UNCLASS_CUI_RE = re.compile(r"\bUNCLASSIFIED\s*//\s*CUI\b|\bU//CUI\b", re.IGNORECASE)
MIXED_DATE_RE = re.compile(
    r"\b\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\.?\s+\d{4}\b"
)
PASSIVE_RE = re.compile(r"\b(?:am|is|are|was|were|be|been|being)\s+[a-z]+(?:ed|en)\b", re.IGNORECASE)
FIRST_DEF_RE = re.compile(r"\b([A-Za-z][A-Za-z0-9&,'\- ]{3,80})\s+\(([A-Z][A-Za-z0-9&/]{1,12})\)")
ACRONYM_TOKEN_RE = re.compile(r"\b([A-Z][A-Z0-9&/]{1,11})(?:s)?\b")
ABBREV_MONTH_RE = re.compile(r"\b(?:Jan|Feb|Aug|Sept|Oct|Nov|Dec)\.\s+\d{1,2}")
DECIMAL_ZERO_RE = re.compile(r"\b\d+\.00\b")
HYPHEN_WRAP_RE = re.compile(r"[A-Za-z]-\s*$")

EXEMPT_UPPER_TOKENS = frozenset({
    "MEMORANDUM", "FOR", "FROM", "SUBJECT", "REFERENCES", "REFERENCE",
    "RECORD", "IN", "TURN", "THROUGH", "ATTENTION", "ATTN", "COMMANDER",
    "ATTACHMENT", "ATTACHMENTS", "DISTRIBUTION", "CC", "SIGNED", "CUI",
    "SUSPENSE", "SUSP", "ON", "THE", "OF", "AND", "OR", "TO", "BY",
    "DEPARTMENT", "AIR", "FORCE", "SPACE", "UNITED", "STATES", "USA",
    "US", "USAF", "USSF", "DAF", "NLT", "IAW", "DSN", "COMM", "NOTAL",
    "PURPOSE", "BACKGROUND", "DISCUSSION", "VIEWS", "OTHERS",
    "RECOMMENDATION", "CONCLUSION", "PROBLEM", "FACTORS", "ACTION",
    "RECOMMENDED", "EDUCATION", "ASSIGNMENTS", "MAJOR", "AWARDS",
    "DECORATIONS", "EFFECTIVE", "DATES", "PROMOTION", "OFFICIAL",
    "PHOTOGRAPH", "BULLET", "PAPER", "POSITION", "TALKING", "POINT",
    "ELECTRONIC", "STAFF", "SUMMARY", "SHEET", "MINUTES", "AGENDA",
    "DO", "NOT", "II", "III", "IV", "VI", "VII", "VIII", "IX", "XI",
    "FED", "ONLY", "FEDCON", "NOCON", "DL", "NOFORN", "REL", "RELIDO",
    "PO", "BOX", "NW", "NE", "SW", "SE", "VA", "DC", "TX", "CA", "FL",
    "OK", "CO", "OH", "NC", "IL", "AL", "AK", "AZ", "GA", "HI", "NM",
    "NV", "NY", "UT", "WA", "APO", "FPO", "DPO", "AE", "AP", "AA",
    "HQ", "AFB", "SFB", "SFS", "JB", "AB", "SQ", "WG", "GP", "FW",
    "OG", "XP", "DO", "CC", "CD", "CV", "JA", "PA", "FM", "FSS",
    "MAJCOM", "FLDCOM", "GS", "SES", "CY", "FY", "AM", "PM",
})


@dataclass
class Issue:
    """Single validation finding emitted by validate_doc.py."""

    severity: str
    code: str
    message: str
    location: str = ""
    suggestion: str = ""


@dataclass
class SlideData:
    """Extracted slide geometry, text lines, tables, charts, and speaker notes."""

    index: int
    title: str = ""
    lines: list[str] = field(default_factory=list)
    font_families: set[str] = field(default_factory=set)
    body_font_sizes_pt: list[float] = field(default_factory=list)
    title_font_sizes_pt: list[float] = field(default_factory=list)
    tables: list[tuple[int, int]] = field(default_factory=list)
    table_texts: list[str] = field(default_factory=list)
    notes: str = ""


@dataclass
class ExtractedDocument:
    """Normalized structural extraction across .txt, .md, .docx, and .pptx."""

    file_path: Path
    file_kind: str
    paragraphs: list[str] = field(default_factory=list)
    table_lines: list[str] = field(default_factory=list)
    header_lines: list[str] = field(default_factory=list)
    footer_lines: list[str] = field(default_factory=list)
    margins_in: dict[str, float] = field(default_factory=dict)
    font_names: set[str] = field(default_factory=set)
    body_font_sizes_pt: list[float] = field(default_factory=list)
    page1_has_page_num: bool = False
    sig_keep_together: bool | None = None
    slides: list[SlideData] = field(default_factory=list)
    slide_width_in: float = 0.0
    slide_height_in: float = 0.0

    @property
    def all_text(self) -> str:
        parts = (
            self.header_lines
            + self.paragraphs
            + self.table_lines
            + self.footer_lines
            + [ln for s in self.slides for ln in s.lines]
            + [s.notes for s in self.slides if s.notes]
        )
        return "\n".join(p for p in parts if p)


def _extract_plaintext(path: Path) -> ExtractedDocument:
    """Extract paragraphs or Markdown slide outlines from .txt or .md."""
    raw = path.read_text(encoding="utf-8")
    lines = raw.splitlines()
    if any(re.match(r"^#{1,3}\s+Slide\s+\d+", ln, re.IGNORECASE) for ln in lines):
        slides: list[SlideData] = []
        current: SlideData | None = None
        for ln in lines:
            s_match = re.match(r"^#{1,3}\s+Slide\s+\d+\s*[:\-]?\s*(.*)$", ln, re.IGNORECASE)
            if s_match:
                current = SlideData(index=len(slides) + 1, title=s_match.group(1).strip())
                slides.append(current)
                continue
            if current is None:
                continue
            clean = ln.strip()
            if clean.lower().startswith("notes:"):
                current.notes = clean.split(":", 1)[1].strip()
            elif clean:
                current.lines.append(re.sub(r"^[-*•]\s+", "", clean))
        return ExtractedDocument(file_path=path, file_kind="md-slides", paragraphs=lines, slides=slides)
    return ExtractedDocument(file_path=path, file_kind="text", paragraphs=lines)


def _extract_docx(path: Path) -> ExtractedDocument:
    """Extract paragraphs, headers, footers, margins, fonts, and tables from .docx."""
    import docx

    doc = docx.Document(str(path))
    ext = ExtractedDocument(file_path=path, file_kind="docx")
    if doc.sections:
        sec = doc.sections[0]
        ext.margins_in = {
            "top": round(sec.top_margin.inches, 2),
            "bottom": round(sec.bottom_margin.inches, 2),
            "left": round(sec.left_margin.inches, 2),
            "right": round(sec.right_margin.inches, 2),
        }
        for hdr in (sec.first_page_header, sec.header):
            ext.header_lines.extend(p.text.strip() for p in hdr.paragraphs if p.text.strip())
            for tbl in hdr.tables:
                ext.header_lines.extend(
                    c.text.strip() for row in tbl.rows for c in row.cells if c.text.strip()
                )
        for ftr in (sec.first_page_footer, sec.footer):
            ext.footer_lines.extend(p.text.strip() for p in ftr.paragraphs if p.text.strip())
        if "PAGE" in sec.first_page_header._element.xml or "PAGE" in sec.first_page_footer._element.xml:
            ext.page1_has_page_num = True

    def inspect_para(p: Any) -> None:
        text = p.text.rstrip()
        if text:
            ext.paragraphs.append(text)
            if "SIGNED" in text or re.search(r",\s*(?:USAF|USSF|USA|USN|USMC|DAF)", text):
                ext.sig_keep_together = bool(p.paragraph_format.keep_together)
        for r in p.runs:
            if r.font.name:
                ext.font_names.add(r.font.name)
            if r.font.size:
                ext.body_font_sizes_pt.append(round(r.font.size.pt, 1))

    for p in doc.paragraphs:
        inspect_para(p)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                ctext = cell.text.strip()
                if ctext:
                    ext.table_lines.append(ctext)
                for p in cell.paragraphs:
                    for r in p.runs:
                        if r.font.name:
                            ext.font_names.add(r.font.name)
                        if r.font.size:
                            ext.body_font_sizes_pt.append(round(r.font.size.pt, 1))
    return ext


def _read_slide_notes(zf: zipfile.ZipFile, slide_num: int) -> str:
    """Resolve and extract speaker notes text for a given slide via stdlib XML."""
    rels_name = f"ppt/slides/_rels/slide{slide_num}.xml.rels"
    if rels_name not in zf.namelist():
        return ""
    rels_root = ET.fromstring(zf.read(rels_name))
    for rel in rels_root:
        if "notesSlide" in rel.attrib.get("Type", ""):
            target = rel.attrib.get("Target", "").lstrip("/")
            if target.startswith("../"):
                target = "ppt/" + target[3:]
            if not target.startswith("ppt/"):
                target = "ppt/slides/" + target
            if target in zf.namelist():
                notes_root = ET.fromstring(zf.read(target))
                texts = [t.text for t in notes_root.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}t") if t.text]
                return " ".join(texts).strip()
    return ""


def _extract_pptx(path: Path) -> ExtractedDocument:
    """Extract slides, shapes, font metrics, tables, and notes via stdlib zipfile/XML."""
    ns_a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    ns_p = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    ext = ExtractedDocument(file_path=path, file_kind="pptx")

    with zipfile.ZipFile(str(path), "r") as zf:
        if "ppt/presentation.xml" in zf.namelist():
            pres_root = ET.fromstring(zf.read("ppt/presentation.xml"))
            sld_sz = pres_root.find(f".//{ns_p}sldSz")
            if sld_sz is not None:
                ext.slide_width_in = round(int(sld_sz.attrib.get("cx", 0)) / 914400.0, 3)
                ext.slide_height_in = round(int(sld_sz.attrib.get("cy", 0)) / 914400.0, 3)

        slide_names = [
            n for n in zf.namelist()
            if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)
        ]
        slide_names.sort(key=lambda n: int(re.search(r"(\d+)", n).group(1)))

        for s_name in slide_names:
            s_num = int(re.search(r"(\d+)", s_name).group(1))
            s_root = ET.fromstring(zf.read(s_name))
            sd = SlideData(index=s_num)
            for sp in s_root.findall(f".//{ns_p}sp"):
                for para in sp.findall(f".//{ns_a}p"):
                    runs_text: list[str] = []
                    for r in para.findall(f"{ns_a}r"):
                        r_pr = r.find(f"{ns_a}rPr")
                        sz_pt = None
                        if r_pr is not None:
                            if "sz" in r_pr.attrib:
                                sz_pt = round(int(r_pr.attrib["sz"]) / 100.0, 1)
                            latin = r_pr.find(f"{ns_a}latin")
                            if latin is not None and "typeface" in latin.attrib:
                                sd.font_families.add(latin.attrib["typeface"])
                                ext.font_names.add(latin.attrib["typeface"])
                        t_el = r.find(f"{ns_a}t")
                        if t_el is not None and t_el.text:
                            runs_text.append(t_el.text)
                            if sz_pt is not None:
                                if sz_pt >= 30 and not sd.title:
                                    sd.title_font_sizes_pt.append(sz_pt)
                                elif sz_pt >= 12 and not t_el.text.strip().startswith(("CUI", "Controlled", "POC:")):
                                    sd.body_font_sizes_pt.append(sz_pt)
                    joined = "".join(runs_text).strip()
                    if joined:
                        if not sd.title and not joined.startswith(("CUI", "-")):
                            sd.title = joined
                        sd.lines.append(joined)
            for tbl in s_root.findall(f".//{ns_a}tbl"):
                rows = tbl.findall(f"{ns_a}tr")
                cols = max((len(r.findall(f"{ns_a}tc")) for r in rows), default=0)
                sd.tables.append((len(rows), cols))
                for t_el in tbl.iter(f"{ns_a}t"):
                    if t_el.text and t_el.text.strip():
                        sd.table_texts.append(t_el.text.strip())
            sd.notes = _read_slide_notes(zf, s_num)
            ext.slides.append(sd)
    return ext


def extract_document(path: Path) -> ExtractedDocument:
    """Dispatch extraction by file extension."""
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return _extract_docx(path)
    if suffix == ".pptx":
        return _extract_pptx(path)
    return _extract_plaintext(path)


def detect_doc_type(ext: ExtractedDocument) -> str:
    """Infer T&Q document type from extracted structure and markers."""
    if ext.file_kind in {"pptx", "md-slides"}:
        return "slides"
    upper = ext.all_text.upper()
    if "ELECTRONIC STAFF SUMMARY SHEET" in upper:
        return "esss"
    if "BULLET BACKGROUND PAPER" in upper:
        return "bbp"
    if "POSITION PAPER" in upper:
        return "position"
    if "TALKING PAPER" in upper:
        return "talking"
    if "POINT PAPER" in upper:
        return "point"
    if "EFFECTIVE DATES OF PROMOTION" in upper or "CURRENT AS OF" in upper:
        return "bio"
    if "MEMORANDUM FOR" in upper:
        if re.search(r"MEMORANDUM FOR\s+RECORD", upper):
            return "mfr"
        if re.search(r"\b1ST\s+IND|\b2D\s+IND|\b3D\s+IND|\b4TH\s+IND", upper):
            return "indorsement"
        if "MINUTES" in upper and "SUBJECT:" in upper:
            return "minutes"
        return "memo"
    if "DEAR " in upper and ("SINCERELY" in upper or "RESPECTFULLY YOURS" in upper):
        return "letter"
    return "memo"


def _check_security_pii_cui(
    ext: ExtractedDocument, doc_type: str, issues: list[Issue]
) -> None:
    """Enforce PII/classified hard stops and DoDI 5200.48 CUI marking rules."""
    text = ext.all_text
    if SSN_RE.search(text):
        issues.append(Issue("CRITICAL", "SEC-SSN", "Potential SSN pattern detected.", "Body", "Remove PII before processing."))
    if EDIPI_RE.search(text):
        issues.append(Issue("CRITICAL", "SEC-EDIPI", "Potential DoD ID / EDIPI detected.", "Body", "Remove PII before processing."))

    for line in text.splitlines():
        if CLASSIFIED_LINE_RE.match(line) or CLASSIFIED_PORTION_RE.search(line):
            issues.append(Issue("CRITICAL", "SEC-CLASSIFIED", f"Classified marking detected: {line.strip()[:40]}", "Security", "Never process classified information in this workflow."))
            break

    if FOUO_RE.search(text):
        issues.append(Issue("ERROR", "CUI-FOUO-OBSOLETE", "Obsolete FOUO marking detected.", "Marking", "Replace FOUO with DoDI 5200.48 CUI markings."))
    if UNCLASS_CUI_RE.search(text):
        issues.append(Issue("ERROR", "CUI-INVALID-BANNER", "Invalid 'UNCLASSIFIED//CUI' banner detected.", "Marking", "Use plain 'CUI' or 'CUI//[CATEGORY]//[LDC]'."))

    cui_banners = [ln.strip() for ln in (ext.header_lines + ext.footer_lines + ext.paragraphs) if re.match(r"^CUI(?://[A-Z0-9/, \-]+)?$", ln.strip())]
    if doc_type == "slides" and ext.slides:
        slide_banners = [
            [ln.strip() for ln in s.lines if re.match(r"^CUI(?://[A-Z0-9/, \-]+)?$", ln.strip())]
            for s in ext.slides
        ]
        if any(slide_banners):
            cui_banners = [b for sub in slide_banners for b in sub]
            missing = [s.index for s, b in zip(ext.slides, slide_banners) if len(b) < 2]
            if missing:
                issues.append(Issue("ERROR", "CUI-SLIDE-BANNER", f"Missing top/bottom CUI banners on slide(s): {missing}.", "Slides", "Include centered bold CUI banner at top and bottom of every slide."))

    if not cui_banners:
        return

    rules = load_style_rules()
    valid_ldcs = set(rules.get("cui_ldcs", {}).keys())
    for b in set(cui_banners):
        parts = b.split("//")
        if len(parts) == 3:
            ldc = parts[2].strip()
            if ldc not in valid_ldcs and not ldc.startswith("REL TO "):
                issues.append(Issue("ERROR", "CUI-BAD-LDC", f"Unrecognized CUI Limited Dissemination Control '{ldc}' in banner '{b}'.", "Banner", "Use DoDI 5200.48 LDCs (FED ONLY, FEDCON, NOCON, DL ONLY, NOFORN, REL TO, RELIDO)."))

    has_transmittal = TRANSMITTAL_NOTICE.lower() in text.lower()
    if not has_transmittal:
        has_ctrl = "Controlled by:" in text
        has_cat = "CUI Category:" in text
        has_poc = "POC:" in text
        if not (has_ctrl and has_cat and has_poc):
            issues.append(Issue("ERROR", "CUI-MISSING-DI", "CUI document lacks complete Page 1 / Slide 1 Designation Indicator block (Controlled by / CUI Category / POC).", "DI Block", "Add DI block or transmittal notice per DoDI 5200.48."))


def _check_dow_dod_consistency(
    ext: ExtractedDocument, dept_mode: str, doc_type: str, issues: list[Issue]
) -> None:
    """Validate Department of War vs. Department of Defense internal consistency."""
    cleaned = re.sub(r"\bDoD[DIM]\s+\d+(?:\.\d+)*\b", "", ext.all_text)
    war_hits = re.findall(r"\b(?:Department of War|Secretary of War|Deputy Secretary of War|SecWar|DoW|DOW)\b", cleaned)
    def_hits = re.findall(r"\b(?:Department of Defense|Secretary of Defense|Deputy Secretary of Defense|SecDef|DoD|DOD)\b", cleaned)

    if dept_mode == "auto" and war_hits and def_hits:
        issues.append(Issue("WARNING", "DEPT-MIXED", "Document mixes Department of War (DoW) and Department of Defense (DoD) terminology.", "Terminology", "Choose one departmental designation unless citing statutory/legacy directives."))
    elif dept_mode == "dow" and def_hits:
        issues.append(Issue("ERROR", "DEPT-MISMATCH", f"Configured --dept dow but found Defense term(s): {def_hits[0]}.", "Terminology", "Update to Department of War / Secretary of War / DoW."))
    elif dept_mode == "dod" and war_hits:
        issues.append(Issue("ERROR", "DEPT-MISMATCH", f"Configured --dept dod but found War term(s): {war_hits[0]}.", "Terminology", "Update to Department of Defense / Secretary of Defense / DoD."))

    if doc_type == "bio" and re.search(r"\bDoW\b", cleaned):
        issues.append(Issue("WARNING", "BIO-DOW-CASE", "Official Biography uses 'DoW'; war.gov AP brand style requires uppercase 'DOW'.", "Bio Style", "Replace 'DoW' with 'DOW' in AP-style public biographies."))


def _check_attachments_and_acronyms(
    ext: ExtractedDocument,
    doc_type: str,
    style_cfg: StyleConfig,
    style_cfg_path: str | None,
    issues: list[Issue],
) -> None:
    """Verify attachment parity and separate acronym vs. abbreviation compliance."""
    lines = ext.paragraphs
    full_text = "\n".join(lines)
    if doc_type in {"memo", "mfr", "indorsement", "letter", "minutes"}:
        declared_count = None
        listed_items = 0
        for idx, ln in enumerate(lines):
            m_multi = re.match(r"^\s*(\d+)\s+Attachments?\s*:\s*$", ln)
            m_single = re.match(r"^\s*Attachment\s*:\s*(.*)$", ln)
            if m_multi:
                declared_count = int(m_multi.group(1))
                for follow in lines[idx + 1 :]:
                    s = follow.strip()
                    if not s or re.match(r"^(?:cc:|DISTRIBUTION:|Controlled by:|When enclosure)", s, re.IGNORECASE):
                        break
                    listed_items += 1
                break
            if m_single:
                declared_count = 1
                if m_single.group(1).strip():
                    listed_items = 1
                else:
                    for follow in lines[idx + 1 :]:
                        if not follow.strip() or follow.strip().startswith(("cc:", "DISTRIBUTION:")):
                            break
                        listed_items += 1
                break
        if declared_count is not None and declared_count != listed_items:
            issues.append(Issue("ERROR", "ATCH-COUNT-MISMATCH", f"Attachment header declares {declared_count} item(s), but {listed_items} listed.", "Attachments", "Match '# Attachments:' count to the exact numbered items."))
        cited_nums = [int(n) for n in re.findall(r"\b(?:Atch|Attachment)\s+(\d+)\b", full_text, re.IGNORECASE)]
        effective_count = min(declared_count, listed_items) if declared_count is not None else listed_items
        if cited_nums and max(cited_nums) > effective_count:
            issues.append(Issue("ERROR", "ATCH-MISSING-REF", f"Body references Atch {max(cited_nums)}, but only {effective_count} attachment(s) listed.", "Attachments", "List every cited attachment in the closing block."))

    lexicon = build_effective_lexicon(style_cfg, style_cfg_path)
    is_formal = doc_type in {"memo", "mfr", "indorsement", "letter", "position", "staff-study", "bio", "minutes"}
    if is_formal:
        for abbr, info in lexicon.abbreviations.items():
            if info.get("formal_allowed", True):
                continue
            pat = re.escape(abbr) if "/" in abbr else rf"\b{re.escape(abbr)}\b"
            if re.search(pat, full_text):
                issues.append(Issue("WARNING", "ABBR-FORMAL", f"Informal abbreviation/truncation '{abbr}' found in formal prose.", "Abbreviations", f"Spell out '{info.get('meaning', abbr)}' in formal correspondence."))
        for m_verb in re.finditer(r"\bcoord(?:s|es)\b", full_text, re.IGNORECASE):
            issues.append(Issue("ERROR", "ABBR-VERB-SUFFIX", f"Invalid verb plural suffix on abbreviation: '{m_verb.group(0)}'.", "Abbreviations", "Spell out 'coordinates' or use valid past/possessive only where approved."))

    meta_prefix = r"^(?:MEMORANDUM FOR|FROM:|SUBJECT:|References?:|//SIGNED//|FOR THE COMMANDER|Controlled by:|CUI Category:|Distribution|POC:|Attachment|When enclosure|CUI(?://[A-Z0-9/, \-]+)?$)"
    if doc_type in {"slides"}:
        body_lines = [
            ln
            for s in ext.slides
            for ln in s.lines
            if not re.match(meta_prefix, ln.strip(), re.IGNORECASE)
        ]
        raw_body = "\n".join(body_lines)
    else:
        body_lines = [ln for ln in lines if not re.match(meta_prefix, ln.strip(), re.IGNORECASE)]
        raw_body = "\n".join(ln for ln in body_lines if not re.search(r",\s*(?:USAF|USSF|USA|USN|USMC|DAF)", ln))
    # Strip numbered-unit office symbols (e.g., '1 FW/XP') and dates ('21 Nov 26') so office symbols aren't flagged as acronyms.
    body_text = re.sub(r"\b\d+\s+[A-Z0-9]{1,6}/[A-Z0-9]{1,6}\b", "", raw_body)

    defs: dict[str, list[int]] = {}
    for m in FIRST_DEF_RE.finditer(body_text):
        acr = m.group(2)
        defs.setdefault(acr, []).append(m.start())
        if len(defs[acr]) > 1:
            issues.append(Issue("WARNING", "ACR-DUP-DEF", f"Acronym '{acr}' is defined {len(defs[acr])} times.", "Acronyms", "Define an acronym once at first use."))

    for acr, positions in defs.items():
        all_uses = [m.start() for m in re.finditer(rf"\b{re.escape(acr)}\b", body_text)]
        if len(all_uses) == 1 and acr not in {"POC", "ETP", "CUI"}:
            issues.append(Issue("WARNING", "ACR-SINGLE-USE", f"Acronym '{acr}' is defined once and never reused (AFH 33-337 p. 334).", "Acronyms", f"Spell out the term cleanly without '({acr})'."))
        elif all_uses and all_uses[0] < positions[0]:
            issues.append(Issue("WARNING", "ACR-USED-BEFORE-DEF", f"Acronym '{acr}' is used before it is defined.", "Acronyms", "Spell out and define at its very first occurrence."))

    used_tokens = {
        m.group(1)
        for m in ACRONYM_TOKEN_RE.finditer(body_text)
        if m.group(1) not in EXEMPT_UPPER_TOKENS
    }
    for tok in sorted(used_tokens):
        if tok in defs:
            continue
        if tok in lexicon.collisions and tok not in style_cfg.acronym_disambiguations:
            meanings = " / ".join(lexicon.collisions[tok])
            issues.append(Issue("WARNING", "ACR-COLLISION", f"Multi-meaning acronym '{tok}' used without definition ({meanings}).", "Acronyms", f"Define '{tok}' at first use or pin meaning in unit_style.json."))
            continue
        if lexicon.mode == "strict" and tok not in lexicon.acronyms:
            issues.append(Issue("ERROR", "ACR-STRICT-UNAPPROVED", f"Acronym '{tok}' is not on the AFPC or strict command list and is undefined.", "Acronyms", "Spell out on first use or use an approved command acronym."))
        elif tok not in lexicon.acronyms:
            issues.append(Issue("WARNING", "ACR-UNDEFINED", f"Acronym '{tok}' appears undefined and outside AFPC baseline.", "Acronyms", "Spell out on first use: Full Term (ACR)."))


def _check_readability_and_deai(
    ext: ExtractedDocument, doc_type: str, issues: list[Issue]
) -> None:
    """Enforce AFH 33-337 readability targets, passive voice, Ch. 7 deadwood, and De-AI."""
    rules = load_style_rules()
    prose = "\n".join(ext.paragraphs)
    if doc_type in {"memo", "mfr", "indorsement", "esss", "position", "staff-study", "bio", "letter", "minutes"}:
        body_paras = [
            p.strip()
            for p in ext.paragraphs
            if len(p.split()) >= 6
            and not re.match(r"^(?:MEMORANDUM FOR|FROM:|SUBJECT:|References?:|Attachment|cc:|DISTRIBUTION:|Controlled by:|CUI)", p.strip())
        ]
        sentences = [s.strip() for p in body_paras for s in re.split(r"[.!?]+(?:\s+|$)", p) if len(s.split()) >= 3]
        if sentences:
            avg_words = sum(len(s.split()) for s in sentences) / len(sentences)
            if avg_words > 20.0:
                issues.append(Issue("WARNING", "STYLE-SENT-LEN", f"Average sentence length is {avg_words:.1f} words (AFH 33-337 p. 89 target <=20, avg 17).", "Readability", "Split multi-clause sentences and lead with active verbs."))
        for p in body_paras:
            s_count = len([s for s in re.split(r"[.!?]+(?:\s+|$)", p) if s.strip()])
            if s_count > 7:
                issues.append(Issue("WARNING", "STYLE-PARA-LEN", f"Paragraph has {s_count} sentences (AFH 33-337 p. 69 limit: <=7).", "Readability", "Break into shorter 3-7 sentence paragraphs."))

    lower_text = ext.all_text.lower()
    for bucket, code in (
        ("wordy_phrases", "STYLE-WORDY"),
        ("smothered_verbs", "STYLE-SMOTHERED"),
        ("redundancies", "STYLE-REDUNDANT"),
        ("ai_buzzwords", "STYLE-AI-BUZZWORD"),
    ):
        for phrase, repl in rules.get(bucket, {}).items():
            if re.search(rf"\b{re.escape(phrase.lower())}\b", lower_text):
                issues.append(Issue("WARNING", code, f"Avoid '{phrase}'.", "Style", f"Replace with: {repl}"))

    passive_hits = PASSIVE_RE.findall(prose)
    if passive_hits:
        issues.append(Issue("INFO", "STYLE-PASSIVE", f"Detected {len(passive_hits)} passive voice construction(s) (e.g., '{passive_hits[0]}').", "Active Voice", "Put the actor before the verb per AFH 33-337 p. 74."))


def _check_format_rules(
    ext: ExtractedDocument, doc_type: str, issues: list[Issue]
) -> None:
    """Validate format-specific geometry and structural rules per document type."""
    text = ext.all_text
    lines = [ln.strip() for ln in ext.paragraphs if ln.strip()]

    if MIXED_DATE_RE.search(text) and doc_type != "bio":
        issues.append(Issue("ERROR", "FMT-MIXED-DATE", "Mixed abbreviated month + 4-digit year detected (e.g., '21 Nov 2026').", "Date", "Use '21 November 2026' or '21 Nov 26' (AFH 33-337 p. 164)."))

    if doc_type in {"memo", "mfr", "indorsement", "minutes", "staff-study"}:
        if re.search(r"^\s*MEMORANDUM FOR:", text, re.MULTILINE):
            issues.append(Issue("ERROR", "MEMO-COLON", "'MEMORANDUM FOR:' must not have a colon.", "Heading", "Remove the colon and leave two spaces after MEMORANDUM FOR."))
        elif "MEMORANDUM FOR" not in text:
            issues.append(Issue("ERROR", "MEMO-NO-FOR", "Missing 'MEMORANDUM FOR' element.", "Heading", "Include MEMORANDUM FOR  <ORG/SYMBOL>."))
        if doc_type != "mfr" and not re.search(r"^\s*FROM:", text, re.MULTILINE):
            issues.append(Issue("ERROR", "MEMO-NO-FROM", "Missing 'FROM:' element.", "Heading", "Include FROM:  <ORG/SYMBOL>."))
        if not re.search(r"^\s*SUBJECT:", text, re.MULTILINE):
            issues.append(Issue("ERROR", "MEMO-NO-SUBJECT", "Missing 'SUBJECT:' element.", "Heading", "Include SUBJECT:  <Title Case Subject>."))
        if re.search(r"\bActing\s+(?:Wing\s+|Group\s+|Squadron\s+|Field\s+|Delta\s+)?Commander\b", text, re.IGNORECASE):
            issues.append(Issue("ERROR", "SIG-ACTING-CMD", "'Acting' is prohibited for command positions (AFH 33-337 p. 171).", "Signature", "Sign under permanent title or designated Commander authority."))

        for idx, ln in enumerate(lines):
            if re.search(r",\s*(?:USAF|USSF|USA|USN|USMC|DAF)", ln):
                name_part = ln.split(",")[0].strip()
                if name_part != name_part.upper():
                    issues.append(Issue("ERROR", "SIG-NAME-CAPS", f"Signer name must be ALL CAPS (rank/duty title remain Mixed Case): '{name_part}'.", "Signature", "Capitalize only FIRST M. LAST (and service)."))
                sig_count = 1
                for follow in lines[idx + 1 :]:
                    if re.match(r"^(?:\d+\s+Attachments?:|Attachment:|cc:|DISTRIBUTION:|Controlled by:|//|Approved)", follow, re.IGNORECASE):
                        break
                    sig_count += 1
                if sig_count > 4:
                    issues.append(Issue("ERROR", "SIG-TOO-MANY-LINES", f"Signature block has {sig_count} lines (hard max 4; target 3).", "Signature", "Condense duty title/org to <=4 lines."))
                break

        if ext.file_kind == "docx":
            if ext.margins_in and any(v < 0.98 for v in ext.margins_in.values()):
                issues.append(Issue("ERROR", "DOCX-MARGINS", f"Margins below 1.0 in: {ext.margins_in}", "Layout", "Maintain >=1.0 inch margins per AFH 33-337 p. 163."))
            if ext.font_names and "Times New Roman" not in ext.font_names:
                issues.append(Issue("ERROR", "DOCX-FONT", f"Expected Times New Roman body font, found: {sorted(ext.font_names)}", "Typography", "Use 12-pt Times New Roman."))
            if ext.page1_has_page_num:
                issues.append(Issue("ERROR", "DOCX-PAGE1-NUM", "Page 1 must never be numbered on Official Memorandums.", "Pagination", "Suppress numbering on Page 1 via different_first_page_header_footer."))
            if ext.sig_keep_together is False:
                issues.append(Issue("WARNING", "DOCX-SIG-ORPHAN", "Signature paragraph lacks keep_together=True.", "Pagination", "Set keep_together=True to prevent orphan signatures."))

    elif doc_type == "bio":
        if re.search(r"\b(?:wife|husband|spouse|married|children|son|daughter)\b", text, re.IGNORECASE):
            issues.append(Issue("ERROR", "BIO-PRIVACY-FAMILY", "Family/marital terms detected in Official Biography (prohibited by Ch. 20 / Privacy Act).", "Privacy", "Remove all family and marital references."))
        if re.search(r"\bover\s+\d[\d,]*\s+(?:flying\s+)?hours\b", text, re.IGNORECASE):
            issues.append(Issue("WARNING", "BIO-MORE-THAN", "Use 'more than' instead of 'over' for flight hours (p. 262).", "Bio Style", "Replace 'over' with 'more than'."))
        if "(Current as of" not in text:
            issues.append(Issue("WARNING", "BIO-CURRENT-AS-OF", "Missing '(Current as of Month Year)' footer.", "Bio Structure", "Add '(Current as of Month YYYY)' one line below promotions."))
        assign_match = re.search(r"ASSIGNMENTS\s*(.*?)(?:SUMMARY OF JOINT|FLIGHT INFORMATION|MAJOR AWARDS|EFFECTIVE DATES|\Z)", text, re.DOTALL | re.IGNORECASE)
        if assign_match and ABBREV_MONTH_RE.search(assign_match.group(1)):
            issues.append(Issue("ERROR", "BIO-ASSIGN-MONTHS", "Months in ASSIGNMENTS must be fully spelled out (AFH 33-337 p. 261).", "Bio Style", "Spell out all month names in ASSIGNMENTS."))
        if ext.file_kind == "docx" and ext.font_names and "Arial" not in ext.font_names:
            issues.append(Issue("ERROR", "BIO-FONT", f"Official Biography requires Arial 9/13.5 pt (found {sorted(ext.font_names)}).", "Typography", "Switch biography text to Arial 9 pt / 13.5 pt Bold."))

    elif doc_type in {"bbp", "talking", "point", "position"}:
        if "ON" not in lines[:5]:
            issues.append(Issue("WARNING", "PAPER-TITLE", "Expected 3-line ALL CAPS title '<TYPE> / ON / <TOPIC>'.", "Header", "Format centered 3-line ALL CAPS title with 'ON' on line 2."))

    elif doc_type == "letter":
        if re.search(r"^\s*Dear\s+.+[:,]\s*$", text, re.MULTILINE):
            issues.append(Issue("ERROR", "LETTER-OPEN-PUNCT", "Personal Letter salutation must use open punctuation (no colon or comma).", "Salutation", "Remove trailing punctuation after 'Dear ...'."))
        if re.search(r"^\s*(?:Sincerely|Respectfully yours)[:,]\s*$", text, re.MULTILINE):
            issues.append(Issue("ERROR", "LETTER-CLOSE-PUNCT", "Complimentary close must use open punctuation (no comma or colon).", "Closing", "Remove trailing punctuation after 'Sincerely'."))
        if "FOR THE COMMANDER" in text.upper():
            issues.append(Issue("ERROR", "LETTER-AUTH-LINE", "Personal Letters never include an authority line ('FOR THE COMMANDER').", "Closing", "Remove FOR THE COMMANDER."))

    elif doc_type == "slides":
        for s in ext.slides:
            content_lines = [
                ln for ln in s.lines
                if ln != s.title and not ln.startswith(("CUI", "Controlled by:", "CUI Category:", "Distribution/", "POC:"))
            ]
            if len(content_lines) > 7:
                issues.append(Issue("ERROR", "SLIDE-7X7-LINES", f"Slide {s.index} has {len(content_lines)} body lines (max 7 per AFH 33-337 p. 119).", f"Slide {s.index}", "Split onto a '(Continued)' slide."))
            for ln in content_lines:
                clean = ln.lstrip("-• ").strip()
                w_count = len(clean.split())
                if w_count > 7:
                    issues.append(Issue("ERROR", "SLIDE-7X7-WORDS", f"Slide {s.index} bullet exceeds 7 words ({w_count}): '{clean[:45]}'", f"Slide {s.index}", "Condense to <=7 words and move details to speaker notes."))
                if len(clean) >= 10 and clean == clean.upper() and any(c.isalpha() for c in clean):
                    issues.append(Issue("WARNING", "SLIDE-ALL-CAPS", f"Slide {s.index} uses ALL CAPS body text ('{clean[:35]}').", f"Slide {s.index}", "Use mixed case so acronyms stand out (p. 119)."))
                if HYPHEN_WRAP_RE.search(clean):
                    issues.append(Issue("WARNING", "SLIDE-HYPHEN-WRAP", f"Slide {s.index} line ends with a hyphen.", f"Slide {s.index}", "Avoid hyphenation at line ends."))
            for r_cnt, c_cnt in s.tables:
                if r_cnt > 4 or c_cnt > 7:
                    issues.append(Issue("ERROR", "SLIDE-TABLE-SIZE", f"Slide {s.index} table is {r_cnt}x{c_cnt} (limit <=4 rows x 7 columns, p. 120).", f"Slide {s.index}", "Simplify table or convert to a chart."))
            for txt in content_lines + s.table_texts:
                if DECIMAL_ZERO_RE.search(txt):
                    issues.append(Issue("WARNING", "SLIDE-DECIMAL-ZERO", f"Slide {s.index} uses trailing '.00' decimal ('{txt[:35]}').", f"Slide {s.index}", "Drop '.00' so numbers aren't misread (p. 120)."))
            for pt in s.body_font_sizes_pt:
                if pt < 20.0:
                    issues.append(Issue("ERROR", "SLIDE-FONT-SMALL", f"Slide {s.index} body text is {pt} pt (AFH 33-337 p. 119 requires 20-26 pt).", f"Slide {s.index}", "Increase body font to 20-26 pt."))
            if not s.notes and ext.file_kind == "pptx":
                issues.append(Issue("INFO", "SLIDE-NO-NOTES", f"Slide {s.index} has no speaker notes.", f"Slide {s.index}", "Add concise speaker notes one level below the slide bullets."))
        if ext.file_kind == "pptx" and len(ext.font_names) > 1:
            issues.append(Issue("WARNING", "SLIDE-MIXED-FONTS", f"Deck mixes multiple fonts: {sorted(ext.font_names)}.", "Typography", "Use one consistent font family throughout (p. 119)."))


def validate_file(
    path: Path,
    *,
    doc_type: str = "auto",
    dept_mode: str = "auto",
    style_config_path: str | None = None,
    strict: bool = False,
) -> dict[str, Any]:
    """Validate any supported .txt, .md, .docx, or .pptx deliverable."""
    style_cfg = load_style_config(style_config_path)
    ext = extract_document(path)
    resolved_type = detect_doc_type(ext) if doc_type == "auto" else doc_type
    resolved_dept = dept_mode if dept_mode != "auto" else (
        style_cfg.department_name if style_cfg.department_name in {"dow", "dod"} else "auto"
    )

    issues: list[Issue] = []
    _check_security_pii_cui(ext, resolved_type, issues)
    _check_dow_dod_consistency(ext, resolved_dept, resolved_type, issues)
    _check_attachments_and_acronyms(ext, resolved_type, style_cfg, style_config_path, issues)
    _check_readability_and_deai(ext, resolved_type, issues)
    _check_format_rules(ext, resolved_type, issues)

    blocking = [i for i in issues if i.severity in {"CRITICAL", "ERROR"}]
    warnings = [i for i in issues if i.severity == "WARNING"]
    passed = not blocking and (not strict or not warnings)
    return {
        "file": str(path),
        "file_kind": ext.file_kind,
        "doc_type": resolved_type,
        "dept_mode": resolved_dept,
        "passed": passed,
        "counts": {
            "critical": sum(1 for i in issues if i.severity == "CRITICAL"),
            "error": sum(1 for i in issues if i.severity == "ERROR"),
            "warning": len(warnings),
            "info": sum(1 for i in issues if i.severity == "INFO"),
        },
        "issues": [asdict(i) for i in issues],
    }


def dump_extracted(path: Path) -> dict[str, Any]:
    """Export structured headers, body, tables, and slide notes for XO Red-Pen review."""
    ext = extract_document(path)
    return {
        "file": str(path),
        "file_kind": ext.file_kind,
        "inferred_type": detect_doc_type(ext),
        "margins_in": ext.margins_in,
        "font_names": sorted(ext.font_names),
        "headers": ext.header_lines,
        "footers": ext.footer_lines,
        "paragraphs": ext.paragraphs,
        "tables": ext.table_lines,
        "slides": [
            {
                "index": s.index,
                "title": s.title,
                "lines": s.lines,
                "tables": s.tables,
                "notes": s.notes,
            }
            for s in ext.slides
        ],
    }


def main(argv: list[str] | None = None) -> int:
    """CLI entrypoint for multi-format DAFH 33-337 linting and text extraction."""
    configure_utf8_stdout()
    ap = argparse.ArgumentParser(description="Validate .txt, .md, .docx, and .pptx against DAFH 33-337.")
    ap.add_argument("file", help="Path to draft or deliverable")
    ap.add_argument("--type", choices=ALL_TYPES, default="auto", dest="doc_type")
    ap.add_argument("--format", choices=("text", "json"), default="text", dest="out_format")
    ap.add_argument("--strict", action="store_true", help="Fail on WARNINGs as well as ERRORs")
    ap.add_argument("--dept", choices=("auto", "dow", "dod"), default="auto")
    ap.add_argument("--style-config", help="Optional unit_style.json path")
    ap.add_argument("--dump-text", action="store_true", help="Dump structured text/slides for XO Red-Pen review")
    args = ap.parse_args(argv)

    target = Path(args.file)
    if args.dump_text:
        dumped = dump_extracted(target)
        if args.out_format == "json":
            print(json.dumps(dumped, indent=2, ensure_ascii=False))
        else:
            print(f"=== {dumped['file']} ({dumped['file_kind']} / {dumped['inferred_type']}) ===")
            for h in dumped["headers"]:
                print(f"[HEADER] {h}")
            for p in dumped["paragraphs"]:
                print(p)
            for s in dumped["slides"]:
                print(f"\n--- Slide {s['index']}: {s['title']} ---")
                for ln in s["lines"]:
                    print(f"  {ln}")
                if s["notes"]:
                    print(f"  Notes: {s['notes']}")
        return 0

    report = validate_file(
        target,
        doc_type=args.doc_type,
        dept_mode=args.dept,
        style_config_path=args.style_config,
        strict=args.strict,
    )
    if args.out_format == "json":
        print(json.dumps(report, indent=2, ensure_ascii=False))
    else:
        status = "PASS" if report["passed"] else "FAIL"
        c = report["counts"]
        print(
            f"[{status}] {report['file']} ({report['doc_type']}) — "
            f"critical:{c['critical']} error:{c['error']} warning:{c['warning']} info:{c['info']}"
        )
        for item in report["issues"]:
            loc = f" [{item['location']}]" if item["location"] else ""
            print(f"  {item['severity']} ({item['code']}){loc}: {item['message']}")
            if item["suggestion"]:
                print(f"    -> {item['suggestion']}")
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
