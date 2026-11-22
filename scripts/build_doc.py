"""Programmatic AFH 33-337 Word (.docx) builder, template & schema generator."""
from __future__ import annotations

import argparse
import copy
import json
import sys
from pathlib import Path
from typing import Any

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor

from tq_common import (
    StyleConfig,
    TRANSMITTAL_NOTICE,
    build_cui_banner,
    build_cui_di_lines,
    configure_utf8_stdout,
    dept_terms,
    load_json,
    load_style_config,
    resolve_seal,
)

CLASSIC_MARKERS = ("1.", "a.", "(1)", "(a)", "1.", "a.", "[1]")

DOC_TYPES = (
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
    "agenda",
)

MEMO_PRESETS = (
    "standard",
    "appointment",
    "policy",
    "etp",
    "assumption-of-command",
    "aar",
)


def _add_page_number_field(paragraph: docx.text.paragraph.Paragraph) -> None:
    """Append a Word XML PAGE field to a header/footer paragraph."""
    run = paragraph.add_run()
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)
    fld_begin = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>')
    instr = parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> PAGE </w:instrText>')
    fld_sep = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>')
    fld_end = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_sep)
    run._r.append(fld_end)


def _set_cell_no_borders(cell: Any) -> None:
    """Remove all table borders from a cell."""
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        '<w:top w:val="none"/><w:left w:val="none"/>'
        '<w:bottom w:val="none"/><w:right w:val="none"/>'
        "</w:tcBorders>"
    )
    tc_pr.append(borders)


def _set_cell_thin_border(cell: Any, color: str = "444444") -> None:
    """Apply a clean single-line border around a placeholder cell."""
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="6" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="single" w:sz="6" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="single" w:sz="6" w:space="0" w:color="{color}"/>'
        "</w:tcBorders>"
    )
    tc_pr.append(borders)


def _add_styled_paragraph(
    container: Any,
    text: str = "",
    *,
    font_name: str = "Times New Roman",
    font_size: float = 12.0,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    align: WD_ALIGN_PARAGRAPH | None = None,
    left_indent_in: float = 0.0,
    first_line_indent_in: float = 0.0,
    space_before_pt: float = 0.0,
    space_after_pt: float = 0.0,
    line_spacing: float = 1.0,
    keep_together: bool = False,
    keep_with_next: bool = False,
    color_rgb: tuple[int, int, int] | None = None,
) -> docx.text.paragraph.Paragraph:
    """Create a paragraph with explicit run typography and Word layout flags."""
    p = container.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.left_indent = Inches(left_indent_in)
    pf.first_line_indent = Inches(first_line_indent_in)
    pf.space_before = Pt(space_before_pt)
    pf.space_after = Pt(space_after_pt)
    pf.line_spacing = line_spacing
    pf.keep_together = keep_together
    pf.keep_with_next = keep_with_next
    if text:
        run = p.add_run(text)
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.bold = bold
        run.italic = italic
        run.underline = underline
        if color_rgb is not None:
            run.font.color.rgb = RGBColor(*color_rgb)
    return p


def _configure_base_style(doc: docx.Document, font_name: str, font_size: float) -> None:
    """Set Normal style font defaults on the document."""
    normal = doc.styles["Normal"]
    normal.font.name = font_name
    normal.font.size = Pt(font_size)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing = 1.0


def _apply_margins(
    section: docx.section.Section,
    *,
    top: float = 1.0,
    bottom: float = 1.0,
    left: float = 1.0,
    right: float = 1.0,
    header_dist: float = 0.5,
    footer_dist: float = 0.5,
) -> None:
    """Apply inch margins and header/footer distances to a document section."""
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    section.top_margin = Inches(top)
    section.bottom_margin = Inches(bottom)
    section.left_margin = Inches(left)
    section.right_margin = Inches(right)
    section.header_distance = Inches(header_dist)
    section.footer_distance = Inches(footer_dist)


def _clear_empty_initial_paragraph(container: Any) -> None:
    """Remove python-docx default blank paragraph if empty."""
    if container.paragraphs and not container.paragraphs[0].text:
        p_el = container.paragraphs[0]._element
        p_el.getparent().remove(p_el)


def _apply_cui_headers_footers(
    section: docx.section.Section,
    cui_banner: str | None,
    *,
    page_number_cont: bool = True,
    cont_footer_note: str = "",
) -> None:
    """Configure Page 1 vs Page 2+ headers/footers with CUI banners and page numbers."""
    section.different_first_page_header_footer = True
    _clear_empty_initial_paragraph(section.first_page_header)
    _clear_empty_initial_paragraph(section.header)
    _clear_empty_initial_paragraph(section.first_page_footer)
    _clear_empty_initial_paragraph(section.footer)

    if cui_banner:
        _add_styled_paragraph(
            section.first_page_header,
            cui_banner,
            bold=True,
            align=WD_ALIGN_PARAGRAPH.CENTER,
            space_after_pt=4,
        )
        _add_styled_paragraph(
            section.header,
            cui_banner,
            bold=True,
            align=WD_ALIGN_PARAGRAPH.CENTER,
            space_after_pt=2,
        )
        _add_styled_paragraph(
            section.first_page_footer,
            cui_banner,
            bold=True,
            align=WD_ALIGN_PARAGRAPH.CENTER,
        )
        _add_styled_paragraph(
            section.footer,
            cui_banner,
            bold=True,
            align=WD_ALIGN_PARAGRAPH.CENTER,
        )

    if page_number_cont:
        p_num = _add_styled_paragraph(section.header, "", align=WD_ALIGN_PARAGRAPH.RIGHT)
        _add_page_number_field(p_num)
    if cont_footer_note:
        _add_styled_paragraph(
            section.first_page_footer,
            cont_footer_note,
            font_size=10.0,
            align=WD_ALIGN_PARAGRAPH.LEFT,
        )


def _add_letterhead_to_first_header(
    section: docx.section.Section,
    letterhead_lines: list[str],
    seal_path: Path | None,
) -> None:
    """Place AFMAN 33-326 / AFH 33-337 Copperplate Gothic Bold letterhead on Page 1."""
    hdr = section.first_page_header
    top_line = letterhead_lines[0] if letterhead_lines else "DEPARTMENT OF THE AIR FORCE"
    sub_lines = letterhead_lines[1:]
    if seal_path and seal_path.exists():
        tbl = hdr.add_table(rows=1, cols=3, width=Inches(6.5))
        widths = (Inches(0.9), Inches(4.7), Inches(0.9))
        for idx, w in enumerate(widths):
            tbl.cell(0, idx).width = w
            _set_cell_no_borders(tbl.cell(0, idx))
        p_img = tbl.cell(0, 0).paragraphs[0]
        p_img.add_run().add_picture(str(seal_path), width=Inches(0.75))
        mid = tbl.cell(0, 1)
        p_top = mid.paragraphs[0]
        p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_top = p_top.add_run(top_line)
        r_top.font.name = "Copperplate Gothic Bold"
        r_top.font.size = Pt(12)
        r_top.bold = True
        for sub in sub_lines:
            _add_styled_paragraph(
                mid,
                sub,
                font_name="Copperplate Gothic Bold",
                font_size=10.5,
                bold=True,
                align=WD_ALIGN_PARAGRAPH.CENTER,
                space_before_pt=3,
            )
        return

    _add_styled_paragraph(
        hdr,
        top_line,
        font_name="Copperplate Gothic Bold",
        font_size=12,
        bold=True,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    for sub in sub_lines:
        _add_styled_paragraph(
            hdr,
            sub,
            font_name="Copperplate Gothic Bold",
            font_size=10.5,
            bold=True,
            align=WD_ALIGN_PARAGRAPH.CENTER,
            space_before_pt=3,
        )


def _compute_memo_margins(paragraphs: list[Any], style_cfg: StyleConfig) -> tuple[float, float]:
    """Return (side_margin_in, top_margin_in) using AFH 33-337 p. 163 short-memo rule."""
    est_lines = sum(
        max(1, len(str(item.get("text", "") if isinstance(item, dict) else item)) // 75)
        for item in paragraphs
    )
    if est_lines == 0 or est_lines >= 20:
        return 1.0, 1.0
    if est_lines >= 10:
        return 1.25, 1.0
    return 1.5, 1.0


def _append_outline_paragraphs(
    doc: docx.Document, paragraphs: list[Any], *, single_unnumbered: bool = False
) -> None:
    """Render Classic Outline paragraphs with AFH 33-337 p. 180 indentation."""
    for idx, item in enumerate(paragraphs):
        if isinstance(item, str):
            item = {"level": 0, "text": item, "marker": f"{idx + 1}."}
        level = int(item.get("level", 0))
        text = str(item.get("text", ""))
        marker = str(item.get("marker", "")).strip()
        if single_unnumbered and level == 0 and len(paragraphs) == 1:
            _add_styled_paragraph(doc, text, space_after_pt=12)
            continue
        if not marker:
            marker = CLASSIC_MARKERS[min(level, len(CLASSIC_MARKERS) - 1)]
        first_indent = round(level * 0.25, 2)
        p = _add_styled_paragraph(
            doc,
            "",
            left_indent_in=0.0,
            first_line_indent_in=first_indent,
            space_after_pt=12,
        )
        r_m = p.add_run(f"{marker}  ")
        r_m.font.name = "Times New Roman"
        r_m.font.size = Pt(12)
        if level in (4, 5):
            r_m.underline = True
        r_t = p.add_run(text)
        r_t.font.name = "Times New Roman"
        r_t.font.size = Pt(12)


def _append_signature_and_closing(
    doc: docx.Document,
    spec: dict[str, Any],
    *,
    left_margin_in: float,
    sig_indent_override: float | None = None,
    cui_banner: str | None = None,
    cui_di_lines: list[str] | None = None,
    cui_transmittal: bool = False,
) -> None:
    """Render authority line, 5th-line 4.5-inch signature block, attachments, cc, and CUI DI."""
    sig_offset = (
        sig_indent_override
        if sig_indent_override is not None
        else max(1.5, 4.5 - left_margin_in)
    )
    authority_line = str(spec.get("authority_line", "")).strip()
    if authority_line:
        _add_styled_paragraph(
            doc,
            authority_line.upper(),
            left_indent_in=sig_offset,
            space_before_pt=12,
            keep_with_next=True,
        )

    sig_lines = [str(x).strip() for x in spec.get("signature_lines", []) if str(x).strip()]
    if sig_lines:
        _add_styled_paragraph(
            doc,
            sig_lines[0],
            left_indent_in=sig_offset,
            space_before_pt=48,
            keep_together=True,
            keep_with_next=len(sig_lines) > 1,
        )
        for idx, line in enumerate(sig_lines[1:]):
            overflow = idx >= 2
            indent = sig_offset + (0.25 if overflow else 0.0)
            _add_styled_paragraph(
                doc,
                line,
                left_indent_in=indent,
                keep_together=True,
                keep_with_next=idx < len(sig_lines) - 2,
            )

    attachments = [str(a).strip() for a in spec.get("attachments", []) if str(a).strip()]
    if len(attachments) == 1:
        _add_styled_paragraph(
            doc, f"Attachment:\n{attachments[0]}", space_before_pt=24, keep_together=True
        )
    elif len(attachments) > 1:
        header = _add_styled_paragraph(
            doc,
            f"{len(attachments)} Attachments:",
            space_before_pt=24,
            keep_with_next=True,
        )
        header.paragraph_format.keep_together = True
        for i, atch in enumerate(attachments, start=1):
            _add_styled_paragraph(doc, f"{i}.  {atch}", keep_together=True)

    cc_list = [str(c).strip() for c in spec.get("cc", []) if str(c).strip()]
    if cc_list:
        _add_styled_paragraph(doc, "cc:", space_before_pt=12, keep_with_next=True)
        for item in cc_list:
            _add_styled_paragraph(doc, item, keep_together=True)

    dist = [str(d).strip() for d in spec.get("distribution", []) if str(d).strip()]
    if dist:
        _add_styled_paragraph(doc, "DISTRIBUTION:", space_before_pt=12, keep_with_next=True)
        for item in dist:
            _add_styled_paragraph(doc, item, keep_together=True)

    if cui_banner and cui_transmittal:
        _add_styled_paragraph(
            doc,
            TRANSMITTAL_NOTICE,
            space_before_pt=24,
            keep_together=True,
        )
    elif cui_banner and cui_di_lines:
        _add_styled_paragraph(
            doc,
            "\n".join(cui_di_lines),
            font_size=10.0,
            left_indent_in=sig_offset,
            space_before_pt=24,
            keep_together=True,
        )


def _render_memo_family(
    doc: docx.Document,
    spec: dict[str, Any],
    *,
    doc_type: str,
    style_cfg: StyleConfig,
    dept: str,
    seal_path: Path | None,
    cui_banner: str | None,
    cui_di_lines: list[str] | None,
    cui_transmittal: bool,
) -> None:
    """Build Official Memo, MFR, Indorsement, Minutes, Staff Study, or Agenda."""
    _configure_base_style(doc, "Times New Roman", 12.0)
    paragraphs = spec.get("paragraphs", [])
    side_m, top_m = (
        _compute_memo_margins(paragraphs, style_cfg)
        if doc_type == "memo"
        else (1.0, 1.0)
    )
    section = doc.sections[0]
    _apply_margins(section, top=top_m, bottom=1.0, left=side_m, right=side_m)
    _apply_cui_headers_footers(section, cui_banner, page_number_cont=True)

    letterhead = spec.get("letterhead_lines") or style_cfg.letterhead_lines
    if doc_type != "agenda":
        _add_letterhead_to_first_header(section, letterhead, seal_path)

    date_str = str(spec.get("date", "[DD Month YYYY]"))
    if doc_type == "indorsement":
        ind_label = str(spec.get("indorsement_header", "1st Ind, ORG/SYMBOL"))
        p = _add_styled_paragraph(doc, f"{ind_label}\t{date_str}", space_before_pt=24)
        p.paragraph_format.tab_stops.add_tab_stop(
            Inches(8.5 - 2 * side_m), WD_ALIGN_PARAGRAPH.RIGHT
        )
    else:
        _add_styled_paragraph(
            doc,
            date_str,
            align=WD_ALIGN_PARAGRAPH.RIGHT,
            space_before_pt=36,
        )

    mem_for = spec.get("memorandum_for", ["ORG/SYMBOL"])
    if doc_type == "mfr":
        mem_for = ["RECORD"]
    if isinstance(mem_for, str):
        mem_for = [mem_for]
    for i, addr in enumerate(mem_for):
        prefix = "MEMORANDUM FOR  " if i == 0 else ""
        _add_styled_paragraph(
            doc,
            f"{prefix}{addr}",
            left_indent_in=0.0 if i == 0 else 1.85,
            space_before_pt=12 if i == 0 else 0,
        )
    if spec.get("in_turn"):
        _add_styled_paragraph(doc, "IN TURN", left_indent_in=1.85)
    if spec.get("through"):
        _add_styled_paragraph(doc, f"THROUGH:  {spec['through']}", space_before_pt=12)

    from_lines = spec.get("from_lines", [])
    if doc_type != "mfr" and from_lines:
        if isinstance(from_lines, str):
            from_lines = [from_lines]
        for i, fline in enumerate(from_lines):
            prefix = "FROM:  " if i == 0 else ""
            _add_styled_paragraph(
                doc,
                f"{prefix}{fline}",
                left_indent_in=0.0 if i == 0 else 0.7,
                space_before_pt=12 if i == 0 else 0,
            )

    subject = str(spec.get("subject", "[Subject in Title Case]"))
    _add_styled_paragraph(doc, f"SUBJECT:  {subject}", space_before_pt=12)

    refs = [str(r).strip() for r in spec.get("references", []) if str(r).strip()]
    if refs:
        label_map = "abcdefghijklmnopqrstuvwxyz"
        for i, ref in enumerate(refs):
            tag = f"({label_map[i % len(label_map)]})"
            prefix = "References:  " if i == 0 else ""
            _add_styled_paragraph(
                doc,
                f"{prefix}{tag} {ref}",
                left_indent_in=0.0 if i == 0 else 1.15,
                space_before_pt=12 if i == 0 else 0,
            )

    _add_styled_paragraph(doc, "", space_before_pt=0, space_after_pt=0)
    _append_outline_paragraphs(doc, paragraphs, single_unnumbered=(doc_type == "memo"))

    if doc_type == "minutes" and spec.get("approving_signature_lines"):
        _append_signature_and_closing(
            doc,
            {
                "signature_lines": spec.get("signature_lines", []),
                "authority_line": "",
            },
            left_margin_in=side_m,
        )
        _add_styled_paragraph(
            doc,
            "Approved as written",
            space_before_pt=24,
            keep_with_next=True,
        )
        _append_signature_and_closing(
            doc,
            {
                "signature_lines": spec.get("approving_signature_lines", []),
                "attachments": spec.get("attachments", []),
                "cc": spec.get("cc", []),
            },
            left_margin_in=side_m,
            cui_banner=cui_banner,
            cui_di_lines=cui_di_lines,
            cui_transmittal=cui_transmittal,
        )
        return

    _append_signature_and_closing(
        doc,
        spec,
        left_margin_in=side_m,
        cui_banner=cui_banner,
        cui_di_lines=cui_di_lines,
        cui_transmittal=cui_transmittal,
    )


def _render_esss(
    doc: docx.Document,
    spec: dict[str, Any],
    *,
    cui_banner: str | None,
    cui_di_lines: list[str] | None,
    cui_transmittal: bool,
) -> None:
    """Build an AFH 33-337 Ch. 18 eSSS (Electronic Staff Summary Sheet)."""
    _configure_base_style(doc, "Times New Roman", 12.0)
    section = doc.sections[0]
    _apply_margins(section, top=1.0, bottom=1.0, left=1.0, right=1.0)
    _apply_cui_headers_footers(section, cui_banner, page_number_cont=True)

    _add_styled_paragraph(
        doc,
        "ELECTRONIC STAFF SUMMARY SHEET",
        bold=True,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after_pt=12,
    )
    meta = [
        ("TO:", spec.get("to", "[APPROVING OFFICIAL / OFFICE]")),
        ("FROM:", spec.get("from", "[OPR / ACTION OFFICER]")),
        ("SUBJECT:", spec.get("subject", "[Staff Package Subject]")),
        ("SUSPENSE:", spec.get("suspense", "[DD Mmm YY]")),
    ]
    for label, val in meta:
        _add_styled_paragraph(doc, f"{label}  {val}", space_after_pt=4)

    coord_rows = spec.get(
        "coordination",
        [
            {"office": "[OPR]", "action": "Coord", "name": "[Name]", "concur": "Concur", "date": "[DD Mmm YY]"},
            {"office": "[CC]", "action": "Sig/Appr", "name": "[Commander]", "concur": "Pending", "date": ""},
        ],
    )
    tbl = doc.add_table(rows=len(coord_rows) + 1, cols=5)
    tbl.style = "Table Grid"
    headers = ("Office", "Action", "Name / Rank", "Status", "Date")
    for c_idx, h_text in enumerate(headers):
        p = tbl.cell(0, c_idx).paragraphs[0]
        r = p.add_run(h_text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)
        r.bold = True
    for r_idx, row in enumerate(coord_rows, start=1):
        vals = (
            str(row.get("office", "")),
            str(row.get("action", "")),
            str(row.get("name", "")),
            str(row.get("concur", "")),
            str(row.get("date", "")),
        )
        for c_idx, val in enumerate(vals):
            p = tbl.cell(r_idx, c_idx).paragraphs[0]
            r = p.add_run(val)
            r.font.name = "Times New Roman"
            r.font.size = Pt(11)

    _add_styled_paragraph(doc, "", space_before_pt=6)
    sections = [
        ("1.  PURPOSE.", spec.get("purpose", "[State single purpose sentence.]")),
        ("2.  BACKGROUND.", spec.get("background", "[State concise context and authority.]")),
        ("3.  DISCUSSION.", spec.get("discussion", "[Summarize analysis, impacts, and options.]")),
        ("4.  VIEWS OF OTHERS.", spec.get("views_of_others", "[Summarize JA/FM/PA/OCR inputs or N/A.]")),
        ("5.  RECOMMENDATION.", spec.get("recommendation", "[State exact action requested.]")),
    ]
    for heading, body in sections:
        p = _add_styled_paragraph(doc, "", space_after_pt=12)
        rh = p.add_run(f"{heading}  ")
        rh.font.name = "Times New Roman"
        rh.font.size = Pt(12)
        rh.bold = True
        rb = p.add_run(str(body))
        rb.font.name = "Times New Roman"
        rb.font.size = Pt(12)

    sig_lines = spec.get(
        "signature_lines", ["//SIGNED//", "FIRST M. LAST, Rank, USAF", "Action Officer, ORG/SYMBOL"]
    )
    _add_styled_paragraph(
        doc, "\n".join(sig_lines), space_before_pt=18, keep_together=True, keep_with_next=True
    )

    tabs = spec.get("tabs", ["Tab 1: Outgoing Memorandum", "Tab 2: Bullet Background Paper"])
    if tabs:
        _add_styled_paragraph(doc, "Tabs:", space_before_pt=12, keep_with_next=True)
        for t in tabs:
            _add_styled_paragraph(doc, str(t), keep_together=True)

    if cui_banner and cui_transmittal:
        _add_styled_paragraph(doc, TRANSMITTAL_NOTICE, space_before_pt=12)
    elif cui_banner and cui_di_lines:
        _add_styled_paragraph(
            doc, "\n".join(cui_di_lines), font_size=10.0, left_indent_in=3.5, space_before_pt=12
        )


def _render_staff_paper(
    doc: docx.Document,
    spec: dict[str, Any],
    *,
    doc_type: str,
    cui_banner: str | None,
    cui_di_lines: list[str] | None,
) -> None:
    """Build Point, Talking, BBP, Background, or Position Paper per AFH 33-337 Ch. 16."""
    _configure_base_style(doc, "Times New Roman", 12.0)
    section = doc.sections[0]
    _apply_margins(section, top=1.0, bottom=1.0, left=1.0, right=1.0)
    id_line = str(spec.get("id_line", "Rank Last/Org/Symbol/DSN 555-0100/initials/21 Nov 26"))
    _apply_cui_headers_footers(
        section, cui_banner, page_number_cont=True, cont_footer_note=id_line
    )

    title_map = {
        "bbp": "BULLET BACKGROUND PAPER",
        "position": "POSITION PAPER",
        "talking": "TALKING PAPER",
        "point": "POINT PAPER",
        "background": "BACKGROUND PAPER",
    }
    paper_heading = str(spec.get("paper_label", title_map.get(doc_type, "STAFF PAPER"))).upper()
    topic = str(spec.get("topic", "[TOPIC IN ALL CAPS]")).upper()
    _add_styled_paragraph(
        doc, paper_heading, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER
    )
    _add_styled_paragraph(doc, "ON", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _add_styled_paragraph(
        doc, topic, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after_pt=18
    )

    for sec in spec.get("sections", []):
        heading = str(sec.get("heading", "PURPOSE")).upper()
        _add_styled_paragraph(
            doc, heading, bold=True, space_before_pt=6, space_after_pt=6, keep_with_next=True
        )
        for item in sec.get("items", []):
            if isinstance(item, str):
                _add_styled_paragraph(doc, item, space_after_pt=12)
                continue
            lvl = int(item.get("level", 0))
            txt = str(item.get("text", ""))
            marker = "-" * (lvl + 1)
            left_in = 0.25 * (lvl + 1)
            _add_styled_paragraph(
                doc,
                f"{marker}  {txt}",
                left_indent_in=left_in,
                first_line_indent_in=-0.25,
                space_after_pt=12,
            )

    if cui_banner and cui_di_lines:
        _add_styled_paragraph(
            doc,
            "\n".join(cui_di_lines),
            font_size=10.0,
            left_indent_in=3.5,
            space_before_pt=18,
            keep_together=True,
        )


def _render_bio(
    doc: docx.Document,
    spec: dict[str, Any],
    *,
    photo_path: str | None,
    cui_banner: str | None,
    cui_di_lines: list[str] | None,
) -> None:
    """Build Official Biography per AFH 33-337 Ch. 20 (Arial 9 / 13.5 pt, 1.15 spacing)."""
    _configure_base_style(doc, "Arial", 9.0)
    section = doc.sections[0]
    _apply_margins(section, top=1.0, bottom=1.0, left=1.0, right=1.0)
    _apply_cui_headers_footers(section, cui_banner, page_number_cont=False)

    top_tbl = doc.add_table(rows=1, cols=2)
    top_tbl.cell(0, 0).width = Inches(3.2)
    top_tbl.cell(0, 1).width = Inches(3.3)
    _set_cell_no_borders(top_tbl.cell(0, 0))
    _set_cell_no_borders(top_tbl.cell(0, 1))

    left_cell = top_tbl.cell(0, 0)
    p_id = left_cell.paragraphs[0]
    p_id.paragraph_format.line_spacing = 1.15
    p_id.paragraph_format.space_after = Pt(8)
    r_id = p_id.add_run(str(spec.get("id_line", "FULL RANK FIRST M. LAST")))
    r_id.font.name = "Arial"
    r_id.font.size = Pt(13.5)
    r_id.bold = True

    narrative = spec.get("narrative_paragraphs", ["[Biographical narrative paragraphs.]"])
    if narrative:
        _add_styled_paragraph(
            left_cell,
            str(narrative[0]),
            font_name="Arial",
            font_size=9.0,
            line_spacing=1.15,
            space_after_pt=8,
        )

    right_cell = top_tbl.cell(0, 1)
    p_pic = right_cell.paragraphs[0]
    p_pic.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if photo_path and Path(photo_path).exists():
        p_pic.add_run().add_picture(str(photo_path), width=Inches(3.2))
    else:
        _set_cell_thin_border(right_cell, "888888")
        r_ph = p_pic.add_run("OFFICIAL PHOTOGRAPH\n(3.2 in x 4.0 in)")
        r_ph.font.name = "Arial"
        r_ph.font.size = Pt(9)
        p_pic.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for para in narrative[1:]:
        _add_styled_paragraph(
            doc,
            str(para),
            font_name="Arial",
            font_size=9.0,
            line_spacing=1.15,
            space_before_pt=6,
            space_after_pt=6,
        )

    for sec in spec.get("sections", []):
        _add_styled_paragraph(
            doc,
            str(sec.get("heading", "EDUCATION")).upper().rstrip(":"),
            font_name="Arial",
            font_size=9.0,
            bold=True,
            line_spacing=1.15,
            space_before_pt=8,
            space_after_pt=3,
            keep_with_next=True,
        )
        for entry in sec.get("entries", []):
            _add_styled_paragraph(
                doc,
                str(entry),
                font_name="Arial",
                font_size=9.0,
                line_spacing=1.15,
                space_after_pt=2,
            )

    as_of = str(spec.get("current_as_of", "(Current as of November 2026)"))
    _add_styled_paragraph(
        doc,
        as_of,
        font_name="Arial",
        font_size=9.0,
        line_spacing=1.15,
        space_before_pt=10,
    )
    if cui_banner and cui_di_lines:
        _add_styled_paragraph(
            doc,
            "\n".join(cui_di_lines),
            font_name="Arial",
            font_size=8.5,
            left_indent_in=3.5,
            space_before_pt=10,
        )


def _render_letter(
    doc: docx.Document,
    spec: dict[str, Any],
    *,
    style_cfg: StyleConfig,
    seal_path: Path | None,
    cui_banner: str | None,
    cui_di_lines: list[str] | None,
) -> None:
    """Build Personal Letter per AFH 33-337 Ch. 15 (1.75 in Page 1 top, open punctuation)."""
    _configure_base_style(doc, "Times New Roman", 12.0)
    section = doc.sections[0]
    _apply_margins(section, top=1.75, bottom=1.0, left=1.0, right=1.0)
    _apply_cui_headers_footers(section, cui_banner, page_number_cont=True)
    letterhead = spec.get("letterhead_lines") or style_cfg.letterhead_lines
    _add_letterhead_to_first_header(section, letterhead, seal_path)

    _add_styled_paragraph(
        doc, str(spec.get("date", "November 21, 2026")), align=WD_ALIGN_PARAGRAPH.RIGHT
    )
    ret_lines = spec.get("return_address", ["Full Rank Full Name", "Duty Title", "Address", "City ST  ZIP"])
    for idx, line in enumerate(ret_lines):
        _add_styled_paragraph(doc, str(line), space_before_pt=12 if idx == 0 else 0)

    rec_lines = spec.get("receiver_address", ["The Honorable Full Name", "Title", "Address", "Washington DC  20301"])
    for idx, line in enumerate(rec_lines):
        _add_styled_paragraph(doc, str(line), space_before_pt=24 if idx == 0 else 0)

    salutation = str(spec.get("salutation", "Dear Mr. Secretary")).rstrip(":,")
    _add_styled_paragraph(doc, salutation, space_before_pt=12, space_after_pt=12)

    for para in spec.get("paragraphs", ["[Unnumbered letter paragraph with 0.5 inch first-line indent.]"]):
        _add_styled_paragraph(doc, str(para), first_line_indent_in=0.5, space_after_pt=12)

    close = str(spec.get("complimentary_close", "Sincerely")).rstrip(":,")
    _add_styled_paragraph(doc, close, left_indent_in=3.5, space_before_pt=12, keep_with_next=True)

    _append_signature_and_closing(
        doc,
        {
            "signature_lines": spec.get("signature_lines", ["FIRST M. LAST, Rank, USAF", "Commander"]),
            "attachments": spec.get("attachments", []),
            "cc": spec.get("cc", []),
        },
        left_margin_in=1.0,
        sig_indent_override=3.5,
        cui_banner=cui_banner,
        cui_di_lines=cui_di_lines,
    )


def _render_resume(doc: docx.Document, spec: dict[str, Any]) -> None:
    """Build Civilian Transition / Federal Resume per AFH 33-337 Ch. 21."""
    _configure_base_style(doc, "Times New Roman", 11.0)
    section = doc.sections[0]
    _apply_margins(section, top=0.75, bottom=0.75, left=0.75, right=0.75)
    _add_styled_paragraph(
        doc,
        str(spec.get("name", "FIRST M. LAST")).upper(),
        font_size=14.0,
        bold=True,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    _add_styled_paragraph(
        doc,
        str(spec.get("contact", "City, ST  |  phone  |  email@example.com")),
        font_size=10.5,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after_pt=10,
    )
    for sec in spec.get("sections", []):
        _add_styled_paragraph(
            doc,
            str(sec.get("heading", "EXPERIENCE")).upper(),
            font_size=11.0,
            bold=True,
            space_before_pt=8,
            space_after_pt=4,
            keep_with_next=True,
        )
        for item in sec.get("items", []):
            _add_styled_paragraph(
                doc, f"-  {item}", font_size=11.0, left_indent_in=0.2, first_line_indent_in=-0.15, space_after_pt=4
            )


def build_document(
    spec: dict[str, Any],
    out_path: Path,
    *,
    doc_type: str | None = None,
    dept: str | None = None,
    seal_arg: str | None = None,
    photo_path: str | None = None,
    cui: bool = False,
    cui_category: str | None = None,
    cui_ldc: str | None = None,
    cui_poc: str | None = None,
    cui_office: str | None = None,
    cui_transmittal: bool = False,
    style_config_path: str | None = None,
) -> Path:
    """Render any supported T&Q document JSON specification into a .docx file."""
    style_cfg = load_style_config(style_config_path)
    resolved_type = (doc_type or str(spec.get("doc_type", "memo"))).strip().lower()
    if resolved_type not in DOC_TYPES:
        raise ValueError(f"Unsupported doc_type: {resolved_type}")

    resolved_dept = (dept or spec.get("dept") or style_cfg.department_name).strip().lower()
    if resolved_dept == "ask":
        resolved_dept = "dod"
    resolved_service = str(spec.get("service", style_cfg.service)).upper()
    seal_choice = seal_arg or spec.get("seal") or style_cfg.seal_path or "auto"
    seal = resolve_seal(seal_choice, resolved_dept, resolved_service, context="doc")

    use_cui = bool(cui or spec.get("cui"))
    cui_cat = cui_category or spec.get("cui_category") or style_cfg.cui_default_category
    cui_ldc_val = cui_ldc or spec.get("cui_ldc") or style_cfg.cui_default_ldc
    cui_banner = build_cui_banner(cui_cat, cui_ldc_val) if use_cui else None
    cui_di = (
        build_cui_di_lines(
            component=None,
            office=cui_office or spec.get("cui_office") or style_cfg.cui_default_office,
            categories=cui_cat,
            ldc=cui_ldc_val,
            poc=cui_poc or spec.get("cui_poc") or style_cfg.cui_default_poc,
            on_letterhead=(resolved_type not in {"bio", "esss", "bbp", "position", "talking", "point"}),
        )
        if use_cui
        else None
    )

    doc = docx.Document()
    if resolved_type in {"memo", "mfr", "indorsement", "minutes", "staff-study", "agenda"}:
        _render_memo_family(
            doc,
            spec,
            doc_type=resolved_type,
            style_cfg=style_cfg,
            dept=resolved_dept,
            seal_path=seal,
            cui_banner=cui_banner,
            cui_di_lines=cui_di,
            cui_transmittal=bool(cui_transmittal or spec.get("cui_transmittal")),
        )
    elif resolved_type == "esss":
        _render_esss(
            doc,
            spec,
            cui_banner=cui_banner,
            cui_di_lines=cui_di,
            cui_transmittal=bool(cui_transmittal or spec.get("cui_transmittal")),
        )
    elif resolved_type in {"bbp", "position", "talking", "point"}:
        _render_staff_paper(
            doc, spec, doc_type=resolved_type, cui_banner=cui_banner, cui_di_lines=cui_di
        )
    elif resolved_type == "bio":
        _render_bio(
            doc,
            spec,
            photo_path=photo_path or spec.get("photo_path"),
            cui_banner=cui_banner,
            cui_di_lines=cui_di,
        )
    elif resolved_type == "letter":
        _render_letter(
            doc,
            spec,
            style_cfg=style_cfg,
            seal_path=seal,
            cui_banner=cui_banner,
            cui_di_lines=cui_di,
        )
    else:
        _render_resume(doc, spec)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path


def get_schema(
    doc_type: str,
    preset: str = "standard",
    dept: str = "dod",
    cui: bool = False,
) -> dict[str, Any]:
    """Return a clean starter JSON schema for a T&Q document or memo preset."""
    terms = dept_terms(dept)
    dt = doc_type.strip().lower()
    if dt in {"memo", "mfr", "indorsement", "minutes", "staff-study", "agenda"}:
        base: dict[str, Any] = {
            "doc_type": dt,
            "dept": terms["choice"],
            "service": "USAF",
            "letterhead_lines": ["DEPARTMENT OF THE AIR FORCE", "1ST FIGHTER WING (ACC)"],
            "date": "21 November 2026",
            "memorandum_for": ["1 FW/CC"] if dt != "mfr" else ["RECORD"],
            "from_lines": ["1 FW/XP", "123 Base Road", "Langley AFB VA  23665-1234"] if dt != "mfr" else [],
            "subject": f"Template Subject ({terms['acronym']} Guidance)",
            "references": [],
            "paragraphs": [
                {"level": 0, "text": f"State bottom-line purpose upfront aligned with {terms['department']} guidance."},
                {"level": 0, "text": "Provide supporting action and point of contact (POC) at DSN 555-0100."},
            ],
            "authority_line": "",
            "signature_lines": ["JANE A. DOE, Colonel, USAF", "Vice Commander"],
            "attachments": [],
            "cc": [],
            "cui": cui,
            "cui_category": "OPSEC" if cui else "",
            "cui_ldc": "FEDCON" if cui else "None",
        }
        if dt == "memo" and preset == "appointment":
            base["subject"] = "Appointment as Unit Government Purchase Card (GPC) Billing Official"
            base["paragraphs"] = [
                {"level": 0, "text": "I appoint SSgt Alex R. Rivera as Primary Billing Official, effective 21 November 2026."},
                {"level": 0, "text": "This letter supersedes all previous GPC Billing Official appointment letters on file."},
                {"level": 0, "text": "Comply with AFI 64-117 and retain this appointment letter until relieved in writing. Direct questions to DSN 555-0142."},
            ]
        elif dt == "memo" and preset == "policy":
            base["subject"] = "Command Policy: Wing Staff Package Turnaround Timelines"
            base["paragraphs"] = [
                {"level": 0, "text": "Effective immediately, all Wing staff packages require backward-planned suspense dates at least 5 duty days before external due dates."},
                {"level": 0, "text": "Review this policy annually. This policy supersedes Policy Memo dated 15 Jan 25."},
            ]
        elif dt == "memo" and preset == "etp":
            base["subject"] = "Exception to Policy Request — DAFI 33-360 Paragraph 4.2"
            base["paragraphs"] = [
                {"level": 0, "text": "Request an Exception to Policy (ETP) under DAFI 33-360 tiering authority through 31 December 2026."},
                {"level": 0, "text": "Justification: Mitigates operational schedule risk while maintaining full audit traceability."},
                {"level": 0, "text": "POC is Capt Jordan Taylor at DSN 555-0199."},
            ]
            base["attachments"] = ["Justification Matrix"]
        elif dt == "memo" and preset == "assumption-of-command":
            base["subject"] = "Assumption of Command"
            base["paragraphs"] = [
                {"level": 0, "text": "Under DAFI 51-509 and Special Order G-024, I assume command of the 1st Operations Support Squadron effective 21 November 2026 at 0900."}
            ]
            base["signature_lines"] = ["ALEXANDER M. GRANT, Lt Col, USAF", "Commander"]
        elif dt == "memo" and preset == "aar":
            base["subject"] = "After-Action Report — Command Communications Summit (20-21 Nov 26)"
            base["paragraphs"] = [
                {"level": 0, "text": "Attendees met with MAJCOM staff at Scott AFB, Illinois, from 20 to 21 November 2026."},
                {"level": 0, "text": "Key Outcomes: Validated unified staff package routing standards across three wings."},
                {"level": 0, "text": "Recommended Follow-Up: Publish updated local supplement by 15 December 2026."},
            ]
        elif dt == "minutes":
            base["subject"] = "Minutes — Wing Safety Council Meeting"
            base["approving_signature_lines"] = ["ROBERT C. KING, Colonel, USAF", "Commander"]
        return base

    if dt == "esss":
        return {
            "doc_type": "esss",
            "dept": terms["choice"],
            "to": "1 FW/CC",
            "from": "1 FW/XP",
            "subject": f"Implement Updated {terms['acronym']} Administrative Standards",
            "suspense": "28 Nov 26",
            "purpose": "Obtain Commander signature on updated Wing administrative standards (Tab 1).",
            "background": "Previous wing policy predated recent department updates.",
            "discussion": "Tab 1 aligns local correspondence templates and mandatory JA/FM staff chops.",
            "views_of_others": "JA and A6 concur; no FM resource impact.",
            "recommendation": "1 FW/CC sign Tab 1 memorandum.",
            "signature_lines": ["//SIGNED//", "TAYLOR R. SMITH, Maj, USAF", "Chief, Plans and Programs (1 FW/XP)"],
            "tabs": ["Tab 1: Outgoing Memorandum", "Tab 2: Bullet Background Paper"],
        }

    if dt in {"bbp", "position", "talking", "point"}:
        return {
            "doc_type": dt,
            "dept": terms["choice"],
            "topic": "UNIFIED STAFF PACKAGE MODERNIZATION",
            "id_line": "Maj Smith/1 FW/XP/DSN 555-0100/trs/21 Nov 26",
            "sections": [
                {"heading": "PURPOSE", "items": [{"level": 0, "text": "Summarize modernization COA for Wing Commander decision."}]},
                {"heading": "BACKGROUND", "items": [{"level": 0, "text": "Current routing takes 9 duty days across three offices"}, {"level": 1, "text": "Digital templates cut rework by 40 percent"}]},
                {"heading": "RECOMMENDATION", "items": [{"level": 0, "text": "Approve standardized package templates effective 1 Dec 26"}]},
            ],
        }

    if dt == "bio":
        return {
            "doc_type": "bio",
            "dept": terms["choice"],
            "id_line": "COLONEL JANE A. DOE",
            "narrative_paragraphs": [
                "Col. Jane A. Doe is Commander, 1st Operations Group, Joint Base Langley-Eustis, Virginia.",
                "Prior to her current position, the colonel was Vice Commander, 4th Fighter Wing, Seymour Johnson Air Force Base, North Carolina.",
            ],
            "sections": [
                {"heading": "EDUCATION", "entries": ["2004 Bachelor of Science, U.S. Air Force Academy, Colorado Springs, Colo."]},
                {"heading": "ASSIGNMENTS", "entries": ["1. June 2004 – May 2006, Student Pilot, Vance AFB, Okla."]},
                {"heading": "MAJOR AWARDS AND DECORATIONS", "entries": ["Legion of Merit", "Meritorious Service Medal with two oak leaf clusters"]},
                {"heading": "EFFECTIVE DATES OF PROMOTION", "entries": ["Second Lieutenant June 2, 2004", "Colonel July 1, 2024"]},
            ],
            "current_as_of": "(Current as of November 2026)",
        }

    if dt == "letter":
        return {
            "doc_type": "letter",
            "dept": terms["choice"],
            "date": "November 21, 2026",
            "return_address": ["Colonel Jane A. Doe", "Commander, 1st Fighter Wing", "123 Base Road", "Langley AFB VA  23665-1234"],
            "receiver_address": ["The Honorable Jordan V. Carter", f"{terms['secretary']}", "1000 Defense Pentagon", "Washington DC  20301-1000"],
            "salutation": "Dear Mr. Secretary",
            "paragraphs": ["Thank you for visiting our Airmen and Guardians during yesterday's operational readiness review."],
            "complimentary_close": "Sincerely",
            "signature_lines": ["JANE A. DOE, Colonel, USAF", "Commander"],
        }

    return {
        "doc_type": "resume",
        "name": "JANE A. DOE",
        "contact": "Hampton, VA  |  757-555-0100  |  jane.doe@example.com",
        "sections": [
            {"heading": "SUMMARY", "items": ["Senior aerospace operations leader with 20 years of joint planning experience."]},
            {"heading": "EXPERIENCE", "items": ["Director of Operations (2022–2026): Led 320-person unit across 4 deployed locations."]},
        ],
    }


def build_staff_package(
    pkg_spec: dict[str, Any],
    out_dir: Path,
    *,
    dept: str | None = None,
    seal_arg: str | None = None,
    cui: bool = False,
    cui_category: str | None = None,
    cui_ldc: str | None = None,
    cui_poc: str | None = None,
    cui_office: str | None = None,
    style_config_path: str | None = None,
) -> list[Path]:
    """Generate a synchronized eSSS, Tab files, and Cover Email in one directory."""
    out_dir.mkdir(parents=True, exist_ok=True)
    shared = {
        "dept": dept or pkg_spec.get("dept", "dod"),
        "subject": pkg_spec.get("subject", "[Package Subject]"),
        "suspense": pkg_spec.get("suspense", "[DD Mmm YY]"),
        "from": pkg_spec.get("from", "[OPR]"),
        "to": pkg_spec.get("to", "[CC]"),
        "cui": bool(cui or pkg_spec.get("cui")),
        "cui_category": cui_category or pkg_spec.get("cui_category", ""),
        "cui_ldc": cui_ldc or pkg_spec.get("cui_ldc", "None"),
        "cui_poc": cui_poc or pkg_spec.get("cui_poc", ""),
        "cui_office": cui_office or pkg_spec.get("cui_office") or pkg_spec.get("from", ""),
    }

    created: list[Path] = []
    esss_spec = copy.deepcopy(pkg_spec.get("esss", get_schema("esss", dept=shared["dept"])))
    esss_spec.update({
        "doc_type": "esss",
        "subject": shared["subject"],
        "suspense": shared["suspense"],
        "from": shared["from"],
        "to": shared["to"],
    })
    tabs_specs = pkg_spec.get("tabs", [])
    esss_tabs = [
        f"Tab {idx + 1}: {t.get('label', t.get('doc_type', 'Attachment').upper())}"
        for idx, t in enumerate(tabs_specs)
    ]
    if esss_tabs:
        esss_spec["tabs"] = esss_tabs

    esss_path = out_dir / "00_eSSS.docx"
    build_document(
        esss_spec,
        esss_path,
        doc_type="esss",
        dept=shared["dept"],
        seal_arg=seal_arg,
        cui=shared["cui"],
        cui_category=shared["cui_category"],
        cui_ldc=shared["cui_ldc"],
        cui_poc=shared["cui_poc"],
        cui_office=shared["cui_office"],
        cui_transmittal=shared["cui"],
        style_config_path=style_config_path,
    )
    created.append(esss_path)

    for idx, tab in enumerate(tabs_specs, start=1):
        tab_spec = copy.deepcopy(tab)
        tab_spec.setdefault("subject", shared["subject"])
        tab_spec.setdefault("dept", shared["dept"])
        fname = str(tab.get("filename", f"Tab{idx}_{tab.get('doc_type', 'doc')}.docx"))
        tab_out = out_dir / fname
        build_document(
            tab_spec,
            tab_out,
            doc_type=tab_spec.get("doc_type", "memo"),
            dept=shared["dept"],
            seal_arg=seal_arg,
            cui=shared["cui"],
            cui_category=shared["cui_category"],
            cui_ldc=shared["cui_ldc"],
            cui_poc=shared["cui_poc"],
            cui_office=shared["cui_office"],
            style_config_path=style_config_path,
        )
        created.append(tab_out)

    cui_prefix = "[CUI]" if shared["cui"] else ""
    banner_line = f"{build_cui_banner(shared['cui_category'], shared['cui_ldc'])}\n\n" if shared["cui"] else ""
    email_text = (
        f"SUBJECT: {cui_prefix}[SIGN][{shared['suspense']}][{shared['subject']}]\n\n"
        f"{banner_line}"
        f"BLUF: Request {shared['to']} review and approve attached staff package NLT {shared['suspense']}.\n\n"
        f"1. Purpose: {esss_spec.get('purpose', '')}\n"
        f"2. Impact/Suspense: Routing due NLT {shared['suspense']}.\n"
        f"3. Tabs Attached: {', '.join(esss_tabs) if esss_tabs else 'See package files'}\n"
        f"4. POC: {shared['cui_poc'] or shared['from']}\n\n"
        "//SIGNED//\n"
        + "\n".join(esss_spec.get("signature_lines", [])[1:])
        + "\n"
    )
    email_path = out_dir / "00_Cover_Email.txt"
    with email_path.open("w", encoding="utf-8", newline="\r\n") as handle:
        handle.write(email_text)
    created.append(email_path)
    return created


def _add_common_flags(parser: argparse.ArgumentParser) -> None:
    """Attach shared department, seal, CUI, and style-config flags."""
    parser.add_argument("--dept", choices=("dow", "dod"), help="Department of War (DoW) or Defense (DoD)")
    parser.add_argument("--seal", help="Seal variant (auto|blue|bw|color|none) or explicit PNG path")
    parser.add_argument("--photo", help="Official portrait path for Official Biography")
    parser.add_argument("--cui", action="store_true", help="Include DoDI 5200.48 CUI banners and DI block")
    parser.add_argument("--cui-category", default="", help="CUI Category codes (e.g. CTI or CTI/EXPT)")
    parser.add_argument("--cui-ldc", default="", help="Limited Dissemination Control (e.g. FEDCON)")
    parser.add_argument("--cui-poc", default="", help="CUI DI block POC line")
    parser.add_argument("--cui-office", default="", help="CUI DI block originating office")
    parser.add_argument("--cui-transmittal", action="store_true", help="Mark as CUI transmittal doc")
    parser.add_argument("--style-config", help="Optional unit_style.json override path")


def main(argv: list[str] | None = None) -> int:
    """CLI entrypoint for Word (.docx) builder, packager, templates, and schemas."""
    configure_utf8_stdout()
    ap = argparse.ArgumentParser(description="Build AFH 33-337 compliant .docx documents and staff packages.")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_build = sub.add_parser("build", help="Build a populated .docx from a JSON specification")
    p_build.add_argument("spec", help="Path to JSON spec file")
    p_build.add_argument("--out", required=True, help="Output .docx path")
    p_build.add_argument("--type", choices=DOC_TYPES, help="Override doc_type in JSON")
    _add_common_flags(p_build)

    p_pkg = sub.add_parser("package", help="Build a synchronized Staff Package folder from JSON")
    p_pkg.add_argument("spec", help="Path to package JSON spec")
    p_pkg.add_argument("--out-dir", required=True, help="Output package directory")
    _add_common_flags(p_pkg)

    p_tpl = sub.add_parser("template", help="Generate a blank T&Q-compliant .docx template")
    p_tpl.add_argument("doc_type", choices=DOC_TYPES, help="Document type")
    p_tpl.add_argument("--out", required=True, help="Output .docx path")
    p_tpl.add_argument("--preset", choices=MEMO_PRESETS, default="standard")
    _add_common_flags(p_tpl)

    p_sch = sub.add_parser("schema", help="Print a starter JSON payload to stdout")
    p_sch.add_argument("doc_type", choices=DOC_TYPES, help="Document type")
    p_sch.add_argument("--preset", choices=MEMO_PRESETS, default="standard")
    p_sch.add_argument("--dept", choices=("dow", "dod"), default="dod")
    p_sch.add_argument("--cui", action="store_true")

    p_init = sub.add_parser("init-style", help="Write a starter unit_style.json configuration")
    p_init.add_argument("out", help="Output JSON path (e.g. unit_style.json)")
    p_init.add_argument("--dept", choices=("dow", "dod", "ask"), default="ask")
    p_init.add_argument("--service", choices=("USAF", "USSF"), default="USAF")

    args = ap.parse_args(argv)
    if args.cmd == "build":
        spec = load_json(Path(args.spec))
        out = build_document(
            spec,
            Path(args.out),
            doc_type=args.type,
            dept=args.dept,
            seal_arg=args.seal,
            photo_path=args.photo,
            cui=args.cui,
            cui_category=args.cui_category,
            cui_ldc=args.cui_ldc,
            cui_poc=args.cui_poc,
            cui_office=args.cui_office,
            cui_transmittal=args.cui_transmittal,
            style_config_path=args.style_config,
        )
        print(f"Built {out}")
        return 0

    if args.cmd == "package":
        spec = load_json(Path(args.spec))
        files = build_staff_package(
            spec,
            Path(args.out_dir),
            dept=args.dept,
            seal_arg=args.seal,
            cui=args.cui,
            cui_category=args.cui_category,
            cui_ldc=args.cui_ldc,
            cui_poc=args.cui_poc,
            cui_office=args.cui_office,
            style_config_path=args.style_config,
        )
        print("Built staff package:\n" + "\n".join(f"  - {p}" for p in files))
        return 0

    if args.cmd == "template":
        spec = get_schema(
            args.doc_type,
            preset=args.preset,
            dept=args.dept or "dod",
            cui=args.cui,
        )
        out = build_document(
            spec,
            Path(args.out),
            doc_type=args.doc_type,
            dept=args.dept,
            seal_arg=args.seal,
            photo_path=args.photo,
            cui=args.cui,
            cui_category=args.cui_category,
            cui_ldc=args.cui_ldc,
            cui_poc=args.cui_poc,
            cui_office=args.cui_office,
            cui_transmittal=args.cui_transmittal,
            style_config_path=args.style_config,
        )
        print(f"Created template {out}")
        return 0

    if args.cmd == "schema":
        print(json.dumps(get_schema(args.doc_type, args.preset, args.dept, args.cui), indent=2))
        return 0

    cfg = StyleConfig(department_name=args.dept, service=args.service)
    out_cfg = Path(args.out)
    out_cfg.parent.mkdir(parents=True, exist_ok=True)
    with out_cfg.open("w", encoding="utf-8", newline="\r\n") as handle:
        handle.write(json.dumps(cfg.to_dict(), indent=2) + "\n")
    print(f"Initialized unit style config -> {out_cfg}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
