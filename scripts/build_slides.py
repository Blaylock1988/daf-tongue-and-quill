"""Programmatic AFH 33-337 Ch. 10 PowerPoint (.pptx) briefing deck builder."""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

try:
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Inches, Pt
except ImportError as exc:
    raise SystemExit(
        "python-pptx is required for build_slides.py. Install with: pip install python-pptx"
    ) from exc

from tq_common import (
    build_cui_banner,
    build_cui_di_lines,
    configure_utf8_stdout,
    dept_terms,
    load_json,
    load_style_config,
    resolve_seal,
)

PALETTES: dict[str, dict[str, tuple[int, int, int]]] = {
    "usaf": {
        "bg": (245, 247, 250),
        "header_bg": (0, 48, 143),
        "title": (0, 48, 143),
        "text": (16, 24, 32),
        "accent": (180, 140, 40),
        "cui": (16, 24, 32),
    },
    "ussf": {
        "bg": (16, 24, 32),
        "header_bg": (0, 38, 58),
        "title": (242, 245, 248),
        "text": (219, 226, 233),
        "accent": (111, 147, 224),
        "cui": (242, 245, 248),
    },
    "dow": {
        "bg": (247, 246, 242),
        "header_bg": (28, 36, 52),
        "title": (28, 36, 52),
        "text": (24, 24, 27),
        "accent": (168, 130, 42),
        "cui": (24, 24, 27),
    },
    "dod": {
        "bg": (245, 246, 248),
        "header_bg": (20, 52, 92),
        "title": (20, 52, 92),
        "text": (20, 24, 32),
        "accent": (65, 105, 165),
        "cui": (20, 24, 32),
    },
    "light": {
        "bg": (250, 250, 248),
        "header_bg": (34, 49, 63),
        "title": (34, 49, 63),
        "text": (30, 30, 30),
        "accent": (0, 48, 143),
        "cui": (30, 30, 30),
    },
}

CHART_MAP = {
    "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "line": XL_CHART_TYPE.LINE,
    "pie": XL_CHART_TYPE.PIE,
}

TRAILING_ZERO_RE = re.compile(r"\b(\d+)\.0+\b")


def _clean_cell_value(raw: Any) -> str:
    """Strip misleading .00 trailing zeros per AFH 33-337 p. 120."""
    return TRAILING_ZERO_RE.sub(r"\1", str(raw))


def _set_slide_bg(slide: Any, rgb: tuple[int, int, int]) -> None:
    """Apply a solid RGB background fill to a slide."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*rgb)


def _add_cui_banners(
    slide: Any,
    cui_banner: str,
    palette: dict[str, tuple[int, int, int]],
    font_family: str,
) -> None:
    """Place centered bold CUI banners at top and bottom of a slide."""
    for y_in in (0.06, 7.08):
        box = slide.shapes.add_textbox(Inches(0.5), Inches(y_in), Inches(12.333), Inches(0.35))
        tf = box.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = cui_banner
        r.font.name = font_family
        r.font.size = Pt(14)
        r.font.bold = True
        r.font.color.rgb = RGBColor(*palette["cui"])


def _add_slide_title(
    slide: Any,
    title_text: str,
    subtitle_text: str,
    palette: dict[str, tuple[int, int, int]],
    font_family: str,
    *,
    is_title_slide: bool = False,
) -> None:
    """Add 40-pt title and optional 30-pt subtitle."""
    top_title = Inches(2.2 if is_title_slide else 0.45)
    t_box = slide.shapes.add_textbox(Inches(0.7), top_title, Inches(11.9), Inches(1.0))
    tf = t_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER if is_title_slide else PP_ALIGN.LEFT
    r = p.add_run()
    r.text = title_text
    r.font.name = font_family
    r.font.size = Pt(40)
    r.font.bold = True
    r.font.color.rgb = RGBColor(*palette["title"])

    if subtitle_text:
        top_sub = Inches(3.45 if is_title_slide else 1.35)
        s_box = slide.shapes.add_textbox(Inches(0.7), top_sub, Inches(11.9), Inches(0.75))
        stf = s_box.text_frame
        stf.word_wrap = True
        sp = stf.paragraphs[0]
        sp.alignment = PP_ALIGN.CENTER if is_title_slide else PP_ALIGN.LEFT
        sr = sp.add_run()
        sr.text = subtitle_text
        sr.font.name = font_family
        sr.font.size = Pt(30)
        sr.font.color.rgb = RGBColor(*palette["accent"])


def _add_bullets_box(
    slide: Any,
    bullets: list[str],
    palette: dict[str, tuple[int, int, int]],
    font_family: str,
    *,
    width_in: float = 11.9,
    top_in: float = 2.0,
) -> None:
    """Add left-justified 24-pt mixed-case bullets (<=7 lines per slide)."""
    box = slide.shapes.add_textbox(Inches(0.7), Inches(top_in), Inches(width_in), Inches(4.7))
    tf = box.text_frame
    tf.word_wrap = True
    for idx, line in enumerate(bullets):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(12)
        r = p.add_run()
        r.text = f"- {line}" if not str(line).lstrip().startswith("-") else str(line)
        r.font.name = font_family
        r.font.size = Pt(24)
        r.font.color.rgb = RGBColor(*palette["text"])


def _add_table_shape(
    slide: Any,
    table_spec: dict[str, Any],
    palette: dict[str, tuple[int, int, int]],
    font_family: str,
    *,
    left_in: float = 0.7,
    top_in: float = 2.1,
    width_in: float = 11.9,
) -> None:
    """Render a native table enforcing AFH 33-337 <=4 rows x <=7 cols limit."""
    rows: list[list[Any]] = list(table_spec.get("rows", []))
    if not rows:
        return
    if len(rows) > 4:
        raise ValueError(
            f"AFH 33-337 Ch. 10 limits slide tables to <=4 rows (got {len(rows)})."
        )
    max_cols = max(len(r) for r in rows)
    if max_cols > 7:
        raise ValueError(
            f"AFH 33-337 Ch. 10 limits slide tables to <=7 columns (got {max_cols})."
        )
    tbl_shape = slide.shapes.add_table(
        len(rows), max_cols, Inches(left_in), Inches(top_in), Inches(width_in), Inches(2.8)
    )
    tbl = tbl_shape.table
    for r_idx, row in enumerate(rows):
        for c_idx in range(max_cols):
            val = _clean_cell_value(row[c_idx]) if c_idx < len(row) else ""
            cell = tbl.cell(r_idx, c_idx)
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            run = p.add_run()
            run.text = val
            run.font.name = font_family
            run.font.size = Pt(20)
            run.font.bold = r_idx == 0
            run.font.color.rgb = RGBColor(*palette["text"])


def _add_chart_shape(
    slide: Any,
    chart_spec: dict[str, Any],
    *,
    left_in: float = 6.5,
    top_in: float = 2.0,
    width_in: float = 6.1,
    height_in: float = 4.6,
) -> None:
    """Render a native bar, line, or pie chart from JSON specification."""
    kind = str(chart_spec.get("type", "bar")).lower()
    xl_type = CHART_MAP.get(kind)
    if xl_type is None:
        raise ValueError(f"Unsupported chart type '{kind}'. Use bar, line, or pie.")
    chart_data = CategoryChartData()
    chart_data.categories = [str(c) for c in chart_spec.get("categories", [])]
    for series in chart_spec.get("series", []):
        chart_data.add_series(str(series.get("name", "Series")), series.get("values", []))
    slide.shapes.add_chart(
        xl_type,
        Inches(left_in),
        Inches(top_in),
        Inches(width_in),
        Inches(height_in),
        chart_data,
    )


def _expand_continued_slides(slides: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Paginate slides exceeding 7 bullet lines into <=7-line Continued slides."""
    expanded: list[dict[str, Any]] = []
    for slide in slides:
        bullets = [str(b) for b in slide.get("bullets", [])]
        if len(bullets) <= 7:
            item = dict(slide)
            if item.get("continued") and "continued" not in str(item.get("title", "")).lower():
                item["title"] = f"{item.get('title', 'Topic')} (Continued)"
            expanded.append(item)
            continue
        base_title = str(slide.get("title", "Topic"))
        for chunk_idx in range(0, len(bullets), 7):
            chunk = bullets[chunk_idx : chunk_idx + 7]
            clone = dict(slide)
            clone["bullets"] = chunk
            clone["title"] = base_title if chunk_idx == 0 else f"{base_title} (Continued)"
            if chunk_idx > 0:
                clone.pop("chart", None)
                clone.pop("table", None)
            expanded.append(clone)
    return expanded


def _populate_notes(slide: Any, notes_text: str) -> None:
    """Write speaker notes (one level deeper than slide bullets)."""
    if not notes_text:
        return
    notes_slide = slide.notes_slide
    notes_tf = notes_slide.notes_text_frame
    notes_tf.text = notes_text


def build_deck(
    spec: dict[str, Any],
    out_path: Path,
    *,
    palette_name: str | None = None,
    dept: str | None = None,
    seal_arg: str | None = None,
    cui: bool = False,
    cui_category: str | None = None,
    cui_ldc: str | None = None,
    cui_poc: str | None = None,
    cui_office: str | None = None,
    style_config_path: str | None = None,
) -> Path:
    """Build a 16:9 AFH 33-337 Ch. 10 PowerPoint deck from a JSON specification."""
    style_cfg = load_style_config(style_config_path)
    resolved_dept = (dept or spec.get("dept") or style_cfg.department_name).strip().lower()
    if resolved_dept == "ask":
        resolved_dept = "dod"
    resolved_service = str(spec.get("service", style_cfg.service)).upper()
    pal_key = (palette_name or spec.get("palette") or resolved_service.lower()).lower()
    palette = PALETTES.get(pal_key, PALETTES["usaf"])
    font_family = str(spec.get("font_family", "Calibri"))

    use_cui = bool(cui or spec.get("cui"))
    cui_cat = cui_category or spec.get("cui_category") or style_cfg.cui_default_category
    cui_ldc_val = cui_ldc or spec.get("cui_ldc") or style_cfg.cui_default_ldc
    cui_banner = build_cui_banner(cui_cat, cui_ldc_val) if use_cui else None
    cui_di = (
        build_cui_di_lines(
            component=dept_terms(resolved_dept)["department"],
            office=cui_office or spec.get("cui_office") or style_cfg.cui_default_office,
            categories=cui_cat,
            ldc=cui_ldc_val,
            poc=cui_poc or spec.get("cui_poc") or style_cfg.cui_default_poc,
            on_letterhead=False,
        )
        if use_cui
        else None
    )
    seal_choice = seal_arg or spec.get("seal") or style_cfg.seal_path or "auto"
    seal_path = resolve_seal(
        seal_choice,
        resolved_dept,
        resolved_service,
        context="slides",
    )

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    raw_slides = _expand_continued_slides(list(spec.get("slides", [])))
    if not raw_slides:
        raw_slides = [{"layout": "title", "title": spec.get("title", "Briefing Title"), "subtitle": spec.get("subtitle", "")}]

    for idx, slide_spec in enumerate(raw_slides):
        slide = prs.slides.add_slide(blank_layout)
        _set_slide_bg(slide, palette["bg"])
        if cui_banner:
            _add_cui_banners(slide, cui_banner, palette, font_family)

        is_title = idx == 0 or slide_spec.get("layout") == "title"
        _add_slide_title(
            slide,
            str(slide_spec.get("title", "")),
            str(slide_spec.get("subtitle", "")),
            palette,
            font_family,
            is_title_slide=is_title,
        )

        if is_title and seal_path and seal_path.exists():
            slide.shapes.add_picture(str(seal_path), Inches(0.6), Inches(0.55), width=Inches(1.1))

        if is_title and cui_banner and cui_di:
            di_box = slide.shapes.add_textbox(
                Inches(7.4), Inches(5.3), Inches(5.3), Inches(1.65)
            )
            dtf = di_box.text_frame
            dtf.word_wrap = True
            p = dtf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            r = p.add_run()
            r.text = "\n".join(cui_di)
            r.font.name = font_family
            r.font.size = Pt(11)
            r.font.color.rgb = RGBColor(*palette["text"])

        bullets = [str(b) for b in slide_spec.get("bullets", [])]
        has_chart = "chart" in slide_spec and bool(slide_spec["chart"])
        has_table = "table" in slide_spec and bool(slide_spec["table"])

        if bullets and has_chart:
            _add_bullets_box(slide, bullets, palette, font_family, width_in=5.5, top_in=2.1)
            _add_chart_shape(slide, slide_spec["chart"], left_in=6.5, top_in=2.0, width_in=6.1)
        elif bullets and has_table:
            _add_bullets_box(slide, bullets, palette, font_family, width_in=11.9, top_in=1.9)
            _add_table_shape(slide, slide_spec["table"], palette, font_family, top_in=4.3)
        elif bullets:
            _add_bullets_box(
                slide,
                bullets,
                palette,
                font_family,
                top_in=4.4 if is_title else 2.1,
            )
        elif has_chart:
            _add_chart_shape(slide, slide_spec["chart"], left_in=1.2, top_in=2.0, width_in=10.9)
        elif has_table:
            _add_table_shape(slide, slide_spec["table"], palette, font_family, top_in=2.1)

        _populate_notes(slide, str(slide_spec.get("notes", "")))

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    return out_path


def get_slides_schema(dept: str = "dod", cui: bool = False) -> dict[str, Any]:
    """Return a Ch. 10 starter briefing deck JSON specification."""
    terms = dept_terms(dept)
    return {
        "dept": terms["choice"],
        "service": "USAF",
        "palette": "usaf",
        "font_family": "Calibri",
        "cui": cui,
        "cui_category": "OPSEC" if cui else "",
        "cui_ldc": "FEDCON" if cui else "None",
        "cui_office": "1 FW/XP" if cui else "",
        "cui_poc": "Maj Smith, DSN 555-0100" if cui else "",
        "slides": [
            {
                "layout": "title",
                "title": "Modernized Staff Routing Proposal",
                "subtitle": f"{terms['department']} · 1st Fighter Wing",
                "notes": "Open with purpose, time limit, and decision requested.",
            },
            {
                "title": "Executive Overview",
                "subtitle": "Bottom Line Up Front",
                "bullets": [
                    "Cut staff routing cycle 40 percent",
                    "Adopt unified digital template suite",
                    "Standardize mandatory JA and FM chops",
                    "Target full cutover by 1 Dec 26",
                ],
                "notes": "Detail: Current average package takes 9.2 duty days across three staff directorates.",
            },
            {
                "title": "Routing Turnaround Comparison",
                "subtitle": "Quarter-by-Quarter Metrics",
                "bullets": [
                    "Legacy manual routing delays approvals",
                    "Automated validation catches format errors early",
                    "Single package spec prevents drift",
                ],
                "chart": {
                    "type": "bar",
                    "categories": ["Q1", "Q2", "Q3", "Q4"],
                    "series": [
                        {"name": "Legacy Days", "values": [10, 9, 11, 10]},
                        {"name": "New Days", "values": [5, 4, 4, 3]},
                    ],
                },
                "notes": "Speak to Q3 spike caused by uncoordinated attachments.",
            },
            {
                "title": "Implementation Milestones",
                "subtitle": "Action Matrix",
                "table": {
                    "rows": [
                        ["Phase", "OPR", "Target Date", "Status"],
                        ["Pilot Wing", "1 FW/XP", "1 Dec 26", "Ready"],
                        ["Command Rollout", "ACC/A6", "15 Jan 27", "Planned"],
                    ]
                },
                "notes": "Approve Phase 1 pilot start on 1 Dec 26.",
            },
        ],
    }


def main(argv: list[str] | None = None) -> int:
    """CLI entrypoint for AFH 33-337 Ch. 10 PowerPoint builder."""
    configure_utf8_stdout()
    ap = argparse.ArgumentParser(description="Build AFH 33-337 Ch. 10 briefing slide decks (.pptx).")
    sub = ap.add_subparsers(dest="cmd", required=True)

    def add_flags(p: argparse.ArgumentParser) -> None:
        p.add_argument("--palette", choices=tuple(PALETTES.keys()))
        p.add_argument("--dept", choices=("dow", "dod"))
        p.add_argument("--seal", help="Seal variant (auto|blue|bw|color|none) or explicit PNG path")
        p.add_argument("--cui", action="store_true")
        p.add_argument("--cui-category", default="")
        p.add_argument("--cui-ldc", default="")
        p.add_argument("--cui-poc", default="")
        p.add_argument("--cui-office", default="")
        p.add_argument("--style-config", help="Optional unit_style.json override")

    p_b = sub.add_parser("build", help="Build a populated .pptx deck from JSON")
    p_b.add_argument("spec", help="Input JSON path")
    p_b.add_argument("--out", required=True, help="Output .pptx path")
    add_flags(p_b)

    p_t = sub.add_parser("template", help="Generate a blank Ch. 10 template deck")
    p_t.add_argument("--out", required=True, help="Output .pptx path")
    add_flags(p_t)

    p_s = sub.add_parser("schema", help="Print starter briefing deck JSON to stdout")
    p_s.add_argument("--dept", choices=("dow", "dod"), default="dod")
    p_s.add_argument("--cui", action="store_true")

    args = ap.parse_args(argv)
    if args.cmd == "schema":
        print(json.dumps(get_slides_schema(args.dept, args.cui), indent=2))
        return 0

    spec = (
        load_json(Path(args.spec))
        if args.cmd == "build"
        else get_slides_schema(args.dept or "dod", args.cui)
    )
    out = build_deck(
        spec,
        Path(args.out),
        palette_name=args.palette,
        dept=args.dept,
        seal_arg=args.seal,
        cui=args.cui,
        cui_category=args.cui_category,
        cui_ldc=args.cui_ldc,
        cui_poc=args.cui_poc,
        cui_office=args.cui_office,
        style_config_path=args.style_config,
    )
    print(f"Built briefing deck -> {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
