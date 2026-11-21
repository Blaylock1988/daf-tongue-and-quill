"""Import MAJCOM/Field Command/Delta acronym & abbreviation guides into JSON."""
from __future__ import annotations

import argparse
import copy
import csv
import json
import re
import sys
import urllib.error
import urllib.request
from datetime import date
from html.parser import HTMLParser
from pathlib import Path
from typing import Any

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from tq_common import DATA_DIR, configure_utf8_stdout, load_afpc_data, load_json

AFPC_URL = "https://www.afpc.af.mil/Career-Management/Acronyms/"
AFPC_DATE_RE = re.compile(
    r"(?:(?:updated|effective|current(?:\s+as\s+of)?|last\s+modified)\s*[:\-]?\s*)?"
    r"(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\s+\d{2,4}"
    r"|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\s+\d{1,2},?\s+\d{4})",
    re.IGNORECASE,
)

SECTION_ACRONYMS_RE = re.compile(
    r"^(?:ATTACHMENT\s+\d+\s*[-:]?\s*)?(?:APPROVED\s+)?ACRONYMS(?:\s+AND\s+INITIALISMS)?\s*:?$",
    re.IGNORECASE,
)
SECTION_ABBREV_RE = re.compile(
    r"^(?:ATTACHMENT\s+\d+\s*[-:]?\s*)?(?:APPROVED\s+)?ABBREVIATIONS(?:\s+AND\s+TRUNCATIONS)?\s*:?$",
    re.IGNORECASE,
)
PAREN_ENTRY_RE = re.compile(r"^([A-Za-z0-9&/.\-]{1,20})\s*\(([^()]+)\)\s*$")
SEP_ENTRY_RE = re.compile(
    r"^([A-Za-z0-9&/.\-]{1,20})\s*(?:[-–—:]|\t)\s*(.+?)\s*$"
)
NOISE_LINE_RE = re.compile(
    r"^(?:page\s+\d+|\d+|the following are|approved list|note:)",
    re.IGNORECASE,
)
VERB_TOKENS = frozenset({
    "achieve", "advise", "approve", "assist", "authorize", "brief", "build",
    "cancel", "certify", "collaborate", "complete", "conduct", "coordinate",
    "deliver", "deploy", "develop", "direct", "educate", "establish", "evaluate",
    "execute", "generate", "improve", "increase", "inspect", "manage", "monitor",
    "organize", "prepare", "publish", "recommend", "review", "schedule", "submit",
    "support", "synchronize", "target", "train", "transfer", "update", "verify",
})


class _HtmlLineExtractor(HTMLParser):
    """Extract readable lines and row pairs from HTML pages and tables."""

    _SKIP_TAGS = frozenset({"script", "style", "noscript", "svg"})
    _BLOCK_TAGS = frozenset({"p", "li", "h1", "h2", "h3", "h4", "h5", "h6", "div", "dt", "dd"})

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._skip_depth = 0
        self._in_tr = False
        self._in_cell = False
        self._row_cells: list[str] = []
        self._cell_buf: list[str] = []
        self._block_buf: list[str] = []
        self.lines: list[str] = []

    def _flush_block(self) -> None:
        text = re.sub(r"\s+", " ", "".join(self._block_buf)).strip()
        self._block_buf.clear()
        if text:
            self.lines.append(text)

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        low = tag.lower()
        if low in self._SKIP_TAGS:
            self._skip_depth += 1
            return
        if self._skip_depth:
            return
        if low == "tr":
            self._flush_block()
            self._in_tr = True
            self._row_cells = []
        elif low in {"td", "th"} and self._in_tr:
            self._in_cell = True
            self._cell_buf = []
        elif low == "br":
            if self._in_cell:
                self._cell_buf.append(" ")
            else:
                self._flush_block()
        elif low in self._BLOCK_TAGS and not self._in_tr:
            self._flush_block()

    def handle_endtag(self, tag: str) -> None:
        low = tag.lower()
        if low in self._SKIP_TAGS and self._skip_depth:
            self._skip_depth -= 1
            return
        if self._skip_depth:
            return
        if low in {"td", "th"} and self._in_cell:
            cell = re.sub(r"\s+", " ", "".join(self._cell_buf)).strip()
            self._cell_buf.clear()
            self._in_cell = False
            if cell:
                self._row_cells.append(cell)
        elif low == "tr" and self._in_tr:
            self._in_tr = False
            if len(self._row_cells) >= 2:
                self.lines.append(f"{self._row_cells[0]} - {self._row_cells[1]}")
            elif self._row_cells:
                self.lines.append(self._row_cells[0])
            self._row_cells = []
        elif low in self._BLOCK_TAGS and not self._in_tr:
            self._flush_block()

    def handle_data(self, data: str) -> None:
        if self._skip_depth:
            return
        if self._in_cell:
            self._cell_buf.append(data)
        else:
            self._block_buf.append(data)


def _extract_html_lines(html_text: str) -> list[str]:
    """Parse HTML into cleaned text lines and table pairs."""
    parser = _HtmlLineExtractor()
    parser.feed(html_text)
    parser.close()
    parser._flush_block()
    return [ln for ln in parser.lines if ln.strip()]


def _extract_detected_date(lines: list[str]) -> str:
    """Detect the newest or explicit update date mentioned in source lines."""
    fallback = ""
    for line in lines:
        match = AFPC_DATE_RE.search(line)
        if not match:
            continue
        candidate = re.sub(r"\s+", " ", match.group(1)).strip()
        if re.search(r"\b(?:updated|effective|current|last\s+modified)\b", line, re.IGNORECASE):
            return candidate
        if not fallback:
            fallback = candidate
    return fallback


def _extract_raw_lines(input_path: Path) -> list[str]:
    """Extract text lines from .pdf, .docx, .html, .txt, or .md inputs."""
    suffix = input_path.suffix.lower()
    if suffix == ".pdf":
        from pypdf import PdfReader

        reader = PdfReader(str(input_path))
        chunks = [page.extract_text() or "" for page in reader.pages]
        return [ln.strip() for ln in "\n".join(chunks).splitlines() if ln.strip()]
    if suffix == ".docx":
        import docx

        doc = docx.Document(str(input_path))
        lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                cells = [c.text.strip() for c in row.cells if c.text.strip()]
                if len(cells) >= 2:
                    lines.append(f"{cells[0]} - {cells[1]}")
        return lines
    if suffix in {".html", ".htm"}:
        return _extract_html_lines(input_path.read_text(encoding="utf-8", errors="replace"))
    return [
        ln.strip()
        for ln in input_path.read_text(encoding="utf-8").splitlines()
        if ln.strip()
    ]


def _infer_kind(term: str, section_kind: str | None) -> str:
    """Classify a parsed entry as 'acronym' or 'abbreviation'."""
    if section_kind in {"acronym", "abbreviation"}:
        return section_kind
    clean = term.rstrip(".")
    if "/" in term or term.endswith(".") or (clean.islower() and len(clean) >= 2):
        return "abbreviation"
    if any(ch.islower() for ch in clean) and not re.match(r"^[A-Z][a-z]+[A-Z]", clean):
        return "abbreviation"
    return "acronym"


def _is_verb_capable(meaning: str) -> bool:
    """Detect whether an abbreviation definition supports verb usage."""
    words = {w.strip(".,;:/()").lower() for w in meaning.split()}
    return bool(words & VERB_TOKENS)


def _add_acronym(acronyms: dict[str, list[str]], term: str, meaning: str) -> None:
    """Append a distinct meaning to an acronym entry."""
    clean_m = meaning.strip()
    if not clean_m:
        return
    bucket = acronyms.setdefault(term, [])
    if clean_m not in bucket:
        bucket.append(clean_m)


def _add_abbrev(abbrevs: dict[str, dict[str, Any]], term: str, meaning: str) -> None:
    """Store a structured abbreviation record."""
    clean_m = meaning.strip()
    if not clean_m:
        return
    abbrevs[term] = {
        "meaning": clean_m,
        "formal_allowed": False,
        "verb_capable": _is_verb_capable(clean_m),
        "note": "Command-imported abbreviation; verify formality and suffix rules.",
    }


def parse_csv_guide(input_path: Path) -> tuple[dict[str, list[str]], dict[str, dict[str, Any]]]:
    """Parse a CSV file with columns term/short, meaning, and optional kind."""
    acronyms: dict[str, list[str]] = {}
    abbrevs: dict[str, dict[str, Any]] = {}
    with input_path.open("r", encoding="utf-8", newline="") as handle:
        reader = csv.reader(handle)
        for row in reader:
            if not row or len(row) < 2:
                continue
            term = row[0].strip()
            meaning = row[1].strip()
            if term.lower() in {"term", "short", "acronym", "abbreviation"}:
                continue
            explicit_kind = row[2].strip().lower() if len(row) > 2 else ""
            kind = "abbreviation" if explicit_kind.startswith("abbr") else (
                "acronym" if explicit_kind.startswith("acr") else _infer_kind(term, None)
            )
            if kind == "abbreviation":
                _add_abbrev(abbrevs, term, meaning)
            else:
                _add_acronym(acronyms, term, meaning)
    return acronyms, abbrevs


def parse_text_lines(
    lines: list[str],
) -> tuple[dict[str, list[str]], dict[str, dict[str, Any]]]:
    """Parse structured acronym and abbreviation entries from extracted lines."""
    acronyms: dict[str, list[str]] = {}
    abbrevs: dict[str, dict[str, Any]] = {}
    section_kind: str | None = None

    for raw_line in lines:
        line = re.sub(r"^(?:#+\s*|[-*•]\s+)", "", raw_line).strip()
        if not line:
            continue
        if SECTION_ACRONYMS_RE.match(line):
            section_kind = "acronym"
            continue
        if SECTION_ABBREV_RE.match(line):
            section_kind = "abbreviation"
            continue
        if NOISE_LINE_RE.match(line):
            continue
        match = PAREN_ENTRY_RE.match(line) or SEP_ENTRY_RE.match(line)
        if not match:
            continue
        term, meaning = match.group(1).strip(), match.group(2).strip()
        if not meaning or (len(term) == 1 and not term.isalpha()):
            continue
        if term.lower() in {"term", "short", "acronym", "abbreviation", "symbol", "initialism"}:
            continue
        if _infer_kind(term, section_kind) == "abbreviation":
            _add_abbrev(abbrevs, term, meaning)
        else:
            _add_acronym(acronyms, term, meaning)
    return acronyms, abbrevs


def _write_json_crlf(out_path: Path, payload: dict[str, Any]) -> None:
    """Write formatted UTF-8 JSON with Windows CRLF line endings."""
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8", newline="\r\n") as handle:
        handle.write(json.dumps(payload, indent=2, ensure_ascii=False) + "\n")


def fetch_afpc_lines(url: str = AFPC_URL, timeout: int = 15) -> tuple[list[str], str]:
    """Fetch the live AFPC acronym page HTML and return extracted lines + update date."""
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"},
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            html_text = resp.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        raise RuntimeError(
            f"AFPC fetch blocked or failed (HTTP {exc.code}: {exc.reason}) at {url}. "
            "Save the page in your browser and rerun with --from-file <saved.html|.pdf|.txt>."
        ) from exc
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        raise RuntimeError(
            f"AFPC fetch failed ({exc}) at {url}. "
            "Save the page in your browser and rerun with --from-file <saved.html|.pdf|.txt>."
        ) from exc

    lines = _extract_html_lines(html_text)
    return lines, _extract_detected_date(lines)


def _load_afpc_source(
    from_file: str | None, url: str
) -> tuple[dict[str, list[str]], dict[str, dict[str, Any]], str, str]:
    """Load parsed acronyms/abbreviations from --from-file or live AFPC URL."""
    if from_file:
        src_path = Path(from_file)
        if not src_path.exists():
            raise FileNotFoundError(f"AFPC offline source file not found: {src_path}")
        if src_path.suffix.lower() == ".csv":
            acronyms, abbrevs = parse_csv_guide(src_path)
            detected_date = ""
        else:
            lines = _extract_raw_lines(src_path)
            acronyms, abbrevs = parse_text_lines(lines)
            detected_date = _extract_detected_date(lines)
        return acronyms, abbrevs, detected_date, str(src_path)

    lines, detected_date = fetch_afpc_lines(url)
    acronyms, abbrevs = parse_text_lines(lines)
    return acronyms, abbrevs, detected_date, url


def diff_afpc_baseline(
    parsed_acronyms: dict[str, list[str]],
    baseline: dict[str, Any],
    detected_date: str = "",
) -> dict[str, Any]:
    """Compare parsed AFPC acronyms against the shipped local baseline."""
    base_map: dict[str, list[str]] = {
        str(k): list(v) for k, v in baseline.get("acronyms", {}).items()
    }
    added: dict[str, list[str]] = {}
    new_meanings: dict[str, list[str]] = {}
    unchanged_count = 0

    for term in sorted(parsed_acronyms):
        meanings = parsed_acronyms[term]
        if term not in base_map:
            added[term] = list(meanings)
            continue
        existing_lower = {m.strip().lower() for m in base_map[term]}
        extras = [m for m in meanings if m.strip().lower() not in existing_lower]
        if extras:
            new_meanings[term] = extras
        else:
            unchanged_count += 1

    removed = sorted(k for k in base_map if k not in parsed_acronyms)
    return {
        "source": baseline.get("source", ""),
        "detected_date": detected_date,
        "parsed_count": len(parsed_acronyms),
        "unchanged_count": unchanged_count,
        "added": added,
        "new_meanings": new_meanings,
        "removed": removed,
        "has_changes": bool(added or new_meanings),
    }


def sync_afpc_baseline(
    parsed_acronyms: dict[str, list[str]],
    target_path: Path,
    *,
    detected_date: str = "",
    prune: bool = False,
) -> dict[str, Any]:
    """Merge new AFPC acronyms/meanings into target JSON while preserving metadata."""
    base_data = load_json(target_path) if target_path.exists() else load_afpc_data()
    data = copy.deepcopy(base_data)
    diff = diff_afpc_baseline(parsed_acronyms, data, detected_date=detected_date)

    merged: dict[str, list[str]] = {
        str(k): list(v) for k, v in data.get("acronyms", {}).items()
    }
    for term, meanings in diff["added"].items():
        merged[term] = list(meanings)
    for term, extras in diff["new_meanings"].items():
        bucket = merged.setdefault(term, [])
        for m in extras:
            if m not in bucket:
                bucket.append(m)

    pruned: list[str] = []
    if prune and diff["removed"]:
        for term in diff["removed"]:
            merged.pop(term, None)
        pruned = list(diff["removed"])

    data["acronyms"] = dict(sorted(merged.items()))
    synced_on = date.today().isoformat()
    data["last_synced"] = synced_on
    if detected_date:
        data["detected_source_date"] = detected_date

    _write_json_crlf(target_path, data)
    return {
        "target": str(target_path),
        "last_synced": synced_on,
        "detected_source_date": detected_date,
        "pruned": pruned,
        "total_acronyms": len(data["acronyms"]),
        "diff": diff,
    }


def _print_afpc_diff(diff: dict[str, Any], source_label: str) -> None:
    """Print a concise human-readable summary of AFPC baseline differences."""
    date_note = f" (source date: {diff['detected_date']})" if diff["detected_date"] else ""
    if not diff["has_changes"]:
        print(
            f"AFPC baseline is up to date against {source_label}{date_note} "
            f"({diff['parsed_count']} parsed)."
        )
        return
    print(
        f"AFPC updates detected from {source_label}{date_note}: "
        f"{len(diff['added'])} new acronym(s), "
        f"{len(diff['new_meanings'])} new meaning(s), "
        f"{len(diff['removed'])} omitted from source."
    )
    for term, meanings in diff["added"].items():
        print(f"  + {term}: {' | '.join(meanings)}")
    for term, extras in diff["new_meanings"].items():
        print(f"  ~ {term} (+meaning): {' | '.join(extras)}")


def detect_collisions(
    cmd_acronyms: dict[str, list[str]], baseline: dict[str, Any]
) -> dict[str, list[str]]:
    """Detect terms whose command meanings differ from AFPC or have >1 meaning."""
    base_map: dict[str, list[str]] = dict(baseline.get("acronyms", {}))
    collisions: dict[str, list[str]] = {}
    for term, meanings in cmd_acronyms.items():
        combined: list[str] = list(base_map.get(term, []))
        for m in meanings:
            if m not in combined:
                combined.append(m)
        if len(combined) > 1:
            collisions[term] = combined
    return collisions


def import_command_guide(
    input_path: Path,
    source_name: str,
    effective_date: str,
    mode: str,
) -> dict[str, Any]:
    """Parse an input command guide and return standardized JSON payload."""
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    if input_path.suffix.lower() == ".csv":
        acronyms, abbrevs = parse_csv_guide(input_path)
    else:
        acronyms, abbrevs = parse_text_lines(_extract_raw_lines(input_path))
    baseline = load_afpc_data()
    collisions = detect_collisions(acronyms, baseline)
    return {
        "kind": "command_acronym_list",
        "source": source_name or input_path.name,
        "effective_date": effective_date,
        "mode": mode,
        "acronyms": dict(sorted(acronyms.items())),
        "abbreviations": dict(sorted(abbrevs.items())),
        "collisions": dict(sorted(collisions.items())),
    }


def _run_afpc_action(args: argparse.Namespace) -> int:
    """Execute --check-afpc or --sync-afpc with exit codes 0/2/3."""
    try:
        parsed_acronyms, _, detected_date, source_label = _load_afpc_source(
            args.from_file, args.url
        )
    except RuntimeError as exc:
        if args.json:
            print(json.dumps({"error": str(exc)}, indent=2, ensure_ascii=False))
        else:
            print(str(exc))
        return 3

    target_path = Path(args.out) if args.out else (DATA_DIR / "afpc_acronyms.json")
    if args.sync_afpc:
        summary = sync_afpc_baseline(
            parsed_acronyms,
            target_path,
            detected_date=detected_date,
            prune=args.prune,
        )
        if args.json:
            print(json.dumps(summary, indent=2, ensure_ascii=False))
        else:
            diff = summary["diff"]
            print(
                f"Synced AFPC baseline -> {summary['target']} "
                f"(+{len(diff['added'])} new, ~{len(diff['new_meanings'])} updated, "
                f"total {summary['total_acronyms']})."
            )
        return 0

    baseline = load_json(target_path) if target_path.exists() else load_afpc_data()
    diff = diff_afpc_baseline(parsed_acronyms, baseline, detected_date=detected_date)
    if args.json:
        print(json.dumps(diff, indent=2, ensure_ascii=False))
    else:
        _print_afpc_diff(diff, source_label)
    return 2 if diff["has_changes"] else 0


def main(argv: list[str] | None = None) -> int:
    """CLI entrypoint for command guide imports and AFPC update checks/syncs."""
    configure_utf8_stdout()
    ap = argparse.ArgumentParser(
        description="Import a MAJCOM/Field Command/Delta acronym guide or check/sync AFPC updates."
    )
    ap.add_argument(
        "input",
        nargs="?",
        help="Path to .pdf, .docx, .html, .txt, .md, or .csv writing guide",
    )
    ap.add_argument(
        "--out",
        help="Output JSON path (required for command-guide import; optional target override for --sync-afpc)",
    )
    ap.add_argument("--source", default="", help="Human-readable guide citation")
    ap.add_argument("--effective-date", default="", help="Guide effective date (e.g., 21 Nov 26)")
    ap.add_argument(
        "--mode",
        choices=("additive", "strict"),
        default="additive",
        help="additive merges on top of AFPC; strict permits only AFPC + this list",
    )
    afpc_grp = ap.add_mutually_exclusive_group()
    afpc_grp.add_argument(
        "--check-afpc",
        action="store_true",
        help="Check AFPC baseline against live site (or --from-file); exits 0=current, 2=updates, 3=fetch blocked",
    )
    afpc_grp.add_argument(
        "--sync-afpc",
        action="store_true",
        help="Merge updates from AFPC URL (or --from-file) into references/data/afpc_acronyms.json",
    )
    ap.add_argument("--from-file", help="Offline saved .html/.pdf/.docx/.txt/.csv source for AFPC check/sync")
    ap.add_argument("--url", default=AFPC_URL, help=f"Override AFPC URL (default: {AFPC_URL})")
    ap.add_argument("--prune", action="store_true", help="With --sync-afpc, drop keys missing from source")
    ap.add_argument("--json", action="store_true", help="Emit structured JSON output for check/sync")
    args = ap.parse_args(argv)

    if args.check_afpc or args.sync_afpc:
        return _run_afpc_action(args)

    if not args.input or not args.out:
        ap.error("input guide path and --out are required unless using --check-afpc or --sync-afpc")

    payload = import_command_guide(
        Path(args.input),
        source_name=args.source,
        effective_date=args.effective_date,
        mode=args.mode,
    )
    out_path = Path(args.out)
    _write_json_crlf(out_path, payload)

    print(
        f"Imported {len(payload['acronyms'])} acronyms and "
        f"{len(payload['abbreviations'])} abbreviations -> {out_path}"
    )
    if payload["collisions"]:
        print(
            f"Multi-meaning collisions detected ({len(payload['collisions'])}): "
            + ", ".join(payload["collisions"].keys())
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
