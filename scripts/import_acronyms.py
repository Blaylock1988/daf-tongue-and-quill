"""Import MAJCOM/Field Command/Delta acronym & abbreviation guides into JSON."""
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from pathlib import Path
from typing import Any

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from tq_common import configure_utf8_stdout, load_afpc_data

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


def _extract_raw_lines(input_path: Path) -> list[str]:
    """Extract text lines from .pdf, .docx, .txt, or .md inputs."""
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
        if not meaning or len(term) == 1 and not term.isalpha():
            continue
        if _infer_kind(term, section_kind) == "abbreviation":
            _add_abbrev(abbrevs, term, meaning)
        else:
            _add_acronym(acronyms, term, meaning)
    return acronyms, abbrevs


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


def main(argv: list[str] | None = None) -> int:
    """CLI entrypoint for importing a command writing guide acronym list."""
    configure_utf8_stdout()
    ap = argparse.ArgumentParser(
        description="Import a MAJCOM/Field Command/Delta acronym & abbreviation guide into JSON."
    )
    ap.add_argument("input", help="Path to .pdf, .docx, .txt, .md, or .csv writing guide")
    ap.add_argument("--out", required=True, help="Output JSON path (e.g., local_acronyms.json)")
    ap.add_argument("--source", default="", help="Human-readable guide citation")
    ap.add_argument("--effective-date", default="", help="Guide effective date (e.g., 21 Nov 26)")
    ap.add_argument(
        "--mode",
        choices=("additive", "strict"),
        default="additive",
        help="additive merges on top of AFPC; strict permits only AFPC + this list",
    )
    args = ap.parse_args(argv)

    payload = import_command_guide(
        Path(args.input),
        source_name=args.source,
        effective_date=args.effective_date,
        mode=args.mode,
    )
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8", newline="\r\n") as handle:
        handle.write(json.dumps(payload, indent=2, ensure_ascii=False) + "\n")

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
