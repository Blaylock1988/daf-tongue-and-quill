"""Thin CLI wrapper invoking validate_doc.py defaulting to --type memo."""
from __future__ import annotations

import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import validate_doc


def main(argv: list[str] | None = None) -> int:
    """Run validate_doc with --type memo unless caller overrides --type."""
    raw = list(sys.argv[1:] if argv is None else argv)
    if "--type" not in raw and "--dump-text" not in raw and "-h" not in raw and "--help" not in raw:
        raw.extend(["--type", "memo"])
    return validate_doc.main(raw)


if __name__ == "__main__":
    raise SystemExit(main())
