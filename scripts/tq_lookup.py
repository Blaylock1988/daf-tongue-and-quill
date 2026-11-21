"""Fast zero-bloat CLI lookup into references/data/*.json."""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from tq_common import (
    build_effective_lexicon,
    configure_utf8_stdout,
    load_forms_data,
    load_ranks_data,
    load_style_config,
    load_style_rules,
)


def lookup_rank(query: str) -> list[dict[str, Any]]:
    """Search USAF, USSF, USAF WO, Army, Navy, and USMC ranks."""
    q = query.strip().lower()
    data = load_ranks_data()
    matches: list[dict[str, Any]] = []
    for service, tiers in data.get("services", {}).items():
        for tier, entries in tiers.items():
            for entry in entries:
                haystack = " ".join(
                    str(entry.get(k, ""))
                    for k in ("grade", "title", "abbrev", "alt_abbrev", "salutation")
                ).lower()
                if q in haystack or q in f"{service} {tier}".lower():
                    matches.append({"service": service, "tier": tier, **entry})
    return matches


def lookup_address(query: str) -> list[dict[str, Any]]:
    """Search Ch. 15 Forms of Address, including paired DoW/DoD titles."""
    q = query.strip().lower()
    data = load_forms_data()
    matches: list[dict[str, Any]] = []
    for entry in data.get("entries", []):
        haystack = " ".join([
            str(entry.get("category", "")),
            str(entry.get("addressee", "")),
            str(entry.get("salutation", "")),
            " ".join(entry.get("address_lines", [])),
        ]).lower()
        if q in haystack:
            matches.append(entry)
    return matches


def lookup_acronym(query: str, style_config_path: str | None) -> dict[str, Any]:
    """Look up an acronym across AFPC, command guide, and collision tables."""
    q_up = query.strip().upper()
    style_cfg = load_style_config(style_config_path)
    lexicon = build_effective_lexicon(style_cfg, style_config_path)
    matches: list[dict[str, Any]] = []
    for term, meanings in lexicon.acronyms.items():
        if q_up == term.upper() or q_up in term.upper() or any(
            q_up in m.upper() for m in meanings
        ):
            coll = lexicon.collisions.get(term, [])
            combined = list(meanings)
            for item in coll:
                if item not in combined:
                    combined.append(item)
            matches.append({
                "term": term,
                "meanings": combined,
                "ambiguous": len(combined) > 1,
                "sources": lexicon.sources.get(term, {}),
            })
    for term, coll in lexicon.collisions.items():
        if q_up == term.upper() and term not in lexicon.acronyms:
            matches.append({
                "term": term,
                "meanings": coll,
                "ambiguous": True,
                "sources": {"collision_table": ", ".join(coll)},
            })
    return {"query": query, "mode": lexicon.mode, "matches": matches}


def lookup_abbrev(query: str, style_config_path: str | None) -> list[dict[str, Any]]:
    """Look up an abbreviation and its formality/verb rules."""
    q = query.strip().lower()
    style_cfg = load_style_config(style_config_path)
    lexicon = build_effective_lexicon(style_cfg, style_config_path)
    matches: list[dict[str, Any]] = []
    for abbr, info in lexicon.abbreviations.items():
        if q in abbr.lower() or q in str(info.get("meaning", "")).lower():
            matches.append({"abbreviation": abbr, **info})
    return matches


def lookup_wordy(query: str) -> list[dict[str, str]]:
    """Search Ch. 7 wordy phrases, smothered verbs, redundancies, and AI words."""
    q = query.strip().lower()
    rules = load_style_rules()
    buckets = ("wordy_phrases", "smothered_verbs", "redundancies", "ai_buzzwords")
    matches: list[dict[str, str]] = []
    for bucket in buckets:
        for phrase, replacement in rules.get(bucket, {}).items():
            if q in phrase.lower() or q in replacement.lower():
                matches.append({
                    "category": bucket,
                    "avoid": phrase,
                    "try": replacement,
                })
    return matches


def lookup_fallacy(query: str) -> list[dict[str, Any]]:
    """Search the 14 AFH 33-337 Ch. 5 informal logical fallacies."""
    q = query.strip().lower()
    rules = load_style_rules()
    matches: list[dict[str, Any]] = []
    for item in rules.get("logical_fallacies", []):
        haystack = " ".join([
            str(item.get("name", "")),
            " ".join(item.get("aliases", [])),
            str(item.get("definition", "")),
        ]).lower()
        if q in haystack:
            matches.append(item)
    return matches


def lookup_cui(query: str) -> dict[str, Any]:
    """Search DoDI 5200.48 CUI Registry categories, LDCs, and distribution statements."""
    q = query.strip().lower()
    rules = load_style_rules()
    cats = [
        {"code": k, **v}
        for k, v in rules.get("cui_categories", {}).items()
        if q in k.lower() or q in str(v.get("name", "")).lower() or q in str(v.get("note", "")).lower()
    ]
    ldcs = [
        {"ldc": k, "rule": v}
        for k, v in rules.get("cui_ldcs", {}).items()
        if q in k.lower() or q in v.lower()
    ]
    dists = [
        {"statement": k, "rule": v}
        for k, v in rules.get("cui_distribution_statements", {}).items()
        if q in k.lower() or q in v.lower()
    ]
    return {"categories": cats, "ldcs": ldcs, "distribution_statements": dists}


def _print_text_results(kind: str, payload: Any) -> None:
    """Format lookup results cleanly for CLI inspection."""
    if kind == "acronym":
        matches = payload.get("matches", [])
        if not matches:
            print("No matching acronym found.")
            return
        for item in matches:
            flag = " [MULTI-MEANING: SPELL OUT OR PIN]" if item["ambiguous"] else ""
            print(f"{item['term']}{flag}: " + " | ".join(item["meanings"]))
        return
    if not payload or (isinstance(payload, dict) and not any(payload.values())):
        print("No matching entries found.")
        return
    print(json.dumps(payload, indent=2, ensure_ascii=False))


def main(argv: list[str] | None = None) -> int:
    """CLI entrypoint for Tongue & Quill structured JSON lookups."""
    configure_utf8_stdout()
    ap = argparse.ArgumentParser(description="Look up DAF Tongue & Quill tables without context bloat.")
    group = ap.add_mutually_exclusive_group(required=True)
    group.add_argument("--rank", help="Rank title, pay grade, abbrev, or service (e.g. Spc3, WO1, Col)")
    group.add_argument("--address", help="Forms of Address query (e.g. 'Secretary of War', 'Ambassador')")
    group.add_argument("--acronym", help="Acronym lookup with multi-meaning collision warnings")
    group.add_argument("--abbrev", help="Standard abbreviation / truncation lookup")
    group.add_argument("--wordy", help="Ch. 7 wordy phrase, smothered verb, redundancy, or AI buzzword")
    group.add_argument("--fallacy", help="Ch. 5 logical fallacy name or keyword")
    group.add_argument("--cui", help="DoDI 5200.48 CUI category, LDC, or Distribution Statement")
    ap.add_argument("--style-config", help="Optional unit_style.json path for command acronyms")
    ap.add_argument("--json", action="store_true", help="Emit raw JSON output")
    args = ap.parse_args(argv)

    if args.rank:
        kind, result = "rank", lookup_rank(args.rank)
    elif args.address:
        kind, result = "address", lookup_address(args.address)
    elif args.acronym:
        kind, result = "acronym", lookup_acronym(args.acronym, args.style_config)
    elif args.abbrev:
        kind, result = "abbrev", lookup_abbrev(args.abbrev, args.style_config)
    elif args.wordy:
        kind, result = "wordy", lookup_wordy(args.wordy)
    elif args.fallacy:
        kind, result = "fallacy", lookup_fallacy(args.fallacy)
    else:
        kind, result = "cui", lookup_cui(args.cui)

    if args.json:
        print(json.dumps(result, indent=2, ensure_ascii=False))
    else:
        _print_text_results(kind, result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
