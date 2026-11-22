"""Shared data loaders, CUI helpers, DoW/DoD nomenclature, and style config."""
from __future__ import annotations

import json
import sys
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Sequence

REPO_ROOT: Path = Path(__file__).resolve().parents[1]
DATA_DIR: Path = REPO_ROOT / "references" / "data"
SEALS_DIR: Path = REPO_ROOT / "templates" / "seals"
TRANSMITTAL_NOTICE: str = (
    "When enclosure is removed, this document is Uncontrolled Unclassified Information"
)


@dataclass
class StyleConfig:
    """Per-unit or per-commander style configuration overrides."""

    department_name: str = "ask"
    dow_acronym: str = "DoW"
    service: str = "USAF"
    letterhead_lines: list[str] = field(
        default_factory=lambda: ["DEPARTMENT OF THE AIR FORCE"]
    )
    seal_path: str = ""
    spaces_after_period: int = 2
    serial_comma: bool = False
    bio_ap_states_spelled_out: bool = False
    acronym_list_path: str = ""
    acronym_mode: str = "additive"
    acronym_disambiguations: dict[str, str] = field(default_factory=dict)
    cui_default_office: str = ""
    cui_default_poc: str = ""
    cui_default_category: str = ""
    cui_default_ldc: str = "None"
    unit_approved_acronyms: dict[str, list[str]] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        """Serialize style config to a plain dictionary."""
        return asdict(self)


@dataclass
class EffectiveLexicon:
    """Merged DAF acronym, abbreviation, and collision rules."""

    acronyms: dict[str, list[str]]
    abbreviations: dict[str, dict[str, Any]]
    collisions: dict[str, list[str]]
    sources: dict[str, dict[str, str]]
    mode: str = "additive"
    command_acronyms_only: dict[str, list[str]] = field(default_factory=dict)


def configure_utf8_stdout() -> None:
    """Ensure UTF-8 output on Windows consoles without failing on non-text streams."""
    reconfigure = getattr(sys.stdout, "reconfigure", None)
    if callable(reconfigure):
        try:
            reconfigure(encoding="utf-8")
        except (OSError, ValueError):
            pass


def load_json(path: Path) -> dict[str, Any]:
    """Load and parse a UTF-8 JSON object from disk."""
    raw = path.read_text(encoding="utf-8")
    data = json.loads(raw)
    if not isinstance(data, dict):
        raise ValueError(f"Expected JSON object in {path}")
    return data


def _coerce_acronym_map(raw: Any) -> dict[str, list[str]]:
    """Normalize acronym map values to lists of distinct meaning strings."""
    if not isinstance(raw, dict):
        return {}
    out: dict[str, list[str]] = {}
    for key, val in raw.items():
        if isinstance(val, list):
            out[str(key)] = [str(item) for item in val if str(item).strip()]
        elif isinstance(val, str) and val.strip():
            out[str(key)] = [val.strip()]
    return out


def load_style_config(path: str | Path | None) -> StyleConfig:
    """Load optional unit_style.json overrides or return defaults."""
    if not path:
        return StyleConfig()
    cfg_path = Path(path)
    data = load_json(cfg_path)
    defaults = StyleConfig()
    return StyleConfig(
        department_name=str(data.get("department_name", defaults.department_name)),
        dow_acronym=str(data.get("dow_acronym", defaults.dow_acronym)),
        service=str(data.get("service", defaults.service)).upper(),
        letterhead_lines=[
            str(line) for line in data.get("letterhead_lines", defaults.letterhead_lines)
        ],
        seal_path=str(data.get("seal_path", defaults.seal_path)),
        spaces_after_period=int(
            data.get("spaces_after_period", defaults.spaces_after_period)
        ),
        serial_comma=bool(data.get("serial_comma", defaults.serial_comma)),
        bio_ap_states_spelled_out=bool(
            data.get("bio_ap_states_spelled_out", defaults.bio_ap_states_spelled_out)
        ),
        acronym_list_path=str(data.get("acronym_list_path", defaults.acronym_list_path)),
        acronym_mode=str(data.get("acronym_mode", defaults.acronym_mode)).lower(),
        acronym_disambiguations={
            str(k): str(v)
            for k, v in data.get("acronym_disambiguations", {}).items()
        },
        cui_default_office=str(data.get("cui_default_office", defaults.cui_default_office)),
        cui_default_poc=str(data.get("cui_default_poc", defaults.cui_default_poc)),
        cui_default_category=str(
            data.get("cui_default_category", defaults.cui_default_category)
        ),
        cui_default_ldc=str(data.get("cui_default_ldc", defaults.cui_default_ldc)),
        unit_approved_acronyms=_coerce_acronym_map(data.get("unit_approved_acronyms")),
    )


def load_afpc_data() -> dict[str, Any]:
    """Load the AFPC acronym baseline, abbreviations, and collisions."""
    return load_json(DATA_DIR / "afpc_acronyms.json")


def load_ranks_data() -> dict[str, Any]:
    """Load USAF, USSF, USAF WO, Army, Navy, and USMC rank tables."""
    return load_json(DATA_DIR / "ranks.json")


def load_forms_data() -> dict[str, Any]:
    """Load Ch. 15 Forms of Address and DoW/DoD leadership entries."""
    return load_json(DATA_DIR / "forms_of_address.json")


def load_style_rules() -> dict[str, Any]:
    """Load Ch. 5/7 wordiness, AI buzzwords, fallacies, and CUI metadata."""
    return load_json(DATA_DIR / "style_rules.json")


def _resolve_command_list_path(
    raw_path: str, style_config_path: str | Path | None
) -> Path:
    """Resolve an imported command acronym path against config dir or repo root."""
    candidate = Path(raw_path)
    if candidate.is_absolute() or candidate.exists():
        return candidate
    if style_config_path:
        rel_cfg = Path(style_config_path).resolve().parent / candidate
        if rel_cfg.exists():
            return rel_cfg
    return REPO_ROOT / candidate


def _merge_meaning(target: dict[str, list[str]], term: str, meanings: list[str]) -> None:
    """Add distinct meanings for an acronym term without duplication."""
    existing = target.setdefault(term, [])
    for meaning in meanings:
        if meaning not in existing:
            existing.append(meaning)


def build_effective_lexicon(
    style_cfg: StyleConfig, style_config_path: str | Path | None = None
) -> EffectiveLexicon:
    """Merge AFPC baseline, command guide JSON, and unit style acronyms."""
    afpc = load_afpc_data()
    acronyms: dict[str, list[str]] = _coerce_acronym_map(afpc.get("acronyms"))
    abbreviations: dict[str, dict[str, Any]] = {
        str(k): dict(v) for k, v in afpc.get("abbreviations", {}).items()
    }
    collisions: dict[str, list[str]] = _coerce_acronym_map(
        afpc.get("multi_meaning_collisions")
    )
    sources: dict[str, dict[str, str]] = {
        k: {"baseline": ", ".join(v)} for k, v in acronyms.items()
    }
    command_only: dict[str, list[str]] = {}
    mode = style_cfg.acronym_mode

    if style_cfg.acronym_list_path:
        cmd_path = _resolve_command_list_path(
            style_cfg.acronym_list_path, style_config_path
        )
        if not cmd_path.exists():
            raise FileNotFoundError(f"Command acronym list not found: {cmd_path}")
        cmd_data = load_json(cmd_path)
        mode = str(cmd_data.get("mode", mode)).lower()
        cmd_acronyms = _coerce_acronym_map(cmd_data.get("acronyms"))
        cmd_source = str(cmd_data.get("source", cmd_path.name))
        for term, meanings in cmd_acronyms.items():
            command_only[term] = list(meanings)
            _merge_meaning(acronyms, term, meanings)
            sources.setdefault(term, {})[cmd_source] = ", ".join(meanings)
            if len(acronyms[term]) > 1:
                collisions[term] = list(acronyms[term])
        for abbr, info in cmd_data.get("abbreviations", {}).items():
            abbreviations[str(abbr)] = dict(info)

    for term, meanings in style_cfg.unit_approved_acronyms.items():
        _merge_meaning(acronyms, term, meanings)
        sources.setdefault(term, {})["unit_style"] = ", ".join(meanings)
        if len(acronyms[term]) > 1:
            collisions[term] = list(acronyms[term])

    return EffectiveLexicon(
        acronyms=acronyms,
        abbreviations=abbreviations,
        collisions=collisions,
        sources=sources,
        mode=mode,
        command_acronyms_only=command_only,
    )


def normalize_cui_categories(categories: str | Sequence[str] | None) -> list[str]:
    """Split, uppercase, deduplicate, and alphabetically sort CUI categories."""
    if not categories:
        return []
    tokens: list[str] = []
    raw_items = [categories] if isinstance(categories, str) else list(categories)
    for item in raw_items:
        for part in str(item).replace(",", "/").split("/"):
            clean = part.strip().upper()
            if clean and clean not in {"BASIC", "CUI", "CUI BASIC"}:
                tokens.append(clean)
    return sorted(set(tokens))


def build_cui_banner(
    categories: str | Sequence[str] | None = None, ldc: str | None = None
) -> str:
    """Build DoDI 5200.48 banner string (e.g., CUI, CUI//CTI//FEDCON)."""
    cats = normalize_cui_categories(categories)
    clean_ldc = (ldc or "").strip().upper()
    parts = ["CUI"]
    if cats:
        parts.append("/".join(cats))
    if clean_ldc and clean_ldc != "NONE":
        parts.append(clean_ldc)
    return "//".join(parts)


def build_cui_di_lines(
    component: str | None,
    office: str | None,
    categories: str | Sequence[str] | None,
    ldc: str | None,
    poc: str | None,
    on_letterhead: bool = True,
) -> list[str]:
    """Build the Page 1 / Cover Slide CUI Designation Indicator block lines."""
    lines: list[str] = []
    if not on_letterhead or component:
        lines.append(f"Controlled by: {component or '[DoD/DoW Component]'}")
    lines.append(f"Controlled by: {office or '[Originating Office]'}")
    cats = normalize_cui_categories(categories)
    lines.append(f"CUI Category: {', '.join(cats) if cats else '[Category]'}")
    clean_ldc = (ldc or "").strip()
    lines.append(
        f"Distribution/Dissemination Control: {clean_ldc if clean_ldc else 'None'}"
    )
    lines.append(f"POC: {poc or '[Name, Phone / Org Email]'}")
    return lines


def dept_terms(dept: str | None, ap_style: bool = False) -> dict[str, str]:
    """Return Department of War or Department of Defense nomenclature strings."""
    choice = (dept or "dod").strip().lower()
    if choice == "dow":
        return {
            "choice": "dow",
            "department": "Department of War",
            "acronym": "DOW" if ap_style else "DoW",
            "secretary": "Secretary of War",
            "deputy_secretary": "Deputy Secretary of War",
            "sec_short": "SecWar",
            "domain": "war.gov",
        }
    return {
        "choice": "dod",
        "department": "Department of Defense",
        "acronym": "DOD" if ap_style else "DoD",
        "secretary": "Secretary of Defense",
        "deputy_secretary": "Deputy Secretary of Defense",
        "sec_short": "SecDef",
        "domain": "defense.gov",
    }


def resolve_seal(
    seal_arg: str | None,
    dept: str | None = None,
    service: str = "USAF",
    *,
    context: str = "doc",
    repo_root: Path = REPO_ROOT,
) -> Path | None:
    """Resolve a seal variant ('auto'|'blue'|'bw'|'color'|'none') or explicit path."""
    if not seal_arg:
        return None
    normalized = seal_arg.strip().lower()
    if normalized in {"none", "off"}:
        return None

    seals_dir = repo_root / "templates" / "seals"
    if normalized in {"blue", "bw", "color"}:
        for name in (f"dow-seal-{normalized}.png", "dow-seal.png"):
            candidate = seals_dir / name
            if candidate.exists():
                return candidate
        return None

    if normalized != "auto":
        explicit = Path(seal_arg)
        if not explicit.exists():
            raise FileNotFoundError(f"Seal image not found: {explicit}")
        return explicit

    candidates: list[str] = []
    dept_norm = (dept or "").strip().lower()
    if dept_norm == "dow":
        if context.strip().lower() == "slides":
            candidates.extend(["dow-seal-color.png", "dow-seal-blue.png", "dow-seal-bw.png", "dow-seal.png"])
        else:
            candidates.extend(["dow-seal-blue.png", "dow-seal-bw.png", "dow-seal-color.png", "dow-seal.png"])
    else:
        candidates.append("dod-seal.png")
        svc_norm = service.strip().lower()
        if svc_norm == "ussf":
            candidates.extend(["ussf-seal.png", "daf-seal.png"])
        else:
            candidates.extend(["daf-seal.png", "dod-seal.png"])

    for name in candidates:
        candidate = seals_dir / name
        if candidate.exists():
            return candidate
    return None
