# DAF Tongue and Quill (`daf-tongue-and-quill`)

A comprehensive agentic skill package covering all 7 Parts and all 28 Chapters of **AFH/DAFH 33-337, *The Tongue and Quill***, paired with authoritative modern **Department of the Air Force (USAF + USSF)** updates, **DoDI 5200.48 / DAFI 16-1403 CUI** marking support, **AFPC DAF-wide acronym baseline + command guide importing**, **EO 14347 Department of War (`DoW`) vs. Department of Defense (`DoD`) user selection**, programmatic **`.docx` / `.pptx` generation**, and multi-format validation.

## Repository File Architecture

![daf-tongue-and-quill File Architecture](docs/file-architecture.svg)

## Installation

### Dependencies
- Python 3.10+ (`python-docx`, `python-pptx`, `pypdf`)

```pwsh
pip install -r requirements.txt
```

### Antigravity / Claude Code Local Skill Registration
Clone or sync the repository root into your global skills directory:

```pwsh
git clone --branch main "https://github.com/Blaylock1988/daf-tongue-and-quill.git" "$env:USERPROFILE\.agents\skills\daf-tongue-and-quill"
```

## Quick CLI Examples

```pwsh
# 1. Print starter JSON (including 5 high-frequency memo presets: appointment, policy, etp, assumption-of-command, aar)
python scripts/build_doc.py schema memo --preset etp --dept dow --cui

# 2. Build a CUI-marked Official Memorandum (.docx) using Department of War (EO 14347) styling
python scripts/build_doc.py build draft.json --out memo.docx --dept dow --cui --cui-category CTI --cui-ldc FEDCON --seal auto

# 3. Bundle a synchronized Staff Package (00_eSSS.docx + Tabs + 00_Cover_Email.txt)
python scripts/build_doc.py package pkg.json --out-dir .\out\pkg --dept dow

# 4. Generate a Ch. 10 16:9 PowerPoint briefing deck
python scripts/build_slides.py schema --dept dow --cui > deck.json
python scripts/build_slides.py build deck.json --out brief.pptx --palette ussf --dept dow --cui

# 5. Validate .txt, .md, .docx, or .pptx (or dump structured text for XO Red-Pen Review)
python scripts/validate_doc.py memo.docx --dept dow --strict
python scripts/validate_doc.py brief.pptx --type slides --dump-text

# 6. Import your MAJCOM / Field Command / Delta acronym writing guide & run unit lookups
python scripts/import_acronyms.py .\command_guide.pdf --out local_acronyms.json --mode strict
python scripts/import_acronyms.py --check-afpc   # or --sync-afpc (use --from-file saved.html if AFPC returns HTTP 403)
python scripts/tq_lookup.py --address "Secretary of War"
python scripts/tq_lookup.py --acronym OPR
python scripts/tq_lookup.py --rank Spc3
```

## Local Seals & Unit Style Profile (`--style-config`)

Generate a unit profile so every draft inherits your commander's preferences automatically:

```pwsh
python scripts/build_doc.py init-style unit_style.json --dept dow --service USSF
```

See [templates/README.md](templates/README.md) for placing local `dow-seal.png`, `dod-seal.png`, `daf-seal.png`, or `ussf-seal.png` assets under `templates/seals/` (`.gitignore`d).

## Running Automated Tests

```pwsh
python -m unittest discover -s tests -v
```
