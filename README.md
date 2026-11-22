# DAF Tongue and Quill (`daf-tongue-and-quill`)

`daf-tongue-and-quill` is an agentic writing skill and Python CLI toolkit for Department of the Air Force (USAF and USSF) communications. It pairs all 7 Parts and 28 Chapters of **AFH/DAFH 33-337 (*The Tongue and Quill*)** with modern **USSF** terminology and ranks, **2024 USAF Warrant Officers**, **DoDI 5200.48 CUI** markings, selectable **Department of War (`DoW`) vs. Department of Defense (`DoD`)** nomenclature (**EO 14347**), **AFPC** acronym standards, programmatic `.docx` and `.pptx` builders, and multi-format validation.

- **Out of Scope:** Formal military decorations and citations (**DAFI 36-2803**) and adverse administrative actions (**DAFI 36-2907**) follow separate directives and are not drafted or validated here.

## Features

### 1. Drafting & Front Office XO Red-Pen Review
- **Mandatory 6-Step Pre-Drafting Gate ([SKILL.md](SKILL.md)):** Confirms unclassified handling, `DoW` vs. `DoD` nomenclature, document format and USAF/USSF affiliation, target audience with mandatory chops (`JA`, `FM`, `PA`, `A6/S6`, `FSS`) and backward-planned suspense dates, signature authority (`FOR THE COMMANDER`, no `Acting` commanders), and local `unit_style.json` / command acronym lists before drafting.
- **Draft & Build Mode (Workflow A):** Applies **FOCUS** (*Focused, Organized, Clear, Understanding, Supported*), **BLUF** (*Bottom-Line Up Front*), and active voice, inspects JSON schemas first, verifies ranks/salutations/acronyms via `tq_lookup.py`, builds `.docx` or `.pptx` deliverables, and validates every output.
- **XO Red-Pen Review Mode (Workflow B):** Extracts structured text from `.txt`, `.md`, `.docx`, or `.pptx` drafts (`--dump-text`), runs AFH 33-337 Ch. 8 **Three-Pass Editing** (Big Picture/BLUF, Paragraph/Logic, Voice/Mechanics/Format), and triages findings into **Showstoppers**, **Substantive**, and **Polish**.

### 2. Modern DAF, USSF, CUI & DoW/DoD Standards
- **USAF + USSF Coverage:** Enforces capitalized `Airman` and `Guardian`, USSF org echelons (Field Command, Space Delta, Space Base Delta, Garrison, Squadron), installations (`AFB`, `SFB`, `SFS`), USSF enlisted grades (`Spc1`–`Spc4`, `Sgt`, `TSgt`–`CMSgt`, `CMSSF`), and 2024 USAF Warrant Officers (`WO1`, `CW2`–`CW5`).
- **Department of War (`DoW`) vs. Department of Defense (`DoD`):** Supports explicit `--dept dow|dod` selection (`DoW` for internal staff work, `DOW` for AP-style public biographies and `war.gov` brand style; statutory `DoD` preserved for legal, congressional, international, and legacy directive IDs such as `DoDI 5200.48`).
- **DoDI 5200.48 / DAFI 16-1403 CUI Markings:** Adds centered bold `CUI//[CATEGORY]//[LDC]` top/bottom banners on every page or slide, Page 1 / Slide 1 Designation Indicator (DI) blocks (`Controlled by`, `CUI Category`, `Distribution/Dissemination Control`, `POC`), and `--cui-transmittal` notices—while rejecting obsolete `FOUO` and `UNCLASSIFIED//CUI` markings.
- **Evaluations Distinction:** Separates AFH 33-337 Ch. 19 briefing/paper bullets from **DAFI 36-2406** EPB/OPB complete-sentence narrative statements.

### 3. Word (`.docx`) & Staff Package Builder ([scripts/build_doc.py](scripts/build_doc.py))
- **14 Document Types (`build`, `template`, `schema`):** Builds `memo`, `mfr`, `indorsement`, `esss`, `bbp`, `position`, `talking`, `point`, `staff-study`, `bio`, `letter`, `resume`, `minutes`, and `agenda`.
- **5 Memorandum Presets (`--preset`):** Ready-to-edit schemas and blank templates for `appointment`, `policy`, `etp`, `assumption-of-command`, and `aar`.
- **Automatic AFH 33-337 Geometry:** Sets 1-inch margins (balancing short `<10`-line memos to 1.5-inch side margins while pinning the signature block at `4.5"` from the left page edge), suppresses Page 1 numbering on memos, places Copperplate Gothic Bold letterheads with embedded seals, sets `keep_together=True` on ALL CAPS name / mixed-case rank signature blocks, builds Arial `9`/`13.5`-pt Official Biographies with photo placeholder frames, and builds open-punctuation Personal Letters (`1.75"` Page 1 top margin).
- **Synchronized Staff Package Bundler (`package`):** Generates `00_eSSS.docx`, numbered attachment tabs (for example `Tab1_Memo.docx`, `Tab2_BBP.docx`), and a matching `00_Cover_Email.txt` from one package JSON spec so subjects, suspense dates, CUI markings, and tab labels never drift.

### 4. Ch. 10 PowerPoint Briefing Builder ([scripts/build_slides.py](scripts/build_slides.py))
- **16:9 Widescreen Decks:** Supports 5 visual palettes (`usaf`, `ussf`, `dow`, `dod`, `light`), `40`-pt titles, `30`-pt subtitles, and `24`-pt bullets.
- **7×7 Enforcement & Auto Pagination:** Automatically splits dense bullet lists across continuation slides (`<Title> (Continued)`).
- **Data Visuals & Speaker Notes:** Embeds native `bar`, `line`, and `pie` charts, enforces `<=4×7` tables while stripping misleading `.00` trailing zeros, populates presenter notes, embeds title-slide seals, and adds CUI slide banners and DI blocks.

### 5. Multi-Format Linter & Extractor ([scripts/validate_doc.py](scripts/validate_doc.py))
- **Multi-Format Inspection:** Validates `.txt`, `.md`, `.docx`, and `.pptx` (`--strict`, `--format json`, `--dept auto|dow|dod`), plus `--dump-text` structured extraction (headers, body, tables, slides, and notes) for red-pen reviews. Also callable via [scripts/validate_memo.py](scripts/validate_memo.py).
- **Security & Markings:** Hard-stops on potential SSN/EDIPI PII and classified banners/portion marks (`SECRET`, `TOP SECRET`, `(TS)`), flags obsolete `FOUO` and `UNCLASSIFIED//CUI`, and validates CUI LDCs, slide banners, and DI blocks.
- **Formatting & Terminology:** Catches mixed `DoW`/`DoD` usage, `MEMORANDUM FOR:` colons, missing `FROM:`/`SUBJECT:`, mixed date styles (`21 Nov 2026`), `Acting` command titles, >4-line or lowercase signer blocks, Word margin/font/Page-1-numbering errors, Personal Letter punctuation errors, Bio family/marital Privacy Act violations, and `<=4×7` / `.00` slide errors.
- **Acronyms, Attachments & Plain Language:** Verifies `# Attachments:` parity against body references, enforces first-use acronym definitions and flags undefined, single-use, or multi-meaning acronyms, catches informal abbreviations (`coord`, `w/`, `hrs`) in formal prose, and flags Ch. 7 wordiness, smothered verbs, redundancies, AI buzzwords, passive voice, >20-word average sentences, and >7-sentence paragraphs.

### 6. Two-Layer Acronym & Abbreviation System ([scripts/import_acronyms.py](scripts/import_acronyms.py))
- **Layer 1 — Shipped AFPC Minimum Baseline ([references/data/afpc_acronyms.json](references/data/afpc_acronyms.json)):** Includes 173 AFPC DAF-wide baseline acronyms, 5 pre-approved categories (`ranks`, `office_symbols`, `organizations`, `weapons_platforms`, `symbols_measurements`), 25 AFH 33-337 Ch. 26 abbreviations with formality/verb flags, and 20 multi-meaning DAF collision entries (`OPR`, `AO`, `PA`, `CC`, `CD`, etc.).
- **Layer 2 — Command Guide Importer:** Parses MAJCOM, Field Command, or Delta writing guides from `.pdf`, `.docx`, `.html`, `.csv`, `.txt`, or `.md` into `local_acronyms.json` under `--mode additive` (merges atop AFPC) or `--mode strict` (permits only AFPC + the imported guide) while keeping acronyms and abbreviations strictly separated and surfacing meaning collisions.
- **AFPC Update Checker & Sync (`--check-afpc` / `--sync-afpc`):** Checks [AFPC Acronyms](https://www.afpc.af.mil/Career-Management/Acronyms/) on demand (exiting `0` if current, `2` if updates exist, `3` if blocked) and supports `--from-file <saved.html|.pdf|.docx|.txt|.csv>` when `afpc.af.mil` returns HTTP 403 to automated HTTP clients.

### 7. Zero-Bloat Reference Lookups ([scripts/tq_lookup.py](scripts/tq_lookup.py))
Queries structured JSON tables under `references/data/` without loading large markdown guides into context:
- `--rank <query>` — Search USAF, USAF Warrant Officer, USSF, USA, USN, and USMC grades, titles, abbreviations, and salutations.
- `--address <query>` — Search 58 Ch. 15 Forms of Address entries (including paired DoW/DoD executive leadership).
- `--acronym <query>` / `--abbrev <query>` — Check definitions, formality/verb rules, and multi-meaning collisions (layered with `unit_style.json` if provided).
- `--wordy <query>` — Search 76 Ch. 7 wordy phrases, 18 smothered verbs, 30 redundancies, and 14 AI buzzwords.
- `--fallacy <query>` — Search the 14 AFH 33-337 Ch. 5 informal logical fallacies and fixes.
- `--cui <query>` — Search 15 DoDI 5200.48 CUI categories, 8 LDCs, and 6 Distribution Statements.

## Bundled Seals & Unit Style Profile (`--style-config`)

All three official Department of War (`DOW`) PNG seals from the [war.gov Brand Guide](https://www.war.gov/Brand-Guide/#DOW-seal) are bundled directly under `templates/seals/`:

- `dow-seal-blue.png` — default (`--seal auto`) for Word `.docx` memorandums and letters per **AFMAN 33-326** ultramarine blue / black letterhead ink standards.
- `dow-seal-bw.png` — Black & White monochrome variant (`--seal bw`).
- `dow-seal-color.png` — Full Color variant, default (`--seal auto`) for `.pptx` briefing slides (`--seal color` also works for `.docx`).

Override any build with `--seal auto|blue|bw|color|none|<path>`, or initialize a reusable `unit_style.json` profile (`dept`, `service`, letterhead, seal, serial comma, local acronyms, CUI defaults):

```pwsh
python scripts/build_doc.py init-style unit_style.json --dept dow --service USSF
```

See [templates/README.md](templates/README.md) for optional `dod-seal.png`, `daf-seal.png`, `ussf-seal.png`, and CAC-gated AF Portal template placement.

## Repository Architecture & Reference Guides

![daf-tongue-and-quill File Architecture](docs/file-architecture.svg)

| Guide | AFH 33-337 Scope & Key Additions |
| :--- | :--- |
| [references/communication-process.md](references/communication-process.md) | Parts I–III (Ch. 1–9): Plain Language, FOCUS, Seven Steps, 14 fallacies, 8 patterns, BLUF, 3-Pass Editing |
| [references/briefings-speaking-and-meetings.md](references/briefings-speaking-and-meetings.md) | Part IV & Ch. 13 (Ch. 10, 11, 13): Speaking, 7×7 slides, 40/30/20–26 pt type, <=4×7 tables, Listening, Meetings |
| [references/official-memorandum.md](references/official-memorandum.md) | Ch. 13, 14, 23: Memos, MFRs, Indorsements, Minutes, DAFI 33-360 Pubs/Forms, 5 presets |
| [references/staff-papers.md](references/staff-papers.md) | Ch. 16, 17, 19: Point, Talking, BBP, Background, Position, Staff Study, Bullets (vs. DAFI 36-2406 EPB/OPB) |
| [references/staff-packages-and-email.md](references/staff-packages-and-email.md) | Ch. 12 & 18: SSS (`DAF Form 1768`), eSSS, Tabs, Chops, Official E-mail, Social Media, MS Teams/IM, Phone |
| [references/official-biography.md](references/official-biography.md) | Ch. 20: Arial 9/13.5 AP-style Official Bio (Military/Civilian, USAF/USSF), Privacy Act/OPSEC |
| [references/personal-letter.md](references/personal-letter.md) | Ch. 15 & 22: Personal Letters (open punctuation), Forms of Address routing, USPS Pub 28 Envelopes |
| [references/resume-and-cover-letter.md](references/resume-and-cover-letter.md) | Ch. 21: 4 Résumé formats, USAJOBS Federal résumés, military-to-civilian translation & Cover Letters |
| [references/mechanics-and-style.md](references/mechanics-and-style.md) | Part VII (Ch. 24–28): Modern DAF/DoW table, DoDI 5200.48 CUI guide, Acronyms vs. Abbrevs, Dates, Numbers, Caps |

## Installation

### Dependencies
- Python 3.10+ (`python-docx`, `python-pptx`, `pypdf`)

```pwsh
pip install -r requirements.txt
```

### Antigravity / Claude Code Local Skill Registration
Clone the repository into your global skills directory:

```pwsh
git clone --branch main "https://github.com/Blaylock1988/daf-tongue-and-quill.git" "$env:USERPROFILE\.agents\skills\daf-tongue-and-quill"
```

## Quick CLI Examples

```pwsh
# 1. Print starter JSON (including 5 memo presets: appointment, policy, etp, assumption-of-command, aar)
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

# 6. Import your MAJCOM / Field Command / Delta acronym guide, check AFPC updates, & run lookups
python scripts/import_acronyms.py .\command_guide.pdf --out local_acronyms.json --mode strict
python scripts/import_acronyms.py --check-afpc   # or --sync-afpc (use --from-file saved.html if AFPC returns HTTP 403)
python scripts/tq_lookup.py --address "Secretary of War"
python scripts/tq_lookup.py --acronym OPR
python scripts/tq_lookup.py --rank Spc3
python scripts/tq_lookup.py --fallacy "post hoc"
python scripts/tq_lookup.py --cui FEDCON
```

## Running Automated Tests

21 unit tests across 5 modules verify Word/PowerPoint geometry, CUI markings, DoW/DoD checks, acronym/abbreviation rules, update checks, and JSON lookups:

```pwsh
python -m unittest discover -s tests -v
```

## Disclaimer

This skill package is an independent tool and is **not** affiliated with, funded by, or officially endorsed by the Department of War (`DoW`), Department of Defense (`DoD`), Department of the Air Force (`DAF`), or the U.S. Government. Bundled official DoW seals are included solely to support authorized personnel preparing official communications and briefings; they remain governed by **10 U.S.C. § 2260**, **18 U.S.C. § 506**, and the DoW Brand Guide and may **not** be used commercially or in any manner that implies official endorsement.

## License

Original code, scripts, prompts, and documentation in this repository are licensed under the [PolyForm Noncommercial License 1.0.0](LICENSE) (`PolyForm-Noncommercial-1.0.0`). Bundled official U.S. Government seals under `templates/seals/` are excluded from this grant and remain U.S. Government property subject to **10 U.S.C. § 2260**, **18 U.S.C. § 506**, and the DoW Brand Guide.
