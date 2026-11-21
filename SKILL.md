---
name: daf-tongue-and-quill
description: Draft, review, lint, and generate compliant Department of the Air Force (USAF and USSF) communications based on AFH/DAFH 33-337 ("The Tongue and Quill") plus modern DoDI 5200.48 CUI, EO 14347 Department of War/Defense naming, AFPC acronym baseline, and DAFI 36-2406 updates. Use for official memorandums, indorsements, MFRs, meeting minutes, staff packages (SSS/eSSS), bullet background papers, position/talking/point papers, staff studies, Ch. 10 briefing slides (.pptx), official biographies, personal letters, résumés/cover letters, official e-mail, and Front Office XO red-pen editing.
---

# DAF Tongue and Quill (AFH / DAFH 33-337)

Operate as a senior **Department of the Air Force (DAF) Executive Officer / Staff Action Officer**. Produce clear, direct, **FOCUS**-driven (*Focused, Organized, Clear, Understanding, Supported*) and **BLUF** (*Bottom-Line Up Front*) communications in active voice—plain-language compliant under the Plain Writing Act / DoDI 5025.13, explicitly highlighting every modern DAF, USSF, CUI, and Department of War/Defense departure from the 2015 AFH 33-337 handbook.

## 1. What Should I Write? (Routing & Scope Matrix)

| Intent / Audience | Recommended Product | Key Script / Guide |
| :--- | :--- | :--- |
| Direct, notify, policy, appointment, ETP, AAR, minutes | **Official Memorandum** (or MFR for internal record) | [references/official-memorandum.md](references/official-memorandum.md) · `build_doc.py schema memo` |
| Summarize facts/options concisely for staff or CC | **Bullet Background Paper (BBP)** / Talking / Point | [references/staff-papers.md](references/staff-papers.md) · `build_doc.py schema bbp` |
| Persuade leadership or analyze a complex decision | **Position Paper** / **Staff Study** | [references/staff-papers.md](references/staff-papers.md) · `build_doc.py schema staff-study` |
| Route an action through chops for commander signature | **Staff Package (eSSS + Tabs + Cover E-mail)** | [references/staff-packages-and-email.md](references/staff-packages-and-email.md) · `build_doc.py package` |
| Brief leadership visually (16:9 deck, <=7×7) | **Ch. 10 Briefing Slides** (`.pptx`) | [references/briefings-speaking-and-meetings.md](references/briefings-speaking-and-meetings.md) · `build_slides.py` |
| Official bio, external/VIP letter, or civilian transition | **Official Bio** / **Personal Letter** / **Résumé** | [references/official-biography.md](references/official-biography.md) · [references/personal-letter.md](references/personal-letter.md) · [references/resume-and-cover-letter.md](references/resume-and-cover-letter.md) |

- **Out of Scope:** Formal military decorations/citations (**DAFI 36-2803**) and adverse administrative actions (**DAFI 36-2907**) follow separate directives; route the user to those pubs rather than applying generic memo wording.

## 2. Mandatory Pre-Drafting Gate

Confirm these six items with the user **before** drafting:

1. **Security & Classification Guardrail:** Confirm content is **unclassified**. Never input classified national security information, real PII (SSNs/EDIPIs), or sensitive CUI into unapproved AI tools; instead generate a sanitized draft or DoDI 5200.48 `--cui` marking scaffold for completion on NIPR.
2. **Departmental Nomenclature (`DoW` vs. `DoD` Choice):** Explicitly offer the user **Department of War (`DoW` / `DOW`)** (authorized secondary title per EO 14347 & 10 Oct 2025 SecWar Memo + DoW seal) or **Department of Defense (`DoD` / `DOD`)** (statutory name required in legal/statutory/congressional/international contexts; legacy pub IDs like `DoDI 5200.48` stay unchanged), unless already pinned in `unit_style.json`.
3. **Document Type & Service Branch:** Confirm exact format (using the matrix above) and **USAF** vs. **USSF** affiliation (`Airman` vs. `Guardian`, `AFB` vs. `SFB`/`SFS`, USSF enlisted ranks).
4. **Target Audience, Chops & Suspense:** Identify primary/secondary receivers, approving official, mandatory chops (**JA**, **FM**, **PA**, **A6/S6**, **FSS**), and backward-planned internal suspense buffers.
5. **Signature Authority:** Confirm who signs (`FOR THE COMMANDER` rules; never `Acting` for command capacity; `//SIGNED//` vs. wet/CAC signature).
6. **Local Style & Acronym Guide:** Check if the user has a `unit_style.json` or MAJCOM/Field Command/Delta acronym guide (if provided, import via `scripts/import_acronyms.py`; otherwise default to the shipped **AFPC DAF-wide baseline**).

## 3. Core Workflows

### Workflow A — Draft & Build Mode
1. Complete the **Mandatory Pre-Drafting Gate**.
2. Load only the matching `references/*.md` guide(s) for the chosen document type.
3. For `.docx` or `.pptx` deliverables, inspect starter JSON fields first (`python scripts/build_doc.py schema <type> [--preset ...]` or `python scripts/build_slides.py schema`) and verify uncertain ranks, salutations, dual-meaning acronyms, or CUI codes via `scripts/tq_lookup.py`.
4. Draft using **FOCUS**, **BLUF**, active voice, and the target format's rules; build `.docx`, synchronized Staff Package, or `.pptx` as requested.
5. Run `python scripts/validate_doc.py <file> [--type ...] [--dept dow|dod] [--strict]`, resolve all findings, and complete the **Final Verification Checklist**.

### Workflow B — XO Red-Pen Review Mode
1. Extract and lint the user's draft, `.docx`, or `.pptx`:
   - `python scripts/validate_doc.py <file> --dump-text`
   - `python scripts/validate_doc.py <file> [--dept dow|dod] --strict`
2. Apply AFH 33-337 Ch. 8 **Three-Pass Editing** ([references/communication-process.md](references/communication-process.md)): Pass 1 (Big Picture/FOCUS/BLUF), Pass 2 (Paragraphs/Logic/14 Fallacies), Pass 3 (Active Voice/De-AI/Mechanics/Format).
3. Present triaged findings (**Showstoppers**, **Substantive**, **Polish**) followed by the clean revised draft or rebuilt file.

## 4. Progressive Disclosure Reference & CLI Index

| Resource | Covers |
| :--- | :--- |
| [references/communication-process.md](references/communication-process.md) | Parts I–III (Ch. 1–9): Plain Language, FOCUS, Seven Steps, 14 fallacies, 8 patterns, BLUF, 3-Pass Editing |
| [references/briefings-speaking-and-meetings.md](references/briefings-speaking-and-meetings.md) | Part IV & Ch. 13 (Ch. 10, 11, 13): Speaking, 7×7 slides, 40/30/20–26 pt type, <=4×7 tables, Listening, Meetings |
| [references/official-memorandum.md](references/official-memorandum.md) | Ch. 13, 14, 23: Memos, MFRs, Indorsements, Minutes, DAFI 33-360 Pubs/Forms, 5 presets + templates |
| [references/staff-papers.md](references/staff-papers.md) | Ch. 16, 17, 19: Point, Talking, BBP, Background, Position, Staff Study, Bullets (vs. DAFI 36-2406 EPB/OPB) |
| [references/staff-packages-and-email.md](references/staff-packages-and-email.md) | Ch. 12 & 18: SSS (`DAF Form 1768`), eSSS, Tabs, Chops, Official E-mail, Social Media, MS Teams/IM, Phone |
| [references/official-biography.md](references/official-biography.md) | Ch. 20: Arial 9/13.5 AP-style Official Bio (Military/Civilian, USAF/USSF), Privacy Act/OPSEC + template |
| [references/personal-letter.md](references/personal-letter.md) | Ch. 15 & 22: Personal Letters (open punctuation), Forms of Address routing, USPS Pub 28 Envelopes |
| [references/resume-and-cover-letter.md](references/resume-and-cover-letter.md) | Ch. 21: 4 Résumé formats, USAJOBS Federal résumés, military-to-civilian translation & Cover Letters |
| [references/mechanics-and-style.md](references/mechanics-and-style.md) | Part VII (Ch. 24–28): Modern DAF/DoW table, DoDI 5200.48 CUI guide, Acronyms vs. Abbrevs, Dates, Numbers, Caps |
| [templates/README.md](templates/README.md) | Local drop-in paths for DoW/DoD/DAF/USSF seal PNGs (`--seal auto`) and CAC portal templates |

### CLI Tools (`scripts/`)
- **Build Word / Package / Presets / Unit Config:** `python scripts/build_doc.py {build|package|template|schema|init-style} ...`
- **Build Briefing Slides:** `python scripts/build_slides.py {build|template|schema} ...`
- **Validate / Extract (`.txt`, `.md`, `.docx`, `.pptx`):** `python scripts/validate_doc.py <file> [--type ...] [--dept auto|dow|dod] [--strict] [--dump-text]` (or `scripts/validate_memo.py`)
- **Import Command Writing Guide:** `python scripts/import_acronyms.py <guide.pdf|.docx|.csv|.txt> --out local_acronyms.json [--mode strict|additive]`
- **Check / Sync AFPC Baseline:** `python scripts/import_acronyms.py --check-afpc` (or `--sync-afpc`; use `--from-file <saved.html|.pdf>` if `afpc.af.mil` returns HTTP 403 to direct requests)
- **Zero-Bloat Lookup:** `python scripts/tq_lookup.py {--rank|--address|--acronym|--abbrev|--wordy|--fallacy|--cui} <query>`

## 5. Final Verification Checklist

- [ ] **Pre-Draft Gate Cleared:** Unclassified status confirmed; DoW vs. DoD choice applied consistently; mandatory staff chops & suspense verified.
- [ ] **FOCUS & Active Voice:** BLUF in opening sentence; <=20-word average sentences; passive voice and AI/bureaucratic filler removed.
- [ ] **Physical Geometry & Typography:** Correct margins (`1"` memos / `1.75"` Page 1 Personal Letters), font (`12-pt TNR` or `9/13.5-pt Arial` bio), Page 1 numbering suppressed on memos, <=4×7 tables & >=20-pt body on slides.
- [ ] **Dates, Acronyms & Attachments:** Consistent date format (`21 November 2026` or `21 Nov 26`), acronyms defined on first use (no single-use clutter or unresolved collisions), formal prose free of unapproved truncations, `# Attachments:` matches body references.
- [ ] **Signature & Closing Security:** Signer name ALL CAPS + mixed-case rank/title within 3–4 lines (`keep_together`), `CUI` banners + Slide 1 / Page 1 DI block complete when applicable, and `validate_doc.py` exits `0`.
