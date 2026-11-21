# Writing Mechanics, Style, CUI & Modern DAF/DoW Departures (Part VII, Ch. 24–28)

## 1. Authoritative Modern DAF / DoW Departures (`[MODERN DAF UPDATE]`)

The 2015 AFH 33-337 PDF (certified 2016) predates the U.S. Space Force, 2024 USAF Warrant Officers, DoDI 5200.48 CUI, EO 14347, and AFPC's DAF-wide acronym list:

| Topic | 2015 AFH 33-337 (PDF) | Authoritative Modern Standard | Governing Source |
| :--- | :--- | :--- | :--- |
| **Department of War vs. Department of Defense** | `Department of Defense (DoD)` / `Secretary of Defense` only | Authorizes **`Department of War`** (`DoW` internal / `DOW` public & `war.gov` brand style) + **`Secretary of War`** (`SecWar`) + new **DoW seal** as official secondary titles for executive comms, briefings & memos; **`Department of Defense (DoD)`** remains statutory name (legal, statutory, congressional, international docs & published directive IDs like `DoDI 5200.48`). **Offer users explicit choice.** | **EO 14347** (5 Sep 2025); **SecWar Memo** (10 Oct 2025); `war.gov` Brand Guide |
| **USAF + USSF Scope** | Air Force only (`USAF`, `Airman`, MAJCOM/Wing/Sq, AB/AFB/JB) | Covers **USAF + USSF**; capitalize **`Airman`** and **`Guardian`**; USSF org echelons (Field Command, Space Delta, Space Base Delta, Garrison, Squadron) & installations (**`SFB`**, **`SFS`**); signature affiliation `USSF`; letterhead Line 1 stays `DEPARTMENT OF THE AIR FORCE`. | AFMAN 33-326 (31 Jul 2019); `spaceforce.mil` |
| **USSF Enlisted Ranks** | Absent | `Spc1`–`Spc4` (*Specialist 1–4*; spoken/written salutation: `Specialist`), `Sgt` (E-5 *Sergeant*), `TSgt`–`CMSgt`, `CMSSF` (`Chief`). | `spaceforce.mil` |
| **USAF Warrant Officers** | Ch. 15 table leaves USAF W-1–W-5 blank | Reintroduced 2024 (IT/Cyber): **`WO1`** (W-1), **`CW2`–`CW5`** (W-2–W-5; check local DAFGM/myFSS for `CW2` vs `CWO2`); commissioned warrant/officers (`Chief Warrant Officer` / `Sir` / `Ma'am`). | 2024 DAF WO Policy; `af.mil` |
| **DAF Acronym Baseline** | Cites `JP 1-02` & unit rules | Shipped minimum baseline: **AFPC Acronym & Abbreviation List** (148 USAF/USSF entries + 5 pre-approved categories), layered under MAJCOM/Field Command/Delta guides via `import_acronyms.py`. | **AFPC Acronyms** (`afpc.af.mil/Career-Management/Acronyms/`) |
| **CUI Replacing FOUO** | Cites **FOUO**, `AFI 31-401`, `DoD 5200.1-R` | **FOUO is obsolete** for new documents; replaced by **DoDI 5200.48 / DAFI 16-1403 CUI**. Classified info governed by **DoDM 5200.01** & **DAFMAN 16-1404**. | **DoDI 5200.48**; **DAFI 16-1403**; **32 CFR 2002** |
| **Evaluations (Ch. 19)** | Compressed EPR/OPR bullets | **DAFI 36-2406** **EPBs/OPBs** use complete-sentence **narrative statements** (~350 chars, Action + Impact, ALQs). Ch. 19 bullets still apply to BBPs & Talking/Point Papers. | **DAFI 36-2406** |
| **Updated Pubs & Media** | `AFMAN 33-326 (2011)`, `AFMAN 33-152`, `AFI 51-604`, Google+/Vine | **AFMAN 33-326** (31 Jul 2019), **DAFMAN 17-1301**, **DAFI 51-509**, **DAFI 33-360**, **DAFMAN 33-363**, **DAFI 33-332**, **DAFI 1-1**, **DAF Form 1768** / TMT & **ETMS2**, **DoD Dictionary**, 16:9 widescreen slides. | DAF e-Publishing |

## 2. DoDI 5200.48 / DAFI 16-1403 CUI Marking Quick-Reference

- **Banner (Top & Bottom of Every Page/Slide):** Bold, centered, ALL CAPS **`CUI`** (DoD omits *"CUI Basic"* labels; never write `UNCLASSIFIED//CUI` or `FOUO`).
  - Full syntax: **`CUI//[CATEGORY(IES)]//[LDC]`** — categories sorted alphabetically joined by `/` (e.g., `CUI//CTI/EXPT//FEDCON`); CUI Specified adds `SP-` where required (`CUI//SP-EXPT`).
- **Designation Indicator (DI) Block (Page 1 / Slide 1 Lower-Right Only):**
  - `Controlled by:` Component *(omit if on official letterhead)*
  - `Controlled by:` Originating Office
  - `CUI Category:` Registry abbreviations (`CTI`, `PRVCY`, `EXPT`, `OPSEC`, `PROPIN`, etc.)
  - `Distribution/Dissemination Control:` `FED ONLY`, `FEDCON`, `NOCON`, `DL ONLY`, `NOFORN`, `REL TO ...`, `RELIDO`, Distribution Statement `B–F`, or `None`
  - `POC:` Name, phone, and/or org email
- **Portion Markings & Transmittals:** Optional on pure unclassified CUI unless locally required, but all-or-nothing when used (`(CUI)`, `(CUI//...)`, `(U)`); mandatory when commingled with CNSI (banner reflects highest classification). Unclassified transmittals forwarding CUI enclosures carry `CUI` banners + `"When enclosure is removed, this document is Uncontrolled Unclassified Information"` and omit the DI block if attachments are marked.
- **Query Registry Codes:** `python scripts/tq_lookup.py --cui FEDCON`

## 3. Acronyms vs. Abbreviations (Ch. 26, pp. 331–340)

- **Acronyms / Initialisms (`AOR`, `OPSEC`, `AFFORGEN`, `DoW`):** Pronounceable words or letter-by-letter brevity codes written in ALL CAPS (unless standardized mixed case like `DoW`, `DoD`, `QoL`, `Spc1`).
  - Spell out on **first use** (`operations security (OPSEC)`—do **not** capitalize expanded words just because the acronym is capped), unless on the **AFPC baseline / 5 pre-approved categories** (`ranks`, `office_symbols`, `organizations`, `weapons_platforms`, `symbols_measurements`) or a unit's strict command list.
  - Pluralize with lowercase `s` without apostrophe (`OPRs`, `AOs`, `LTRs`).
  - **Single-use clutter (p. 334):** If an uncommon term appears only once, spell it out without defining an acronym.
  - **Multi-meaning collisions (`OPR`, `AO`, `PA`, `CC`, `CD`, `AI`):** Always disambiguate on first use or pin via `unit_style.json`.
- **Abbreviations / Truncations (`coord`, `w/`, `w/o`, `ltr`, `hrs`, `in.`, `no.`):** Never define in parentheses like an acronym. Keep periods only where confusion would occur (`in.`, `no.`). Keep informal truncations out of formal memorandums, personal letters, and bios unless expressly authorized by command supplement, and never add `-s`/`-es` to abbreviated verbs.
- **Query / Import Commands:**
  - `python scripts/tq_lookup.py --acronym OPR` · `--abbrev coord`
  - `python scripts/import_acronyms.py <command_guide.pdf> --out local_acronyms.json --mode strict`

## 4. Dates, Numbers, Capitalization & Punctuation (Ch. 24–28)

- **Dates (Ch. 14, 26, 28):** Military: **`21 November 2026`** or **`21 Nov 26`** (never mix abbreviated month with 4-digit year). Civilian/Letters: `November 21, 2026`. Bios use AP style (`Jan.`, `Feb.`, `Aug.`, `Sept.`, `Oct.`, `Nov.`, `Dec.` with day only; `March`–`July` always spelled out; `ASSIGNMENTS` spells out ALL months).
- **Numbers (Ch. 28, pp. 355–362):**
  1. **General Rule:** Spell out **one through nine** (`1–9`); use numerals for **10 and above** (`10+`).
  2. **Technical / Measurements:** Always numerals for **dates**, **clock time (`0800`, never `0800 hours`/`hrs`)**, **money (`$5`, `$2M`)**, **measurements (`3 mi`, `50 lb`)**, **ages**, **percentages (`6 percent` prose / `6%` technical)**, **ratios/scores**, **document parts (`paragraph 3`, `attachment 2`, `tab 1`)**, and **mixed series containing >=10** (`"3 aircraft, 12 trucks, and 4 trailers"`).
  3. **Formal / Sentence Starters:** Always spell out numbers beginning a sentence; spell out 1–100 and round numbers in formal proclamations/decorations. Never write `"twelve (12)"` outside legal contracts.
  - **Unit Designations & Ordinals:** Spell out `1st`–`9th` Air Force (`Ninth Air Force`) or abbreviate `9 AF`; use `10th+` numerals (`24th Air Force` / `24 AF`); **drop ordinals when unit is abbreviated** (`42d Air Base Wing` vs. `42 ABW`). Use `1st`, **`2d`**, **`3d`**, `4th` (not `2nd`/`3rd` except USMC `2ndLt`).
- **Capitalization (Ch. 27, pp. 341–354):** Always capitalize **`Airman` / `Airmen`** and **`Guardian` / `Guardians`**. Capitalize military rank **only immediately before a proper name or in direct address** (`Colonel Grant`, `the colonel`). Capitalize nouns with numbers only when part of a proper title (`Chapter 14, The Official Memorandum` vs. `paragraph 3`). Use `United States` (noun) vs. `US` (adjective, no periods). Headings/titles capitalize first/last word, all words >=4 letters, and short verbs (`Is`, `Be`); lowercase articles, short conjunctions, and short prepositions (`to` in infinitives lowercase).
- **Punctuation & Terminology (Ch. 24–25, pp. 287–330):** Air Force favors **open punctuation** (salutation/close/lists), while using the closed **serial (Oxford) comma** whenever ambiguity would result (or when set in `unit_style.json`). Use em dash `—` without spaces for parenthetical breaks; en dash `–` for ranges without `from`/`between`. Avoid dangling participles, split infinitives that hurt rhythm, and irregular-verb errors (`laid`/`lain`, `hanged`/`hung`).
