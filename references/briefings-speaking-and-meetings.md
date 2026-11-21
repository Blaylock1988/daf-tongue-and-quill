# Briefings, Air Force Speaking, Listening & Meetings (Part IV & Ch. 13)

## 1. Air Force Speaking & Delivery (Ch. 10, pp. 111–117)

- **Delivery Formats:**
  - *Impromptu:* No advance prep—quickly organize into 1) Main Point, 2) 1–2 Supporting Reasons, 3) Restate Point.
  - *Prepared (Extemporaneous):* Outline + key notes/slides; conversational, adaptable to leadership interruptions (standard DAF briefing style).
  - *Manuscript:* Word-for-word reading reserved for strict policy statements, congressional testimony, or public ceremonies.
- **Briefing Anatomy:**
  1. *Attention / Opening:* Crisp greeting, state **purpose** and **BLUF** upfront, time constraint awareness.
  2. *Body:* Match audience decision level; anticipate executive interruptions by placing the decision/ask on Slide 2 (right after Title Slide).
  3. *Conclusion & Q&A:* Restate exact action/decision requested, invite questions cleanly, answer directly without guessing (`"I will verify that with Logistics and respond NLT 1500 today"`).
- **15 Do's & 14 Don'ts Highlights (pp. 121–122):** Stand aside (never block screen or talk at screen), give the audience silent reading time before speaking, never read slides verbatim, use a cover slide or black screen rather than leaving an empty projector glaring, consider **skipping slides entirely** for 5–15 minute short briefings, never wing it or apologize for lack of prep.

## 2. Briefing Slides Design Rules (Ch. 10, pp. 118–120)

| Element | AFH 33-337 Ch. 10 Rule | Practical Detail / Modern Note |
| :--- | :--- | :--- |
| **Aspect Ratio** | Hardware-dependent in 2015 | `[MODERN DAF UPDATE]` Default to **16:9 widescreen** (`13.333" x 7.5"`). |
| **Color Palette** | **4–5 colors max**, consistent across deck | Light text on dark OR dark text on light; **avoid red-green, blue-red, and red lettering**; test projector contrast. |
| **7×7 Rule** | **<=7 words per line**, **<=7 lines per slide** | Slides reinforce spoken points—never self-explanatory prose walls. |
| **Typography** | Same font family throughout; mixed case | **Title: 40 pt · Subtitle: 30 pt · Body text: 20–26 pt**; left-justify text; avoid ALL CAPS (harder to read & masks acronyms); no end-of-line hyphenation. |
| **Continuation** | One thought per slide | Append **`Continued`** on subsequent slides covering the same main point. |
| **Charts** | Use visuals instead of dense numbers | **Line** = trends over time; **Bar** = compare values; **Pie** = parts of a whole. Place graphics off-center to guide eye to key text. |
| **Tables** | Avoid if verbal/chart works; **max 4 rows × 7 cols** | Strip **`.00`** decimals on whole numbers (p. 120: `"100.00"` misreads as `"10,000"`); round numbers; move notes to footnotes. |
| **Speaker Notes** | Know **1 level of detail below** every slide bullet | Put data sources, anticipated Q&A, and verbatim supporting metrics in Speaker Notes, not on slide face. |
| **Classification / CUI** | Legacy FOUO slide banners | `[MODERN DAF UPDATE]` Use bold centered **`CUI`** (or `CUI//[CAT]//[LDC]`) banners at **top and bottom of every slide** plus lower-right **DI block on Slide 1** (DoDI 5200.48 / DAFI 16-1403). |

```pwsh
# Generate starter briefing JSON, build 16:9 .pptx, and lint it
python scripts/build_slides.py schema --dept dow --cui > deck.json
python scripts/build_slides.py build deck.json --out brief.pptx --palette ussf --dept dow --cui
python scripts/validate_doc.py brief.pptx --type slides --dept dow
```

## 3. Effective Listening Strategies (Ch. 11, pp. 123–131)

- **Hearing vs. Listening:** Hearing is passive physiology; listening is active attention, comprehension, and retention.
- **Three Listening Modes:**
  1. *Informative:* Focus on facts, dates, instructions, clarify ambiguous acronyms immediately.
  2. *Critical:* Evaluate argument logic, evidence sufficiency, and informal fallacies (Ch. 5) before buying in.
  3. *Empathic:* Understand speaker perspective and morale concerns without interrupting or rushing to fix prematurely.
- **Overcoming Barriers:** Silence multitasking/devices, hold questions until patterns emerge, paraphrase critical taskings back (`"To confirm, you want the COA matrix by 1400 Friday?"`).

## 4. Effective Meetings (Ch. 13, pp. 153–160)

- **Before Calling a Meeting:** Confirm an email or phone call cannot settle it. Define clear **purpose**, invite only **essential attendees**, distribute a **read-ahead agenda** (time-boxed topics + OPRs) >=24 hours ahead.
- **Running the Meeting:** Start on time, enforce agenda time-boxes, assign a **recorder**, capture explicit **Action Items (`Task — OPR — Suspense`)**, close with verbal recap and date of next meeting (or termination).
- **Meeting Minutes:** Prepared in Official Memorandum format per Ch. 13 (p. 158)—see [official-memorandum.md](official-memorandum.md) and `build_doc.py schema minutes`.

## 5. Markdown Slide Outline Template (For Rapid Drafting Before `.pptx`)

```markdown
## Slide 1: Modernized Staff Routing Proposal
- Department of War · 1st Fighter Wing
- 21 November 2026
Notes: 10-minute decision brief to approve wing digital template standard effective 1 Dec 26.

## Slide 2: Executive Overview
- Cut routing turnaround 40 percent
- Adopt unified template and linter suite
- Standardize mandatory JA and FM chops
- Begin pilot rollout 1 Dec 26
Notes: Current average wing SSS takes 9.2 duty days across 3 directorates due to manual rework.
```
