# Communication Process, Argumentation & Editing (Parts I–III, Ch. 1–9)

## 1. Plain Language & FOCUS Principles (pp. 2–5)

- **Legal Mandate:** Public Law 111-274 (*Plain Writing Act of 2010*), implemented by DoDI 5025.13 and DAFI 33-360 (*Publications and Forms Management*).
  - **Be Clear:** Use plain language, avoid jargon, establish acronyms at first use, prefer active voice, format visually, use tables/figures for data.
  - **Be Concise:** Cut deadwood, keep sentences <=20 words (target avg **17 words**, p. 89), keep paragraphs **3–7 sentences** focused on one point (p. 69).
  - **Be Specific:** Include only details the reader needs with precise language.
- **FOCUS:**
  - **F**ocused: Answer the question, the whole question, and nothing but the question.
  - **O**rganized: Structure logically so busy leaders find the bottom line immediately.
  - **C**lear: Free of ambiguity, passive evasion, and bureaucratic fog.
  - **U**nderstanding: Tailored to primary, secondary, gatekeeper, and decision-maker audiences.
  - **S**upported: Backed by verified evidence, sound warrants, and counterargument awareness.
- **Four Communication Purposes (Ch. 1, 3):** **Direct**, **Inform** (educate), **Persuade**, **Inspire**.

## 2. Seven Steps to Effective Communication (Ch. 2–9)

1. **Analyze Purpose & Audience (Ch. 3):** Draft a single-sentence purpose statement with a specific action verb. Map four audience roles (pp. 19–20):
   - *Primary receiver* (person addressed or listening directly)
   - *Secondary audience* (observers, FOIA/record reviewers, downstream implementers)
   - *Key decision makers* (who actually approves/signs/funds)
   - *Gatekeepers* (XOs, chiefs of staff, JA/FM/A6 reviewers who can kick back the package)
2. **Research Your Topic (Ch. 4):** Gather authoritative primary sources; evaluate **Authority**, **Accuracy**, **Precision**, **Relevance**, and **Sufficiency** (pp. 44–46).
3. **Support Your Ideas (Ch. 5):** Build Toulmin-style logical arguments (**Claim**, **Evidence**, **Warrant**, **Qualification**) and eliminate all **14 Informal Logical Fallacies** (query via `python scripts/tq_lookup.py --fallacy <name>`):
   - *Asserted Conclusion* (circular reasoning / loaded question), *Character Attack* (ad hominem), *Emotional Appeals* (force, pity, bandwagon/tradition/stirring symbols), *False Authority* (incl. primacy of print), *False Cause* (*post hoc*), *Single Cause*, *Faulty Analogy*, *Faulty Dilemma*, *Hasty Generalization*, *Non Sequitur*, *Slippery Slope*, *Red Herring*, *Stacked Evidence*, *Straw Man*.
4. **Organize & Outline (Ch. 6):** Lead **deductively / BLUF** for busy or receptive leaders; use **inductive** order only when the audience is hostile and needs neutral common ground first. Apply the **Cardinal Rule**: any divided topic must have **at least two parts** (`1.` requires `2.`; `a.` requires `b.`). Select one of **8 Organizational Patterns** (pp. 59–62):
   1. Topical / Classification
   2. Comparison and Contrast
   3. Chronological (good for controversial topics)
   4. Sequential (step-by-step)
   5. Spatial / Geographical
   6. Problem and Solution (Staff Study & eSSS pattern)
   7. Reasoning or Logic (inductive/deductive)
   8. Cause and Effect
5. **Draft (Ch. 7):** Lead with **Bottom-Line Up Front (BLUF)** and maintain **Three Cures for Passive Voice** (p. 74):
   - *Cure 1 — Put Actor Before Verb:* `"The handlers broke the part"` (not `"The part was broken by the handlers"`).
   - *Cure 2 — Drop Part of the Verb:* `"The results are in the attachment"` (not `"are listed in"`).
   - *Cure 3 — Change the Verb:* `"The replacement has not arrived"` (not `"has not been received"`).
   - Eliminate **smothered `-ion`/`-ment` verbs** (`"decide"` vs. `"make a decision"`), **doubleheaders** (`"test and evaluate"`), **repetitive redundancies** (`"advance planning"`), **clichés**, and AI filler words (lookup via `python scripts/tq_lookup.py --wordy <phrase>`).
6. **Edit — Three-Pass Approach (Ch. 8, pp. 93–102):**
   - **Pass 1 (Big Picture):** Purpose clear in opening? BLUF visible? Pattern fits audience? Conclusion matches intro?
   - **Pass 2 (Paragraphs & Logic):** Topic sentence per paragraph? One idea per paragraph (3–7 sentences)? Smooth transitions? Any logical fallacies or stacked evidence?
   - **Pass 3 (Sentences & Mechanics):** Active verbs dominant? Deadwood cut? Avg sentence <=20 words? Parallel list structure? Acronyms defined/abbreviations proper? Numbers/dates/signature block match Ch. 14/25–28?
7. **Fight for Feedback & Staff Coordination (Ch. 9, pp. 103–109):** Solicit targeted feedback early ("Does the recommendation clearly state what the WG/CC signs?"). Route parallel chops early rather than surprising JA, FM, PA, or A6/S6 at suspense minus 24 hours.
