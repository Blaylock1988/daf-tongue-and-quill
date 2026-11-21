# Official Biography (Ch. 20, pp. 257–264)

## 1. Typography, Layout & Security Guardrails

- **Typography (Unique PA Standard):** **Arial 9 pt** throughout body, section headings, and entries; **Arial 13.5 pt Bold** Identification Line; bold ALL CAPS headings **without colons**; **1.15 line spacing**; **single space after periods**; max **2 pages**.
- **Official Photo:** Centered upper-right, **`3.2" x 4.0"`** (`build_doc.py --photo <path>` or sized placeholder frame).
- **Privacy Act & OPSEC Hard Rules (pp. 257–260):**
  - **NEVER** include family member names or marital references (*wife, husband, spouse, children, son, daughter, married*).
  - **NEVER** list unit personnel headcounts, aircraft/spacecraft totals, or equipment quantities.
  - Follow current Public Affairs (PA) guidance on contingency deployment locations—use broad theater names (e.g., *Southwest Asia*) when exact locations remain sensitive.
- **AP Style Rules (pp. 258–259):**
  - **Months with a specific day:** Abbreviate **`Jan.`**, **`Feb.`**, **`Aug.`**, **`Sept.`**, **`Oct.`**, **`Nov.`**, **`Dec.`** (`Jan. 1, 2024`). **Always spell out `March`, `April`, `May`, `June`, `July`.** Without a day, spell out all months (`January 2024`).
  - **Dates Exception (`ASSIGNMENTS` / `CAREER CHRONOLOGY`):** Spell out **ALL months completely** regardless of length (p. 261).
  - **States:** AFH 33-337 prints pre-2014 AP state abbreviations with cities (`Colo.`, `Va.`); eight states never abbreviate (`Alaska`, `Hawaii`, `Idaho`, `Iowa`, `Maine`, `Ohio`, `Texas`, `Utah`). `[MODERN DAF UPDATE]` Post-2014 AP Stylebook spells out all 50 states in narrative text—verify local PA / unit preference (`bio_ap_states_spelled_out` in `unit_style.json`). Omit duplicate state when already part of school name (*University of Texas, Austin*).
  - **`[MODERN DAF UPDATE]` Departmental AP Acronym (`DOW` vs. `DOD`):** Per `war.gov` Brand Guide & EO 14347, use uppercase **`DOW`** (never `DoW`) in AP-style public biographies when referencing the Department of War, or **`DOD`** when referencing Department of Defense.
  - **`[MODERN DAF UPDATE]` USSF & Bases:** Capitalize **`Guardian`** alongside **`Airman`**; spell out **Air Force Base (`AFB`)**, **Space Force Base (`SFB`)**, **Space Force Station (`SFS`)**, and **Joint Base (`JB`)** on first use.

## 2. Section-by-Section Rules

1. **Identification Line (13.5 pt Arial Bold ALL CAPS):** Military: `FULL RANK FIRST M. LAST` (`COLONEL JANE A. DOE` / `CHIEF MASTER SERGEANT ALEX R. SMITH`). Civilian: `FIRST M. LAST` (**no rank**).
2. **Narrative Biography (Third Person):** 2–3 paragraphs (military) or up to 5–6 (civilian). Paragraph 1 uses abbreviated rank + full name (`Col. Jane A. Doe`); subsequent references without the name use **lowercase generic rank** (`the colonel`). Close the narrative with: `"Prior to his/her current position, the [rank] was [duty title, organization, base, location]."`
3. **`EDUCATION`:** Chronological (**oldest first**); include degrees, major/school, PME (`by correspondence` noted for PME only, never for civilian degrees), and year.
4. **`ASSIGNMENTS` (Military) / `CAREER CHRONOLOGY` (Civilian):** Numbered chronologically (oldest first), full month names (`June 2004 – May 2006`), **Capitalize duty titles**, list contingency deployments in parentheses.
5. **`SUMMARY OF JOINT ASSIGNMENTS` (Officers Only):** Chronological joint billets ending each line with `, as a [rank]` (e.g., `, as a major`).
6. **`FLIGHT INFORMATION` (Rated Officers Only):** Rating, **More than** (never `"over"`) X flying hours (never round up), primary aircraft flown.
7. **`MAJOR AWARDS AND DECORATIONS`:** Descending order of precedence; **Title Case** medal names; **lowercase** devices (`with two oak leaf clusters`); **never write `one`** before a medal name.
8. **Optional Sections (as applicable):** `OTHER ACHIEVEMENTS`, `EFFECTIVE DATES OF PROMOTION` (chronological, **full spelled-out rank + AP date**; write `Frocked` with no date if frocked; **CMSAF/CMSSF use singular `EFFECTIVE DATE OF PROMOTION`**).
9. **Closing Line (1 blank line below last entry):** `(Current as of Month Year)`.

## 3. Clean Military Biography Template

```text
COLONEL JANE A. DOE

Col. Jane A. Doe is Commander, 1st Operations Group, Joint Base Langley-Eustis,
Virginia. She leads aerospace operations, readiness training, and force generation for
three operational squadrons.

Col. Doe commissioned through the U.S. Air Force Academy in 2004 and completed
undergraduate pilot training in 2006. She has deployed in support of Operations
Enduring Freedom and Inherent Resolve. Prior to her current position, the colonel was
Vice Commander, 4th Fighter Wing, Seymour Johnson Air Force Base, North Carolina.

EDUCATION
2004 Bachelor of Science in Aeronautical Engineering, U.S. Air Force Academy, Colo.
2011 Squadron Officer School, Maxwell AFB, Ala.
2016 Master of Military Operational Art and Science, Air Command and Staff College, by correspondence

ASSIGNMENTS
1. June 2004 – May 2006, Student Pilot, 71st Flying Training Wing, Vance AFB, Okla.
2. June 2006 – August 2010, F-22 Evaluator Pilot, 27th Fighter Squadron, Langley AFB, Va.

SUMMARY OF JOINT ASSIGNMENTS
1. September 2016 – July 2018, Action Officer, J3 Operations, U.S. Indo-Pacific Command, Camp H.M. Smith, Hawaii, as a major

FLIGHT INFORMATION
Rating:  Command pilot
Flight hours:  More than 2,400
Aircraft flown:  T-38A, F-22A

MAJOR AWARDS AND DECORATIONS
Legion of Merit
Meritorious Service Medal with two oak leaf clusters

EFFECTIVE DATES OF PROMOTION
Second Lieutenant June 2, 2004
Captain June 2, 2008
Colonel July 1, 2024

(Current as of November 2026)
```
