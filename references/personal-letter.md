# The Personal Letter, Forms of Address & Envelopes (Ch. 15 & 22)

## 1. Personal Letter Layout & Format (Ch. 15, pp. 191–218)

- **When to Use:** Formal thank-you letters, condolences, letters to elected officials, foreign diplomats, academia, and private citizens where an Official Memorandum is inappropriate. **Target <=1 page**; **avoid acronyms/abbreviations** in prose.
- **Physical Geometry:** **Page 1 top margin `1.75"`** (`1"` left/right/bottom; **`1"` all around** on continuation pages); **12-pt Times New Roman**; Page 1 letterhead centered `5/8"` from top (optional pre-printed or computer-generated).
- **Heading & Addresses:**
  - **Date:** Flush right, `1.75"` from top, civilian format preferred (`November 21, 2026`).
  - **Return Address (Sender):** 2nd line below date, flush left — **full spelled-out rank** + full name, duty title, organization, street, `City ST  ZIP+4` (**no comma** between city & state, **2 spaces** before ZIP).
  - **Receiver Address:** **3rd line below** return address, flush left — **full spelled-out rank/honorific** (`The Honorable`, `Brigadier General`), do not duplicate titles (`Dr. Name` OR `Name, MD`, never both), use **`US`** / **`UN`** without periods.
- **Salutation & Complimentary Close (Open Punctuation Rule):**
  - **Salutation:** 2nd line below receiver address; **NO colon or comma** (`Dear Colonel Smith`, `Dear Mr. Secretary`, `Dear Madam Ambassador`, `Dear Specialist Doe`). Retain periods in `Mr.`, `Mrs.`, `Ms.`, `Dr.`.
  - **Body:** Starts 2nd line below salutation; **unnumbered** paragraphs with **`0.5"` first-line indent**; single-spaced within, double-spaced between.
  - **Complimentary Close:** 2nd line below text at **`4.5"` from left edge**; **NO trailing comma/colon** (`Sincerely`; use `Respectfully yours` for the President of the United States).
  - **No Authority Line:** Never use `FOR THE COMMANDER` on Personal Letters.
- **Signature & Closing:** **5th line below** close at `4.5"` from left edge (`FIRST M. LAST, Rank, Service`; duty title **optional**). `Attachment:` on 3rd line below signature; `cc:` on 2nd line below attachment (or 3rd below signature). Never endorse (`1st Ind`) a Personal Letter.

> [!TIP]
> Never bloat context copying the 17-page Ch. 15 protocol table—query any civilian, diplomatic, legislative, military, or DoW/DoD title directly:
> `python scripts/tq_lookup.py --address "Secretary of War"` or `"Foreign Ambassador"`

## 2. Paired Executive Forms of Address (`[MODERN DAF UPDATE]`)

Per **EO 14347** (5 Sep 2025) and **10 Oct 2025 SecWar Memo**, offer users either **Department of War** (secondary title for executive/public correspondence) or **Department of Defense** (statutory title for legal/congressional/treaty letters):

| Addressee | Address Block | Salutation & Close |
| :--- | :--- | :--- |
| **Secretary of War** *(EO 14347)* | `The Honorable (full name)` / `Secretary of War` / `1000 Defense Pentagon` / `Washington DC  20301-1000` | `Dear Mr./Madam Secretary` / `Sincerely` |
| **Secretary of Defense** *(Statutory)* | `The Honorable (full name)` / `Secretary of Defense` / `1000 Defense Pentagon` / `Washington DC  20301-1000` | `Dear Mr./Madam Secretary` / `Sincerely` |
| **Secretary of the Air Force** | `The Honorable (full name)` / `Secretary of the Air Force` / `1670 Air Force Pentagon` / `Washington DC  20330-1670` | `Dear Mr./Madam Secretary` / `Sincerely` |

## 3. Envelopes & USPS OCR Standards (Ch. 22, pp. 275–282)

- **Sizing:** Rectangular (`3.5" x 5"` min to `6.125" x 11.5"` max); match window alignment to `MEMORANDUM FOR` if using window envelopes.
- **Return Address & Official Mail:** Upper-left `1/8"` from top/left; place `OFFICIAL BUSINESS` at least 2 lines below return address.
- **USPS Pub 28 OCR Read Area:** `5/8"` to `2 3/4"` from bottom edge, `1/2"` margins on both sides. Format in **ALL CAPS**, `12-pt` Courier New or Times New Roman, **open punctuation** (**only** hyphen allowed is inside ZIP+4), **max 5 address lines**, last line `CITY ST  ZIP+4` (2 spaces before ZIP).

## 4. Clean Personal Letter Template

```text
                                                    November 21, 2026

Colonel Jane A. Doe
Commander, 1st Fighter Wing
123 Base Road
Langley AFB VA  23665-1234


The Honorable Jordan V. Carter
Secretary of War
1000 Defense Pentagon
Washington DC  20301-1000

Dear Mr. Secretary

    Thank you for visiting our Airmen and Guardians during yesterday's operational
readiness review at Joint Base Langley-Eustis. Your engagement reinforced our wing's
focus on rapid combat generation.

    Please convey our gratitude to your staff for supporting the mission brief.

                                                    Sincerely


                                                    JANE A. DOE, Colonel, USAF
                                                    Commander
```
