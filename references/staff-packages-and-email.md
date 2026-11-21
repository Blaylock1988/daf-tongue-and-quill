# Staff Packages (SSS / eSSS), Official E-mail & Workplace Media (Ch. 12 & 18)

## 1. Staff Summary Sheet (SSS) & eSSS (Ch. 18, pp. 237–246)

`[MODERN DAF UPDATE]` AF Form 1768 is now **DAF Form 1768**, though electronic routing primarily occurs inside **Task Management Tool (TMT)**, **ETMS2**, or **eSSS** email workflows.

- **Coordination Action Codes:** `Coord` (review/input), `Appr` (decision/approval authority), `Sig` (signature authority), `Info` (situational awareness only), `File` (official records retention per AFRIMS / DAFMAN 33-363).
- **Mandatory Staff Chops (Front Office Kickback Shield):**
  - **JA (Legal):** Any legal authority citation, waiver/ETP, discipline, fiscal/gift, or MOA/MOU.
  - **FM (Comptroller):** Any manpower, funding, TDY, or GPC obligation.
  - **PA (Public Affairs):** Public release, external media, community outreach, or Official Biographies.
  - **A6/S6 (Cyber/Comm):** IT/network changes, software approval, COMPUSEC (DAFMAN 17-1301).
  - **FSS/A1 (Personnel):** Military/civilian manpower, training requirements, or duty position changes.
- **Backward Suspense Planning:** Add internal coordination buffers (typically 3–5 duty days per chop tier + 2 duty days for XO/Front Office scrub) backward from the external HHQ due date.
- **5 Standard Body Headings:**
  1. `1. PURPOSE:` Single sentence naming what action is requested and why.
  2. `2. BACKGROUND:` Brief history/directive authority.
  3. `3. DISCUSSION:` Key facts, COA trade-offs, resource/risk impacts.
  4. `4. VIEWS OF OTHERS:` Summarize each coordinated office's concur/concur-w/comment/non-concur. Never bury a **non-concurrence**—route with the non-concurrence + OPR's point-by-point **rebuttal placed at the last Tab**.
  5. `5. RECOMMENDATION:` Exact signature/decision action (`"1 FW/CC sign Tab 1 memorandum."`).
- **Signature & Tabs (Ch. 18):** eSSS signature sits **flush left** (`//SIGNED//` matching AO/OPR; **never** use `FOR THE COMMANDER` on an SSS/eSSS). Paper Form 1768 centers `Tabs`; eSSS places `Tabs` flush left 1 blank line below signature (`Tab 1` = outgoing action + Atch 1/2; `Tab 2` = incoming tasker; `Tab 3` = background/research; last Tab = non-concurrence/rebuttal).
- **CUI Transmittal Packages:** An unclassified eSSS forwarding CUI Tabs carries top/bottom `CUI` banners plus the conspicuous separation notice (`"When enclosure is removed, this document is Uncontrolled Unclassified Information"`); full DI block sits on the CUI Tabs themselves.

```pwsh
# Bundle a synchronized eSSS + Tabs + Cover E-mail folder from one JSON spec
python scripts/build_doc.py package pkg.json --out-dir .\out\package --dept dow --cui
```

## 2. Official E-mail & Electronic Comms (Ch. 12, pp. 133–152)

`[MODERN DAF UPDATE]` Governed by **DAFMAN 17-1301** (COMPUSEC, superseding AFMAN 33-152) and **DAFMAN 33-363** (Records Management). Legacy references to Google+ and Vine are obsolete; enterprise collaboration uses **MS Teams / NIPR**.

- **Subject Line Convention:** `[CUI/Marking][KEYWORD][Suspense DD Mmm YY][5–7 Word Subject]` (e.g., `[CUI][SIGN][28 Nov 26][Wing Staff Package Policy]`).
  - *Keywords:* `ACTION`, `SIGN`, `COORD`, `INFO`, `URGENT`. Strip meaningless `FW:`/`RE:` chains.
- **Body Organization:** Target **<=24 lines** using one of three AFH 33-337 methods: **BLUF**, **BLIND** (*Bottom Line, Impact, Next Steps, Details*), or embedded **eSSS**.
- **CUI E-mail Requirements (DoDI 5200.48):** Put `[CUI]`/`CUI:` in subject, place full **`CUI` banner** at very top of body, preserve markings when forwarding/replying, mark CUI attachment filenames and files independently, and **digitally sign & encrypt (S/MIME)** per DoD/DAF policy.
- **E-mail Signature Block:** Include `//SIGNED//` (when required) over plain text only: Full Name, Rank, Service, Duty Title, Org/Symbol, DSN/Comm, Official E-mail. Omit motivational quotes, ASCII art, slogans, and personal social links. Prefer **Organizational Mailboxes (OMs)** over individual accounts for enduring taskers.
- **Social Media, IM/Teams & Telephone/Voice Mail (Ch. 12):**
  - *Social Media (`[MODERN DAF UPDATE]` DAFI 35-107 / PA):* Strictly enforce OPSEC and PERSEC; never speak for the DAF/DoW unless authorized; include personal-opinion disclaimers where appropriate.
  - *IM / MS Teams:* Official decisions made in Teams chat remain official federal records—capture formal taskings via e-mail or MFR; protect CUI inside approved enterprise tenants only.
  - *Telephone & Voice Mail:* Answer with unit + rank + name + `"May I help you?"`; keep voicemails concise with name, number, BLUF reason, and best callback window; never transmit CUI/PII over unencrypted commercial lines.
