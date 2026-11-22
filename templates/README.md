# Official Seals & Portal Templates

This folder provides bundled Department of War (`DoW`) seal assets and optional local drop-in slots for `DoD` / `DAF` / `USSF` seals and CAC-gated AF Portal templates (`templates/portal/` remains `.gitignore`d).

> **Disclaimer:** This open-source repository is **not** affiliated with, funded by, or officially endorsed by the Department of War, Department of Defense, Department of the Air Force, or the U.S. Government. Official seals included here are strictly for authorized personnel generating official correspondence and briefings and may **not** be used commercially or to imply endorsement (**10 U.S.C. § 2260**; **18 U.S.C. § 506**; [war.gov Brand Guide](https://www.war.gov/Brand-Guide/#DOW-seal)).

## 1. Bundled & Optional Official Seals (`templates/seals/`)

Both `build_doc.py` and `build_slides.py` default to `--seal auto` and accept `--seal auto|blue|bw|color|none|<path>`.

### Bundled Department of War (`--dept dow`) Seals (PNG)

Downloaded from `https://www.war.gov/Brand-Guide/#DOW-seal`. **PNG** is bundled because `python-docx` and `python-pptx` embed RGBA PNGs natively with transparent backgrounds and lossless crispness (no white box on dark slides or rasterizer dependencies).

| Filename | Variant | Auto Default & Governing Standard |
| :--- | :--- | :--- |
| `templates/seals/dow-seal-blue.png` | **Blue** | Default for Word `.docx` memos/letters under **AFMAN 33-326** (ultramarine blue / black ink for official letterhead) |
| `templates/seals/dow-seal-bw.png` | **Black & White** | Monochrome option (`--seal bw`) for black-ink printing per **AFMAN 33-326** |
| `templates/seals/dow-seal-color.png` | **Full Color** | Default for `.pptx` briefing slides (`--seal auto`); selectable via `--seal color` |

### Optional Local Drop-In Seals (`--dept dod` / Service Fallbacks)

Add PNGs using these filenames if needed:

| Filename | When `--seal auto` Picks It | Recommended Governing Reference |
| :--- | :--- | :--- |
| `templates/seals/dod-seal.png` | `--dept dod` (statutory DoD usage) | DoD Visual Information Style Guide / AFMAN 33-326 |
| `templates/seals/daf-seal.png` | Fallback for USAF when dept seal absent | AFMAN 33-326 / `trademark.af.mil` |
| `templates/seals/ussf-seal.png` | Fallback for USSF when dept seal absent | AFMAN 33-326 / `trademark.af.mil` |

Or pass an explicit image path at any time with `--seal C:\path\to\seal.png` (or `--seal none` to suppress letterhead images).

### Placement Standards (AFMAN 33-326 & AFH 33-337)
- **Official Memorandum / Letterhead:** `0.5"` from the top edge and `0.5"` from the left edge of Page 1 (`0.8"` diameter typical).
- **No seal image present:** When no matching seal exists (e.g., `--dept dod` without `dod-seal.png`) or `--seal none` is passed, `build_doc.py` and `build_slides.py` render clean, T&Q-compliant text-only letterhead (`DEPARTMENT OF THE AIR FORCE` / organizational sub-line).

## 2. Optional CAC-Gated Portal Templates (`templates/portal/`)

If you download official Microsoft Word `.dotx` / `.docx` writing templates from the CAC-gated AF Portal link (`https://www.my.af.mil/gcss-af/USAF/site/HAF/Quill`) or e-Publishing (`https://www.e-publishing.af.mil`), store them under `templates/portal/` for side-by-side comparison. Note that `build_doc.py template <type>` and `build_slides.py template` already generate AFH 33-337-compliant `.docx` and `.pptx` templates directly from code.
