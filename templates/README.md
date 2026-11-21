# Local Official Seals & Portal Templates

Official U.S. government, Department of War (`DoW`), Department of Defense (`DoD`), Department of the Air Force (`DAF`), and U.S. Space Force (`USSF`) seals and CAC-gated AF Portal templates are **not** bundled in this repository due to legal seal-usage protections and portal access restrictions.

You can optionally drop local files here so `build_doc.py` and `build_slides.py` embed them automatically. Everything in `templates/seals/` and `templates/portal/` is `.gitignore`d.

## 1. Local Official Seals (`templates/seals/`)

Place PNG seal assets using these exact filenames to enable `--seal auto`:

| Filename | When `--seal auto` Picks It | Recommended Governing Reference |
| :--- | :--- | :--- |
| `templates/seals/dow-seal.png` | `--dept dow` (Executive Order 14347 secondary title) | `war.gov` Brand Guide / Visual Information Style Guide |
| `templates/seals/dod-seal.png` | `--dept dod` (statutory DoD usage) | DoD Visual Information Style Guide / AFMAN 33-326 |
| `templates/seals/daf-seal.png` | Fallback for USAF when dept seal absent | AFMAN 33-326 / `trademark.af.mil` |
| `templates/seals/ussf-seal.png` | Fallback for USSF when dept seal absent | AFMAN 33-326 / `trademark.af.mil` |

Or pass an explicit image path at any time with `--seal C:\path\to\seal.png`.

### Placement Standards (AFMAN 33-326 & AFH 33-337)
- **Official Memorandum / Letterhead:** `0.5"` from the top edge and `0.5"` from the left edge of Page 1 (`0.8"` diameter typical).
- **No seal image present:** Both `build_doc.py` and `build_slides.py` gracefully render clean, T&Q-compliant text-only letterhead (`DEPARTMENT OF THE AIR FORCE` / organizational sub-line).

## 2. Optional CAC-Gated Portal Templates (`templates/portal/`)

If you download official Microsoft Word `.dotx` / `.docx` writing templates from the CAC-gated AF Portal link (`https://www.my.af.mil/gcss-af/USAF/site/HAF/Quill`) or e-Publishing (`https://www.e-publishing.af.mil`), store them under `templates/portal/` for side-by-side comparison. Note that `build_doc.py template <type>` and `build_slides.py template` already generate AFH 33-337-compliant `.docx` and `.pptx` templates directly from code.
