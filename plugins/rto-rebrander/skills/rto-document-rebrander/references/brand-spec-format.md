# Brand spec format

`assets/brand_spec.yaml` is the single source of truth for what the rebrander finds and replaces. Edit this file (and swap the matching assets in `assets/`) to drive a different rebrand without changing any code.

## Top-level structure

```yaml
old_brand:    { ... }
new_brand:    { ... }
preserve:     { ... }
options:      { ... }
```

## `old_brand`

Defines what the rebrander is looking for in the old documents.

| Field | Type | Purpose |
|---|---|---|
| `name` | string | The primary organisation name to find (e.g. `"BSI Learning"`). Matched verbatim in `<w:t>`/`<a:t>` text nodes. |
| `name_variants` | list of strings | Alternative spellings to also catch. Listed first in replacement order so longer/more-specific variants win over `name`. |
| `footer_pattern` | regex string | Regex to find the old footer line. The whole match is replaced with `new_brand.footer`. Use this when the old footer has a year that varies across documents. |
| `colours` | list of 6-char hex strings | Old brand hex colours to look for. Used for inspection/reporting only — the actual replacement is driven by `new_brand.colour_map`. Uppercase, no `#`. |

## `new_brand`

Defines the replacement values and assets.

| Field | Type | Purpose |
|---|---|---|
| `name` | string | Replacement for `old_brand.name` and any variant matches. |
| `footer` | string | The new footer string that replaces matches of `old_brand.footer_pattern`. |
| `theme_xml` | filename | Filename in `assets/` to swap into `word/theme/theme1.xml` (and `ppt/theme/*.xml` for pptx). Should be a fully-formed OOXML theme XML with the new accent palette. |
| `logos.full_colour_png` | filename | Main full-colour logo (PNG). Used to replace BSI logo PNG matches. |
| `logos.full_colour_svg` | filename | Same logo as SVG. Used to replace BSI SVG matches. |
| `logos.monogram_svg` | filename | Smaller monogram. Currently unused by the swap logic but available for future variants. |
| `colour_map` | dict of hex → hex | Old hex → new hex replacement map applied to `w:val`/`w:fill`/`w:color` attributes throughout XML parts. |
| `old_logo_reference` | filename | Reference image used for perceptual-hash matching of logo candidates inside documents. |

## `preserve`

Patterns the rebrander must never touch.

| Field | Type | Purpose |
|---|---|---|
| `text_patterns` | list of regex strings | Each pattern is matched in text content; matches are replaced with sentinels before brand-name replacement runs, then restored. This protects unit codes, ASQA references, ABN, RTO code, etc. from accidental matches. |
| `version_history_cutoff_date` | YYYY-MM-DD | Currently informational. The intended behaviour is: in tables labelled "Version History", rows whose Date column is before this cutoff are preserved as historical record. The current implementation relies on the unit-code and date carveouts to protect these — see `safety-carveouts.md`. |

## `options`

Behaviour toggles.

| Field | Type | Purpose |
|---|---|---|
| `swap_theme_xml` | bool | If true, replace `theme/theme1.xml` with the new theme. Default true. |
| `swap_logo_images` | bool | If true, scan and swap embedded logo images. Default true. |
| `logo_match_threshold` | int 0–64 | Perceptual-hash distance threshold. Lower = stricter match. 12–16 is a good range for typical logos. Default 14. |
| `swap_colours` | bool | If true, apply `colour_map` to all OOXML colour attributes. Default true. |
| `swap_text` | bool | If true, apply text replacement (with `preserve` carveouts). Default true. |

## Adding a new rebrand

To use this skill for a future rebrand (say, aEX → SomethingElse):

1. Drop the new theme XML, logos, and a reference of the *current* aEX logo into `assets/`.
2. Edit `brand_spec.yaml`:
   - `old_brand.name` → `"aEX Institute"`
   - `old_brand.footer_pattern` → regex for the aEX footer
   - `old_brand.colours` → `["E1E146", "194641"]`
   - `new_brand.*` → fill in the new values
   - `new_brand.colour_map` → `{ "E1E146": "<new accent1>", "194641": "<new accent2>" }`
3. Update `preserve.text_patterns` if any new acronyms or codes need protecting.
4. Run `inspect_doc.py` on a sample document to verify the new spec catches what you expect, then run `rebrand.py` on the folder.
