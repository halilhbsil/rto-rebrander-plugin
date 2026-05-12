---
name: rto-document-rebrander
description: Rebrand BSI Learning RTO documents (Word .docx, PowerPoint .pptx, and PDFs) to the new aEX Institute brand — swapping logos, colour palette, organisation name, and footer text whilst preserving RTO compliance metadata (unit codes, ASQA references, ABN, RTO code 21371, version-history entries). Use this skill whenever the user mentions rebranding RTO documents, BSI Learning to aEX Institute, swapping logos in Word documents, applying the new brand template, updating footers across a document library, or running a bulk find-and-replace across training and assessment materials. Trigger even if the user only mentions one of those things — for example, "update our learner guides with the new logo" or "the policies still say BSI Learning, can we fix that".
---

# RTO Document Rebrander

This skill rebrands a folder of RTO documents from **BSI Learning** to **aEX Institute** branding. It handles `.docx`, `.pptx`, and (best-effort) `.pdf` files, and produces a per-file change log plus a folder-level summary report.

The skill is configured for the BSI → aEX rebrand out of the box, but it's data-driven: every brand-specific value lives in `assets/brand_spec.yaml`, so the same machinery can be reused for any future rebrand by editing that file and swapping the assets.

## When to use this skill

Trigger this skill whenever the user is dealing with rebranding RTO documents. Common phrasings:

- "Rebrand these documents" / "Apply the new brand to..."
- "Replace the BSI logo with aEX" / "Swap the colours in our policies"
- "The footer still says BSI Learning, can you fix it across the folder"
- "Update our learner guides / assessment tools / handbooks to the new template"
- "Bulk find-and-replace BSI Learning to aEX Institute"

If the user only uploads a single file and asks for a rebrand, still use this skill — it works on a single file or a whole folder.

## How it works

The skill operates **directly on the OOXML zip structure** of `.docx`/`.pptx` files. This is more reliable than `python-docx`/`python-pptx` for theme swaps and image replacement, which those libraries don't fully expose.

There are **two rebranding strategies** for `.docx` files. Choose based on what the user wants the output to look like.

### Strategy 1 — Find-and-replace (default, minimal-touch)

This is the default. The original document's layout is preserved; only brand elements are swapped:

1. **Replaces `theme/theme1.xml`** with the aEX-correct theme (already bundled in `assets/aex_theme1.xml`). Theme-coloured content auto-rebrands.
2. **Scans `media/` images** and replaces any matching the old BSI logo (perceptual hash for PNG/JPG/EMF, keyword check for SVG).
3. **Replaces hardcoded brand colours** in `styles.xml`, `document.xml`, headers, and footers using the colour map (handles VML `fillcolor`, OOXML `w:fill`, theme-tint suffixes, and hash-prefixed values).
4. **Replaces text** ("BSI Learning" → "aEX Institute") with safety carveouts (see below).
5. **Replaces the old footer string** ("© BSI Learning YYYY | RTO 21371") with the new footer ("© aEX Institute 2026 | RTO 21371").

**Use this when:** the user just wants the BSI references gone, the original layout is fine to keep, or the docs are heterogeneous (every doc has different cover art / structure).

### Strategy 2 — Template transplant (`--template` flag)

This rebuilds each `.docx` against a target template, taking the template's chrome (cover, header, footer, theme, styles) and the source's body content. The output looks like the template, populated with the source's unit content:

1. **Extract the source's unit code + title** (e.g. "BSBCMM511" / "Communicate with Influence") from its first paragraphs. Falls back to filename parsing.
2. **Extract the template's own unit code + title** (e.g. "FSKNUM017" / "Use familiar and routine maps and plans for work"). These become the find-targets for substitution.
3. **Extract the source's footer metadata** — version (e.g. `v 1.0`), dates (e.g. `16/8/21`) — so the rebranded doc keeps the source's actual version history rather than picking up the template's defaults.
4. **Splice the template's cover, header, footer1, footer2, theme, styles, fontTable, embedded fonts, and page geometry (pgSz/pgMar)** into the source doc's structure, preserving the source's body content. The template's unit code/title and version/dates are substituted with the source's (case-insensitive matching for the title to handle title-case in cover and sentence-case in footer).
5. **Migrate template chrome images** (lime monogram, header logo, blue cover wordmark) to the new doc with renamed file paths and renumbered relationship IDs.
6. **Inject an explicit cover-only section** with `<w:titlePg/>` + first-page footer reference. This is necessary because LibreOffice doesn't reliably honour `titlePg` in single-section documents — the cover must be in its own section that ends before the body begins.
7. **Run a find-replace pass** on the result to clean up any BSI references that survived in the body content. The chrome media files are protected via `skip_paths` from being misidentified by the perceptual-hash logo matcher.

**What gets matched perfectly with template-transplant:** cover layout (forest-green page, lime monogram, unit code/title, candidate name box, brand wordmark + copyright at bottom), header logo on body pages, footer text + dates + version on all pages, page margins, theme colours, embedded fonts, **body heading colours** (Heading 1 → orange, Heading 2 → blue, Heading 3 → bold black, matching the gold-standard template).

**How body heading restyling works.** Source docs typically use **direct character formatting** for headings (`<w:rPr><w:b/><w:sz val="32"/></w:rPr>`) rather than named paragraph styles, so transplanting `styles.xml` alone has no effect on body content. The transplant therefore runs a **heuristic body-style remapper** that walks top-level paragraphs, profiles the first run's formatting, and reassigns named styles based on size + bold patterns:

- `sz >= 32` (16pt+) bold → Heading 1 (orange in gold template)
- `sz 28` (14pt) bold → Heading 2 (blue in gold template)
- `sz 24` (12pt) bold short text → Heading 3 (bold black)
- A whitelist of common section heading text patterns (`Version History`, `Acknowledgement of Country`, `Table of Contents`, `Unit Description`, `Application of Unit`, `Performance Criteria`, `Skills and Knowledge`, etc.) → Heading 1 regardless of size, since these are semantic top-level sections that source authors often used 12pt bold for

The remapper:
- Skips paragraphs already carrying a named style
- Skips paragraphs inside `<w:tbl>` and `<w:txbxContent>` (table cells and shape textboxes have specific layout-driven formatting that should be preserved)
- Guards against false positives (survey scale labels like "5 definitely applies to me", single-word fragments)
- After classifying, injects `<w:pStyle w:val="HeadingN"/>` into the pPr and **strips direct formatting from rPr** so the named style fully takes over

Final styles/theme/fontTable transplanted from the template are protected from the body-cleanup pass via `skip_paths`, since the gold-standard's Heading 1 orange (`#FF9655`) is also in the BSI legacy palette and would otherwise be replaced with the new aEX lime, defeating the styling.

**What still needs manual polish:** content inside textboxes/shapes (decorative-positioned content), tables (table styles aren't restyled), and the few size-24 bold paragraphs that are semantically headings but don't appear in the section-heading whitelist. The user can address these per-doc; the heuristic gets ~90% of the way there automatically.

**Older Word doc compatibility (BSBOPS502 2010-era format).** Source docs from older versions of Word need extra handling:

- **Namespace merge.** Older docs declare ~15 namespaces in `<w:document>` while the template uses ~35 (`wp14`, `cx`, `cx1..cx8`, `wpi`, `oel`, `aink`, `am3d`, `w15`, `w16*`). Splicing template cover content with `wp14:anchorId`, `mc:Choice Requires="wps"`, etc. into a doc without those declarations produces "unbound prefix" XML parse errors. The transplant merges missing xmlns declarations and `mc:Ignorable` token sets from template into source's `<w:document>` root before splicing.

- **Numbering.xml preserved, NOT transplanted.** Source body paragraphs reference list IDs (`numId`, `abstractNumId`) defined in source's own `numbering.xml`. Replacing it with template's makes those IDs collide — source paragraphs that ask for "numId 3" suddenly get a different list format from template's `abstractNumId 3`, producing nonsense like Greek-letter bullets (φ, κ, λ, μ) instead of `•` bullets.

- **Compensator-indent stripping.** BSI Learning's older docs pair `pgMar left=0 right=0` with paragraph-level `<w:ind w:left="1440"/>` and `<w:ind w:right="1481"/>` "compensators" to fake margins, plus table-level `<w:tblInd w:w="1382"/>` for the same purpose. When we override pgMar to template's non-zero values, those compensators stack on top, producing double-indented body text and tables overflowing the right margin. The transplant runs `_strip_compensator_indents` which strips:
  - `<w:ind w:left="1440"/>` exactly (the canonical 1-inch compensator; other values like 2159 or 2880 for nested list levels are preserved)
  - `<w:ind w:right="X"/>` where 1000 ≤ X ≤ 2000 (covers values like 1434, 1435, 1481, 1523 that source authors used as ~1in right-margin compensators)
  - `<w:ind w:right="X"/>` where X ≥ 4000 (narrow-column reservations from source's wider geometry that would crush text in the new geometry)
  - `<w:tblInd w:w="X"/>` where 1000 ≤ X ≤ 2000 (table-level left compensator; reset to `w="0"` to keep the element valid)
  - The function skips paragraphs inside tables and textboxes (their indents are usually intentional)

- **Body typography is enforced from the template, normalised to a strict two-font system.** The transplant lifts the template's `styles.xml`, then runs `_normalise_styles_to_two_fonts` to collapse every `<w:rFonts>` to one of two brand fonts:

  - **Obviously Narw Semi** (display) — applied to any `<w:style>` whose `styleId` case-insensitively matches `/heading|title|subtitle/`, OR that already referenced `Obviously Narw Semi` in any rFonts (catches `VersionStyleChar`). Covers Heading 1/2/3, Heading1excludeTOC, Subtitle, Heading1NumberedList, all `*HeadingNChar` link styles, and VersionStyleChar.
  - **Edu Favorit Light** (body) — applied to everything else: Normal, NoSpacing, HeaderChar, FooterChar, ListParagraph, Bullets/Bullets2, TableParagraph, all TOC* styles, NormalBoldChar, footertextChar, and `<w:rPrDefault>` (docDefaults).

  Source's `<w:docDefaults>` and `Normal` are discarded; the template's embedded fonts (`word/fonts/font*.odttf`) and `fontTable.xml` are copied across so machines without the brand fonts installed still resolve the correct metrics.

  The rewrite collapses all four font attributes (`w:ascii`, `w:hAnsi`, `w:eastAsia`, `w:cs`) on each `<w:rFonts/>` to the chosen font and drops any `w:asciiTheme`/`w:hAnsiTheme`/etc. (theme-font references) so Word can't substitute the theme's major/minor font over our explicit choice. This is what cleans up the template's own legacy contamination: `Avenir` in docDefaults, `Georgia` in Subtitle, and `Times New Roman` / `Arial` / `Arial Unicode MS` / `Calibri` references baked into the legacy list, bullet, table, and `Heading1NumberedList` styles.

  **Trade-off:** the `<w:spacing w:after="120"/>` in the May2026 template's `Normal` adds 6pt after every paragraph. Inside multi-bullet shapes this stacks (3 bullets × 6pt = 18pt extra height) and can clip source-authored SmartArt — BSBOPS502 page 7's chevron textbox is the known regression case. Spot-check shape contents after rebranding; affected shapes need to be resized manually.

- **Body textbox autofit (overflow backstop).** Source docs often use `<a:noAutofit/>` in textbox `<wps:bodyPr>` (text doesn't shrink to fit). With brand fonts now enforced, content authored against the source's narrower font (e.g. Calibri 11pt) is now rendered at Edu Favorit Light 12pt with extra after-spacing — almost always wider and taller. The transplant rewrites `<a:noAutofit/>` to `<a:normAutofit/>` in body content (only in `src_body_updated`, never in cover content), which lets Word/LibreOffice shrink text to fit the textbox. `<a:normAutofit/>` is written without an explicit `fontScale`, so Word recomputes the scale on first render. This is now the **primary** defence against shape overflow, since the source-typography preservation that previously protected shapes has been removed.

- **Internal-whitespace unit codes.** Some older docs split unit codes with spaces ("BSB OPS502", "FSK NUM017") because the cover banner spans multiple runs that the text extractor joined with whitespace. The unit-code matcher normalises (strips internal whitespace) before testing against the strict regex.

- **Composite version-token cleanup in filename fallback.** Filenames like `BSBOPS502_Manage_Bus_Op_Plans_LG__v1_0.docx` need `v\d+(?:[._]\d+)*` patterns stripped BEFORE replacing underscores — otherwise `v1` is removed but the trailing `_0` survives the underscore-to-space pass and becomes a stray `0` in the title.

- **Variant ordering in name replacement (longest-first).** The split-run text-replacement pass iterates `[*variants, old_name]` (variants first) — never the other way around. Otherwise "BSI Learning" matches inside "BSI Learning Institute Pty Ltd" and gets replaced before the variant can fire, producing nonsense like "aEX Institute Institute Pty Ltd".

- **Footer metadata extraction handles letter-by-letter run splits.** Word frequently splits footer dates and version tokens across many `<w:t>` runs (each digit/dot in its own run because the original author applied per-letter formatting). The extractor produces both 'spaced' (joined with space, catches contiguous text) and 'tight' (joined with no separator, catches letter-by-letter splits) views per paragraph and tries both. Per-paragraph isolation prevents cross-line bleed (e.g. `RTO 21371` adjacent to `24.04.2026` would otherwise tight-join to `2137124.04.2026` and lose the date's word boundary).

- **Field-cache awareness in the substitutor.** The per-paragraph collapse-and-rewrite pass (Pass 2 in `_substitute_text_in_xml`) joins all `<w:t>` bodies in a paragraph, performs the replacement on the joined string, then writes the result back into the *first* `<w:t>` and empties the rest. This must skip `<w:t>` runs that sit inside a field's cached-result region — the span between `<w:fldChar w:fldCharType="separate"/>` and `<w:fldChar w:fldCharType="end"/>`. That cached text is Word's evaluation of a `PAGE`/`DATE`/etc. field, not user-authored content. Without the skip, two things go wrong: (a) the cached value (e.g. a page number `2`) gets joined into the substitution input and pollutes the visible static text — the template's footer reads `…Learner Guide V4.0[tab]<PAGE field cached "2">`, and substituting `V4.0 → v 1.0` over the joined string `…V4.02` produces a literal `v 1.02` in the static run; (b) emptying the field's cache slot doesn't break the field, but Word still re-evaluates and renders the page number, so the rebranded footer ends up showing both an orphan `2` (in static text) AND a correct page number (in the field), with a tab between them. The substitutor scans for `fldChar` open/separate/end markers, builds a list of cached-text ranges, and excludes any `<w:t>` whose start falls inside one — for both the join *and* the rewrite.

**Header logo defaults to template's blue.** The bundled gold-standard template (`FSKNUM017_Learner_Guide_V4_0_aEX.docx`) already has the correct blue (`#4C85ED` accent6) wordmark in its header. Earlier code substituted a green-themed `aex_landscape_green.svg` from the asset bundle for the header logo — that was needed when the template was a generic learner-guide template without brand styling. With the gold-standard template bundled, no substitution is needed; keeping the template's original header media gives the right colour and dimensions automatically.

**Use this when:** the user has a target template (e.g. their college's `FSKNUM017_Learner_Guide_V4_0_aEX.docx`) and wants the rebranded docs to look like that template — clean white pages, cover with brand wordmark, header logo, properly chrome'd. This is the strategy for "make these old docs look like our new brand template", not just "remove BSI references".

For PDFs, the skill prefers to **rebrand the source `.docx`/`.pptx`** when one is available alongside the PDF (this covers ~90% of typical RTO doc libraries, where PDFs are exports of source files). For PDF-only documents, it does a best-effort scan and **flags the file for manual review** — PDF text replacement is unreliable and the safest path is to either re-create the source or hand-edit.

## Running the rebrander

The skill ships three scripts in `scripts/`:

- **`rebrand.py`** — main entry point; processes a file or a folder.
- **`inspect_doc.py`** — preview what would change in a single file (sanity check).
- **`rebrand_docx.py` / `rebrand_pptx.py` / `rebrand_pdf.py`** — per-format modules.

### Output destination

By default — when the user does **not** specify an output folder — every rebranded file is written to a `rebranded/` subfolder sitting **next to the source file**. So a source at `…/FSKOCM001/Learner/Guide.docx` becomes `…/FSKOCM001/Learner/rebranded/Guide.docx`. This is the canonical behaviour: rebrands stay locally adjacent to the originals on the shared drive, never overwrite the source, and surface obviously where the new copy is.

The combined `rebrand_report.{md,json}` lands in `<input>/rebranded/` for directory walks, or in `<source.parent>/rebranded/` for single-file runs.

Files already sitting inside a `rebranded/` folder are automatically skipped during a directory walk, so re-running on a parent folder won't pick up previous outputs and rebrand them again.

The CLI still accepts an explicit output folder as a second positional argument for the legacy "consolidate everything into one place" workflow.

### Typical workflow

When the user asks for a rebrand, follow this sequence:

1. **Confirm scope.** Ask which file/folder. Output destination usually doesn't need confirming — the default beside-source `rebranded/` folder is what most users want. Only ask if they explicitly want everything in one consolidated folder.
2. **Run inspect_doc.py on one representative file first** to show the user what would change. Builds confidence before a 50–200 file batch.
3. **Run rebrand.py with `--dry-run`** if the user wants a preview without writing anything.
4. **Run rebrand.py for real**, optionally with `--also-export-pdf` if they want PDFs alongside the rebranded `.docx` files.
5. **Show the user the Markdown report**. For a directory walk it's at `<input>/rebranded/rebrand_report.md`; for a single file it's in that file's sibling `rebranded/`. Lists every file with change counts, flags needs-review items, lists failures.
6. **Use `present_files`** to give the user the rebranded files plus the report.

### Command examples

```bash
# Preview a single file
python scripts/inspect_doc.py "/path/to/Guide.docx"

# Single file -> sibling rebranded/ folder beside the source
python scripts/rebrand.py "/path/to/FSKOCM001/Learner/Guide.docx"

# Directory walk -> each file gets its own sibling rebranded/ folder
python scripts/rebrand.py "/path/to/FSK10119 EDM Review"

# With the bundled aEX template (template-transplant strategy)
python scripts/rebrand.py "/path/to/Guide.docx" \
    --template assets/aex_target_template.docx

# With user-supplied template + Word PDF export
python scripts/rebrand.py "/path/to/Guide.docx" \
    --template /path/to/users_new_template.docx --also-export-pdf

# Dry run — produces a report without writing files
python scripts/rebrand.py "/path/to/Guide.docx" --dry-run

# Legacy: consolidate every output into one explicit folder
python scripts/rebrand.py /path/to/old_docs /path/to/output_folder
```

## Safety carveouts (RTO compliance metadata)

These patterns are **never replaced** even if they match the rebrand rules. They're protected by sentinel-wrapping during the text-replacement pass. The full list is in `assets/brand_spec.yaml` under `preserve.text_patterns`, but the categories are:

- **Unit codes** (e.g. `BSBWHS411`, `FSKNUM017`, `CHCAGE001`). The pattern `[A-Z]{3,5}\d{2,3}[A-Z]?` would catch most of these including ones starting with "BSB" — important so the BSI replacement doesn't accidentally hit unit codes.
- **Qualification codes** (e.g. `BSB30120`, `CHC30121`, `FSK10219`).
- **Training package references**, ASQA, training.gov.au, tga.gov.au.
- **Standards for RTOs** clause numbers.
- **ABN** in the format `ABN NN NNN NNN NNN` — preserved as historical record.
- **RTO code** `RTO 21371` — explicitly preserved in the new footer.
- **Historical version-history author entries** — patterns `BSI Learning Curriculum` and `BSI Learning (Name)` are preserved so audit-trail rows like "1.0 | 20.07.2021 | BSI Learning Curriculum | Spellchecked" stay intact. If you see other historical author phrasings (e.g. "BSI Learning Compliance Team"), add them to `preserve.text_patterns` in the spec.

For the **version history table** preservation: the regex carveouts (`BSI Learning Curriculum` and `BSI Learning (Name)`) cover the two common shapes of historical author attributions. If a particular document uses a different historical phrasing that gets rewritten in error, see `references/safety-carveouts.md` for how to extend the protection rules.

## Brand spec configuration

The whole rebrand is driven by `assets/brand_spec.yaml`. To run a different rebrand later, edit this file and swap the assets in `assets/`. The spec covers:

- `old_brand.name` and `name_variants` — text to find
- `old_brand.footer_pattern` — regex for the old footer
- `old_brand.colours` — old hex colours to look for
- `new_brand.name` and `new_brand.footer` — replacement text
- `new_brand.theme_xml` — filename in `assets/` to use as the new theme
- `new_brand.logos.*` — filenames in `assets/` for the new logos
- `new_brand.colour_map` — old hex → new hex mapping
- `preserve.text_patterns` — regex patterns to never touch
- `options.*` — toggles (swap_theme_xml, swap_logo_images, swap_colours, swap_text)

See `references/brand-spec-format.md` for full details on each field.

## Bundled assets

In `assets/`:

- `brand_spec.yaml` — the configuration (above)
- `aex_theme1.xml` — the aEX OOXML theme (extracted from the user's `Blank_Template.dotx`); accent1 = lime `#E1E146`, accent4 = forest green `#194641`
- `aex_logo.png` / `aex_logo.svg` — main aEX logo (the lime monogram, 2478×2478 PNG and matching SVG)
- `aex_logo.emf` — wide forest-green "aEX institute" wordmark (for older docs with wide EMF logo placeholders ~4:1)
- `aex_logo_alt1.*` / `aex_logo_alt2.*` — secondary aEX logos
- `aex_landscape.svg` — the user-supplied landscape "aEX institute" wordmark (viewBox 828.62×304.13, ~2.72:1)
- `aex_landscape_green.svg` / `.png` / `.emf` — same wordmark recoloured to forest green; used in template transplant header
- `aex_target_template.docx` — the user's target reference (`Learner_Guide_Document.docx`); used as the default `--template` if the user doesn't provide one
- `aex_monogram.svg` — the simple monogram outline
- `EduFavoritVariable.ttf` — the new brand font
- `old_bsi_logo.png` — reference image for perceptual-hash matching of old BSI logos in documents
- `aex_reference_template.dotx` — the user's original blank template (kept for reference)

## Outputs

After a real run (not dry-run), the output folder contains:

- The rebranded files, mirroring the input folder structure
- `rebrand_report.md` — human-readable summary
- `rebrand_report.json` — machine-readable detail (every file, every change)

## When to flag for manual review

The skill auto-flags files in the report when:

- A `.pdf` had no source `.docx`/`.pptx` sibling (best-effort overlay only)
- Conversion to PDF failed (LibreOffice not available)
- An exception was raised processing the file

For other concerns — for example, a stylised cover page where text is actually rasterized into an image — the rebrander won't catch it automatically. Recommend the user spot-check the first few files in each document family (learner guide, assessment tool, policy, handbook) before treating the rest as done.

## Edge cases worth knowing

The skill has been validated against documents from two distinct BSI Learning eras. Here's what to expect:

**Older docs (pre-2020, e.g. `BSBCMM511 Comm LG v1.0.docx`):**
- Use Office 2007-style themes (`#1F497D` blue, `#F79646` orange) — when the theme is swapped to aEX, content using theme colours auto-rebrands.
- Embed the BSI logo in **EMF** (Windows Enhanced Metafile) format. The skill detects EMF logos by rendering them via LibreOffice and hashing — slow (~1-2s per file) but reliable. Only EMFs under `emf_size_limit_bytes` (default 512KB) are checked, since logos are always small.
- Use **VML legacy shape syntax** (`<v:shape>`, `<v:textbox>`) for cover-page banners with `fillcolor="#xxxxxx"` attributes. The colour replacer handles both OOXML and VML attribute styles.
- Have an older brand palette: `#000099` (deep blue), `#000066`, `#254061` (cover banner), `#F7941E`, `#E36C0A`. All mapped in `brand_spec.yaml`.
- Often abbreviate as **"BSIL"** in places like "permission of BSIL" — added as a name variant.
- Have malformed footer years (e.g. `"© BSI Learning 201 | RTO 21371"` — typo). Footer regex allows 2-4 digit years.

**Newer docs (2020+, e.g. the `Sample_Doc.docx` style):**
- Use the BSI Learning 2020-refresh palette: `#003469` navy, `#DE761C` orange. Mapped in the spec.
- Use modern OOXML DrawingML throughout — image swap finds standard PNG/SVG logos.
- Have the standard "© BSI Learning YYYY | RTO 21371" footer.

**Other things to be aware of:**

- **Decorative orange elements may persist.** Older docs sometimes include vector decorations (e.g. dashed lines with orange play-arrow accents) embedded in EMF or DrawingML shapes that the rebrander doesn't try to colour-swap (it's not a logo). These will still render with their original orange. The user can replace these manually with brand-correct decorations or accept them as residual brand colour. Spot-check the first page of the first doc in each family to find them.
- **Theme accent2 = orange.** The bundled `aex_theme1.xml` (extracted from the user's template) has accent2 = `#FF6100` orange. Anywhere the original doc used theme accent2 (often "Learner Guide"-style subtitle text or accent bars), the rebranded doc will still render as orange. This is a brand-design choice from the user's template, not a rebrand bug. To force everything to lime/forest green, edit `aex_theme1.xml` and change accent2/accent5/accent6 to brand colours.
- **Embedded fonts.** Word can embed fonts in `word/fonts/font*.odttf`. The bundled aEX template already has EduFavorit embedded; if the user's old documents reference different fonts, they'll need to re-embed manually or accept the document falling back to the system font on machines without EduFavorit installed.
- **Smart-quote variants.** "BSI Learning" with a stray non-breaking space or smart apostrophe won't match. The variants list in the spec covers obvious alternatives. If the user reports misses, add new entries.
- **Cover-page logos that are SVG-with-text.** SVG matching uses a keyword check (`bsi` in the file). If a BSI cover logo SVG doesn't contain that string in its identifiers, it won't be detected — recommend the user replace those manually or add a hash-based SVG matcher.
- **Track changes / comments.** The skill leaves these intact. If the user has docs with unaccepted track changes containing "BSI Learning", the visible text after acceptance may differ from what the rebrander processed. Recommend accepting all changes before rebranding.

## Reference files

- `references/brand-spec-format.md` — full schema for `brand_spec.yaml`
- `references/safety-carveouts.md` — how the preservation logic works and how to extend it
- `references/pdf-strategy.md` — detailed reasoning on the PDF approach
