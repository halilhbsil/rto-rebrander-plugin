# PDF strategy

PDFs are the awkward format in this rebrand because they're not designed for editing. The rebrander uses a **two-tier strategy** to get the best result.

## Tier 1 — use the source file (preferred, ~90% of cases)

The user has source `.docx`/`.pptx` files for ~90% of their PDFs (these PDFs are exports of source files). For these, the cleanest path is:

1. Rebrand the source `.docx`/`.pptx`.
2. Re-export to PDF.

The `rebrand.py` script does this automatically when `--also-export-pdf` is passed: after rebranding each `.docx`, it invokes LibreOffice headless to produce a matching PDF in the same output folder. The resulting PDFs are **professional-quality** because the new theme, logos, and colours are applied at the source-file level and the PDF is fresh-rendered.

The PDF rebrander module (`rebrand_pdf.py`) also has a `find_source_file()` helper. When called on a PDF, it looks for a same-named `.docx` or `.pptx` sibling. If found, it returns a `method: "use_source_file"` log entry so the caller knows to handle this PDF via its source rather than directly.

### LibreOffice notes

- LibreOffice must be installed on the system. The image used here has it at `/usr/bin/libreoffice`.
- Conversion runs headless (no GUI required).
- The output PDF will use whatever fonts are installed on the system. For best results the EduFavorit font (bundled in `assets/EduFavoritVariable.ttf`) should be installed system-wide before conversion. Otherwise LibreOffice will substitute a similar font, and the result may not exactly match Word's render. The bundled aEX template embeds EduFavorit into the docx itself (via `word/fonts/font*.odttf`), but LibreOffice may or may not honour the embedded fonts depending on version.

## Tier 2 — best-effort overlay (~10% of cases, PDFs without source)

For PDFs where no source file exists, direct PDF editing is tried via `pypdf`. **This is unreliable** and the rebrander deliberately conservative:

- It scans each page's text layer for occurrences of the old brand name and footer pattern, **counts them**, and reports them in the change log.
- It does **not** attempt in-place text replacement, because:
  - PDF fonts are often subset-encoded — the new replacement string may include characters that aren't in the embedded font subset, causing garbled output.
  - Text in PDFs is positioned absolutely, so replacing "BSI Learning" with the longer "aEX Institute" causes overflow that pypdf can't cleanly repaginate.
  - Embedded raster logos can't be swapped without recompositing the page (out of scope).
- The output PDF is a copy of the original, plus a `needs_manual_review: true` flag and a warning explaining what was detected.

The user then has three options for tier-2 PDFs:

1. **Recreate the source.** If the document is important and/or large, retype/import it into Word using the aEX template, then re-export.
2. **Edit the PDF directly** in Acrobat Pro or similar GUI tool. Use the bundled aEX assets (`assets/aex_logo.png`, the colour palette in `brand_spec.yaml`) as references.
3. **Live with the unbranded PDF** if the document is low-stakes and rarely seen.

The `rebrand_report.md` lists every tier-2 PDF in the "Files needing manual review" section so the user can plan the work.

## Why not try harder on direct PDF editing?

There are tools that can do more aggressive PDF editing — for example, regenerating each page from extracted text + reapplied styles. They tend to:

- Lose fidelity (especially for documents with tables, sidebars, callout boxes).
- Strip accessibility tags (a problem for RTO compliance with disability standards).
- Mangle embedded forms.

For the volume here (~10% of 50–200 = 5–20 documents), it's faster and safer to handle them by hand than to build a heuristic that gets it wrong some of the time.
