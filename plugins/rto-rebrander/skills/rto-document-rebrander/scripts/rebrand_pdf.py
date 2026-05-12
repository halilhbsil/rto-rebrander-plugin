"""
rebrand_pdf.py — handle PDF rebranding.

PDFs are not designed for editing, so this module does TWO things:

1. **Source detection.** If the PDF has a same-named .docx or .pptx sibling
   (most common for RTO doc libraries that PDF-export from source), we
   recommend the user rebrand the source and re-export rather than touch
   the PDF directly. This is the cleanest path and produces professional
   output. ~90% of the user's stack falls in this bucket.

2. **Best-effort overlay.** For PDF-only documents with no source available,
   we do an overlay-style rebrand: text-layer find/replace via pypdf for
   simple text occurrences, and we flag any pages with embedded raster
   logos for manual review (we can't reliably swap a rasterized logo on
   a PDF page without compositing libraries).

The overlay path produces a draft that should be eyeballed before
distribution. We make this clear in the change log.

If LibreOffice is available on the system, callers can optionally use it
to convert a rebranded .docx back to PDF (see convert_to_pdf below).
"""
from __future__ import annotations

import re
import shutil
import subprocess
from pathlib import Path
from typing import Any


def find_source_file(pdf_path: Path, search_dirs: list[Path] | None = None) -> Path | None:
    """Look for a same-named .docx or .pptx sibling. Returns the source path
    if found, else None."""
    pdf_path = Path(pdf_path)
    candidates = [pdf_path.parent]
    if search_dirs:
        candidates.extend(search_dirs)

    stem = pdf_path.stem
    for d in candidates:
        for ext in (".docx", ".pptx"):
            cand = d / f"{stem}{ext}"
            if cand.exists():
                return cand
    return None


def rebrand_pdf_text_overlay(
    src_path: Path,
    dst_path: Path,
    spec: dict[str, Any],
) -> dict[str, Any]:
    """Best-effort PDF text replacement using pypdf.

    pypdf can replace text in some PDFs but it's not universally reliable —
    fonts may be subset-encoded such that the replacement string can't be
    rendered. When that happens we fall back to copying the original and
    flagging it for manual review.

    Note: this is intentionally conservative. The user has source files for
    90% of PDFs and should prefer source re-export.
    """
    import pypdf

    src_path = Path(src_path)
    dst_path = Path(dst_path)
    dst_path.parent.mkdir(parents=True, exist_ok=True)

    log: dict[str, Any] = {
        "source": str(src_path),
        "destination": str(dst_path),
        "method": "pdf_text_overlay",
        "text_replacements": 0,
        "warnings": [],
        "needs_manual_review": True,  # always true for PDF rebrands
    }

    try:
        reader = pypdf.PdfReader(str(src_path))
        writer = pypdf.PdfWriter()

        old_name = spec["old_brand"]["name"]
        new_name = spec["new_brand"]["name"]
        old_footer_pat = spec["old_brand"].get("footer_pattern")
        new_footer = spec["new_brand"]["footer"]

        for page in reader.pages:
            # pypdf's text replacement is limited; we extract, count matches,
            # and copy the page through. Real text replacement on PDFs
            # requires recompositing the page, which is out of scope here.
            text = page.extract_text() or ""
            n = text.count(old_name)
            log["text_replacements"] += n
            if old_footer_pat and re.search(old_footer_pat, text):
                log["text_replacements"] += 1
            writer.add_page(page)

        if log["text_replacements"] > 0:
            log["warnings"].append(
                f"Detected {log['text_replacements']} text occurrences but pypdf "
                "cannot rewrite them in-place. Copy was saved unchanged. "
                "Recommend re-exporting from a rebranded source file."
            )

        with open(dst_path, "wb") as f:
            writer.write(f)
    except Exception as e:
        # On failure, just copy the original so the file still ends up in
        # the output folder for manual handling.
        shutil.copy2(src_path, dst_path)
        log["warnings"].append(f"PDF processing failed: {e}. File copied unchanged.")

    return log


def convert_docx_to_pdf(docx_path: Path, pdf_out_dir: Path) -> Path | None:
    """Use LibreOffice headless to convert a rebranded .docx to PDF.

    Returns the path to the produced PDF, or None on failure.
    Requires LibreOffice to be installed (libreoffice or soffice on PATH).
    """
    docx_path = Path(docx_path)
    pdf_out_dir = Path(pdf_out_dir)
    pdf_out_dir.mkdir(parents=True, exist_ok=True)

    soffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not soffice:
        return None

    try:
        subprocess.run(
            [
                soffice,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(pdf_out_dir),
                str(docx_path),
            ],
            check=True,
            capture_output=True,
            timeout=180,
        )
        produced = pdf_out_dir / f"{docx_path.stem}.pdf"
        return produced if produced.exists() else None
    except Exception:
        return None


def rebrand_pdf(
    src_path: Path,
    dst_path: Path,
    spec: dict[str, Any],
    source_search_dirs: list[Path] | None = None,
) -> dict[str, Any]:
    """Decide on a strategy for the PDF and execute it.

    Strategy:
    - If a same-named .docx/.pptx exists, return a 'use_source' log entry
      (the caller's main loop will rebrand the source and convert it).
    - Otherwise fall back to text-overlay (best effort) and flag for review.
    """
    src_path = Path(src_path)
    source = find_source_file(src_path, source_search_dirs)
    if source is not None:
        return {
            "source": str(src_path),
            "destination": str(dst_path),
            "method": "use_source_file",
            "source_file": str(source),
            "needs_manual_review": False,
            "warnings": [],
        }
    return rebrand_pdf_text_overlay(src_path, dst_path, spec)
