"""
rebrand_docx_template.py — template-transplant rebrander.

This is a different rebrand strategy from rebrand_docx.py. Where rebrand_docx
does a "minimal touch" find-and-replace that preserves the original document's
structure and only swaps brand elements, this module does a "chrome transplant"
where it replaces the cover page, header, footer, and theme with the aEX
template's versions while preserving the original body content.

Used when the user wants the rebranded output to *look like* the new aEX
template (forest green cover, lime monogram, landscape logo header on body
pages) rather than just have BSI references swapped out on the original
layout.

Process per document:

1. Extract unit code + unit title from the source doc. We look at the cover
   page's first non-trivial text node group: typically two consecutive
   paragraphs where the first is the unit code (e.g. "BSBCMM511") and the
   second is the title (e.g. "Communicate with Influence"). Variants are
   tolerated — the extractor falls back to filename parsing.

2. Build the output by:
   a. Starting from the source doc as the base (preserves all body content,
      styles, numbering, body images intact)
   b. Replacing word/theme/theme1.xml with the template's
   c. Replacing word/footer1.xml with template's (with unit code/title
      substituted into "Unit Code - Unit Name" placeholders)
   d. Adding word/footer2.xml and word/header1.xml from template (with
      template's media files migrated and rIds renumbered)
   e. Splicing the template's cover-page XML region into the source doc's
      document.xml, replacing the source's cover region (from <w:body> open
      to the start of the "Version History" paragraph)
   f. Updating section properties so the cover uses footer2 and the body
      uses header1 + footer1
   g. Adding/updating relationship entries in word/_rels/document.xml.rels
      and [Content_Types].xml

3. After chrome transplant, run the existing find-replace pass from
   rebrand_docx for any BSI references that survived in the body content
   (text, colours, in-body logos).

The two passes are complementary: chrome transplant fixes the visual
template, and find-replace cleans up brand residue inside the body.
"""
from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from rebrand_docx import (
    BODY_FONT,
    DISPLAY_FONT,
    _normalise_runs_to_two_fonts,
    _normalise_styles_to_two_fonts,
    rebrand_docx,
)


# Namespace map for OOXML
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG_RELS_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"

# Two-font brand policy: DISPLAY_FONT (headings/title/subtitle) +
# BODY_FONT (everything else). Constants and the styles-normaliser live
# in rebrand_docx.py so both strategies share the same policy. Imported
# above.


def _read_text_from_xml(xml: str, max_paragraphs: int = 20) -> list[str]:
    """Extract text content from <w:p> paragraphs in XML order."""
    paragraphs = re.findall(r"<w:p[\s>].*?</w:p>", xml, flags=re.DOTALL)
    out = []
    for p in paragraphs[:max_paragraphs]:
        texts = re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", p, flags=re.DOTALL)
        joined = "".join(texts).strip()
        if joined:
            out.append(joined)
    return out


def extract_unit_info(src_zip: zipfile.ZipFile, fallback_filename: str) -> tuple[str, str]:
    """Extract (unit_code, unit_title) from any learner-guide cover.

    Universal algorithm — no per-layout strategies. It exploits three
    invariants that hold across every Australian RTO learner guide
    regardless of author or template:

    1. **The unit code appears on the cover.** Pattern
       `[A-Z]{3,8}\\d{2,3}[A-Z]?` (FSKOCM001, BSBCMM511, CHCDIV001,
       TAEDES501). May be inline in a sentence ("FSKOCM001
       Participate..."), standalone ("BSBOPS502"), or split across
       paragraphs ("BSB" + "OPS502" — happens when per-letter
       formatting on the cover banner makes Word emit multiple runs).
       The code always appears before the "Version History" or
       similar section boundary.

    2. **The title is adjacent to the code.** Either in the same
       paragraph (after a dash, tab, or space) or in the following
       paragraph. Title text may wrap across multiple paragraphs if
       the author used a line break inside the heading.

    3. **The title ends at a known marker.** Doc-type markers
       ("Learner Guide", "Trainer Guide", "Assessment Tool", etc.),
       section headings ("Version History", "Acknowledgement of
       Country", "Candidate Name", "Unit Description", etc.), another
       unit code, or a blank gap mark the end of the title.

    Algorithm:

    a. Slice cover region: everything before the first paragraph
       matching a section-boundary marker.
    b. Find the unit code: scan paragraphs left-to-right. For each
       paragraph try an embedded match first ("...CODE..."); if none,
       try concatenating with the next paragraph (split-code case).
    c. Build the title: take any text *after* the code in the same
       paragraph (stripping a leading dash if present), then absorb
       following paragraphs until hitting a stop marker, another unit
       code, an empty paragraph, an absorption cap (4 paragraphs),
       or a length cap (~150 chars).
    d. If no code found anywhere on the cover, fall back to filename.

    This handles every cover layout I've seen in the BSI/aEX
    curriculum set: dashed, no-dash, code-alone-then-title,
    split-across-runs, and wrap-across-paragraphs — without
    file-specific branches.
    """
    doc_xml = src_zip.read("word/document.xml").decode("utf-8")
    paragraphs = _read_text_from_xml(doc_xml, max_paragraphs=40)

    # --- patterns ---------------------------------------------------------
    # Embedded code. Two alpha forms accepted:
    #   (i)  contiguous block of 3–8 uppercase letters ("BSBOPS", "FSKNUM"),
    #   (ii) two segments of 1–4 letters joined by whitespace
    #        ("BSB OPS", "FSK NUM").
    # Either alpha form is then followed by optional whitespace and 2–3
    # digits with an optional trailing letter suffix.
    #
    # The alternation is ordered with the longer (multi-segment) form
    # first so for "BSB OPS502" the engine prefers "BSB OPS" over the
    # shorter "BSB" → fail → restart at "OPS" which would mis-extract
    # the bare "OPS502". For "BSBOPS502" the multi-segment form fails
    # immediately (no internal whitespace) and the contiguous form
    # matches.
    code_embedded_re = re.compile(
        r"\b("
        r"[A-Z]{1,4}\s+[A-Z]{1,4}"  # multi-segment "BSB OPS"
        r"|[A-Z]{3,8}"              # contiguous "BSBOPS"
        r")\s*(\d{2,3}[A-Z]?)\b"
    )
    # Strict standalone form (whole paragraph is the code).
    code_alone_re = re.compile(r"^([A-Z]{3,8}\s*\d{2,3}[A-Z]?)$")

    # Two-tier boundary detection:
    #
    #   cover_end_re — markers that ONLY appear after the cover banner.
    #   Used to slice the cover region. "Candidate Name" / "Date of
    #   Commencement" are NOT here because they're cover-page form
    #   fields that often appear above the title banner (FSKDIG001,
    #   FSKLRG002, FSKLRG005 all have them on paragraphs 0-3).
    #
    #   title_stop_re — superset of cover_end_re plus those cover form
    #   fields. Used to stop title absorption. If we're walking forward
    #   from the code paragraph collecting title text, hitting any of
    #   these means the title is done.
    cover_end_re = re.compile(
        r"^("
        r"Version History|Acknowledgement of Country|Table of Contents|"
        r"Unit Description|Application of Unit|Performance Criteria|"
        r"Skills and Knowledge|Foundation Skills|Range of Conditions|"
        r"Assessment Conditions|Assessment Requirements|"
        r"Elements and Performance Criteria|"
        r"Course Information|Section\s+\d"
        r")\b",
        re.IGNORECASE,
    )
    title_stop_re = re.compile(
        r"^("
        r"Version History|Acknowledgement of Country|Table of Contents|"
        r"Unit Description|Application of Unit|Performance Criteria|"
        r"Skills and Knowledge|Foundation Skills|Range of Conditions|"
        r"Assessment Conditions|Assessment Requirements|"
        r"Elements and Performance Criteria|"
        r"Candidate Name|Date of Commencement|Candidate Details|"
        r"Trainer Assessor|Trainer/Assessor|Course Information|"
        r"Section\s+\d"
        r")\b",
        re.IGNORECASE,
    )
    # Doc-type marker at the START of a paragraph. NOT anchored at end:
    # FSKRDG004's workbook cover has a "Workbook &amp; Assessment"
    # paragraph that's the doc-type label — requiring full-string match
    # would let "&amp; Assessment" bleed into the title. Any paragraph
    # leading with one of these markers (optionally followed by combine
    # text like "& Assessment", "and Assessment Tool", or a version
    # marker) is treated as a doc-type label and ends title absorption.
    doc_type_re = re.compile(
        r"^(Learner Guide|Trainer Guide|Assessment Tool|Assessor Guide|"
        r"Workbook|Student Guide|Trainer Manual|Marking Guide|"
        r"Learner Workbook|Assessment Workbook|Practical Assessment)"
        r"\b",
        re.IGNORECASE,
    )
    title_tail_strip_re = re.compile(
        r"\s*(?:[-–—]\s*)?(Learner Guide|Trainer Guide|Assessment Tool|"
        r"Workbook|Assessor Guide|Student Guide|Marking Guide)"
        r"(?:\s*v?\d+\.?\d*)?\s*$",
        re.IGNORECASE,
    )
    # Leading garbage on the title fragment: stray dash, version, etc.
    title_lead_strip_re = re.compile(
        r"^(?:[-–—:\s]+|v?\d+(?:\.\d+)*\s+)",
        re.IGNORECASE,
    )

    def _normalize_code(prefix: str, digits: str) -> str:
        return f"{prefix}{digits}".replace(" ", "")

    # --- (a) cover boundary ----------------------------------------------
    cover_end = len(paragraphs)
    for i, p in enumerate(paragraphs):
        if cover_end_re.match(p.strip()):
            cover_end = i
            break
    cover = paragraphs[:cover_end]

    # --- (b) find the unit code ------------------------------------------
    code_para_idx: int | None = None
    code_end_in_para: int | None = None  # None if code consumed whole para(s)
    consumed_paras = 1
    unit_code = ""

    for i, p in enumerate(cover):
        # Try embedded match within this paragraph first.
        m = code_embedded_re.search(p)
        if m:
            unit_code = _normalize_code(m.group(1), m.group(2))
            code_para_idx = i
            code_end_in_para = m.end()
            # If the code IS the whole paragraph, treat title as starting
            # in the next paragraph (not the empty remainder).
            if code_alone_re.match(p.strip()):
                code_end_in_para = None
            break

        # Try concat with next paragraph for split-across-runs covers.
        if i + 1 < len(cover):
            joined = (p + cover[i + 1]).strip()
            if code_alone_re.match(joined):
                unit_code = re.sub(r"\s+", "", joined)
                code_para_idx = i
                code_end_in_para = None
                consumed_paras = 2
                break

    if code_para_idx is None:
        # No code found on cover — filename fallback handles both code
        # and title.
        return _filename_fallback(fallback_filename)

    # --- (c) build the title ---------------------------------------------
    title_parts: list[str] = []

    if code_end_in_para is not None:
        rest = cover[code_para_idx][code_end_in_para:].strip()
        rest = title_lead_strip_re.sub("", rest).strip()
        if rest and not doc_type_re.match(rest):
            title_parts.append(rest)

    start_absorb = code_para_idx + consumed_paras
    char_budget = 150
    blank_run = 0
    for k in range(start_absorb, min(start_absorb + 5, len(cover))):
        nxt = cover[k].strip()
        if not nxt:
            blank_run += 1
            # Two consecutive blanks → cover layout shifted; stop.
            if blank_run >= 2 and title_parts:
                break
            continue
        blank_run = 0
        if doc_type_re.match(nxt):
            break
        if title_stop_re.match(nxt):
            break
        if code_embedded_re.search(nxt):
            # Repeat or unrelated code reference — title is done.
            break
        title_parts.append(nxt)
        if sum(len(t) for t in title_parts) > char_budget:
            break

    unit_title = " ".join(title_parts).strip()
    unit_title = title_tail_strip_re.sub("", unit_title).strip()

    if not unit_title:
        # Code found but no usable title — try filename for the title
        # part only, keep the cover-found code.
        _, fb_title = _filename_fallback(fallback_filename)
        unit_title = fb_title

    return unit_code, unit_title


def _filename_fallback(filename: str) -> tuple[str, str]:
    """Derive (unit_code, unit_title) from a filename when the cover gives
    nothing usable. Best-effort — strips version markers, doc-type tokens,
    review-cycle suffixes, and orphan digits."""
    stem = Path(filename).stem

    code = ""
    m = re.match(r"([A-Z]{3,8}\d{2,3}[A-Z]?)", stem)
    if m:
        code = m.group(1)

    cleaned = re.sub(r"^[A-Z]{3,8}\d{2,3}[A-Z]?[\s_-]*", "", stem)
    # Strip composite version markers ("v1_0", "_v1.0", "V4.0") before
    # underscore-to-space so the trailing digit doesn't survive the
    # token strip later.
    cleaned = re.sub(r"_?v\d+(?:[._]\d+)*", "", cleaned, flags=re.IGNORECASE)
    cleaned = cleaned.replace("_", " ").replace("-", " ").strip()
    # Strip common doc-type/version/review tokens. Review-cycle suffixes
    # like "EDM", "NRT", "ED" survive otherwise as the entire title — we
    # treat any all-caps 2-4 letter token at the tail as a review marker
    # rather than a title word.
    cleaned = re.sub(
        r"\b(Comm|LG|Learner Guide|Trainer Guide|Assessment Tool|"
        r"Workbook|Student Guide|Marking Guide|"
        r"v\d+\.?\d*|V\d+\.?\d*|aEX|BSI)\b",
        "",
        cleaned,
        flags=re.IGNORECASE,
    ).strip()
    cleaned = " ".join(
        tok for tok in cleaned.split() if not re.fullmatch(r"\d+", tok)
    )
    # Trailing 2-4 letter all-caps tokens are almost always review-cycle
    # markers (EDM, NRT, ED, R1, R2). Drop them.
    while True:
        toks = cleaned.split()
        if toks and re.fullmatch(r"[A-Z]{2,4}\d?", toks[-1]):
            cleaned = " ".join(toks[:-1])
        else:
            break
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return code, cleaned or "Untitled Unit"


def _walk_top_level_elements(inner: str) -> list[tuple[int, int, str]]:
    """Return list of (start, end, tag) tuples for top-level elements in inner XML.

    Top-level means depth-0 — paragraphs nested inside textboxes are skipped.
    """
    depth = 0
    current_top_start: int | None = None
    elements: list[tuple[int, int, str]] = []
    for m in re.finditer(r"<(/?)([A-Za-z_][A-Za-z0-9_:.-]*)[\s/>]", inner):
        tag_name = m.group(2)
        is_close = bool(m.group(1))
        gt_pos = inner.find(">", m.end() - 1)
        if gt_pos < 0:
            continue
        is_self_closing = inner[gt_pos - 1] == "/"

        if is_close:
            depth -= 1
            if depth == 0 and current_top_start is not None:
                elements.append((current_top_start, gt_pos + 1, tag_name))
                current_top_start = None
        else:
            if depth == 0:
                current_top_start = m.start()
            if not is_self_closing:
                depth += 1
            elif depth == 0:
                elements.append((m.start(), gt_pos + 1, tag_name))
                current_top_start = None
    return elements


def _split_cover_and_body(doc_xml: str) -> tuple[str, str, str, str]:
    """Split a document.xml into (prefix, cover_inner, body_inner, suffix).

    Two patterns are handled:

    A. **Multi-paragraph cover** (e.g. BSBCMM511): the cover is several
       top-level paragraphs/tables, then a separate top-level paragraph
       starts with "Version History". Boundary: at the start of the
       Version History element.

    B. **Single-paragraph cover** (e.g. Blank_Template.dotx): the cover is
       anchored shapes attached to one paragraph, followed by a page break
       inside that same paragraph, followed by the "Version History" run.
       Boundary: at the page break inside that paragraph; cover ends with
       a synthetic </w:p>, body starts with a synthetic <w:p> wrapper that
       carries the "Version History" content.
    """
    body_open_match = re.search(r"<w:body[\s>]", doc_xml)
    if not body_open_match:
        raise ValueError("Document has no <w:body> open tag")
    body_open_end = doc_xml.find(">", body_open_match.start()) + 1
    body_close = doc_xml.rfind("</w:body>")
    if body_close < 0:
        raise ValueError("Document has no </w:body> close tag")
    prefix = doc_xml[:body_open_end]
    suffix = doc_xml[body_close:]
    inner = doc_xml[body_open_end:body_close]

    elements = _walk_top_level_elements(inner)

    # Find which top-level element (if any) contains "Version History"
    # OUTSIDE of textboxes.
    boundary_idx: int | None = None
    for i, (s, e, tag) in enumerate(elements):
        chunk = inner[s:e]
        stripped = re.sub(
            r"<w:txbxContent>.*?</w:txbxContent>",
            "",
            chunk,
            flags=re.DOTALL,
        )
        texts = re.findall(
            r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", stripped, flags=re.DOTALL
        )
        if "Version History" in "".join(texts):
            boundary_idx = i
            break

    if boundary_idx is None:
        # Fallback: split at the last sectPr in inner, or treat all as cover
        for i, (s, e, tag) in enumerate(elements):
            if tag in ("sectPr", "w:sectPr"):
                boundary_idx = i
                break
        if boundary_idx is None:
            return prefix, inner, "", suffix
        boundary_start = elements[boundary_idx][0]
        return prefix, inner[:boundary_start], inner[boundary_start:], suffix

    # We found "Version History" in element [boundary_idx]. Determine if it's
    # at the start of that element (Pattern A) or buried inside (Pattern B).
    s, e, tag = elements[boundary_idx]
    chunk = inner[s:e]
    stripped = re.sub(
        r"<w:txbxContent>.*?</w:txbxContent>", "", chunk, flags=re.DOTALL
    )
    # Where in the stripped chunk does "Version History" appear?
    vh_idx_in_stripped = stripped.find("Version History")

    # Heuristic: Pattern B if there's a page break BEFORE Version History
    # within this element.
    pgbrk_pat = re.compile(r"<w:br\s+w:type=\"page\"\s*/>")
    pgbrk_match = None
    for m in pgbrk_pat.finditer(stripped):
        if m.end() <= vh_idx_in_stripped:
            pgbrk_match = m  # last page break before Version History
    if pgbrk_match is None:
        # Pattern A: simple split
        return prefix, inner[:s], inner[s:], suffix

    # Pattern B: split inside this element at the page break.
    # We need to do the split on the ORIGINAL chunk (which includes
    # textboxes), not the stripped version. Find the same page break in the
    # original chunk: search for the page break that ends at the same content
    # position in the chunk.
    # Simplest: find the page break position in chunk that has Version History
    # AFTER it.
    vh_in_chunk = chunk.find("Version History")
    pgbrks_in_chunk = list(pgbrk_pat.finditer(chunk))
    # Take the last page break that's before vh_in_chunk
    target_pgbrk = None
    for m in pgbrks_in_chunk:
        if m.end() <= vh_in_chunk:
            target_pgbrk = m
    if target_pgbrk is None:
        # Fallback: use the first page break
        target_pgbrk = pgbrks_in_chunk[0] if pgbrks_in_chunk else None
    if target_pgbrk is None:
        # No page break found — fall back to Pattern A
        return prefix, inner[:s], inner[s:], suffix

    # Split chunk at the page break. We need to:
    # - Keep everything from chunk[0] to chunk[pgbrk.end()] as cover_part_in_chunk
    # - Add a synthetic </w:r></w:p> close (after the run that contains the page break)
    # - For body, prepend a synthetic <w:p> open and pick up from after the page break
    # We need to walk back to find the enclosing <w:r> that contains the page break.

    pgbrk_end = target_pgbrk.end()
    pgbrk_start = target_pgbrk.start()

    # Walk back to find the <w:r> that contains the page break
    r_open_idx = chunk.rfind("<w:r", 0, pgbrk_start)
    # Verify it's an <w:r ...> opening (not <w:rPr> or similar)
    while r_open_idx >= 0 and not re.match(
        r"<w:r[\s>]", chunk[r_open_idx:r_open_idx + 5]
    ):
        r_open_idx = chunk.rfind("<w:r", 0, r_open_idx)
    # Find the matching </w:r> after the page break
    r_close_idx = chunk.find("</w:r>", pgbrk_end)
    if r_close_idx < 0:
        # Should not happen in well-formed XML
        return prefix, inner[:s], inner[s:], suffix
    r_close_end = r_close_idx + len("</w:r>")

    # cover_part: chunk from start to r_close_end (the run containing page break)
    cover_part = chunk[:r_close_end]
    # Add a closing </w:p> to make a valid paragraph
    cover_part_closed = cover_part + "</w:p>"

    # body_part: synthetic <w:p> + chunk content from after the closed run to original close
    # Take the original <w:p ... > attributes from the start of chunk, replicate
    p_open_match = re.match(r"<w:p[\s>][^>]*>", chunk)
    p_open_tag = p_open_match.group(0) if p_open_match else "<w:p>"
    body_part = p_open_tag + chunk[r_close_end:]

    # Final assembly:
    # - cover_inner = inner from 0 to s + cover_part_closed
    # - body_inner = body_part + inner from e to end
    cover_inner = inner[:s] + cover_part_closed
    body_inner = body_part + inner[e:]

    return prefix, cover_inner, body_inner, suffix


def _substitute_text_in_xml(
    xml: str,
    replacements: list[tuple[str, str]],
    case_insensitive: bool = False,
) -> str:
    """Replace literal text phrases in a Word XML fragment.

    Word splits text across runs, so a phrase like 'FSKLRG013' or
    'Apply Strategies to Respond' might appear as a single <w:t> body
    OR be split across many <w:t> elements (often when the original
    author dropped run-level formatting on individual letters — Word
    preserves this in the OOXML).

    We handle both cases by:

    1. Direct per-<w:t> body replacement (catches contiguous text)
    2. Per-paragraph collapse-and-rewrite (catches split runs by
       concatenating all <w:t> text in a paragraph, doing string
       replacement, then re-distributing the result back into the
       runs — putting all replaced text into the first run and
       emptying subsequent runs that contained the matched portion)

    When `case_insensitive=True`, the matching is case-insensitive
    but the replacement preserves the source's exact case. Use this
    for substituting unit titles where the template might have the
    title in title-case (cover) and sentence-case (footer) but you
    want the source's canonical-case title in both places.

    Replacements are applied in order, so longer/more-specific phrases
    should come first if they overlap with shorter ones.
    """
    flags = re.IGNORECASE if case_insensitive else 0

    for old, new in replacements:
        if not old or old == new:
            continue

        old_pattern = re.compile(re.escape(old), flags=flags)

        # Pass 1: per-<w:t>-node replacement (handles contiguous text)
        def _sub_in_t(m: re.Match, _pat=old_pattern, _new=new) -> str:
            prefix, body, suffix = m.group(1), m.group(2), m.group(3)
            return f"{prefix}{_pat.sub(_new, body)}{suffix}"

        xml = re.sub(
            r"(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)",
            _sub_in_t,
            xml,
            flags=re.DOTALL,
        )

        # Pass 2: per-paragraph collapse-and-rewrite (handles split runs).
        # Skip <w:t> that sit inside a field's cached-result region (between
        # <w:fldChar fldCharType="separate"/> and <w:fldChar fldCharType="end"/>).
        # That cached text is not user-authored content — it's Word's evaluation
        # of a PAGE/DATE/etc. field. Joining it into the substitution input would
        # mix dynamic field cache (e.g. a page number "2") into the static text;
        # rewriting it would clobber the cache without telling the field-evaluator
        # to refresh, leaving the static slot polluted on render.
        def _collapse_para(m: re.Match, _pat=old_pattern, _new=new) -> str:
            para = m.group(0)
            field_spans = []
            depth = 0
            cache_start = None
            for fm in re.finditer(
                r'<w:fldChar\s[^>]*w:fldCharType="(begin|separate|end)"',
                para,
            ):
                kind = fm.group(1)
                if kind == "begin":
                    depth += 1
                elif kind == "separate" and depth > 0 and cache_start is None:
                    cache_start = fm.end()
                elif kind == "end" and depth > 0:
                    if cache_start is not None:
                        field_spans.append((cache_start, fm.start()))
                        cache_start = None
                    depth -= 1

            def _in_field_cache(pos: int) -> bool:
                return any(s <= pos < e for s, e in field_spans)

            t_matches = list(re.finditer(
                r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", para, flags=re.DOTALL
            ))
            t_matches = [tm for tm in t_matches if not _in_field_cache(tm.start())]
            if not t_matches:
                return para
            concat = "".join(tm.group(1) for tm in t_matches)
            if not _pat.search(concat):
                return para
            new_concat = _pat.sub(_new, concat)
            first_done = [False]

            def _rewrite_t(tm: re.Match) -> str:
                if _in_field_cache(tm.start()):
                    return tm.group(0)
                pfx_full = tm.group(0)
                open_match = re.match(r"<w:t(?:\s[^>]*)?>", pfx_full)
                if not open_match:
                    return pfx_full
                open_tag = open_match.group(0)
                if not first_done[0]:
                    first_done[0] = True
                    if "xml:space" not in open_tag:
                        open_tag = open_tag[:-1] + ' xml:space="preserve">'
                    return f"{open_tag}{_xml_escape(new_concat)}</w:t>"
                else:
                    return f'<w:t xml:space="preserve"></w:t>'

            return re.sub(
                r"<w:t(?:\s[^>]*)?>.*?</w:t>",
                _rewrite_t,
                para,
                flags=re.DOTALL,
            )

        xml = re.sub(r"<w:p[\s>].*?</w:p>", _collapse_para, xml, flags=re.DOTALL)

    return xml


def _xml_escape(s: str) -> str:
    """Escape ampersand and angle brackets for safe XML body content."""
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _extract_footer_metadata(zf: zipfile.ZipFile) -> dict[str, list[str] | str | None]:
    """Pull version and date strings out of a doc's footer1.xml.

    RTO learner guide footers conventionally have three lines:
    1. "<UnitCode> <Title> – <DocType> <Version>"  e.g. "BSBCMM511 ... v 1.0"
    2. "© <Brand> <Year> | RTO <Code>"             e.g. "© aEX Institute 2026 | RTO 21371"
    3. "<Date> | Review Date: <Date>"              e.g. "16/8/21 | Review Date: 16/8/21"

    We extract:
    - **version**: a string like "v 1.0", "V4.0", "v1.2.3"
    - **dates**: a list of date-like strings ("16/8/21", "26.03.2026", "1-Jan-24")
      preserving order of appearance.
    - **year**: the four-digit year from the copyright line if present.

    Returns {} if no footer1.xml or no markers found. Used by the template-
    transplant flow to substitute the template's metadata with the source's.

    Word frequently splits version/date tokens letter-by-letter into separate
    `<w:t>` runs (each digit/dot in its own run because the original author
    applied per-letter formatting). To handle this, we extract each <w:t>
    body, then reconstruct two views:
      - 'spaced': joined with ' ' (works for tokens that aren't split)
      - 'tight':  joined with '' (works for letter-by-letter splits)
    and search both, preferring whichever matches first.
    """
    if "word/footer1.xml" not in zf.namelist():
        return {}
    xml = zf.read("word/footer1.xml").decode("utf-8", errors="replace")

    # Build per-paragraph text views. This preserves the separation between
    # different lines in the footer so e.g. "RTO 21371" on the copyright
    # line and "24.04.2026" on the dates line don't fuse into "2137124.04.2026"
    # (which loses the date's word boundary). Within each paragraph we
    # produce both 'spaced' (joined with ' ') and 'tight' (no whitespace)
    # views to handle both contiguous and letter-by-letter run splits.
    paragraphs = re.findall(r"<w:p[\s>].*?</w:p>", xml, flags=re.DOTALL)
    para_views: list[tuple[str, str]] = []
    for p in paragraphs:
        bodies = re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", p, flags=re.DOTALL)
        if not bodies:
            continue
        spaced = re.sub(r"\s+", " ", " ".join(bodies)).strip()
        tight = re.sub(r"\s+", "", "".join(bodies))
        para_views.append((spaced, tight))

    out: dict[str, list[str] | str | None] = {}

    # Search across all paragraph views in order. First match wins for
    # version + year (each appears once). Dates accumulate from any
    # paragraph that matches.
    ver_re = r"\b[vV]\s*\.?\s*\d+(?:\.\d+)*\b"
    date_re = r"\b\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b"
    year_re = r"©.{0,30}?(\d{4})"

    all_dates: list[str] = []
    for spaced, tight in para_views:
        if "version" not in out:
            ver_m = re.search(ver_re, tight) or re.search(ver_re, spaced)
            if ver_m:
                out["version"] = ver_m.group(0).strip()
        # Per-paragraph dates: prefer the view with more matches
        d_tight = re.findall(date_re, tight)
        d_spaced = re.findall(date_re, spaced)
        d = d_tight if len(d_tight) >= len(d_spaced) else d_spaced
        all_dates.extend(d)
        if "year" not in out:
            year_m = re.search(year_re, tight) or re.search(year_re, spaced)
            if year_m:
                # Years should be plausible (1990-2100); guard against
                # accidentally matching the RTO code "21371" → "2137".
                candidate = year_m.group(1)
                try:
                    if 1990 <= int(candidate) <= 2100:
                        out["year"] = candidate
                except ValueError:
                    pass
    if all_dates:
        out["dates"] = all_dates

    return out


def _substitute_unit_in_text_xml(xml: str, unit_code: str, unit_title: str) -> str:
    """Legacy wrapper kept for backward-compat. Prefer _substitute_text_in_xml.

    Replaces literal placeholders 'Unit Code' and 'Unit Name' if a template
    is using those (older approach). Most real templates have the actual
    unit code/title hardcoded, so the caller should use _substitute_text_in_xml
    with extracted template values instead.
    """
    return _substitute_text_in_xml(
        xml,
        [("Unit Code", unit_code), ("Unit Name", unit_title)],
    )


def _list_relationships(rels_xml: str) -> list[tuple[str, str, str]]:
    """Parse a .rels XML and return list of (Id, Type, Target) tuples."""
    out = []
    for m in re.finditer(
        r'<Relationship\s+Id="([^"]+)"\s+Type="([^"]+)"\s+Target="([^"]+)"',
        rels_xml,
    ):
        out.append((m.group(1), m.group(2), m.group(3)))
    return out


def _next_rid(used_rids: set[str]) -> str:
    """Return the next free rIdN id."""
    n = 1
    while f"rId{n}" in used_rids:
        n += 1
    return f"rId{n}"


def _next_image_index(used_image_names: set[str]) -> int:
    """Return next imageN.<ext> integer not yet used in word/media/."""
    n = 1
    while any(name.startswith(f"image{n}.") for name in used_image_names):
        n += 1
    return n


def _strip_compensator_indents(body_xml: str) -> tuple[str, dict[str, int]]:
    """Strip "compensator" left/right indents from body paragraphs.

    BSI Learning's older RTO doc library uses an unusual geometry pattern:
    section `pgMar` is set to `left=0 right=0`, and individual paragraphs
    carry `<w:ind w:left="1440"/>` (and sometimes `<w:ind w:right="X"/>`)
    to manually create the apparent margins. When we transplant the
    template's chrome and override the section pgMar to template's
    non-zero values, those per-paragraph indents would double-stack on
    top of the new page margins, producing body text that's indented
    too far in.

    This function strips:
    - `w:left="1440"` exactly — the canonical compensator value. Other
      values (e.g. 720, 2159, 2880) represent legitimate nested-list
      indents and are preserved.
    - `w:right="X"` where X is large (≥ 4000) — these were narrow-column
      reservations the source author made for layouts that depended on
      source's wider effective text column (e.g. for anchored images
      with `wrapTopAndBottom`). With the new geometry, those huge right
      indents would crush text to a sliver. We strip them.
    - `w:hanging` and `w:firstLine` are NOT stripped — they're typically
      list-level positioning that should be preserved.

    Stashes table and textbox content so paragraphs nested inside those
    aren't touched (table cells often have intentional indentation).

    Returns (rewritten_xml, {'left': N, 'right': M}) — counts of how
    many indents of each type were stripped.
    """
    stats = {"left": 0, "right": 0}

    # Stash tables and textboxes — same approach as the heading remapper
    placeholders: list[str] = []

    def _stash(m: re.Match) -> str:
        idx = len(placeholders)
        placeholders.append(m.group(0))
        return f"<!--TPL_INDSTASH_{idx}-->"

    stashed = re.sub(
        r"<w:tbl[\s>][^>]*?>.*?</w:tbl>"
        r"|<w:txbxContent[\s>][^>]*?>.*?</w:txbxContent>"
        r"|<mc:AlternateContent[\s>][^>]*?>.*?</mc:AlternateContent>",
        _stash,
        body_xml,
        flags=re.DOTALL,
    )

    def _normalize_ind(m: re.Match) -> str:
        ind_attrs = m.group(1)
        new_attrs = ind_attrs
        # Strip canonical 1440-twip left compensator
        left_m = re.search(r'\bw:left="(\d+)"', ind_attrs)
        if left_m and left_m.group(1) == "1440":
            new_attrs = re.sub(r'\s*w:left="1440"', "", new_attrs)
            stats["left"] += 1
        # Strip right indents that look like "right margin compensators".
        # Two ranges qualify:
        # - 1000-2000 twips (~0.7-1.4 inch) — the canonical ~1 inch
        #   compensator paired with `pgMar right=0`. Common values
        #   observed in BSI Learning's library: 1434, 1435, 1481, 1523.
        # - ≥ 4000 twips (~2.78 inch) — narrow-column reservations from
        #   source's wider effective text area; would crush text in the
        #   new geometry.
        # Values 0-999 (small adjustments) and 2001-3999 (deliberate
        # medium reservations, e.g. for adjacent floating images) are
        # preserved.
        right_m = re.search(r'\bw:right="(\d+)"', ind_attrs)
        if right_m:
            r = int(right_m.group(1))
            if (1000 <= r <= 2000) or r >= 4000:
                new_attrs = re.sub(r'\s*w:right="\d+"', "", new_attrs)
                stats["right"] += 1
        # If we stripped everything (no attributes left), drop the
        # whole <w:ind/> element. Otherwise rebuild.
        new_attrs = new_attrs.strip()
        if not new_attrs:
            return ""
        return f"<w:ind {new_attrs}/>"

    # Match <w:ind ... /> with attributes and rewrite
    rewritten = re.sub(
        r"<w:ind\s+([^/]*)/>",
        _normalize_ind,
        stashed,
    )

    # Restore stashed regions (tables come back with their original
    # tblInd values intact)
    rewritten = re.sub(
        r"<!--TPL_INDSTASH_(\d+)-->",
        lambda m: placeholders[int(m.group(1))],
        rewritten,
    )

    # Now strip table-level left indent compensators. BSI Learning's
    # tables are positioned with `<w:tblInd w:w="1382"/>` (or similar
    # value in 1000-2000 twip range) to provide a ~1 inch left offset
    # paired with `pgMar left=0`. With template's `pgMar left=1134`,
    # the tblInd stacks on top of the page margin, pushing tables 1.75
    # inches in from the left and often overflowing the right margin.
    # We do this AFTER the stash-restore because tables were stashed
    # during the paragraph indent pass; processing tblInd here catches
    # them after they've come back.
    def _normalize_tblind(m: re.Match) -> str:
        attrs = m.group(1)
        w_m = re.search(r'\bw:w="(\d+)"', attrs)
        if w_m and 1000 <= int(w_m.group(1)) <= 2000:
            stats["tblInd"] = stats.get("tblInd", 0) + 1
            new_attrs = re.sub(r'\bw:w="\d+"', 'w:w="0"', attrs)
            return f"<w:tblInd {new_attrs}/>"
        return m.group(0)

    rewritten = re.sub(
        r"<w:tblInd\s+([^/]*)/>",
        _normalize_tblind,
        rewritten,
    )
    return rewritten, stats


def _remap_body_heading_styles(body_xml: str) -> tuple[str, dict[str, int]]:
    """Heuristically classify direct-formatted headings in body content
    and replace direct formatting with named-style references so the
    template's `styles.xml` governs their appearance.

    The motivating problem: source docs from BSI Learning's RTO library
    use **direct character formatting** for headings — e.g. "Version
    History" is `<w:r><w:rPr><w:b/><w:sz w:val="24"/></w:rPr><w:t>...</w:t></w:r>`,
    not `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>...`. Because
    of that, transplanting the template's `styles.xml` (which redefines
    Heading 1 as orange, Heading 2 as blue, etc.) has zero effect on
    body content — Word looks at the named style only if the paragraph
    references one.

    This function fixes that by walking top-level paragraphs in the
    body, profiling the first run's formatting, classifying the
    paragraph as a heading by size/bold pattern, and rewriting it to:

    - Inject `<w:pStyle w:val="HeadingN"/>` into the pPr
    - Strip direct formatting (`<w:sz>`, `<w:b/>`, `<w:color>`,
      `<w:rFonts>`) from rPr so the named style's formatting takes
      over

    **Heuristic** (size in OOXML half-points; divide by 2 for pt):

    - sz >= 36 + bold + short text → Heading 1 (chapter heading, e.g. 18pt+)
    - sz 32 + bold + short text    → Heading 1 (section heading, 16pt)
    - sz 28 + bold + short text    → Heading 2 (subsection, 14pt)
    - sz 24 + bold + short, doesn't look like a list label → Heading 3 (12pt)
    - everything else                 → leave alone

    **Risks** of misclassification:

    - "5 definitely applies to me" — a survey scale label that happens
      to be 12pt bold. We guard against this by checking the text doesn't
      start with a digit followed by a space.
    - Table header rows (12pt bold) — we skip paragraphs inside <w:tbl>
      and <w:txbxContent> regions entirely (those have specific
      formatting that should be preserved).
    - Paragraphs already carrying a named style — we leave alone.

    Returns (rewritten_xml, stats) where stats counts headings reassigned
    per level.
    """
    stats = {"Heading1": 0, "Heading2": 0, "Heading3": 0, "skipped_already_styled": 0}

    # Stash table and textbox regions so we don't process paragraphs
    # nested inside them. Innermost-first matching by re's non-greedy
    # behaviour is sufficient for the typical Learner Guide structure
    # (tables aren't nested inside textboxes or vice versa in our docs).
    placeholders: list[str] = []

    def _stash(m: re.Match) -> str:
        idx = len(placeholders)
        placeholders.append(m.group(0))
        return f"<!--TPL_STASH_{idx}-->"

    stashed = re.sub(
        r"<w:tbl[\s>][^>]*?>.*?</w:tbl>"
        r"|<w:txbxContent[\s>][^>]*?>.*?</w:txbxContent>"
        r"|<mc:AlternateContent[\s>][^>]*?>.*?</mc:AlternateContent>",
        _stash,
        body_xml,
        flags=re.DOTALL,
    )

    def _classify_and_restyle(m: re.Match) -> str:
        para = m.group(0)
        # Skip if already styled (paragraph references a named style)
        if "<w:pStyle" in para:
            stats["skipped_already_styled"] += 1
            return para
        # Get text content
        texts = re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", para, flags=re.DOTALL)
        text = "".join(texts).strip()
        if not text:
            return para
        # Skip very long text (body paragraphs can have a bold word at the
        # start that would otherwise trigger heading classification)
        if len(text) > 100:
            return para
        # Get first run's formatting
        first_run = re.search(r"<w:r[\s>][^>]*?>(.*?)</w:r>", para, flags=re.DOTALL)
        if not first_run:
            return para
        rpr_block = first_run.group(1)
        sz_m = re.search(r'<w:sz\s+w:val="(\d+)"', rpr_block)
        size = int(sz_m.group(1)) if sz_m else None
        bold = bool(re.search(r"<w:b\s*/>|<w:b\s[^>]*/>", rpr_block))
        # Classify (only bold paragraphs are candidates for headings)
        if not bold or size is None:
            return para

        # Whitelist of known top-level section headings — these should be
        # Heading 1 regardless of source size, because the source author
        # often used 12pt bold for what's semantically a top-level heading.
        # Match case-insensitively because docs vary (ALL CAPS vs sentence
        # case). The list is conservative — entries should be unambiguous
        # heading text in any RTO learner-guide context.
        SECTION_HEADING_PATTERNS = (
            "version history",
            "acknowledgement of country",
            "acknowledgment of country",
            "table of contents",
            "unit description",
            "application of unit",
            "performance criteria",
            "skills and knowledge",
            "performance evidence",
            "knowledge evidence",
            "assessment conditions",
            "elements and performance criteria",
            "foundation skills",
            "unit mapping information",
        )
        text_lower = text.lower().strip().rstrip(":")
        is_known_heading = text_lower in SECTION_HEADING_PATTERNS

        target_style: str | None = None
        if is_known_heading:
            target_style = "Heading1"
        elif size >= 36:
            target_style = "Heading1"
        elif size >= 32:
            target_style = "Heading1"
        elif size >= 28:
            target_style = "Heading2"
        elif size >= 24 and len(text) < 60:
            # 12pt bold short text — be careful of false positives
            if re.match(r"^\d\s", text):
                # Survey scale label: "5 definitely applies to me"
                return para
            if len(text) < 4:
                # Too short — likely a fragment ("And", "Or")
                return para
            target_style = "Heading3"
        if target_style is None:
            return para
        # Apply: inject pStyle and strip direct formatting
        return _apply_named_style(para, target_style, stats)

    rewritten = re.sub(
        r"<w:p(?:\s[^>]*)?(?:/>|>.*?</w:p>)",
        _classify_and_restyle,
        stashed,
        flags=re.DOTALL,
    )

    # Restore stashed regions
    def _restore(m: re.Match) -> str:
        return placeholders[int(m.group(1))]

    rewritten = re.sub(r"<!--TPL_STASH_(\d+)-->", _restore, rewritten)
    return rewritten, stats


def _apply_named_style(
    para_xml: str, style_name: str, stats: dict[str, int]
) -> str:
    """Inject <w:pStyle w:val="style_name"/> into a paragraph's pPr and
    strip direct formatting from its rPr blocks so the named style takes
    full effect.
    """
    pstyle_xml = f'<w:pStyle w:val="{style_name}"/>'

    # Inject pStyle at the start of pPr (creating pPr if needed)
    if "<w:pPr>" in para_xml:
        para_xml = para_xml.replace("<w:pPr>", f"<w:pPr>{pstyle_xml}", 1)
    elif re.search(r"<w:pPr\s[^>]*>", para_xml):
        para_xml = re.sub(
            r"(<w:pPr\s[^>]*>)", rf"\1{pstyle_xml}", para_xml, count=1
        )
    else:
        # No pPr — insert one right after the opening <w:p ...>
        para_xml = re.sub(
            r"(<w:p(?:\s[^>]*)?>)",
            rf"\1<w:pPr>{pstyle_xml}</w:pPr>",
            para_xml,
            count=1,
        )

    # Strip direct formatting from every rPr block in the paragraph.
    # This is what makes the named style fully take over — without
    # stripping, the direct sz/b/color win over the style's definition.
    def _strip_direct_fmt(m: re.Match) -> str:
        block = m.group(0)
        block = re.sub(r"<w:sz\s+w:val=\"\d+\"\s*/>", "", block)
        block = re.sub(r"<w:szCs\s+w:val=\"\d+\"\s*/>", "", block)
        block = re.sub(r"<w:b\s*/>", "", block)
        block = re.sub(r"<w:bCs\s*/>", "", block)
        block = re.sub(r"<w:color\s[^/]*/>", "", block)
        block = re.sub(r"<w:rFonts\s[^/]*/>", "", block)
        return block

    para_xml = re.sub(
        r"<w:rPr>.*?</w:rPr>", _strip_direct_fmt, para_xml, flags=re.DOTALL
    )
    para_xml = re.sub(
        r"<w:rPr\s[^>]*>.*?</w:rPr>",
        _strip_direct_fmt,
        para_xml,
        flags=re.DOTALL,
    )

    stats[style_name] = stats.get(style_name, 0) + 1
    return para_xml


def transplant_chrome(
    src_path: Path,
    template_path: Path,
    dst_path: Path,
    unit_code: str | None = None,
    unit_title: str | None = None,
    spec: dict[str, Any] | None = None,
    asset_dir: Path | None = None,
) -> dict[str, Any]:
    """Transplant the template's cover/header/footer/theme onto src, preserving body.

    Pipeline:
    1. Extract unit_code/unit_title from src if not provided.
    2. Read template's chrome XML (theme1, header1, footer1, footer2,
       cover region of document.xml) and media files referenced.
    3. Build the output: source body + template chrome with substitutions.
    4. After zip is written, run the standard rebrand_docx pass on it for
       any leftover BSI references in the body.
    """
    src_path = Path(src_path)
    template_path = Path(template_path)
    dst_path = Path(dst_path)
    dst_path.parent.mkdir(parents=True, exist_ok=True)

    log: dict[str, Any] = {
        "source": str(src_path),
        "destination": str(dst_path),
        "unit_code": unit_code or "",
        "unit_title": unit_title or "",
        "method": "template_transplant",
        "actions": [],
        "warnings": [],
    }

    # -- Stage 1: extract data from both inputs ---------------------------
    with zipfile.ZipFile(src_path, "r") as zsrc, zipfile.ZipFile(template_path, "r") as ztpl:
        # Extract unit info if not given
        if not unit_code or not unit_title:
            extracted_code, extracted_title = extract_unit_info(zsrc, src_path.name)
            unit_code = unit_code or extracted_code
            unit_title = unit_title or extracted_title
            log["unit_code"] = unit_code
            log["unit_title"] = unit_title
        log["actions"].append(f"Extracted unit_code='{unit_code}', unit_title='{unit_title}'")

        # Read template chrome XMLs
        # Also load ALL template files into a dict so later stages can pull
        # styles.xml, fontTable.xml, embedded fonts, etc. without needing
        # to keep the zip handle open.
        tpl_files: dict[str, bytes] = {n: ztpl.read(n) for n in ztpl.namelist()}
        tpl_doc = tpl_files["word/document.xml"].decode("utf-8")
        tpl_theme = tpl_files["word/theme/theme1.xml"]
        tpl_header1 = tpl_files["word/header1.xml"].decode("utf-8")
        tpl_footer1 = tpl_files["word/footer1.xml"].decode("utf-8")
        tpl_footer2 = tpl_files["word/footer2.xml"].decode("utf-8")
        tpl_styles = tpl_files["word/styles.xml"]
        tpl_numbering = tpl_files.get("word/numbering.xml")

        # Read template's chrome relationship files
        tpl_doc_rels = ztpl.read("word/_rels/document.xml.rels").decode("utf-8")
        tpl_header1_rels = (
            ztpl.read("word/_rels/header1.xml.rels").decode("utf-8")
            if "word/_rels/header1.xml.rels" in ztpl.namelist()
            else ""
        )
        tpl_footer2_rels = (
            ztpl.read("word/_rels/footer2.xml.rels").decode("utf-8")
            if "word/_rels/footer2.xml.rels" in ztpl.namelist()
            else ""
        )

        # Find which media files the template's header/footer reference
        tpl_chrome_media: dict[str, bytes] = {}
        for rels_xml in (tpl_header1_rels, tpl_footer2_rels):
            for _, _, target in _list_relationships(rels_xml):
                # Targets are relative to header1.xml's location, i.e. word/header1.xml
                # so target "media/image3.png" → "word/media/image3.png"
                media_path = "word/" + target.lstrip("/")
                if media_path in ztpl.namelist():
                    tpl_chrome_media[media_path] = ztpl.read(media_path)

        # Historical note: an earlier version of this code substituted the
        # template's header logo with a green-themed `aex_landscape_green.svg`
        # bundled in assets. That was needed when the template was a generic
        # learner-guide template without brand styling. Now that the bundled
        # template (FSKNUM017_Learner_Guide_V4_0_aEX.docx) is the user's
        # college's gold-standard with the correct BLUE accent6 wordmark
        # already in place, we should NOT substitute — keeping the template's
        # original header media gives the right colour and dimensions.

        # Extract template's cover XML region
        _, tpl_cover, _, _ = _split_cover_and_body(tpl_doc)

        # Capture template's page geometry (size + margins) from its first
        # sectPr. Some template features — particularly cover footers with
        # anchored elements positioned by negative offsets — depend on the
        # template's smaller bottom margin and footer offset to render
        # correctly. Without copying these, anchored cover-footer content
        # (e.g. blue landscape logo + copyright) lands off-page.
        tpl_first_sect = re.search(
            r"<w:sectPr[\s>].*?</w:sectPr>", tpl_doc, flags=re.DOTALL
        )
        tpl_pgSz_xml: str | None = None
        tpl_pgMar_xml: str | None = None
        if tpl_first_sect:
            sect_text = tpl_first_sect.group(0)
            m_sz = re.search(r"<w:pgSz[^/>]*/?>", sect_text)
            m_mar = re.search(r"<w:pgMar[^/>]*/?>", sect_text)
            if m_sz:
                tpl_pgSz_xml = m_sz.group(0)
                # Self-close if needed
                if not tpl_pgSz_xml.endswith("/>"):
                    tpl_pgSz_xml = tpl_pgSz_xml.rstrip(">") + "/>"
            if m_mar:
                tpl_pgMar_xml = m_mar.group(0)
                if not tpl_pgMar_xml.endswith("/>"):
                    tpl_pgMar_xml = tpl_pgMar_xml.rstrip(">") + "/>"
        log["template_pgSz"] = tpl_pgSz_xml
        log["template_pgMar"] = tpl_pgMar_xml

        # The template's cover XML references images (the lime monogram on
        # the cover top-right) via rIds that resolve through the TEMPLATE'S
        # document.xml.rels. We need to:
        # 1. Find the template's cover image rIds
        # 2. Map each rId → media file path via template's rels
        # 3. Migrate those media files into the source doc (with renamed
        #    filenames to avoid conflicts)
        # 4. Generate NEW rIds in the source doc's rels pointing to the
        #    migrated files
        # 5. Rewrite the template cover XML to use the new rIds
        tpl_cover_rids = set(re.findall(r'r:embed="([^"]+)"', tpl_cover))
        tpl_doc_rels_map = {
            rid: target
            for rid, _, target in _list_relationships(tpl_doc_rels)
        }
        # Read template's cover-image bytes (keyed by source rId)
        tpl_cover_image_data: dict[str, tuple[str, bytes]] = {}
        for rid in tpl_cover_rids:
            if rid in tpl_doc_rels_map:
                target = tpl_doc_rels_map[rid]
                media_path = "word/" + target.lstrip("/")
                if media_path in ztpl.namelist():
                    tpl_cover_image_data[rid] = (media_path, ztpl.read(media_path))

        # Substitute the template's unit code/title with the source's.
        # We extract the template's own unit_code+title (typically the
        # FSKLRG013 sample text or "Unit Code"/"Unit Name" placeholders),
        # then use them as the find-target. This makes the script work
        # with any template — placeholder-based, sample-data, or already-
        # filled-in templates from past rebrands.
        tpl_unit_code, tpl_unit_title = extract_unit_info(ztpl, template_path.name)
        log["template_unit_code"] = tpl_unit_code
        log["template_unit_title"] = tpl_unit_title

        # The title gets case-insensitive substitution because templates
        # often use Title Case in the cover and sentence-case in footers.
        # Both will be replaced with the source's canonical title.
        title_subs: list[tuple[str, str]] = []
        if tpl_unit_title and tpl_unit_title.lower() != unit_title.lower():
            title_subs.append((tpl_unit_title, unit_title))

        # Unit code stays case-sensitive (it's an identifier — accidentally
        # matching "BSBCMM" mid-word would be bad).
        code_subs: list[tuple[str, str]] = []
        if tpl_unit_code and tpl_unit_code != unit_code:
            code_subs.append((tpl_unit_code, unit_code))

        # Legacy explicit placeholders some templates use
        legacy_subs = [("Unit Code", unit_code), ("Unit Name", unit_title)]

        def _apply_all_subs(xml: str) -> str:
            # Title (case-insensitive), then code (case-sensitive), then
            # legacy placeholders (case-sensitive), then footer metadata
            # (version, dates) extracted from the source doc.
            xml = _substitute_text_in_xml(xml, title_subs, case_insensitive=True)
            xml = _substitute_text_in_xml(xml, code_subs)
            xml = _substitute_text_in_xml(xml, legacy_subs)
            xml = _substitute_text_in_xml(xml, footer_meta_subs)
            return xml

        # Footer metadata: extract version + dates from the source's own
        # footer (so e.g. BSBCMM511 v 1.0's footer keeps its v 1.0 / 16/8/21
        # rather than picking up the template's V4.0 / 26.03.2026 defaults).
        # If the source has no footer or doesn't have these markers, we
        # leave the template's values alone — they're easy to spot and
        # update manually after rebrand.
        src_footer_meta = _extract_footer_metadata(zsrc)
        tpl_footer_meta = _extract_footer_metadata(ztpl)
        log["source_footer_metadata"] = src_footer_meta
        log["template_footer_metadata"] = tpl_footer_meta

        footer_meta_subs: list[tuple[str, str]] = []
        # Version: "V4.0" → "v 1.0" (longer/specific token; case-sensitive
        # because we want to preserve the source's exact capitalisation)
        if (
            src_footer_meta.get("version")
            and tpl_footer_meta.get("version")
            and src_footer_meta["version"] != tpl_footer_meta["version"]
        ):
            footer_meta_subs.append(
                (str(tpl_footer_meta["version"]), str(src_footer_meta["version"]))
            )
        # Dates: position-mapped (first → first, second → second, etc.).
        # If the source has fewer dates than the template, the extras stay
        # as template defaults — better than nuking them with empty strings.
        src_dates = src_footer_meta.get("dates") or []
        tpl_dates = tpl_footer_meta.get("dates") or []
        if isinstance(src_dates, list) and isinstance(tpl_dates, list):
            for i, tpl_date in enumerate(tpl_dates):
                if i < len(src_dates) and src_dates[i] != tpl_date:
                    footer_meta_subs.append((tpl_date, src_dates[i]))

        tpl_cover_filled = _apply_all_subs(tpl_cover)
        tpl_footer1_filled = _apply_all_subs(tpl_footer1)
        if tpl_footer2:
            tpl_footer2_filled = _apply_all_subs(tpl_footer2)
        else:
            tpl_footer2_filled = tpl_footer2
        # Header may also contain the unit info (some templates put it there)
        if tpl_header1:
            tpl_header1_filled = _apply_all_subs(tpl_header1)
        else:
            tpl_header1_filled = tpl_header1

        # Read source contents (will mutate)
        src_files: dict[str, bytes] = {}
        for name in zsrc.namelist():
            src_files[name] = zsrc.read(name)

    # -- Stage 2: rebuild the source doc with chrome transplant ------------

    # 2a. Replace theme + styles + numbering with template's. Without
    # styles.xml, the cover's text references like <w:pStyle val="Heading1excludeTOC"/>
    # resolve to default formatting (small left-aligned), losing the
    # large centered title styling. The template's styles.xml is also
    # what gives body headings the orange/blue treatment seen in the target.
    src_files["word/theme/theme1.xml"] = tpl_theme
    log["actions"].append("Replaced word/theme/theme1.xml")
    if "word/styles.xml" in src_files:
        # Transplant the template's styles.xml, then normalise every
        # <w:rFonts> to the two-font brand system:
        #   * Display (Heading*/Title/Subtitle, plus any style already
        #     referencing Obviously Narw Semi) -> Obviously Narw Semi
        #   * Body (Normal, lists, tables, footers, docDefaults, …)
        #     -> Edu Favorit Light
        # The bundled May2026 template still carries legacy font
        # references (Avenir in docDefaults, Georgia in Subtitle, and
        # Times New Roman/Arial/Arial Unicode MS/Calibri inside list,
        # bullet, and "Heading1NumberedList" styles). The user's brand
        # standard is two fonts only, so we collapse those leftovers
        # to the brand pair here. Embedded fonts in word/fonts/
        # font*.odttf (copied below) ensure these render correctly on
        # machines without the brand fonts installed.
        styles_text = tpl_styles.decode("utf-8")
        styles_text, font_stats = _normalise_styles_to_two_fonts(styles_text)
        src_files["word/styles.xml"] = styles_text.encode("utf-8")
        log["font_normalisation"] = font_stats
        log["actions"].append(
            "Replaced word/styles.xml with template's; normalised to two-font system "
            f"(display={font_stats['display']} body={font_stats['body']} "
            f"rFonts rewritten={font_stats['rfonts_rewritten']})"
        )
    # NB: We deliberately do NOT replace word/numbering.xml. Source body
    # paragraphs reference list IDs (numId, abstractNumId) that are
    # defined in the source's own numbering.xml. If we replace it with
    # the template's, the IDs collide — source paragraphs that ask for
    # "numId 3" suddenly get a totally different list format from the
    # template's abstractNumId 3, producing nonsense like Greek-letter
    # bullets instead of • bullets, missing list indents, and so on.
    # Keep the source's numbering.xml; the cover content we splice in
    # from the template is plain text without list references, so no
    # template list IDs are needed in the output.

    # 2b. Plan: migrate template's chrome media files to src's word/media/
    # with renamed file names, and build a remap dict for rId/path updates.
    src_media_names = {
        Path(n).name for n in src_files if n.startswith("word/media/")
    }
    chrome_media_remap: dict[str, str] = {}  # tpl_path → src_path (renamed)
    for tpl_media_path, data in tpl_chrome_media.items():
        ext = Path(tpl_media_path).suffix
        idx = _next_image_index(src_media_names)
        new_name = f"image{idx}{ext}"
        new_path = f"word/media/{new_name}"
        src_files[new_path] = data
        src_media_names.add(new_name)
        chrome_media_remap[tpl_media_path] = new_path
        log["actions"].append(f"Migrated {tpl_media_path} → {new_path}")

    # 2c. Migrate template cover images (placeholder — done in 2c.2 below
    # after we've parsed src's existing rels and can allocate new rIds).
    cover_rid_remap: dict[str, str] = {}  # tpl_rid → src_rid

    # 2c.1 Build source's document.xml.rels: parse current, ensure header1 and
    # footer2 relationships exist (footer1 likely already exists from src).
    src_doc_rels = src_files.get(
        "word/_rels/document.xml.rels",
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        b'</Relationships>',
    ).decode("utf-8")

    src_doc_relids = {rid for rid, _, _ in _list_relationships(src_doc_rels)}
    src_rel_targets = {target: rid for rid, _, target in _list_relationships(src_doc_rels)}

    # We need rIds in document.xml.rels for: header1.xml, footer1.xml, footer2.xml,
    # theme/theme1.xml. Add any that are missing.
    chrome_doc_relationships_needed = [
        ("header", "header1.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"),
        ("footer1", "footer1.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"),
        ("footer2", "footer2.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"),
        ("theme", "theme/theme1.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"),
    ]
    chrome_rids: dict[str, str] = {}  # name → rId
    for name, target, type_url in chrome_doc_relationships_needed:
        if target in src_rel_targets:
            chrome_rids[name] = src_rel_targets[target]
        else:
            new_rid = _next_rid(src_doc_relids)
            src_doc_relids.add(new_rid)
            new_rel = (
                f'<Relationship Id="{new_rid}" Type="{type_url}" Target="{target}"/>'
            )
            src_doc_rels = src_doc_rels.replace(
                "</Relationships>", f"{new_rel}</Relationships>"
            )
            chrome_rids[name] = new_rid
            log["actions"].append(f"Added rels: {new_rid} → {target}")
    src_files["word/_rels/document.xml.rels"] = src_doc_rels.encode("utf-8")

    # 2c.2 Migrate the template's COVER images (the lime monogram etc.)
    # into source's media folder. The template's cover XML references these
    # images via rIds that resolve through the TEMPLATE'S rels — but those
    # same rIds in the source doc point to entirely different files (e.g.
    # template's rId12 = lime monogram SVG, but BSI's rId12 = orange-arrow
    # decoration EMF). If we splice without remapping, the wrong image
    # renders. Solution:
    #   - Migrate each template cover image with a new filename
    #   - Allocate a new rId in source's rels
    #   - Build a remap dict: tpl_rid → src_rid
    #   - Rewrite the cover XML to use the new rIds
    image_rel_type = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    )
    new_cover_rels_xml = ""
    cover_image_paths: set[str] = set()  # for skip_paths in body cleanup
    for tpl_rid, (tpl_media_path, data) in tpl_cover_image_data.items():
        ext = Path(tpl_media_path).suffix
        idx = _next_image_index(src_media_names)
        new_name = f"image{idx}{ext}"
        new_path = f"word/media/{new_name}"
        src_files[new_path] = data
        src_media_names.add(new_name)
        cover_image_paths.add(new_path)
        # Allocate new rId in src
        new_rid = _next_rid(src_doc_relids)
        src_doc_relids.add(new_rid)
        cover_rid_remap[tpl_rid] = new_rid
        new_cover_rels_xml += (
            f'<Relationship Id="{new_rid}" Type="{image_rel_type}" '
            f'Target="media/{new_name}"/>'
        )
        log["actions"].append(
            f"Migrated cover image {tpl_media_path} → {new_path} (rId {tpl_rid}→{new_rid})"
        )
    if new_cover_rels_xml:
        # Append to document.xml.rels
        src_doc_rels = src_files["word/_rels/document.xml.rels"].decode("utf-8")
        src_doc_rels = src_doc_rels.replace(
            "</Relationships>", f"{new_cover_rels_xml}</Relationships>"
        )
        src_files["word/_rels/document.xml.rels"] = src_doc_rels.encode("utf-8")

    # Rewrite the spliced cover XML to use the new rIds
    def _remap_cover_rids(xml: str) -> str:
        for tpl_rid, src_rid in cover_rid_remap.items():
            # Replace r:embed="tpl_rid" with r:embed="src_rid"
            # Be careful to use word boundaries
            xml = xml.replace(f'r:embed="{tpl_rid}"', f'r:embed="{src_rid}"')
            xml = xml.replace(f'r:link="{tpl_rid}"', f'r:link="{src_rid}"')
            xml = xml.replace(f'r:id="{tpl_rid}"', f'r:id="{src_rid}"')
        return xml

    tpl_cover_filled = _remap_cover_rids(tpl_cover_filled)

    # 2d. Write template's footer/header XMLs into src
    src_files["word/footer1.xml"] = tpl_footer1_filled.encode("utf-8")
    log["actions"].append("Replaced word/footer1.xml with template (substituted)")
    src_files["word/footer2.xml"] = tpl_footer2_filled.encode("utf-8")
    log["actions"].append("Added/replaced word/footer2.xml from template")

    # The template's header dimensions match the template's own logo
    # (which we are NOT substituting any more); no patching needed.
    tpl_header1_patched = tpl_header1_filled

    src_files["word/header1.xml"] = tpl_header1_patched.encode("utf-8")
    log["actions"].append("Added/replaced word/header1.xml from template")

    # 2e. Build the rels for header1 and footer2 inside src, with rIds that
    # reference the migrated media. The template's header1.xml uses rId1/rId2,
    # so we rebuild a fresh rels file with rId1/rId2 pointing to the src's
    # renamed media files.
    def _build_header_footer_rels(tpl_rels_xml: str) -> str:
        """Take template's header1/footer2 rels, rewrite Target paths to point
        to the migrated files in src's media folder."""
        out = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        )
        for rid, type_url, target in _list_relationships(tpl_rels_xml):
            tpl_full_path = "word/" + target.lstrip("/")
            new_full_path = chrome_media_remap.get(tpl_full_path, tpl_full_path)
            # Convert back to relative target (relative to word/header1.xml)
            new_target = new_full_path.replace("word/", "", 1)
            out += f'<Relationship Id="{rid}" Type="{type_url}" Target="{new_target}"/>'
        out += "</Relationships>"
        return out

    if tpl_header1_rels:
        src_files["word/_rels/header1.xml.rels"] = _build_header_footer_rels(
            tpl_header1_rels
        ).encode("utf-8")
    if tpl_footer2_rels:
        src_files["word/_rels/footer2.xml.rels"] = _build_header_footer_rels(
            tpl_footer2_rels
        ).encode("utf-8")

    # 2f. Update [Content_Types].xml — ensure footer2 / header1 are declared
    ct_xml = src_files["[Content_Types].xml"].decode("utf-8")
    additions = []
    if 'PartName="/word/footer2.xml"' not in ct_xml:
        additions.append(
            '<Override PartName="/word/footer2.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
        )
    if 'PartName="/word/header1.xml"' not in ct_xml:
        additions.append(
            '<Override PartName="/word/header1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        )
    # Make sure PNG and SVG default extensions are declared
    if 'Extension="png"' not in ct_xml:
        additions.append('<Default Extension="png" ContentType="image/png"/>')
    if 'Extension="svg"' not in ct_xml:
        additions.append('<Default Extension="svg" ContentType="image/svg+xml"/>')
    if additions:
        ct_xml = ct_xml.replace("</Types>", "".join(additions) + "</Types>")
        src_files["[Content_Types].xml"] = ct_xml.encode("utf-8")
        log["actions"].append(f"Added {len(additions)} entries to [Content_Types].xml")

    # 2g. Transplant template's styles.xml so the source's body content
    # (which uses named styles like "Heading1") picks up the template's
    # styling — orange Heading 1, blue Heading 2, etc. The template's
    # styles.xml is a strict superset of the source's heading style names,
    # so existing references resolve correctly. Without this step the
    # body keeps source's BSI-orange Heading 1 styling.
    # Fallback: if the source doc had no styles.xml at all (rare — most .docx
    # files have one even if minimal), drop in the template's so heading
    # references in the cover and body still resolve. Stage 2a above is the
    # authoritative path when source has its own styles.xml. Either path
    # runs the two-font normaliser so output is consistent.
    if "word/styles.xml" in tpl_files and "word/styles.xml" not in src_files:
        styles_text = tpl_files["word/styles.xml"].decode("utf-8")
        styles_text, font_stats = _normalise_styles_to_two_fonts(styles_text)
        src_files["word/styles.xml"] = styles_text.encode("utf-8")
        log["font_normalisation"] = font_stats
        log["actions"].append(
            "Transplanted word/styles.xml from template (source had none); "
            f"normalised to two-font system (display={font_stats['display']} "
            f"body={font_stats['body']} rFonts rewritten={font_stats['rfonts_rewritten']})"
        )

    # Also transplant fontTable.xml + embedded fonts so heading-text fonts
    # are available. If the template has fonts/font*.odttf files that the
    # source doesn't, we copy them over.
    if "word/fontTable.xml" in tpl_files:
        src_files["word/fontTable.xml"] = tpl_files["word/fontTable.xml"]
        log["actions"].append("Transplanted word/fontTable.xml from template")
        # Copy embedded font files
        font_count = 0
        for name, data in tpl_files.items():
            if name.startswith("word/fonts/") and name.endswith(".odttf"):
                src_files[name] = data
                font_count += 1
        if font_count:
            log["actions"].append(f"Copied {font_count} embedded font(s) from template")
        # Copy fontTable rels if present
        if "word/_rels/fontTable.xml.rels" in tpl_files:
            src_files["word/_rels/fontTable.xml.rels"] = tpl_files[
                "word/_rels/fontTable.xml.rels"
            ]
        # Update [Content_Types].xml for .odttf
        ct_xml = src_files["[Content_Types].xml"].decode("utf-8")
        if 'Extension="odttf"' not in ct_xml:
            ct_xml = ct_xml.replace(
                "</Types>",
                '<Default Extension="odttf" '
                'ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/>'
                "</Types>",
            )
            src_files["[Content_Types].xml"] = ct_xml.encode("utf-8")

    # Same for theme1.xml (the template's theme defines accent colours used
    # by named styles). Without this swap, heading colours referencing
    # theme accents would resolve to source's old theme.
    if "word/theme/theme1.xml" in tpl_files:
        src_files["word/theme/theme1.xml"] = tpl_files["word/theme/theme1.xml"]
        log["actions"].append("Transplanted word/theme/theme1.xml from template")

    # 2h. Surgery on document.xml: replace cover region with template's
    # filled cover, and update sectPr to reference correct header/footer rIds.
    src_doc = src_files["word/document.xml"].decode("utf-8")

    # Merge any namespace declarations the template uses but the source's
    # `<w:document>` root doesn't declare. Older Word docs (2010-era and
    # earlier) carry only a subset of namespaces — typically missing
    # `wp14`, `cx`, `cx1..cx8`, `wpi`, `oel`, `aink`, `am3d`, `w15`, `w16*`.
    # The template's cover content uses some of these (`wp14:anchorId`,
    # `mc:AlternateContent` with `Choice Requires="wps"`, etc.), and
    # splicing it into a doc without those declarations produces an
    # "unbound prefix" XML parse error that LibreOffice and Word both
    # reject. Merging is safe: extra namespace declarations are harmless,
    # and any prefix the source already uses keeps its source URI.
    src_root_m = re.search(r"<w:document([^>]*)>", src_doc)
    tpl_root_m = re.search(r"<w:document([^>]*)>", tpl_doc)
    if src_root_m and tpl_root_m:
        src_attrs = src_root_m.group(1)
        tpl_attrs = tpl_root_m.group(1)
        src_prefixes = set(re.findall(r'xmlns:(\w+)=', src_attrs))
        tpl_xmlns = re.findall(r'(xmlns:\w+="[^"]+")', tpl_attrs)
        additions = []
        for decl in tpl_xmlns:
            prefix_m = re.match(r'xmlns:(\w+)=', decl)
            if prefix_m and prefix_m.group(1) not in src_prefixes:
                additions.append(decl)
        # Also merge mc:Ignorable values so the IGnore list covers the new prefixes
        src_ignor = re.search(r'mc:Ignorable="([^"]*)"', src_attrs)
        tpl_ignor = re.search(r'mc:Ignorable="([^"]*)"', tpl_attrs)
        merged_ignor = None
        if src_ignor and tpl_ignor:
            tokens = set(src_ignor.group(1).split()) | set(tpl_ignor.group(1).split())
            merged_ignor = " ".join(sorted(tokens))
        elif tpl_ignor:
            merged_ignor = tpl_ignor.group(1)
        if additions or merged_ignor:
            new_attrs = src_attrs
            if additions:
                new_attrs = new_attrs.rstrip() + " " + " ".join(additions)
            if merged_ignor:
                if 'mc:Ignorable=' in new_attrs:
                    new_attrs = re.sub(
                        r'mc:Ignorable="[^"]*"',
                        f'mc:Ignorable="{merged_ignor}"',
                        new_attrs,
                    )
                else:
                    new_attrs = new_attrs.rstrip() + f' mc:Ignorable="{merged_ignor}"'
            src_doc = src_doc.replace(
                src_root_m.group(0), f"<w:document{new_attrs}>", 1
            )
            log["actions"].append(
                f"Merged {len(additions)} namespace declarations from template "
                f"into source's <w:document> root"
            )

    prefix, _src_cover, src_body, suffix = _split_cover_and_body(src_doc)

    # Update sectPr references in src_body to point to the chrome rIds.
    # In the source body there'll be a final <w:sectPr> that references
    # headerReference/footerReference rIds. We rewrite those to the new ones.
    def _update_sect_refs(xml: str) -> str:
        # Replace any <w:headerReference w:type="default" r:id="..."/> with our header rId
        xml = re.sub(
            r'<w:headerReference[^/]*r:id="[^"]*"[^/]*/>',
            f'<w:headerReference w:type="default" r:id="{chrome_rids["header"]}"/>',
            xml,
        )
        # And footerReference for default type
        xml = re.sub(
            r'<w:footerReference\s+w:type="default"[^/]*r:id="[^"]*"[^/]*/>',
            f'<w:footerReference w:type="default" r:id="{chrome_rids["footer1"]}"/>',
            xml,
        )
        # And footerReference for first-page (cover) — points to footer2 which
        # has the cover's blue landscape logo + copyright line.
        xml = re.sub(
            r'<w:footerReference\s+w:type="first"[^/]*r:id="[^"]*"[^/]*/>',
            f'<w:footerReference w:type="first" r:id="{chrome_rids["footer2"]}"/>',
            xml,
        )

        # Older BSI Learning docs have a peculiar geometry pattern: the
        # body section's `pgMar` has `left=0 right=0`, and individual
        # paragraphs (and heading styles) carry `<w:ind w:left="1440">`
        # to compensate. The apparent 1-inch left margin in the source's
        # render is built from per-paragraph indents, NOT page margins.
        # When we transplant template's `styles.xml`, the heading styles
        # lose their compensating indent (template's heading styles have
        # no `<w:ind>`), and headings collapse to the page edge while
        # body text — which has its own paragraph-level compensators —
        # stays indented. Result: misaligned headings vs body, and
        # chrome (header/footer) sits at the page edge.
        #
        # To produce a clean, uniformly-margined output that matches the
        # gold standard, we:
        # 1. Override the body sectPr's `left/right` to template's values
        #    (gives consistent page margins for headings AND chrome).
        # 2. Strip the compensator indent (`<w:ind w:left="1440">`)
        #    from body paragraphs, so they don't double-indent on top of
        #    the page margin.
        # The other pgMar attributes (top, bottom, header, footer) come
        # from template via the per-attribute merge below.
        if tpl_pgMar_xml:
            tpl_attrs: dict[str, str] = {}
            for attr_m in re.finditer(r'w:(\w+)="(\d+)"', tpl_pgMar_xml):
                tpl_attrs[attr_m.group(1)] = attr_m.group(2)

            def _merge_pgmar(m: re.Match) -> str:
                src_pgmar = m.group(0)
                merged = src_pgmar
                # Override every margin attribute from template, except
                # source-specific ones we want to preserve (currently:
                # none — full override, since the compensator stripping
                # below makes paragraph wrap independent of section
                # left/right).
                for attr in ("top", "bottom", "left", "right", "header", "footer", "gutter"):
                    if attr in tpl_attrs:
                        if re.search(rf'\bw:{attr}="\d+"', merged):
                            merged = re.sub(
                                rf'\bw:{attr}="\d+"',
                                f'w:{attr}="{tpl_attrs[attr]}"',
                                merged,
                            )
                        else:
                            merged = merged.replace(
                                "/>", f' w:{attr}="{tpl_attrs[attr]}"/>'
                            )
                return merged

            xml = re.sub(r"<w:pgMar[^/>]*/?>", _merge_pgmar, xml)

        # If the doc is missing required references inside any sectPr, inject
        # them right after the opening tag. Order matters in OOXML — header
        # references should come before footer references, but Word is
        # forgiving in practice.
        def _ensure_refs(m: re.Match) -> str:
            sect = m.group(0)
            if "headerReference" not in sect:
                sect = re.sub(
                    r"(<w:sectPr[^>]*>)",
                    rf'\1<w:headerReference w:type="default" r:id="{chrome_rids["header"]}"/>',
                    sect,
                    count=1,
                )
            if 'footerReference w:type="default"' not in sect:
                sect = re.sub(
                    r"(<w:sectPr[^>]*>)",
                    rf'\1<w:footerReference w:type="default" r:id="{chrome_rids["footer1"]}"/>',
                    sect,
                    count=1,
                )
            # Add a first-page footer reference so the cover gets the
            # cover-specific footer (with the brand wordmark + copyright).
            if 'footerReference w:type="first"' not in sect:
                sect = re.sub(
                    r"(<w:sectPr[^>]*>)",
                    rf'\1<w:footerReference w:type="first" r:id="{chrome_rids["footer2"]}"/>',
                    sect,
                    count=1,
                )
            # NOTE: We do NOT auto-add <w:titlePg/> here. The cover-specific
            # sectPr is injected explicitly at the end of the cover content
            # by the splice code below; the body sectPr (the one this
            # function is editing) should NOT have titlePg, otherwise the
            # first page of the body section (which is page 2 of the doc)
            # gets the cover footer instead of the body footer.
            return sect

        xml = re.sub(
            r"<w:sectPr[\s>][^>]*>.*?</w:sectPr>",
            _ensure_refs,
            xml,
            flags=re.DOTALL,
        )
        return xml

    src_body_updated = _update_sect_refs(src_body)

    # Heuristic body-style remap: detect direct-formatted headings
    # ("Version History", chapter titles, etc.) and replace direct
    # formatting with named-style references so the template's
    # styles.xml governs their appearance (orange Heading 1, blue
    # Heading 2, etc.). Without this step the body keeps its source's
    # plain bold-black headings even after styles.xml is transplanted.
    src_body_updated, heading_stats = _remap_body_heading_styles(src_body_updated)
    log["heading_remap"] = heading_stats
    log["actions"].append(
        f"Heuristic heading remap: H1={heading_stats.get('Heading1', 0)}, "
        f"H2={heading_stats.get('Heading2', 0)}, H3={heading_stats.get('Heading3', 0)}"
    )

    # Strip "compensator" indents from body paragraphs. BSI Learning's
    # older docs paired `pgMar left=0 right=0` with `<w:ind w:left="1440"/>`
    # on most body paragraphs to fake a 1-inch left margin. Now that
    # we've overridden the body section's `left/right` to template's
    # non-zero values, those compensators would produce double
    # indentation. Strip ONLY the canonical 1440-twip left compensator
    # (legitimate non-1440 indents — e.g. nested list levels at 2159 or
    # 2880 — are preserved). Right indents are also normalised: any
    # right indent ≥ 4000 twips is treated as a column-narrowing
    # leftover from the source's wider effective text area, and
    # stripped so text wraps to the new (smaller) effective column.
    src_body_updated, indent_stats = _strip_compensator_indents(src_body_updated)
    log["compensator_strip"] = indent_stats
    log["actions"].append(
        f"Stripped compensator indents: left={indent_stats.get('left', 0)}, "
        f"right={indent_stats.get('right', 0)}, "
        f"tblInd={indent_stats.get('tblInd', 0)}"
    )

    # Switch body-textbox autofit from noAutofit to normAutofit so text
    # shrinks to fit when the substitute font is wider than the
    # template's Avenir (e.g. Liberation Sans on Linux, system default
    # on machines without Avenir installed). Source authors used
    # `<a:noAutofit/>` because they were designing for a known font;
    # rebranded docs may render on machines with different fonts and
    # need autofit to avoid overflow.
    autofit_count = src_body_updated.count("<a:noAutofit/>")
    src_body_updated = src_body_updated.replace(
        "<a:noAutofit/>", "<a:normAutofit/>"
    )
    log["actions"].append(
        f"Body textboxes: changed {autofit_count} <a:noAutofit/> → <a:normAutofit/>"
    )

    # Build new document.xml: prefix + filled cover + body + suffix
    #
    # Inject an explicit section break at the end of the cover content so
    # the cover sits in its own section. The cover section gets titlePg
    # + first-page footer reference; this is structurally identical to
    # what the gold-standard template does (4 sections, titlePg on the
    # cover one). LibreOffice/Word reliably honours type=first footer
    # only when the cover is in a section that ends BEFORE the body
    # begins — single-section docs with titlePg don't always trigger the
    # type=first footer in LibreOffice's renderer.
    cover_sect_xml = (
        '<w:p><w:pPr><w:sectPr>'
        f'<w:headerReference w:type="default" r:id="{chrome_rids["header"]}"/>'
        f'<w:footerReference w:type="default" r:id="{chrome_rids["footer1"]}"/>'
        f'<w:footerReference w:type="first" r:id="{chrome_rids["footer2"]}"/>'
    )
    if tpl_pgSz_xml:
        cover_sect_xml += tpl_pgSz_xml
    if tpl_pgMar_xml:
        cover_sect_xml += tpl_pgMar_xml
    cover_sect_xml += "<w:titlePg/></w:sectPr></w:pPr></w:p>"

    new_doc = prefix + tpl_cover_filled + cover_sect_xml + src_body_updated + suffix
    src_files["word/document.xml"] = new_doc.encode("utf-8")
    log["actions"].append(
        f"Spliced new cover ({len(tpl_cover_filled)}b) + cover-section sectPr into document.xml"
    )

    # -- Stage 3: write final zip ------------------------------------------
    with zipfile.ZipFile(dst_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in src_files.items():
            zout.writestr(name, data)

    # -- Stage 4: run the standard find-replace pass on the result for
    # any BSI residue in the body (text, colours, in-body logos).
    #
    # We tell the body-cleanup pass to skip the chrome media we just
    # placed (header/cover landscape logo + monogram). Average-hash
    # perceptual matching can wrongly identify the green landscape
    # wordmark as the BSI logo (both are wide horizontal text logos —
    # ahash distance comes in around 13, just under the threshold of
    # 14) and swap it back to the lime monogram. The skip_paths set
    # protects everything we transplanted.
    if spec is not None and asset_dir is not None:
        log["actions"].append("Running standard rebrand_docx pass for body cleanup")
        # Build the skip set: every media file we migrated as part of
        # chrome transplant (header logo + cover monogram + footer media),
        # plus the structural XML parts we transplanted from the template
        # (styles.xml, theme1.xml, fontTable.xml). The cleanup pass would
        # otherwise rewrite colours in these — particularly painful for
        # styles.xml because the gold-standard Heading 1 orange (#FF9655)
        # is also in the BSI legacy palette and gets replaced with the
        # new aEX lime, defeating the purpose of transplanting the
        # template's heading styling.
        skip_paths: set[str] = set()
        skip_paths.update(chrome_media_remap.values())  # header/footer chrome
        skip_paths.update(cover_image_paths)            # cover-page images
        skip_paths.update({
            "word/styles.xml",
            "word/theme/theme1.xml",
            "word/fontTable.xml",
        })
        log["body_cleanup_skip_paths"] = sorted(skip_paths)
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            tmp_path = Path(tmp.name)
        shutil.copy2(dst_path, tmp_path)
        cleanup_log = rebrand_docx(tmp_path, dst_path, spec, asset_dir, skip_paths=skip_paths)
        tmp_path.unlink(missing_ok=True)
        log["body_cleanup"] = cleanup_log

    return log
