"""
rebrand_docx.py — rebrand a .docx file in-place from old branding to new.

Operates on the zipped XML structure directly for maximum fidelity, because
python-docx doesn't expose theme XML, embedded fonts, or some style/colour
nodes. Steps performed (each is independently togglable via the brand spec):

1. Replace word/theme/theme1.xml with the new theme (aEX-correct palette).
2. Scan word/media/ and replace any image that visually matches the old
   logo (by perceptual hash) with the new logo.
3. Replace hardcoded old brand hex colours in styles.xml, document.xml,
   header*.xml, footer*.xml.
4. Replace old organisation name with new in body/headers/footers, with
   safety carveouts (unit codes, ASQA refs, version-history rows etc.).
5. Replace old footer string with new footer string.

Returns a structured change-log dict so callers can summarise what happened.
"""
from __future__ import annotations

import io
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from PIL import Image
import imagehash


SENTINEL_OPEN = "\x00\x01PRESERVE\x00\x01"
SENTINEL_CLOSE = "\x00\x02PRESERVE\x00\x02"


# ---------------------------------------------------------------------------
# Two-font brand policy.
#
# Every <w:rFonts> the rebrander emits (in styles.xml AND in body content)
# resolves to one of these two fonts. Theme-font attributes (w:asciiTheme
# etc.) are stripped during rewrite so Word can't substitute a theme font
# over our explicit choice. See references/safety-carveouts.md for why this
# is absolute (e.g. a stray Calibri inside a list paragraph defeats the
# brand even after styles.xml is clean).
# ---------------------------------------------------------------------------
DISPLAY_FONT = "Obviously Narw Semi"
BODY_FONT = "Edu Favorit Light"

# Style-id substrings that identify heading/title/subtitle paragraphs.
_DISPLAY_STYLE_RE = re.compile(r"heading|title|subtitle", re.IGNORECASE)


def _rebuild_rfonts(font_name: str) -> str:
    """Emit a clean <w:rFonts/> element pinning all four script slots to
    one typeface and dropping any theme-font attributes."""
    return (
        f'<w:rFonts w:ascii="{font_name}" w:hAnsi="{font_name}" '
        f'w:eastAsia="{font_name}" w:cs="{font_name}"/>'
    )


def _normalise_styles_to_two_fonts(styles_xml: str) -> tuple[str, dict[str, int]]:
    """Rewrite every <w:rFonts> in styles.xml so only DISPLAY_FONT or
    BODY_FONT remain.

    Classification per <w:style>:
      * styleId case-insensitively contains 'heading'/'title'/'subtitle' -> display
      * style body already references DISPLAY_FONT -> display
        (catches VersionStyleChar and similar manually-styled display runs)
      * otherwise -> body
    <w:rPrDefault> (docDefaults) -> body.
    """
    stats = {"display": 0, "body": 0, "rfonts_rewritten": 0}

    def rewrite_block(block: str, font_name: str) -> str:
        def sub_one(_m: "re.Match[str]") -> str:
            stats["rfonts_rewritten"] += 1
            return _rebuild_rfonts(font_name)
        return re.sub(r"<w:rFonts\b[^/]*/>", sub_one, block)

    def sub_default(m: "re.Match[str]") -> str:
        return rewrite_block(m.group(0), BODY_FONT)
    styles_xml = re.sub(
        r"<w:rPrDefault>.*?</w:rPrDefault>",
        sub_default,
        styles_xml,
        flags=re.DOTALL,
    )

    def sub_style(m: "re.Match[str]") -> str:
        full = m.group(0)
        sid_match = re.search(r'\bw:styleId="([^"]+)"', full)
        sid = sid_match.group(1) if sid_match else ""
        is_display = bool(_DISPLAY_STYLE_RE.search(sid)) or (DISPLAY_FONT in full)
        if is_display:
            stats["display"] += 1
            return rewrite_block(full, DISPLAY_FONT)
        stats["body"] += 1
        return rewrite_block(full, BODY_FONT)

    styles_xml = re.sub(
        r"<w:style\b.*?</w:style>",
        sub_style,
        styles_xml,
        flags=re.DOTALL,
    )

    return styles_xml, stats


def _normalise_runs_to_two_fonts(xml: str) -> tuple[str, dict[str, int]]:
    """Rewrite every <w:rFonts> in a body-content XML part (document.xml,
    header*.xml, footer*.xml) so only DISPLAY_FONT or BODY_FONT remain.

    Word's font-resolution order is: direct run formatting > linked
    character style > paragraph style > docDefaults > theme. So even
    with a perfect styles.xml + theme, a single direct <w:rFonts
    w:ascii="Calibri"/> on a run produces Calibri on screen. This pass
    scrubs those.

    Classification is paragraph-scoped: any <w:rFonts> inside a <w:p>
    whose <w:pStyle w:val="..."/> matches Heading/Title/Subtitle gets
    DISPLAY_FONT; otherwise BODY_FONT. Implemented as a single-pass
    token scanner so it handles paragraphs nested inside tables and
    textboxes correctly (a naive per-<w:p> regex would mismatch nested
    structures).

    DOES NOT touch numbering.xml — bullet glyphs there use Symbol /
    Wingdings / Courier New legitimately; rewriting them to text fonts
    produces letter bullets.
    """
    stats = {"display": 0, "body": 0, "rfonts_rewritten": 0}

    token = re.compile(
        r"(<w:p(?:\s[^>]*)?>)"                 # 1: paragraph open
        r"|(<w:pStyle\s+w:val=\"([^\"]+)\"\s*/>)"  # 2 full, 3 val
        r"|(<w:rFonts\b[^/]*/>)"               # 4: rFonts self-closing
    )

    parts: list[str] = []
    last = 0
    current_is_heading = False
    for m in token.finditer(xml):
        parts.append(xml[last:m.start()])
        last = m.end()
        if m.group(1):  # <w:p open
            current_is_heading = False
            parts.append(m.group(1))
        elif m.group(2):  # <w:pStyle
            current_is_heading = bool(_DISPLAY_STYLE_RE.search(m.group(3)))
            parts.append(m.group(2))
        else:  # <w:rFonts
            font = DISPLAY_FONT if current_is_heading else BODY_FONT
            if current_is_heading:
                stats["display"] += 1
            else:
                stats["body"] += 1
            stats["rfonts_rewritten"] += 1
            parts.append(_rebuild_rfonts(font))
    parts.append(xml[last:])
    return "".join(parts), stats


# Fonts in numbering.xml that draw the bullet glyph itself rather than
# laying out a digit/letter prefix in a text font. Rewriting these to a
# brand text font produces letter bullets instead of the intended ● ▪ ○
# glyphs.
_NUMBERING_GLYPH_FONTS = frozenset({
    "Symbol",
    "Wingdings",
    "Wingdings 2",
    "Wingdings 3",
    "Webdings",
    "Courier New",
})


def _normalise_numbering_to_two_fonts(numbering_xml: str) -> tuple[str, dict[str, int]]:
    """Rewrite <w:rFonts> in numbering.xml, leaving bullet-glyph fonts alone.

    Number-prefix runs ("1.", "a)", "i.") use a text font and should be
    BODY_FONT to match the surrounding text. Bullet-glyph runs (●, ▪, ○)
    use Symbol/Wingdings/etc. and must stay — substituting them produces
    letter bullets.

    Per-slot rewrite: each of the four font slots (w:ascii, w:hAnsi,
    w:eastAsia, w:cs) is examined independently. If that slot names a
    glyph font (_NUMBERING_GLYPH_FONTS), it stays; otherwise it is
    rewritten to BODY_FONT. This handles the mixed-slot edge case
    (e.g. ascii=Symbol, eastAsia=Times New Roman, cs=Arial) where the
    bullet glyph slot must be preserved but the fallback slots should
    NOT silently carry text-font names through.

    The w:hint attribute (and any other non-font attrs) is preserved.
    """
    stats = {"rewritten": 0, "preserved_glyph": 0, "slots_rewritten": 0}

    slot_pat = re.compile(r'\bw:(ascii|hAnsi|eastAsia|cs)="([^"]+)"')

    def sub_one(m: "re.Match[str]") -> str:
        block = m.group(0)
        slots = slot_pat.findall(block)
        if not slots:
            # No font slot at all (e.g. <w:rFonts w:hint="default"/>) —
            # nothing to decide; leave untouched.
            return block

        new_block = block
        any_rewritten = False
        any_preserved = False
        for slot_name, slot_value in slots:
            if slot_value in _NUMBERING_GLYPH_FONTS:
                any_preserved = True
                continue
            # Replace just this slot's value with BODY_FONT
            new_block = re.sub(
                r'\bw:' + slot_name + r'="' + re.escape(slot_value) + r'"',
                f'w:{slot_name}="{BODY_FONT}"',
                new_block,
                count=1,
            )
            stats["slots_rewritten"] += 1
            any_rewritten = True

        if any_rewritten:
            stats["rewritten"] += 1
        if any_preserved and not any_rewritten:
            stats["preserved_glyph"] += 1
        return new_block

    new_xml = re.sub(r"<w:rFonts\b[^/]*/>", sub_one, numbering_xml)
    return new_xml, stats


def _strip_secondary_cover_logos(
    contents: dict[str, bytes],
) -> dict[str, Any]:
    """Strip anchored brand-logo drawings from "first-page" footers belonging
    to sections 2+.

    Source BSI docs often have multiple sections, each with titlePg=True
    and a "first" footer containing a cover-style anchored logo. The
    rebrander correctly swaps that logo for the aEX one — but under
    the new aEX brand convention only the actual cover (section 1's
    first page) should display that footer logo. On documents with two
    or more sections this produces stray logos on the first physical
    page of each downstream section (typically page 2+).

    Strategy: parse document.xml's <w:sectPr> blocks in order, collect
    the relId of each section's "first" footerReference, treat the
    first one as the cover (keep it), and from every other "first"
    footer remove all <w:drawing>...</w:drawing> blocks (text content
    survives so the © aEX line still appears).

    Returns a log entry: {"footers_cleaned": [list of part names]}.
    """
    log: dict[str, Any] = {"footers_cleaned": []}

    doc_bytes = contents.get("word/document.xml")
    rels_bytes = contents.get("word/_rels/document.xml.rels")
    if not doc_bytes or not rels_bytes:
        return log

    doc_xml = doc_bytes.decode("utf-8", errors="replace")
    rels_xml = rels_bytes.decode("utf-8", errors="replace")

    # Collect "first" footer relIds in section order
    first_footer_rids: list[str] = []
    for sect in re.finditer(r"<w:sectPr\b[^>]*>.*?</w:sectPr>", doc_xml, re.DOTALL):
        m = re.search(
            r'<w:footerReference\s+w:type="first"\s+r:id="(rId\d+)"',
            sect.group(0),
        )
        if m:
            first_footer_rids.append(m.group(1))

    if len(first_footer_rids) < 2:
        return log  # Only the cover (or no first footer at all) — nothing to do

    # Map rId -> footer part path
    rid_to_target: dict[str, str] = {}
    for m in re.finditer(
        r'<Relationship\b[^>]*?Id="(rId\d+)"[^>]*?Target="(footer\d+\.xml)"',
        rels_xml,
    ):
        rid_to_target[m.group(1)] = "word/" + m.group(2)

    secondary = first_footer_rids[1:]  # all but the cover
    for rid in secondary:
        part = rid_to_target.get(rid)
        if not part or part not in contents:
            continue
        body = contents[part].decode("utf-8", errors="replace")
        if "<w:drawing>" not in body:
            continue
        new_body = re.sub(
            r"<w:drawing>.*?</w:drawing>",
            "",
            body,
            flags=re.DOTALL,
        )
        # If the surviving body still contains an <w:r> that now has an
        # empty <w:rPr> wrapped around nothing, leave it — Word ignores
        # empty runs. Cleaner output isn't worth a second pass here.
        if new_body != body:
            contents[part] = new_body.encode("utf-8")
            log["footers_cleaned"].append(part)

    return log


def _read_zip_member(zf: zipfile.ZipFile, name: str) -> str:
    return zf.read(name).decode("utf-8")


def _ahash_emf(emf_bytes: bytes) -> imagehash.ImageHash | None:
    """Render an EMF to PNG via LibreOffice, then average-hash it.

    Older Word docs (pre-2016 or so) embed logos as EMF (Windows Enhanced
    Metafile) — a vector format Pillow can't read directly. We shell out
    to LibreOffice headless for the conversion. Slow (~1-2s per call) but
    reliable. Returns None if LibreOffice is unavailable or conversion fails.

    LibreOffice tends to render EMFs onto a page-sized canvas with white
    padding around the actual graphic. To get a hash that matches a
    tightly-cropped reference logo, we auto-crop white margins before
    hashing.
    """
    soffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not soffice:
        return None
    try:
        with tempfile.TemporaryDirectory() as td:
            td_path = Path(td)
            emf_path = td_path / "in.emf"
            emf_path.write_bytes(emf_bytes)
            subprocess.run(
                [
                    soffice,
                    "--headless",
                    "--convert-to",
                    "png",
                    "--outdir",
                    str(td_path),
                    str(emf_path),
                ],
                check=True,
                capture_output=True,
                timeout=30,
            )
            png_path = td_path / "in.png"
            if not png_path.exists():
                return None
            # Open, flatten alpha on white, auto-crop white margins, hash.
            with Image.open(png_path) as im:
                im = im.convert("RGBA")
                bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
                bg.paste(im, mask=im.split()[-1])
                rgb = bg.convert("RGB")
                # Auto-crop: invert the image so non-content (white) becomes
                # black, getbbox() returns the bbox of non-black pixels.
                from PIL import ImageChops, ImageOps
                inverted = ImageOps.invert(rgb)
                bbox = inverted.getbbox()
                if bbox:
                    rgb = rgb.crop(bbox)
                # Skip ahash if the cropped image is unreasonably small
                # (likely just noise rather than a real logo render).
                if rgb.size[0] < 16 or rgb.size[1] < 16:
                    return None
                return imagehash.average_hash(rgb)
    except Exception:
        return None


def _ahash(path: Path) -> imagehash.ImageHash | None:
    """Average-hash an image; returns None if not openable."""
    try:
        with Image.open(path) as im:
            im = im.convert("RGBA")
            # Flatten transparency on white so hashes of transparent and
            # white-background versions of the same logo match.
            bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
            bg.paste(im, mask=im.split()[-1])
            return imagehash.average_hash(bg.convert("RGB"))
    except Exception:
        return None


def _protect_preserved_spans(text: str, patterns: list[str]) -> tuple[str, list[str]]:
    """Wrap regex matches in sentinels so subsequent replacements skip them.

    Returns the protected text plus the list of preserved spans (in order).
    """
    preserved: list[str] = []

    def _stash(m: re.Match) -> str:
        idx = len(preserved)
        preserved.append(m.group(0))
        return f"{SENTINEL_OPEN}{idx}{SENTINEL_CLOSE}"

    out = text
    for pat in patterns:
        out = re.sub(pat, _stash, out)
    return out, preserved


def _restore_preserved_spans(text: str, preserved: list[str]) -> str:
    def _unstash(m: re.Match) -> str:
        return preserved[int(m.group(1))]

    return re.sub(
        re.escape(SENTINEL_OPEN) + r"(\d+)" + re.escape(SENTINEL_CLOSE),
        _unstash,
        text,
    )


def _replace_text_in_runs(
    xml: str,
    old: str,
    new: str,
    preserve_patterns: list[str] | None = None,
) -> tuple[str, int]:
    """Replace `old` with `new` inside <w:t>...</w:t> nodes only.

    Word splits text across multiple <w:r><w:t> runs even within a phrase
    (because of formatting boundaries, spell-check insertions, etc.). To
    catch runs split across <w:t> nodes within a single paragraph, we do
    two passes:

      Pass 1: simple replacement inside individual <w:t> nodes.
      Pass 2: collapse adjacent <w:t> runs in each <w:p>, run replacement,
              and rewrite. This pass only fires if pass 1 doesn't catch
              everything (we re-search after pass 1).

    `preserve_patterns` (if provided) protects matching spans from
    replacement via sentinel-wrapping, in both passes.
    """
    preserve_patterns = preserve_patterns or []
    count = 0

    # Pass 1: per-<w:t>-node replacement, with carveouts
    def _sub_in_t(m: re.Match) -> str:
        nonlocal count
        prefix, body, suffix = m.group(1), m.group(2), m.group(3)
        protected, preserved = _protect_preserved_spans(body, preserve_patterns)
        if old in protected:
            count += protected.count(old)
            protected = protected.replace(old, new)
        return f"{prefix}{_restore_preserved_spans(protected, preserved)}{suffix}"

    xml = re.sub(
        r"(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)",
        _sub_in_t,
        xml,
        flags=re.DOTALL,
    )

    # Pass 2: split-run handling. Collapse runs within each paragraph.
    def _collapse_in_paragraph(m: re.Match) -> str:
        nonlocal count
        para = m.group(0)
        # Concatenate all <w:t> contents to test if a split match exists.
        concat = "".join(re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", para, flags=re.DOTALL))
        # Apply carveouts to the concatenated text first
        protected_concat, preserved = _protect_preserved_spans(concat, preserve_patterns)
        if old not in protected_concat:
            return para
        # Replace inside the protected (carveout-aware) concatenation
        new_concat_protected = protected_concat.replace(old, new)
        new_concat = _restore_preserved_spans(new_concat_protected, preserved)
        # Count how many actual replacements happened
        count_local = protected_concat.count(old)
        first_replaced = {"done": False}

        def _collapse_t(tm: re.Match) -> str:
            prefix, body, suffix = tm.group(1), tm.group(2), tm.group(3)
            if not first_replaced["done"]:
                first_replaced["done"] = True
                return f"{prefix}{new_concat}{suffix}"
            return f"{prefix}{suffix}"  # empty out remaining runs in this para

        new_para = re.sub(
            r"(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)",
            _collapse_t,
            para,
            flags=re.DOTALL,
        )
        nonlocal_count_inc(count_local)
        return new_para

    def nonlocal_count_inc(n: int) -> None:
        nonlocal count
        count += n

    # Only run pass 2 if pass 1 didn't catch all instances. Detect by
    # concatenating per paragraph and checking against carveouts.
    def _has_split_match(p: str) -> bool:
        concat = "".join(re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", p, flags=re.DOTALL))
        protected_concat, _ = _protect_preserved_spans(concat, preserve_patterns)
        return old in protected_concat

    if any(_has_split_match(p) for p in re.findall(r"<w:p[\s>].*?</w:p>", xml, flags=re.DOTALL)):
        xml = re.sub(
            r"<w:p[\s>].*?</w:p>",
            _collapse_in_paragraph,
            xml,
            flags=re.DOTALL,
        )

    return xml, count


def _replace_colours(xml: str, colour_map: dict[str, str]) -> int:
    """Replace hardcoded hex colours in OOXML attributes. Returns number replaced.

    Handles three attribute styles seen in real Word/Office documents:

    1. **OOXML form**: `w:val="003469"`, `w:fill="003469"`, `w:color="003469"`,
       and the un-prefixed equivalents `val=`, `fill=`, `color=` used in
       DrawingML (`<a:srgbClr val="003469"/>`).

    2. **VML/legacy form**: `fillcolor="#003469"`, `strokecolor="#003469"`.
       Note the hash prefix and the concatenated attribute name. These
       appear in older docs (pre-2016) using Microsoft VML for shapes.

    3. **VML theme-reference form**: `fillcolor="#4f81bd [3204]"`. The hex
       is the original theme colour pre-rendered, with a tint/shade index
       in brackets. We replace the hex part if it matches a brand colour
       and preserve the bracketed suffix.
    """
    count = 0

    def _sub_colour(m: re.Match) -> str:
        nonlocal count
        attr_name = m.group(1)       # e.g. "fillcolor", "w:fill", "val"
        eq_quote = m.group(2)        # `="`
        hash_prefix = m.group(3)     # "#" or ""
        old_hex = m.group(4).upper() # 6-char hex
        suffix = m.group(5)          # ` [3204]` or ""
        new_hex = colour_map.get(old_hex)
        if new_hex:
            count += 1
            return f'{attr_name}{eq_quote}{hash_prefix}{new_hex}{suffix}"'
        return m.group(0)

    # Match attribute names ending in val|fill|color (with optional namespace
    # prefix and optional inner words like "fill" in "fillcolor"), with
    # optional "#" on the value, and optional " [NNNN]" theme-tint suffix.
    pattern = (
        r'(\b(?:[\w-]+:)?\w*(?:val|fill|color))'  # attribute name
        r'(\s*=\s*")'                              # equals + opening quote
        r'(#?)'                                    # optional hash prefix
        r'([A-Fa-f0-9]{6})'                        # 6 hex chars
        r'((?:\s*\[\d+\])?)'                       # optional theme suffix
        r'"'                                       # closing quote
    )
    new_xml = re.sub(pattern, _sub_colour, xml)
    return new_xml, count  # type: ignore[return-value]


def _process_xml_part(
    xml: str,
    spec: dict[str, Any],
    is_footer: bool = False,
) -> tuple[str, dict[str, int]]:
    """Run text + colour replacements on one OOXML part. Returns updated XML
    and a per-action change count."""
    counts = {"text_replacements": 0, "colour_replacements": 0, "footer_replacements": 0}
    options = spec.get("options", {})

    # 1. Footer string replacement (regex, anchored on old footer pattern).
    if is_footer:
        old_footer_pat = spec["old_brand"].get("footer_pattern")
        new_footer = spec["new_brand"]["footer"]
        if old_footer_pat:
            # We need to do this against the visible text only. Because the
            # footer text is typically split across many <w:t> runs, the
            # safest approach is: extract concatenated text, regex-match,
            # if matched then collapse and rewrite the first <w:p> that
            # contained the match.
            paragraphs = re.findall(r"<w:p[\s>].*?</w:p>", xml, flags=re.DOTALL)
            for para in paragraphs:
                concat = "".join(
                    re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", para, flags=re.DOTALL)
                )
                if re.search(old_footer_pat, concat):
                    new_concat = re.sub(old_footer_pat, new_footer, concat)
                    first = {"done": False}

                    def _collapse(tm: re.Match) -> str:
                        prefix, body, suffix = tm.group(1), tm.group(2), tm.group(3)
                        if not first["done"]:
                            first["done"] = True
                            return f"{prefix}{new_concat}{suffix}"
                        return f"{prefix}{suffix}"

                    new_para = re.sub(
                        r"(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)",
                        _collapse,
                        para,
                        flags=re.DOTALL,
                    )
                    xml = xml.replace(para, new_para)
                    counts["footer_replacements"] += 1

    # 2. Text replacement (organisation name) with safety carveouts.
    if options.get("swap_text", True):
        old_name = spec["old_brand"]["name"]
        new_name = spec["new_brand"]["name"]
        variants = spec["old_brand"].get("name_variants", [])

        # Protect preserved patterns by wrapping them in sentinels first.
        # We can't apply sentinels to the binary XML because they'd break
        # XML structure, so we apply protection per <w:t> body only.
        preserve_patterns = spec.get("preserve", {}).get("text_patterns", [])

        def _protected_replace_in_t(m: re.Match) -> str:
            prefix, body, suffix = m.group(1), m.group(2), m.group(3)
            protected, preserved = _protect_preserved_spans(body, preserve_patterns)
            for variant in [*variants, old_name]:  # variants first (longer matches)
                if variant in protected:
                    n = protected.count(variant)
                    protected = protected.replace(variant, new_name)
                    counts["text_replacements"] += n
            return f"{prefix}{_restore_preserved_spans(protected, preserved)}{suffix}"

        xml = re.sub(
            r"(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)",
            _protected_replace_in_t,
            xml,
            flags=re.DOTALL,
        )

        # Then handle split-run cases via the helper that already handles
        # paragraph-level collapsing. We pass preserve_patterns so the
        # split-run pass also respects the carveouts (otherwise it would
        # rewrite e.g. "BSI Learning Curriculum" inside a version-history
        # row even though the per-node pass above had protected it).
        # Pass 1 of the helper is a no-op here because per-node replacement
        # already happened above; only pass 2 (collapse) does new work.
        #
        # Variants must come BEFORE old_name (longer/more-specific phrases
        # first) — otherwise "BSI Learning" matches inside "BSI Learning
        # Institute Pty Ltd" and gets replaced before the variant can,
        # producing nonsense like "aEX Institute Institute Pty Ltd".
        for variant in [*variants, old_name]:
            xml, n = _replace_text_in_runs(
                xml, variant, new_name, preserve_patterns=preserve_patterns
            )
            counts["text_replacements"] += n

    # 3. Colour replacement.
    if options.get("swap_colours", True):
        new_xml, n = _replace_colours(xml, spec["new_brand"]["colour_map"])
        xml, counts["colour_replacements"] = new_xml, n

    return xml, counts


def _replace_logo_images(
    src_zip: zipfile.ZipFile,
    dst_zip: zipfile.ZipFile,
    spec: dict[str, Any],
    asset_dir: Path,
    skip_paths: set[str] | None = None,
) -> dict[str, Any]:
    """Stream files from src_zip → dst_zip, swapping logo media files.

    Strategy: hash every image in word/media/ and compare to the old logo
    reference hash. If within threshold, swap with the new logo of the
    closest type (PNG↔PNG, SVG↔SVG).

    Files in `skip_paths` are stream-copied unchanged (used by the
    template-transplant flow to protect already-placed brand chrome
    media from being misidentified as old-brand logos).
    """
    skip_paths = skip_paths or set()
    options = spec.get("options", {})
    threshold = options.get("logo_match_threshold", 14)
    swap_logos = options.get("swap_logo_images", True)
    swap_theme = options.get("swap_theme_xml", True)
    swap_emf_logos = options.get("swap_emf_logos", True)
    emf_size_limit = options.get("emf_size_limit_bytes", 524288)

    # Compute reference hash for old logo
    old_logo_path = asset_dir / spec["new_brand"]["old_logo_reference"]
    ref_hash = _ahash(old_logo_path) if old_logo_path.exists() else None

    new_logo_png = asset_dir / spec["new_brand"]["logos"]["full_colour_png"]
    new_logo_svg = asset_dir / spec["new_brand"]["logos"]["full_colour_svg"]
    new_logo_emf_path = asset_dir / spec["new_brand"]["logos"].get(
        "full_colour_emf", "aex_logo.emf"
    )
    new_logo_emf_bytes = (
        new_logo_emf_path.read_bytes() if new_logo_emf_path.exists() else None
    )
    new_monogram_svg = asset_dir / spec["new_brand"]["logos"]["monogram_svg"]
    new_theme_xml = (asset_dir / spec["new_brand"]["theme_xml"]).read_bytes()

    swapped: list[str] = []

    for item in src_zip.infolist():
        data = src_zip.read(item.filename)

        # Files explicitly marked to skip (e.g. transplanted chrome media)
        # are stream-copied without inspection. Theme swap is also skipped
        # for these — caller is in charge of the file.
        if item.filename in skip_paths:
            dst_zip.writestr(item, data)
            continue

        # Theme swap
        if (
            swap_theme
            and item.filename in {"word/theme/theme1.xml", "ppt/theme/theme1.xml"}
        ):
            data = new_theme_xml
            swapped.append(item.filename)

        # Image swap
        elif swap_logos and ref_hash and item.filename.startswith(("word/media/", "ppt/media/")):
            ext = Path(item.filename).suffix.lower()
            if ext in {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif"}:
                # Hash this image
                try:
                    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
                        tmp.write(data)
                        tmp_path = Path(tmp.name)
                    h = _ahash(tmp_path)
                    tmp_path.unlink(missing_ok=True)
                    if h is not None and (h - ref_hash) <= threshold:
                        # Replace with new logo PNG
                        data = new_logo_png.read_bytes()
                        swapped.append(item.filename)
                except Exception:
                    pass
            elif ext == ".svg":
                # SVG comparison is harder — we use a heuristic: if the SVG
                # contains "bsi" or "BSI" in its text/identifiers, swap it.
                svg_text = data.decode("utf-8", errors="ignore")
                if re.search(r"\bbsi\b", svg_text, flags=re.IGNORECASE):
                    data = new_logo_svg.read_bytes()
                    swapped.append(item.filename)
            elif ext == ".emf" and swap_emf_logos and new_logo_emf_bytes:
                # EMF logos (older Word docs). Skip very large EMFs — those
                # are content images, not logos. Render to PNG via
                # LibreOffice, hash, compare.
                if len(data) <= emf_size_limit:
                    h = _ahash_emf(data)
                    if h is not None and (h - ref_hash) <= threshold:
                        # Replace with the bundled aEX EMF (preserves the
                        # original image's content type so Word's _rels stays valid)
                        data = new_logo_emf_bytes
                        swapped.append(item.filename)

        dst_zip.writestr(item, data)

    return {"images_swapped": swapped}


def rebrand_docx(
    src_path: Path,
    dst_path: Path,
    spec: dict[str, Any],
    asset_dir: Path,
    skip_paths: set[str] | None = None,
) -> dict[str, Any]:
    """Rebrand one .docx file. Returns a change-log dict.

    `skip_paths` (optional) is a set of zip-internal paths that should
    NOT be touched by image-swap logic. Use this when the caller has
    already placed brand-correct media at those paths (e.g. the
    template-transplant flow has already migrated the new chrome
    media and the perceptual-hash matching might wrongly swap them
    back). Other passes (theme swap, text/colour replacement) still
    run on those files because they affect XML, not media bytes.
    """
    src_path = Path(src_path)
    dst_path = Path(dst_path)
    asset_dir = Path(asset_dir)
    dst_path.parent.mkdir(parents=True, exist_ok=True)
    skip_paths = skip_paths or set()

    # Stage 1: stream-copy the zip, swapping theme + logo images.
    intermediate = dst_path.with_suffix(dst_path.suffix + ".intermediate")
    with zipfile.ZipFile(src_path, "r") as src_zip:
        with zipfile.ZipFile(intermediate, "w", zipfile.ZIP_DEFLATED) as dst_zip:
            log = _replace_logo_images(src_zip, dst_zip, spec, asset_dir, skip_paths)

    # Stage 2: open intermediate and rewrite XML parts (text + colours).
    final_log: dict[str, Any] = {
        "source": str(src_path),
        "destination": str(dst_path),
        "images_swapped": log["images_swapped"],
        "text_replacements": 0,
        "colour_replacements": 0,
        "footer_replacements": 0,
        "parts_modified": [],
    }

    # Read all members into memory, transform relevant ones, write to final.
    with zipfile.ZipFile(intermediate, "r") as zin:
        members = zin.namelist()
        contents: dict[str, bytes] = {n: zin.read(n) for n in members}

    xml_targets = []
    for name in contents:
        # Files explicitly marked to skip are left untouched. The transplant
        # flow uses this to protect already-rebranded chrome (styles.xml,
        # theme1.xml, fontTable.xml) from having brand colours like the
        # gold-standard's Heading 1 orange (#FF9655) replaced by the
        # colour-map, since FF9655 also appears in the BSI legacy palette.
        if name in skip_paths:
            continue
        if name == "word/document.xml":
            xml_targets.append((name, False))
        elif name.startswith("word/header") and name.endswith(".xml"):
            xml_targets.append((name, False))
        elif name.startswith("word/footer") and name.endswith(".xml"):
            xml_targets.append((name, True))
        elif name == "word/styles.xml":
            xml_targets.append((name, False))
        elif name.startswith("word/footnotes") and name.endswith(".xml"):
            xml_targets.append((name, False))
        elif name.startswith("word/endnotes") and name.endswith(".xml"):
            xml_targets.append((name, False))

    for name, is_footer in xml_targets:
        xml = contents[name].decode("utf-8")
        new_xml, counts = _process_xml_part(xml, spec, is_footer=is_footer)
        if new_xml != xml:
            contents[name] = new_xml.encode("utf-8")
            final_log["parts_modified"].append(name)
        final_log["text_replacements"] += counts["text_replacements"]
        final_log["colour_replacements"] += counts["colour_replacements"]
        final_log["footer_replacements"] += counts["footer_replacements"]

    # Two-font enforcement — runs after text/colour passes so its rewrites
    # are not overwritten. Applies to both rebrand strategies; when called
    # from the transplant flow, styles.xml is usually in skip_paths
    # (already normalised during transplant) and we just normalise body
    # content here.
    font_log = {
        "styles": {"display": 0, "body": 0, "rfonts_rewritten": 0},
        "runs": {"display": 0, "body": 0, "rfonts_rewritten": 0},
    }
    if "word/styles.xml" in contents and "word/styles.xml" not in skip_paths:
        s_xml = contents["word/styles.xml"].decode("utf-8")
        new_s, s_stats = _normalise_styles_to_two_fonts(s_xml)
        if new_s != s_xml:
            contents["word/styles.xml"] = new_s.encode("utf-8")
            if "word/styles.xml" not in final_log["parts_modified"]:
                final_log["parts_modified"].append("word/styles.xml")
        font_log["styles"] = s_stats

    for name in list(contents.keys()):
        if name in skip_paths:
            continue
        if name == "word/document.xml" or (
            (
                name.startswith("word/header")
                or name.startswith("word/footer")
                or name.startswith("word/footnotes")
                or name.startswith("word/endnotes")
                or name.startswith("word/comments")
            )
            and name.endswith(".xml")
        ):
            xml = contents[name].decode("utf-8")
            new_xml, r_stats = _normalise_runs_to_two_fonts(xml)
            if new_xml != xml:
                contents[name] = new_xml.encode("utf-8")
                if name not in final_log["parts_modified"]:
                    final_log["parts_modified"].append(name)
            for k, v in r_stats.items():
                font_log["runs"][k] += v

    # Numbering: handled separately so we can spare bullet-glyph fonts.
    if "word/numbering.xml" in contents and "word/numbering.xml" not in skip_paths:
        n_xml = contents["word/numbering.xml"].decode("utf-8")
        new_n, n_stats = _normalise_numbering_to_two_fonts(n_xml)
        if new_n != n_xml:
            contents["word/numbering.xml"] = new_n.encode("utf-8")
            if "word/numbering.xml" not in final_log["parts_modified"]:
                final_log["parts_modified"].append("word/numbering.xml")
        font_log["numbering"] = n_stats

    final_log["font_normalisation"] = font_log

    # Strip stray cover-style logos from secondary sections' first-page
    # footers (the source-doc artifact where every section's first page
    # carried a logo). Cover footer is preserved.
    cover_log = _strip_secondary_cover_logos(contents)
    if cover_log["footers_cleaned"]:
        final_log["secondary_cover_logos_stripped"] = cover_log["footers_cleaned"]
        for p in cover_log["footers_cleaned"]:
            if p not in final_log["parts_modified"]:
                final_log["parts_modified"].append(p)

    with zipfile.ZipFile(dst_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in members:
            zout.writestr(name, contents[name])

    intermediate.unlink(missing_ok=True)
    return final_log
