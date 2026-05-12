"""
rebrand_pptx.py — rebrand a .pptx file.

PowerPoint's structure is similar to Word's: a zip of XML parts. We swap:

- ppt/theme/theme1.xml (and other themes) → aex_theme1.xml
- ppt/media/* logo images → new logo (by perceptual hash match)
- Text content in ppt/slides/slideN.xml, ppt/slideLayouts/, ppt/slideMasters/,
  ppt/notesSlides/, ppt/handoutMasters/ → org name + footer replacements
- Hardcoded brand colours throughout

Reuses helpers from rebrand_docx where possible.
"""
from __future__ import annotations

import re
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from PIL import Image
import imagehash

from rebrand_docx import (
    _ahash,
    _ahash_emf,
    _protect_preserved_spans,
    _restore_preserved_spans,
    _replace_colours,
)


def _replace_text_in_pptx_runs(
    xml: str,
    spec: dict[str, Any],
) -> tuple[str, dict[str, int]]:
    """Replace text in <a:t> nodes (PowerPoint's text-run nodes)."""
    counts = {"text_replacements": 0, "footer_replacements": 0}
    options = spec.get("options", {})

    if not options.get("swap_text", True):
        return xml, counts

    old_name = spec["old_brand"]["name"]
    new_name = spec["new_brand"]["name"]
    variants = spec["old_brand"].get("name_variants", [])
    preserve_patterns = spec.get("preserve", {}).get("text_patterns", [])
    old_footer_pat = spec["old_brand"].get("footer_pattern")
    new_footer = spec["new_brand"]["footer"]

    def _sub_in_t(m: re.Match) -> str:
        prefix, body, suffix = m.group(1), m.group(2), m.group(3)
        protected, preserved = _protect_preserved_spans(body, preserve_patterns)

        # Footer replacement first (more specific than name)
        if old_footer_pat and re.search(old_footer_pat, protected):
            n = len(re.findall(old_footer_pat, protected))
            protected = re.sub(old_footer_pat, new_footer, protected)
            counts["footer_replacements"] += n

        # Then organisation name (variants first, then base name)
        for variant in [*variants, old_name]:
            if variant in protected:
                n = protected.count(variant)
                protected = protected.replace(variant, new_name)
                counts["text_replacements"] += n

        return f"{prefix}{_restore_preserved_spans(protected, preserved)}{suffix}"

    xml = re.sub(
        r"(<a:t(?:\s[^>]*)?>)(.*?)(</a:t>)",
        _sub_in_t,
        xml,
        flags=re.DOTALL,
    )
    return xml, counts


def _process_pptx_xml_part(
    xml: str,
    spec: dict[str, Any],
) -> tuple[str, dict[str, int]]:
    counts = {"text_replacements": 0, "colour_replacements": 0, "footer_replacements": 0}
    options = spec.get("options", {})

    xml, text_counts = _replace_text_in_pptx_runs(xml, spec)
    counts["text_replacements"] = text_counts["text_replacements"]
    counts["footer_replacements"] = text_counts["footer_replacements"]

    if options.get("swap_colours", True):
        xml, n = _replace_colours(xml, spec["new_brand"]["colour_map"])
        counts["colour_replacements"] = n

    return xml, counts


def rebrand_pptx(
    src_path: Path,
    dst_path: Path,
    spec: dict[str, Any],
    asset_dir: Path,
) -> dict[str, Any]:
    """Rebrand one .pptx file. Returns a change-log dict."""
    src_path = Path(src_path)
    dst_path = Path(dst_path)
    asset_dir = Path(asset_dir)
    dst_path.parent.mkdir(parents=True, exist_ok=True)

    options = spec.get("options", {})
    threshold = options.get("logo_match_threshold", 14)
    swap_logos = options.get("swap_logo_images", True)
    swap_theme = options.get("swap_theme_xml", True)
    swap_emf_logos = options.get("swap_emf_logos", True)
    emf_size_limit = options.get("emf_size_limit_bytes", 524288)

    old_logo_path = asset_dir / spec["new_brand"]["old_logo_reference"]
    ref_hash = _ahash(old_logo_path) if old_logo_path.exists() else None

    new_logo_png = (asset_dir / spec["new_brand"]["logos"]["full_colour_png"]).read_bytes()
    new_logo_svg = (asset_dir / spec["new_brand"]["logos"]["full_colour_svg"]).read_bytes()
    new_logo_emf_path = asset_dir / spec["new_brand"]["logos"].get(
        "full_colour_emf", "aex_logo.emf"
    )
    new_logo_emf = (
        new_logo_emf_path.read_bytes() if new_logo_emf_path.exists() else None
    )
    new_theme_xml = (asset_dir / spec["new_brand"]["theme_xml"]).read_bytes()

    final_log: dict[str, Any] = {
        "source": str(src_path),
        "destination": str(dst_path),
        "images_swapped": [],
        "text_replacements": 0,
        "colour_replacements": 0,
        "footer_replacements": 0,
        "parts_modified": [],
    }

    with zipfile.ZipFile(src_path, "r") as zin:
        members = zin.namelist()
        contents: dict[str, bytes] = {n: zin.read(n) for n in members}

    # Pass 1: theme + logo image swaps
    for name in members:
        if swap_theme and name.startswith("ppt/theme/") and name.endswith(".xml"):
            contents[name] = new_theme_xml
            final_log["images_swapped"].append(name)
        elif swap_logos and ref_hash and name.startswith("ppt/media/"):
            ext = Path(name).suffix.lower()
            if ext in {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif"}:
                try:
                    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
                        tmp.write(contents[name])
                        tmp_path = Path(tmp.name)
                    h = _ahash(tmp_path)
                    tmp_path.unlink(missing_ok=True)
                    if h is not None and (h - ref_hash) <= threshold:
                        contents[name] = new_logo_png
                        final_log["images_swapped"].append(name)
                except Exception:
                    pass
            elif ext == ".svg":
                svg_text = contents[name].decode("utf-8", errors="ignore")
                if re.search(r"\bbsi\b", svg_text, flags=re.IGNORECASE):
                    contents[name] = new_logo_svg
                    final_log["images_swapped"].append(name)
            elif ext == ".emf" and swap_emf_logos and new_logo_emf:
                if len(contents[name]) <= emf_size_limit:
                    h = _ahash_emf(contents[name])
                    if h is not None and (h - ref_hash) <= threshold:
                        contents[name] = new_logo_emf
                        final_log["images_swapped"].append(name)

    # Pass 2: text + colour replacements in slides/layouts/masters/notes
    xml_targets = [
        n
        for n in members
        if n.endswith(".xml")
        and (
            n.startswith("ppt/slides/")
            or n.startswith("ppt/slideLayouts/")
            or n.startswith("ppt/slideMasters/")
            or n.startswith("ppt/notesSlides/")
            or n.startswith("ppt/notesMasters/")
            or n.startswith("ppt/handoutMasters/")
        )
    ]

    for name in xml_targets:
        xml = contents[name].decode("utf-8")
        new_xml, counts = _process_pptx_xml_part(xml, spec)
        if new_xml != xml:
            contents[name] = new_xml.encode("utf-8")
            final_log["parts_modified"].append(name)
        final_log["text_replacements"] += counts["text_replacements"]
        final_log["colour_replacements"] += counts["colour_replacements"]
        final_log["footer_replacements"] += counts["footer_replacements"]

    with zipfile.ZipFile(dst_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in members:
            zout.writestr(name, contents[name])

    return final_log
