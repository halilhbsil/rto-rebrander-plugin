"""
inspect_doc.py — preview what a rebrand would change in a single document.

Prints:
- What old-brand text occurrences exist
- What old-brand colours appear in styles/document XML
- What images in word/media/ look like logo candidates
- Any version-history rows that would be preserved

Useful as a sanity check before running rebrand.py on a whole folder.

Usage:
    python inspect_doc.py <file.docx|file.pptx>
"""
from __future__ import annotations

import re
import sys
import tempfile
import zipfile
from collections import Counter
from pathlib import Path

import yaml
from PIL import Image
import imagehash

SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))

# Reuse the EMF rendering helper from rebrand_docx
from rebrand_docx import _ahash_emf  # noqa: E402


def _ahash(data: bytes, ext: str):
    try:
        with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
            tmp.write(data)
            tmp_path = Path(tmp.name)
        with Image.open(tmp_path) as im:
            im = im.convert("RGBA")
            bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
            bg.paste(im, mask=im.split()[-1])
            return imagehash.average_hash(bg.convert("RGB"))
    except Exception:
        return None


def inspect(path: Path) -> None:
    asset_dir = SCRIPT_DIR.parent / "assets"
    spec = yaml.safe_load((asset_dir / "brand_spec.yaml").read_text())

    old_name = spec["old_brand"]["name"]
    variants = spec["old_brand"].get("name_variants", [])
    old_colours = set(c.upper() for c in spec["old_brand"].get("colours", []))
    old_footer_pat = spec["old_brand"].get("footer_pattern")

    ref_logo = asset_dir / spec["new_brand"]["old_logo_reference"]
    ref_hash = None
    if ref_logo.exists():
        with open(ref_logo, "rb") as f:
            ref_hash = _ahash(f.read(), ref_logo.suffix)

    print(f"Inspecting: {path}")
    print(f"Old brand: {old_name}")
    print(f"Variants: {variants}")
    print()

    is_pptx = path.suffix.lower() == ".pptx"
    text_node = "a:t" if is_pptx else "w:t"

    text_hits: dict[str, int] = Counter()
    colour_hits: dict[str, int] = Counter()
    footer_hits = 0
    logo_candidates: list[tuple[str, int | None]] = []

    with zipfile.ZipFile(path) as zf:
        for info in zf.infolist():
            data = zf.read(info.filename)

            # Image inspection
            if info.filename.startswith(("word/media/", "ppt/media/")):
                ext = Path(info.filename).suffix.lower()
                if ext in {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif"} and ref_hash:
                    h = _ahash(data, ext)
                    if h is not None:
                        dist = h - ref_hash
                        if dist <= spec["options"]["logo_match_threshold"] + 8:
                            logo_candidates.append((info.filename, int(dist)))
                elif ext == ".svg":
                    txt = data.decode("utf-8", errors="ignore")
                    if re.search(r"\bbsi\b", txt, flags=re.IGNORECASE):
                        logo_candidates.append((info.filename, None))
                elif (
                    ext == ".emf"
                    and ref_hash
                    and spec["options"].get("swap_emf_logos", True)
                    and len(data) <= spec["options"].get("emf_size_limit_bytes", 524288)
                ):
                    h = _ahash_emf(data)
                    if h is not None:
                        dist = h - ref_hash
                        if dist <= spec["options"]["logo_match_threshold"] + 8:
                            logo_candidates.append((info.filename, int(dist)))
                continue

            # XML inspection
            if not info.filename.endswith(".xml"):
                continue
            try:
                xml = data.decode("utf-8")
            except UnicodeDecodeError:
                continue

            # Text occurrences
            text_bodies = re.findall(
                rf"<{text_node}(?:\s[^>]*)?>(.*?)</{text_node}>",
                xml,
                flags=re.DOTALL,
            )
            joined = " ".join(text_bodies)
            for v in [old_name, *variants]:
                n = joined.count(v)
                if n:
                    text_hits[v] += n
            if old_footer_pat:
                footer_hits += len(re.findall(old_footer_pat, joined))

            # Hex colours
            for m in re.finditer(
                r'(?:w:val|w:fill|w:color|val|fill|color)\s*=\s*"([A-Fa-f0-9]{6})"',
                xml,
            ):
                hex_ = m.group(1).upper()
                if hex_ in old_colours:
                    colour_hits[hex_] += 1

    print(f"== Text occurrences ==")
    for k, v in text_hits.most_common():
        print(f"  {k!r}: {v}")
    if not text_hits:
        print("  (none)")
    print()
    print(f"== Footer pattern matches ==")
    print(f"  {footer_hits}")
    print()
    print(f"== Old brand colours found ==")
    for k, v in colour_hits.most_common():
        print(f"  #{k}: {v} occurrences")
    if not colour_hits:
        print("  (none)")
    print()
    print(f"== Logo image candidates ==")
    for fn, dist in logo_candidates:
        if dist is None:
            print(f"  {fn} (SVG matched on 'bsi' keyword)")
        else:
            print(f"  {fn} (perceptual distance: {dist})")
    if not logo_candidates:
        print("  (none)")


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: python inspect_doc.py <file.docx|file.pptx>")
        return 1
    inspect(Path(sys.argv[1]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
