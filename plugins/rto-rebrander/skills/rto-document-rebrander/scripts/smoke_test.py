"""
smoke_test.py — verify the rebrander is installed correctly.

Run this once after copying the skill folder into your Claude skills dir
and running `pip install -r requirements.txt`. It checks:

    1. Python version is 3.10+
    2. PyYAML, Pillow, imagehash imports succeed
    3. All bundled assets exist and are non-empty
    4. brand_spec.yaml parses and has the expected top-level keys
    5. The rebrand engine modules import without errors
    6. (Best-effort) LibreOffice is detectable on PATH

Exits 0 on success with `OK — skill is ready.`, exits 1 on the first
failure with a clear message about what to fix.
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parent.parent
ASSET_DIR = SKILL_ROOT / "assets"
SCRIPTS_DIR = SKILL_ROOT / "scripts"


def fail(msg: str) -> None:
    print(f"FAIL: {msg}", file=sys.stderr)
    print()
    print("See README.md → Troubleshooting for fixes.")
    sys.exit(1)


def ok(msg: str) -> None:
    print(f"  ok  {msg}")


# 1. Python version
print("Python version...")
if sys.version_info < (3, 10):
    fail(f"Python 3.10+ required; you have {sys.version_info.major}.{sys.version_info.minor}")
ok(f"Python {sys.version_info.major}.{sys.version_info.minor} (>= 3.10)")

# 2. Required imports
print("Python dependencies...")
try:
    import yaml  # noqa: F401
    ok("PyYAML")
except ImportError:
    fail("PyYAML not installed. Run: pip install -r requirements.txt")

try:
    from PIL import Image  # noqa: F401
    ok("Pillow")
except ImportError:
    fail("Pillow not installed. Run: pip install -r requirements.txt")

try:
    import imagehash  # noqa: F401
    ok("imagehash")
except ImportError:
    fail("imagehash not installed. Run: pip install -r requirements.txt")

# 3. Bundled assets
print("Bundled assets...")
required_assets = [
    "brand_spec.yaml",
    "aex_theme1.xml",
    "aex_target_template.docx",
    "aex_logo.png",
    "aex_logo.svg",
    "aex_logo.emf",
    "old_bsi_logo.png",
]
for name in required_assets:
    p = ASSET_DIR / name
    if not p.exists():
        fail(f"Missing asset: {p}")
    if p.stat().st_size == 0:
        fail(f"Empty asset: {p}")
    ok(name)

# 4. brand_spec.yaml parses
print("brand_spec.yaml...")
try:
    with open(ASSET_DIR / "brand_spec.yaml", encoding="utf-8") as f:
        spec = yaml.safe_load(f)
except Exception as e:  # noqa: BLE001
    fail(f"brand_spec.yaml failed to parse: {e}")
for key in ("old_brand", "new_brand", "preserve", "options"):
    if key not in spec:
        fail(f"brand_spec.yaml is missing top-level key: {key}")
    ok(f"key: {key}")

# 5. Engine modules import
print("Engine modules...")
sys.path.insert(0, str(SCRIPTS_DIR))
for module in (
    "rebrand_docx",
    "rebrand_docx_template",
    "rebrand_pptx",
    "rebrand_pdf",
    "rebrand",
    "inspect_doc",
):
    try:
        __import__(module)
        ok(module)
    except Exception as e:  # noqa: BLE001
        fail(f"Failed to import {module}: {e}")

# 6. LibreOffice (optional)
print("LibreOffice (optional)...")
soffice = shutil.which("libreoffice") or shutil.which("soffice")
if soffice:
    ok(f"found at {soffice}")
else:
    print("  --  LibreOffice not on PATH. EMF logo detection and --also-export-pdf")
    print("       will be skipped. Install LibreOffice if you need them; otherwise ignore.")

print()
print("OK - skill is ready.")
