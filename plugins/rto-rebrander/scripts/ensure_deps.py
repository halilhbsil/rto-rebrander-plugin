"""Auto-install the rebrander's Python dependencies on first session.

Runs as a SessionStart hook. Fast path (deps already present) finishes in
~80ms. Slow path runs `pip install -r requirements.txt --user` once, then
all future sessions hit the fast path.

Output is brief: silent on the fast path, progress lines on first install.
Non-fatal on failure - prints a friendly message and the rebrander still
loads (the find-and-replace + template-transplant strategies don't need
external binaries to work).
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

REQUIRED = ("yaml", "PIL", "imagehash")


def fast_check() -> bool:
    """Cheap import probe - returns True if every required package is importable."""
    for mod in REQUIRED:
        try:
            __import__(mod)
        except ImportError:
            return False
    return True


def find_requirements_txt() -> Path | None:
    """Locate the bundled skill's requirements.txt relative to this script."""
    here = Path(__file__).resolve().parent.parent  # plugins/rto-rebrander/
    req = here / "skills" / "rto-document-rebrander" / "requirements.txt"
    return req if req.exists() else None


def pip_install(requirements: Path) -> tuple[bool, str]:
    """Run pip install --user and return (success, stdout+stderr)."""
    try:
        proc = subprocess.run(
            [sys.executable, "-m", "pip", "install", "--user", "-r", str(requirements)],
            capture_output=True,
            text=True,
            timeout=300,
        )
        return proc.returncode == 0, (proc.stdout + proc.stderr)
    except FileNotFoundError:
        return False, "python executable not found on PATH"
    except subprocess.TimeoutExpired:
        return False, "pip install timed out after 5 minutes"
    except Exception as exc:
        return False, f"unexpected error: {exc!r}"


def check_optional_word() -> str | None:
    """Best-effort check that Microsoft Word is installed (for docx2pdf)."""
    if os.name != "nt":
        return None  # Word check is Windows-only
    candidates = [
        Path(os.environ.get("ProgramFiles", "C:/Program Files")) / "Microsoft Office",
        Path(os.environ.get("ProgramFiles(x86)", "C:/Program Files (x86)")) / "Microsoft Office",
    ]
    for base in candidates:
        if base.exists() and any(base.rglob("WINWORD.EXE")):
            return None  # Found Word
    return (
        "Microsoft Word not detected. The rebrander still works, but the "
        "--also-export-pdf flag will be unavailable until Word is installed."
    )


def main() -> int:
    if fast_check():
        return 0

    print("[rto-rebrander] First-run setup: installing Python dependencies...", flush=True)
    req = find_requirements_txt()
    if req is None:
        print("[rto-rebrander] Warning: could not locate requirements.txt; skipping auto-install.", flush=True)
        return 0

    ok, log = pip_install(req)
    if ok:
        print("[rto-rebrander] Dependencies installed successfully.", flush=True)
    else:
        # Non-fatal: print guidance and let the user resolve manually.
        print("[rto-rebrander] Auto-install of dependencies failed.", flush=True)
        print(f"[rto-rebrander] Run manually: python -m pip install -r \"{req}\"", flush=True)
        print(f"[rto-rebrander] pip output (last 500 chars): ...{log[-500:]}", flush=True)

    word_warning = check_optional_word()
    if word_warning:
        print(f"[rto-rebrander] Note: {word_warning}", flush=True)

    return 0


if __name__ == "__main__":
    sys.exit(main())
