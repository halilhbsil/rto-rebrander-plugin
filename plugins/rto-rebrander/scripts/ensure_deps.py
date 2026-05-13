"""Auto-install the rebrander's Python dependencies on first session.

Runs as a SessionStart hook. Fast path (everything already installed)
finishes in ~150ms by querying pip's installed-package database without
importing the actual packages. Slow path runs `pip install -r requirements.txt --user`
once, then all future sessions hit the fast path.

Behavior:
- Python version too old (<3.10): print clear guidance, exit 0 (rebrander
  still loads; user gets a deterministic error if they try to invoke it).
- Any pip dep missing: run pip install, report result.
- Microsoft Word / LibreOffice: detect and warn (don't auto-install; both
  require UAC elevation and would pop a prompt every session).

Always exits 0. The rebrander's find-and-replace and template-transplant
strategies need only stdlib + PyYAML for full operation; auto-install failure
isn't fatal to the rest of the skill loading.
"""
from __future__ import annotations

import importlib.metadata
import os
import re
import subprocess
import sys
from pathlib import Path

MIN_PYTHON = (3, 10)


def check_python_version() -> str | None:
    """Return a warning string if Python is too old, else None."""
    if sys.version_info < MIN_PYTHON:
        have = ".".join(str(p) for p in sys.version_info[:3])
        need = ".".join(str(p) for p in MIN_PYTHON)
        return (
            f"Python {have} detected; rebrander needs {need}+. "
            f"Install a newer Python from https://www.python.org/downloads/"
        )
    return None


def find_requirements_txt() -> Path | None:
    """Locate the bundled skill's requirements.txt relative to this script."""
    here = Path(__file__).resolve().parent.parent  # plugins/rto-rebrander/
    req = here / "skills" / "rto-document-rebrander" / "requirements.txt"
    return req if req.exists() else None


def parse_requirements(req_file: Path) -> list[str]:
    """Extract package names from a pip requirements file.

    Returns the bare package name for each line (drops version specifiers
    and environment markers). Skips comments, blank lines, and entries
    whose environment marker doesn't apply to the current platform.
    """
    names: list[str] = []
    for raw in req_file.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue

        # Split off environment marker (e.g. "docx2pdf>=0.1 ; sys_platform == 'win32'")
        if ";" in line:
            spec, marker = line.split(";", 1)
            if not _marker_applies(marker.strip()):
                continue
        else:
            spec = line

        # Strip version specifier to get just the package name.
        name = re.split(r"[<>=!~\[]", spec, maxsplit=1)[0].strip()
        if name:
            names.append(name)
    return names


def _marker_applies(marker: str) -> bool:
    """Best-effort check of a pip environment marker.

    Only handles the markers we actually use in this project (sys_platform);
    falls back to True for anything else so we don't accidentally skip a
    legitimate dep.
    """
    # Handles: sys_platform == "win32"  /  sys_platform != "linux"  etc.
    m = re.match(r"""sys_platform\s*(==|!=)\s*["']([^"']+)["']""", marker)
    if m:
        op, value = m.group(1), m.group(2)
        return (sys.platform == value) if op == "==" else (sys.platform != value)
    return True  # Unknown marker: don't skip


def fast_check(packages: list[str]) -> list[str]:
    """Return the list of packages that are NOT currently installed."""
    missing: list[str] = []
    for pkg in packages:
        try:
            importlib.metadata.distribution(pkg)
        except importlib.metadata.PackageNotFoundError:
            missing.append(pkg)
    return missing


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


def detect_windows_executable(exe_glob: str, *bases: str) -> bool:
    """Best-effort search for a Windows executable under common install bases."""
    if os.name != "nt":
        return False
    for base in bases:
        root = Path(base) if base else None
        if root and root.exists() and any(root.rglob(exe_glob)):
            return True
    return False


def check_optional_binaries() -> list[str]:
    """Return human-readable warnings for missing optional system binaries."""
    if os.name != "nt":
        return []
    warnings: list[str] = []

    pf = os.environ.get("ProgramFiles", "C:/Program Files")
    pfx = os.environ.get("ProgramFiles(x86)", "C:/Program Files (x86)")

    word_found = detect_windows_executable(
        "WINWORD.EXE",
        f"{pf}/Microsoft Office",
        f"{pfx}/Microsoft Office",
    )
    if not word_found:
        warnings.append(
            "Microsoft Word not detected. Rebranding still works; "
            "--also-export-pdf will be unavailable until Word is installed."
        )

    lo_found = detect_windows_executable(
        "soffice.exe",
        f"{pf}/LibreOffice",
        f"{pfx}/LibreOffice",
    )
    if not lo_found:
        warnings.append(
            "LibreOffice not detected. Rebranding still works; EMF logo "
            "detection in pre-2020 BSI docs will be skipped. To enable: "
            "winget install TheDocumentFoundation.LibreOffice"
        )

    return warnings


def main() -> int:
    py_warning = check_python_version()
    if py_warning:
        print(f"[rto-rebrander] WARNING: {py_warning}", flush=True)
        # Continue anyway — older Python may still work for basic operations.

    req = find_requirements_txt()
    if req is None:
        print("[rto-rebrander] requirements.txt not found; skipping dep check.", flush=True)
        return 0

    packages = parse_requirements(req)
    if not packages:
        return 0  # nothing to check

    missing = fast_check(packages)

    if not missing:
        return 0  # Fast path: everything installed.

    # Slow path: install missing deps.
    print(
        f"[rto-rebrander] First-run setup: installing {len(missing)} Python "
        f"package(s) ({', '.join(missing)})...",
        flush=True,
    )
    ok, log = pip_install(req)
    if ok:
        # Re-check to confirm install succeeded.
        still_missing = fast_check(packages)
        if still_missing:
            print(
                f"[rto-rebrander] pip reported success but these are still missing: "
                f"{', '.join(still_missing)}. You may need to restart your shell.",
                flush=True,
            )
        else:
            print("[rto-rebrander] All dependencies installed successfully.", flush=True)
    else:
        print("[rto-rebrander] Auto-install failed. Run manually:", flush=True)
        print(f"[rto-rebrander]   python -m pip install --user -r \"{req}\"", flush=True)
        print(f"[rto-rebrander] pip output (last 500 chars): ...{log[-500:]}", flush=True)

    for warning in check_optional_binaries():
        print(f"[rto-rebrander] Note: {warning}", flush=True)

    return 0


if __name__ == "__main__":
    sys.exit(main())
