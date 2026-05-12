"""
rebrand.py — main entry point for rebranding RTO documents.

Usage:
    python rebrand.py <input> [<output>] [--spec PATH] [--assets PATH]
                      [--also-export-pdf] [--dry-run] [--template PATH]

`input` is either a single file (.docx/.pptx/.pdf) OR a directory to
walk. If `output` is omitted (recommended), each rebranded file is
written to a `rebranded/` folder alongside its source — so a source
at `…/FSKOCM001/Learner/Guide.docx` ends up at
`…/FSKOCM001/Learner/rebranded/Guide.docx`. This keeps rebranded
outputs locally adjacent to the originals on shared drives.

If `output` is given, all rebranded files go into that single folder
preserving relative-to-input paths (the old behaviour).

Reports are written next to the rebranded files: into `<input>/rebranded/`
for directory inputs, or into the source's parent `rebranded/` for
single-file inputs, or into `output/` when explicit.

Examples:
    # Single file -> beside-source rebranded/ folder
    python rebrand.py "C:/…/FSKOCM001 Learner Guide.docx"

    # Directory walk -> each file gets a sibling rebranded/ folder
    python rebrand.py "C:/…/FSK10119 EDM Review"

    # Use the bundled aEX template (template-transplant strategy)
    python rebrand.py "C:/…/Guide.docx" --template ../assets/aex_target_template.docx

    # Also export Word PDFs alongside .docx outputs
    python rebrand.py "C:/…/Guide.docx" --also-export-pdf

    # Old-style explicit output folder still works
    python rebrand.py ./old_docs ./rebranded_out

    # Dry run — produce a report without writing files
    python rebrand.py "C:/…/Guide.docx" --dry-run
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

import yaml

# Make sibling modules importable when run as `python rebrand.py ...`
SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))

from rebrand_docx import rebrand_docx  # noqa: E402
from rebrand_pptx import rebrand_pptx  # noqa: E402
from rebrand_pdf import rebrand_pdf, convert_docx_to_pdf  # noqa: E402


SUPPORTED_EXTS = {".docx", ".pptx", ".pdf"}

# Subfolder name used when --output is not specified. Each source file's
# rebranded copy lands in this folder beside the original, e.g.
# .../FSKOCM001/Learner/<this>/FSKOCM001 Learner Guide V4.0 EDM.docx
DEFAULT_OUTPUT_SUBDIR = "rebranded"


def load_spec(spec_path: Path) -> dict[str, Any]:
    with open(spec_path) as f:
        return yaml.safe_load(f)


def _is_supported_file(p: Path) -> bool:
    if p.suffix.lower() not in SUPPORTED_EXTS:
        return False
    if p.name.startswith("~$") or p.name.startswith("."):
        return False
    # Skip anything ALREADY in a `rebranded/` folder so re-running on a
    # parent dir doesn't pick up previous outputs and rebrand them again.
    if DEFAULT_OUTPUT_SUBDIR in p.parts:
        return False
    return True


def collect_files(root: Path) -> list[Path]:
    """Return supported files. `root` may be a single file or a directory."""
    if root.is_file():
        return [root] if _is_supported_file(root) else []
    files = []
    for p in root.rglob("*"):
        if p.is_file() and _is_supported_file(p):
            files.append(p)
    return sorted(files)


def resolve_destination(src: Path, input_root: Path, output_root: Path | None) -> Path:
    """Decide where the rebranded copy of `src` should be written.

    - With explicit `output_root`: preserve the source's path relative
      to `input_root` under `output_root` (old behaviour).
    - Without `output_root`: write to `<src.parent>/rebranded/<src.name>`,
      i.e. a `rebranded/` folder sitting next to the source file.
    """
    if output_root is not None:
        if src == input_root:
            return output_root / src.name
        try:
            rel = src.relative_to(input_root)
        except ValueError:
            rel = Path(src.name)
        return output_root / rel
    return src.parent / DEFAULT_OUTPUT_SUBDIR / src.name


def resolve_report_dir(input_root: Path, output_root: Path | None) -> Path:
    """Where the rebrand_report.{json,md} files should land."""
    if output_root is not None:
        return output_root
    if input_root.is_file():
        return input_root.parent / DEFAULT_OUTPUT_SUBDIR
    return input_root / DEFAULT_OUTPUT_SUBDIR


def process_one(
    src: Path,
    dst: Path,
    spec: dict[str, Any],
    asset_dir: Path,
    source_search_dirs: list[Path],
    template_path: Path | None = None,
) -> dict[str, Any]:
    ext = src.suffix.lower()
    if ext == ".docx":
        if template_path is not None:
            # Template-transplant: rebuild source's chrome from the template,
            # preserving body content. Then run find-replace pass on the
            # result to clean up any residual brand text/colours.
            from rebrand_docx_template import transplant_chrome
            log = transplant_chrome(
                src, template_path, dst, spec=spec, asset_dir=asset_dir
            )
            # Surface body-cleanup totals at top level for the report
            body = log.get("body_cleanup", {})
            log.setdefault("text_replacements", body.get("text_replacements", 0))
            log.setdefault("colour_replacements", body.get("colour_replacements", 0))
            log.setdefault("footer_replacements", body.get("footer_replacements", 0))
            log.setdefault("images_swapped", body.get("images_swapped", []))
            return log
        return rebrand_docx(src, dst, spec, asset_dir)
    if ext == ".pptx":
        return rebrand_pptx(src, dst, spec, asset_dir)
    if ext == ".pdf":
        return rebrand_pdf(src, dst, spec, source_search_dirs)
    raise ValueError(f"Unsupported extension: {ext}")


def write_summary(report: dict[str, Any], out_dir: Path) -> tuple[Path, Path]:
    """Write JSON + Markdown summary reports. Returns paths."""
    out_dir.mkdir(parents=True, exist_ok=True)
    json_path = out_dir / "rebrand_report.json"
    md_path = out_dir / "rebrand_report.md"

    with open(json_path, "w") as f:
        json.dump(report, f, indent=2, default=str)

    # Markdown summary
    lines = []
    lines.append(f"# Rebrand report\n")
    lines.append(f"_Generated {report['timestamp']}_\n")
    lines.append(f"**Old brand:** {report['old_brand']}  ")
    lines.append(f"**New brand:** {report['new_brand']}\n")
    lines.append(f"## Summary\n")
    lines.append(f"- Files processed: **{report['total_files']}**")
    lines.append(f"- Successful: **{report['successful']}**")
    lines.append(f"- Skipped (need manual review): **{report['needs_review']}**")
    lines.append(f"- Failed: **{report['failed']}**")
    lines.append(f"- Total text replacements: **{report['totals']['text_replacements']}**")
    lines.append(f"- Total colour replacements: **{report['totals']['colour_replacements']}**")
    lines.append(f"- Total footer replacements: **{report['totals']['footer_replacements']}**")
    lines.append(f"- Total logo images swapped: **{report['totals']['images_swapped']}**\n")

    needs_review = [r for r in report["files"] if r.get("needs_manual_review")]
    if needs_review:
        lines.append(f"## Files needing manual review ({len(needs_review)})\n")
        for r in needs_review:
            lines.append(f"- `{r['source']}`")
            for w in r.get("warnings", []):
                lines.append(f"    - ⚠️ {w}")
        lines.append("")

    failed = [r for r in report["files"] if r.get("error")]
    if failed:
        lines.append(f"## Failures ({len(failed)})\n")
        for r in failed:
            lines.append(f"- `{r['source']}` — {r['error']}")
        lines.append("")

    lines.append(f"## All files\n")
    lines.append("| File | Text | Colours | Footers | Images | Notes |")
    lines.append("|---|---:|---:|---:|---:|---|")
    for r in report["files"]:
        notes_parts = []
        if r.get("needs_manual_review"):
            notes_parts.append("⚠️ review")
        if r.get("error"):
            notes_parts.append(f"❌ {r['error']}")
        if r.get("method") == "use_source_file":
            notes_parts.append(f"used source: `{Path(r['source_file']).name}`")
        notes = "; ".join(notes_parts) or "ok"
        lines.append(
            f"| `{Path(r['source']).name}` "
            f"| {r.get('text_replacements', 0)} "
            f"| {r.get('colour_replacements', 0)} "
            f"| {r.get('footer_replacements', 0)} "
            f"| {len(r.get('images_swapped', []))} "
            f"| {notes} |"
        )
    md_path.write_text("\n".join(lines), encoding="utf-8")

    return json_path, md_path


def main() -> int:
    parser = argparse.ArgumentParser(description="Rebrand RTO documents")
    parser.add_argument(
        "input",
        type=Path,
        help="Input file or folder (.docx/.pptx/.pdf). Folders are walked recursively.",
    )
    parser.add_argument(
        "output",
        type=Path,
        nargs="?",
        default=None,
        help=(
            "Optional output folder. If omitted, each rebranded file is "
            "written to a 'rebranded/' subfolder next to its source. "
            "Pass an explicit path to consolidate all outputs into one "
            "folder (preserves paths relative to input)."
        ),
    )
    parser.add_argument(
        "--spec",
        type=Path,
        default=SCRIPT_DIR.parent / "assets" / "brand_spec.yaml",
        help="Path to brand_spec.yaml",
    )
    parser.add_argument(
        "--assets",
        type=Path,
        default=SCRIPT_DIR.parent / "assets",
        help="Path to assets folder (theme XML, logos)",
    )
    parser.add_argument(
        "--also-export-pdf",
        action="store_true",
        help="After rebranding .docx files, also export PDFs via LibreOffice",
    )
    parser.add_argument(
        "--template",
        type=Path,
        default=None,
        help=(
            "Path to a target template .docx (e.g. Learner_Guide_Document.docx). "
            "When provided, .docx files will undergo a 'template transplant' — "
            "source body content is preserved but the cover page, header, "
            "footer, theme, and styling are replaced with the template's. "
            "Use this when the rebranded output should LOOK LIKE the new "
            "brand template, not just have BSI references swapped on the "
            "original layout. Without --template, the standard find-and-"
            "replace strategy is used (preserves original layout)."
        ),
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Don't write files; just produce a report of what would change",
    )
    args = parser.parse_args()

    spec = load_spec(args.spec)
    asset_dir = args.assets
    input_root = args.input
    output_root = args.output
    files = collect_files(input_root)

    print(f"Found {len(files)} files to process")
    print(f"Old brand: {spec['old_brand']['name']}")
    print(f"New brand: {spec['new_brand']['name']}")
    if output_root is None:
        print(f"Output mode: per-source 'rebranded/' folder beside each file")
    else:
        print(f"Output folder: {output_root}")
    print()

    # For PDF source-detection: search the input root for sibling .docx/
    # .pptx that may carry the same brand content as a PDF-only input.
    pdf_search_root = input_root if input_root.is_dir() else input_root.parent
    source_search_dirs = [pdf_search_root]

    report: dict[str, Any] = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "old_brand": spec["old_brand"]["name"],
        "new_brand": spec["new_brand"]["name"],
        "total_files": len(files),
        "successful": 0,
        "failed": 0,
        "needs_review": 0,
        "totals": {
            "text_replacements": 0,
            "colour_replacements": 0,
            "footer_replacements": 0,
            "images_swapped": 0,
        },
        "files": [],
    }

    for src in files:
        dst = resolve_destination(src, input_root, output_root)
        try:
            rel_display = src.relative_to(input_root) if input_root.is_dir() else src.name
        except ValueError:
            rel_display = src.name
        print(f"  {rel_display}  ->  {dst}")
        try:
            if args.dry_run:
                # Process to a temp file just to count changes, then discard
                import tempfile

                with tempfile.NamedTemporaryFile(
                    suffix=src.suffix, delete=False
                ) as tmp:
                    tmp_path = Path(tmp.name)
                log = process_one(
                    src, tmp_path, spec, asset_dir, source_search_dirs,
                    template_path=args.template,
                )
                tmp_path.unlink(missing_ok=True)
            else:
                log = process_one(
                    src, dst, spec, asset_dir, source_search_dirs,
                    template_path=args.template,
                )

            # Optional PDF export for rebranded docx
            if (
                args.also_export_pdf
                and not args.dry_run
                and src.suffix.lower() == ".docx"
            ):
                pdf_out = convert_docx_to_pdf(dst, dst.parent)
                if pdf_out:
                    log["pdf_exported"] = str(pdf_out)
                else:
                    log.setdefault("warnings", []).append(
                        "PDF export failed (LibreOffice not available or conversion error)"
                    )

            report["files"].append(log)
            report["successful"] += 1
            if log.get("needs_manual_review"):
                report["needs_review"] += 1
            report["totals"]["text_replacements"] += log.get("text_replacements", 0)
            report["totals"]["colour_replacements"] += log.get("colour_replacements", 0)
            report["totals"]["footer_replacements"] += log.get("footer_replacements", 0)
            report["totals"]["images_swapped"] += len(log.get("images_swapped", []))
        except Exception as e:
            report["files"].append({"source": str(src), "error": str(e)})
            report["failed"] += 1
            print(f"    FAILED: {e}")

    report_dir = (
        Path.cwd() if args.dry_run else resolve_report_dir(input_root, output_root)
    )
    json_path, md_path = write_summary(report, report_dir)
    print()
    print(f"Done. Reports written to:")
    print(f"  {md_path}")
    print(f"  {json_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
