# RTO Document Rebrander — install & quick start

This is a Claude **Skill** that automates rebranding our RTO documents from **BSI Learning** to **aEX Institute** — Word `.docx`, PowerPoint `.pptx`, and (best-effort) PDFs. Logos, colours, organisation name, and footer text are swapped while RTO compliance metadata (unit codes, ASQA references, ABN, RTO 21371, historical version-history entries) is preserved.

> **Audience:** internal aEX Institute staff. The bundled template and brand assets are aEX-licensed; do not redistribute outside the organisation.

## What this skill does

When you ask Claude something like *"rebrand the docs in `C:\…\OldGuides` and put the output in `C:\…\Rebranded`"*, the skill:

1. Walks the input folder for `.docx`, `.pptx`, `.pdf` files.
2. Either does a minimal-touch find-and-replace (default) or rebuilds each `.docx` against the bundled aEX template (template-transplant strategy) — Claude picks based on the request.
3. Preserves compliance metadata via regex carveouts.
4. Writes the rebranded files plus a `rebrand_report.md` summarising every change.

See `SKILL.md` for the full feature description (it's also what Claude reads to decide when to use the skill).

## System requirements

| Requirement | Why | How to check |
|---|---|---|
| **Python 3.10+** on PATH | runs the rebrand scripts | `python --version` |
| **PyYAML, Pillow, imagehash** Python packages | OOXML manipulation + perceptual-hash logo matching | `pip install -r requirements.txt` |
| **Microsoft Word** (Windows) or **LibreOffice** *(optional)* | only needed for: `.docx → .pdf` export (`--also-export-pdf` flag) AND detection of EMF-format BSI logos in older docs | Word: Start menu. LibreOffice: `soffice --version` |

**If you skip the optional step:** the skill still works for the find-and-replace and template-transplant strategies. You'll lose detection of EMF logos (rare — only affects pre-2020 BSI docs) and PDF export (you can export to PDF from Word manually after the rebrand).

## Install

### 1. Place the folder in your Claude skills directory

Both **Claude Code** and **Claude Desktop** load skills from the same location and run them the same way (locally on your machine, with direct filesystem access). Copy the entire `rto-document-rebrander/` folder to:

| OS | Path |
|---|---|
| **Windows** | `%USERPROFILE%\.claude\skills\rto-document-rebrander\` |
| **macOS / Linux** | `~/.claude/skills/rto-document-rebrander/` |

If the `skills` folder doesn't exist yet, create it.

### 2. Install the Python dependencies

From inside the skill folder:

```powershell
# Windows PowerShell
python -m pip install -r requirements.txt
```

```bash
# macOS / Linux
python3 -m pip install -r requirements.txt
```

### 3. Run the smoke test

```powershell
# Windows
python scripts\smoke_test.py
```

```bash
# macOS / Linux
python scripts/smoke_test.py
```

You should see `OK — skill is ready.` If anything fails, the script prints which dependency or asset is missing.

### 4. (Optional) Install LibreOffice

Only needed if you want PDF export or EMF-logo detection:

- **Windows:** `winget install TheDocumentFoundation.LibreOffice`
- **macOS:** `brew install --cask libreoffice`
- **Linux:** your distro's package manager (`apt install libreoffice`, `dnf install libreoffice`, etc.)

After install, restart your terminal so `soffice` is on PATH.

## How to use it

Just ask Claude. Some examples it'll recognise:

- *"Rebrand `C:\BSI Documents\Learner Guides\FSKOCM001 Learner Guide.docx`."*
- *"Rebrand every document in `C:\BSI Documents\Learner Guides`."*
- *"Replace the BSI logo with the aEX logo across this folder."*
- *"Make these old docs look like our new aEX learner guide template."* (template-transplant strategy)

**Where rebranded files land.** By default, every rebranded file is written to a `rebranded/` subfolder sitting *next to* the source. So a learner guide at `…/FSKOCM001/Learner/Guide.docx` gets a rebranded copy at `…/FSKOCM001/Learner/rebranded/Guide.docx`. Source files are never overwritten. If you want all outputs consolidated in one place, ask Claude to write them to a specific folder.

Claude will run the rebrand scripts, summarise what changed, and flag any files that need manual review.

## Verifying output

After a rebrand, always:

1. Open `rebrand_report.md` — for a single-file run it's in the same `rebranded/` folder as the output; for a directory walk it's at `<input-folder>/rebranded/rebrand_report.md`. Read the per-file change counts and look at the "Files needing manual review" section.
2. Spot-check at least one document from each family (learner guide, assessment tool, policy, handbook). The `BSBOPS502` page-with-chevron-diagram is the canonical regression-test anchor — if its three textboxes (Strategic Plan / Business Plans / Operational Plans) each show all three bullets, the typography preservation is working.
3. Open one or two PDFs to verify visual fidelity.

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `ModuleNotFoundError: No module named 'yaml'` | Step 2 not run | `pip install -r requirements.txt` |
| `'WindowsPath' object has no attribute 'lower'` | Old version of the skill — bug fixed 2026-05-07 | Replace your skill folder with this version |
| Footer reads `Learner Guide v 1.02` (orphan `2`) | Old version — substitutor field-cache fix landed 2026-05-07 | Replace your skill folder with this version |
| Textboxes/SmartArt clip text in BSBOPS502-style docs | Old version — typography preservation fix landed 2026-05-07 | Replace your skill folder with this version |
| `LibreOffice not available` warning in the report | Optional dependency missing | Install LibreOffice (step 4) or ignore the warning |
| EMF logos in old docs not detected | Optional dependency missing | Install LibreOffice |
| PDF export skipped | Optional dependency missing | Install LibreOffice or export from Word manually |

## What's in the box

```
rto-document-rebrander/
├── SKILL.md            ← what Claude reads to know what this skill does
├── README.md           ← this file (for humans)
├── requirements.txt    ← Python dependencies
├── scripts/            ← the rebrand engine (don't edit unless you know OOXML)
│   ├── rebrand.py
│   ├── inspect_doc.py
│   ├── rebrand_docx.py
│   ├── rebrand_docx_template.py
│   ├── rebrand_pptx.py
│   ├── rebrand_pdf.py
│   └── smoke_test.py
├── assets/             ← brand-specific data (logos, theme XML, target template)
│   ├── brand_spec.yaml ← single source of truth for what to find/replace
│   ├── aex_target_template.docx  ← gold-standard learner guide template
│   └── … (logos, theme, fonts)
└── references/         ← deeper docs on brand spec, PDF strategy, carveouts
```

## Support

- For skill bugs or edge-case docs: file an issue with the team that owns this skill.
- For Claude / Claude Code questions: see the official Anthropic docs.
- For Word/LibreOffice rendering issues that aren't fixed by the skill: the rebrand engine handles ~90% of cases; the remaining 10% (rasterised cover-page text, complex SmartArt) need a human pass — see the "When to flag for manual review" section in `SKILL.md`.
