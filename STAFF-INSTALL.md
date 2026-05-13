# RTO Rebrander — Staff Install Guide

> **For aEX Institute staff only.** This tool rebrands BSI Learning documents to our new aEX Institute branding. The repository is public for distribution convenience, but the contents are aEX-specific and have no use to anyone outside our organisation.

This guide walks you through installing the **RTO Document Rebrander** as a Claude Code plugin on your work laptop. The installer does almost all the work for you — you'll just need to double-click one file, then type one short line inside Claude Code.

> **Estimated time:** about 5 minutes.

---

## What you need

| Requirement | Already on your laptop? |
|---|---|
| **Windows 10 or 11** | Yes (work laptop) |
| **Microsoft Word** | Yes (work laptop) |
| **Claude Code Desktop** | Yes (you're reading this in it, presumably) |
| **Python 3.10+** | The installer will set this up for you if missing |

You **do not** need a GitHub account. You **do not** need to know any commands.

---

## Step 1 — Run the installer

1. Halil has shared a file called **`install.bat`** with you (via SharePoint, Teams, or email).
2. Save it somewhere you can find, such as your Desktop.
3. **Double-click `install.bat`**.
4. Windows will pop up a security prompt ("User Account Control"). Click **Yes** to allow it to run.
5. A black PowerShell window will appear. Press any key when prompted.

The installer will:
- Check if you have Python 3.10+, and install it for you if you don't (this can take a minute)
- Remove any older version of the rebrander you may have had
- Tell Claude Code where to find the new plugin

When it's finished, you'll see a yellow message that says **"Setup complete!"** with one last step to do inside Claude Code.

---

## Step 2 — Install the plugin in Claude Code

1. **Open Claude Code Desktop** (if it's already open, close and reopen it first so it picks up the new settings).
2. In the chat input area at the bottom, type or paste this single line and press Enter:

```
/plugin install rto-rebrander@aex-internal
```

3. Claude Code will download the plugin and confirm when it's done.
4. **Close and reopen Claude Code Desktop** one more time. The first time the plugin runs, it quietly installs a few small Python helpers behind the scenes (you may see a one-time progress message — this only happens once).

That's it. The rebrander is now ready to use.

---

## Step 3 — Try it out

In any Claude Code conversation, just ask Claude in plain English. You don't need to remember any commands. Examples:

- *"Rebrand all the docs in `C:\BSI Documents\Learner Guides` to aEX branding."*
- *"Apply the new aEX template to this file: `C:\Documents\BSBOPS502 Learner Guide.docx`"*
- *"Replace the BSI Learning logo with the aEX logo in this folder."*

Claude reads your request, picks up that the rebrander applies, and runs it. Output goes to a folder called **`rebranded/`** next to each source file.

After every batch, please:
1. Open the **`rebrand_report.md`** file in the output folder
2. Spot-check at least one document per family (learner guide, assessment tool, policy, handbook) before sending the batch on

---

## Updating to a new version (later)

When a new version is released, you don't need to download anything. Just type two short lines in Claude Code:

```
/plugin marketplace update
/plugin update rto-rebrander
```

Then close and reopen Claude Code Desktop. Done.

---

## Troubleshooting

| Problem | What to try |
|---|---|
| Double-clicking `install.bat` doesn't do anything | Right-click it, choose **Run as administrator**. |
| Installer says **"winget unavailable"** | Your Windows is older than expected. Manually install Python from https://www.python.org/downloads/ (tick **"Add Python to PATH"** on the first screen), then re-run `install.bat`. |
| `/plugin install` says **"plugin not found"** | The installer didn't finish properly. Run `install.bat` again first. |
| Claude doesn't recognise the rebrander when you ask it to | Type `/plugin list` in Claude Code. If `rto-rebrander` is listed but not enabled, type `/plugin enable rto-rebrander`. If it's not listed at all, run `install.bat` again. |
| Rebrander says **"docx2pdf failed"** | Microsoft Word isn't installed or isn't licenced. The basic rebrand still works; only the `--also-export-pdf` option is unavailable. Contact IT. |
| Rebrander warns about **"LibreOffice not detected"** | This is just a warning, not an error. You only need LibreOffice for very old (pre-2020) BSI docs with EMF-format logos. Most modern docs don't need it. |
| Anything else | Email **halil.houssein@aexinstitute.com.au** |

---

## What this plugin contains (for the curious)

- The rebranding engine (Python scripts that surgically edit Word/PowerPoint XML)
- The aEX brand template, logos, theme colours, and fonts
- Compliance carve-outs that protect unit codes, ASQA references, our RTO number (21371), our ABN, and version-history entries
- A reference template (`aEX Institute LG Word Template May2026.docx`) used by the template-transplant strategy

The plugin is **internal to aEX Institute**. Please don't share `install.bat` or the GitHub repository link with anyone outside aEX — while the contents are publicly accessible, they're tailored to our organisation and not useful elsewhere.
