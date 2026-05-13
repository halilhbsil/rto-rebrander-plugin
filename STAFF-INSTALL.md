# RTO Rebrander — Staff Install Guide

This guide walks you through installing the **RTO Document Rebrander** as a Claude Code plugin on your work laptop. Once installed, you can ask Claude to rebrand BSI Learning documents to aEX Institute branding directly inside any Claude Code session, and the tool stays up to date automatically.

> **Estimated time:** 10–15 minutes the first time.
> **Audience:** aEX Institute staff. The plugin is internal-only and the GitHub repo it lives in is private.

---

## What you need before you start

Make sure all four of these are true on your laptop. If anything is missing, the relevant step below shows you how to install it.

| Requirement | How to check | If missing |
|---|---|---|
| **Windows 10 or 11** | `winver` in Start menu | (You're fine if you have a recent work laptop) |
| **Microsoft Word** | Already on staff laptops | Contact IT |
| **Python 3.10 or newer** | Open PowerShell, run `python --version` | See Step 0 below |
| **Claude Code Desktop** | You're reading this inside it, presumably | https://claude.com/code |

You **do not** need to know how to code. Every step in this guide is copy-paste.

---

## Step 0 — Install Python (skip if you already have it)

Open PowerShell (press **Win + X**, choose *Terminal* or *Windows PowerShell*) and run:

```powershell
python --version
```

If you see something like `Python 3.12.x`, you can skip to Step 1.

If you see *"not recognized"* or a version below 3.10:

1. Go to https://www.python.org/downloads/
2. Click the big yellow **Download Python** button
3. Run the installer
4. **IMPORTANT:** On the first screen, tick the box that says **"Add Python to PATH"** before clicking Install
5. Click *Install Now* and wait until it says *Setup was successful*
6. Close PowerShell and reopen it, then run `python --version` again to confirm

---

## Step 1 — Install the GitHub CLI

The GitHub CLI is a tiny helper that lets Claude Code download the plugin from our internal GitHub. You install it once and never think about it again.

1. Go to https://cli.github.com
2. Click **Download for Windows**
3. Run the installer (accept all defaults)
4. Close PowerShell and reopen it
5. Verify it worked:

```powershell
gh --version
```

You should see `gh version 2.x.x` or similar.

---

## Step 2 — Sign in to GitHub

Still in PowerShell:

```powershell
gh auth login
```

Answer the prompts:

| Prompt | Choose |
|---|---|
| Where do you use GitHub? | **GitHub.com** |
| Preferred protocol for Git operations? | **HTTPS** |
| Authenticate Git with your GitHub credentials? | **Y** (Yes) |
| How would you like to authenticate? | **Login with a web browser** |

A code will appear in PowerShell (something like `XXXX-XXXX`). PowerShell will then open your default browser to GitHub. Paste the code and click *Authorize*. Once you see *Congratulations, you're all set!* you can close the browser and return to PowerShell — you'll see a green tick `✓ Logged in as <your-github-username>`.

> **Don't have a GitHub account yet?** Sign up at https://github.com/signup first (free, takes 30 seconds), then tell Halil your username so he can add you as a collaborator before you continue with Step 3.

---

## Step 3 — (Migration only) Remove the old skill

**Only do this step if you previously installed the rebrander as a "skill" by unzipping it into your `.claude/skills/` folder.** If you've never installed the rebrander before, skip to Step 4.

To remove the old version, paste this into PowerShell and press Enter:

```powershell
Remove-Item -Recurse -Force "$env:USERPROFILE\.claude\skills\rto-document-rebrander"
```

You won't see any output — that means it worked. If you see *"Cannot find path..."*, the old version wasn't there, which is also fine.

---

## Step 4 — Install the plugin in Claude Code

Open Claude Code Desktop. In the chat input area, type these two commands one at a time (press Enter after each):

```
/plugin marketplace add halilhbsil/rto-rebrander-plugin
```

You should see a confirmation that the marketplace was added. Then:

```
/plugin install rto-rebrander@aex-internal
```

You'll see a brief installation message. **Now close and reopen Claude Code Desktop** — this is important. The first time the plugin loads, it installs a few small Python helpers in the background. You may see a one-time message that reads:

```
[rto-rebrander] First-run setup: installing Python package(s)...
[rto-rebrander] All dependencies installed successfully.
```

This takes about 30 seconds the first time and never appears again on subsequent sessions.

---

## Step 5 — Verify it worked

In Claude Code, type `/` and look at the menu of available commands. You should see:

```
/rto-rebrander:rto-document-rebrander
```

You can also just ask Claude in plain English:

> *"What does the rto-rebrander plugin do?"*

Claude should respond with a summary of the rebrander's features.

---

## Using the rebrander

You **don't need to type slash commands** to use it. Just ask Claude in plain English. Examples:

- *"Rebrand all the docs in `C:\BSI Documents\Learner Guides` to aEX branding."*
- *"Apply the new aEX template to this file: `C:\Documents\BSBOPS502 Learner Guide.docx`"*
- *"Replace the BSI Learning logo with the aEX logo across this folder."*

Claude reads your request, picks up that the rebrander applies, and runs it. Output goes to a folder called `rebranded/` next to each source file by default.

After the rebrand finishes, **always**:
1. Open the `rebrand_report.md` file in the output folder
2. Spot-check at least one document per family (learner guide, assessment tool, policy, handbook) before treating the whole batch as done

---

## Updating to a new version (later)

When a new version of the rebrander is released, you'll get a heads-up from Halil. To update, just run these two commands inside Claude Code:

```
/plugin marketplace update
/plugin update rto-rebrander
```

Then close and reopen Claude Code Desktop. That's it — no re-downloading, no re-extracting zips.

---

## Troubleshooting

| Problem | What to try |
|---|---|
| `/plugin marketplace add` says **"repository not found"** | You probably don't have access to the private GitHub repo yet. Email Halil your GitHub username so he can add you as a collaborator. |
| `gh auth login` doesn't open a browser | Try `gh auth login --web` explicitly. If that still fails, ask Halil for a Personal Access Token method. |
| Rebrander gives **"docx2pdf failed"** errors | Microsoft Word isn't installed or isn't licenced. The basic rebrand still works; only the `--also-export-pdf` option is unavailable. |
| Rebrander says **"LibreOffice not detected"** | This is just a warning, not an error. You only need LibreOffice if you're processing very old (pre-2020) BSI docs with EMF logos. Most modern docs don't need it. |
| Claude doesn't seem to recognise the rebrander | Type `/plugin list` to confirm `rto-rebrander` is enabled. If it isn't, run `/plugin enable rto-rebrander`. |
| Anything else | Email Halil: halil.houssein@aexinstitute.com.au |

---

## What this plugin contains (for the curious)

- The rebranding engine (Python scripts that surgically edit Word/PowerPoint XML)
- The aEX brand template, logos, theme colours, and fonts
- Compliance carve-outs that protect unit codes, ASQA references, RTO number, ABN, and version-history entries from being rewritten
- A reference template (`aEX Institute LG Word Template May2026.docx`) used by the template-transplant strategy

The plugin is internal-only and the bundled assets are aEX-licensed. **Do not share the GitHub repository link with anyone outside aEX Institute.**
