# aEX Internal Plugin Marketplace

> **⚠️ aEX Institute internal tool — not useful to anyone else.**
>
> This repository is **public** purely so aEX Institute staff can install the plugin without needing a GitHub account. The contents (rebranding engine, bundled aEX logos, theme XML, brand template, fonts, and brand specification) are **specific to aEX Institute** and have **no general utility**. The find-and-replace direction is hardcoded BSI Learning → aEX Institute, the brand assets are aEX-licensed, and the compliance metadata refers to aEX's RTO number (21371) and ABN. If you arrived here from a search, this is not a generic RTO rebrander you can adapt — it's purpose-built for one organisation and one migration.
>
> The underlying *technique* (OOXML find-and-replace + template transplant with compliance-aware sentinels) might inform similar projects, but the actual code is not reusable as-is. Please do not fork.

---

Internal Claude Code plugin marketplace for aEX Institute. Currently hosts one plugin: **rto-rebrander**, which rebrands BSI Learning RTO documents to aEX Institute branding.

## For staff

If you're aEX staff installing this for the first time, follow the **[STAFF-INSTALL.md](STAFF-INSTALL.md)** guide — it covers prerequisites (Python, GitHub CLI), GitHub access, migration from the old skill-based install, and verification.

The rest of this document is a quick reference for the maintainer.

## Install (one-time)

```
/plugin marketplace add <git-url-of-this-repo>
/plugin install rto-rebrander@aex-internal
```

After install, restart Claude Code once. The plugin's SessionStart hook will detect missing Python packages and install them automatically (`PyYAML`, `Pillow`, `imagehash`). You'll see progress lines in the terminal during first run, then subsequent sessions start silently.

## Update

When a new version is released:

```
/plugin marketplace update
/plugin update rto-rebrander
```

Both can also be triggered through the `/plugin` UI.

## Requirements on the user's machine

| Required | Auto-handled by plugin? |
|---|---|
| Python 3.10+ | No — install from [python.org](https://www.python.org/downloads/) once |
| Pure-Python deps (PyYAML, Pillow, imagehash) | Yes — auto-installed via SessionStart hook |
| Microsoft Word | No — already installed on staff machines; only required if you use `--also-export-pdf` |
| LibreOffice (optional, EMF logo support) | No — optional; install with `winget install TheDocumentFoundation.LibreOffice` if processing pre-2020 BSI docs |

## Usage

Once installed, just ask Claude in any session:

> "Rebrand the docs in `C:\BSI Documents\Learner Guides` to aEX branding."

Claude loads the skill automatically. To invoke directly:

```
/rto-rebrander:rto-document-rebrander
```

Output goes to a `rebranded/` subfolder next to each source file unless an output dir is specified. See `plugins/rto-rebrander/skills/rto-document-rebrander/SKILL.md` for the full feature set.

## Layout

```
.
├── .claude-plugin/
│   └── marketplace.json              # this marketplace catalog
├── plugins/
│   └── rto-rebrander/                # the plugin
│       ├── .claude-plugin/
│       │   └── plugin.json
│       ├── hooks/
│       │   └── hooks.json            # SessionStart auto-installer wiring
│       ├── scripts/
│       │   └── ensure_deps.py        # idempotent pip-install on first session
│       └── skills/
│           └── rto-document-rebrander/   # the actual skill (engine + assets)
└── README.md
```

## Releasing a new version

1. Edit code under `plugins/rto-rebrander/skills/rto-document-rebrander/`
2. Bump `version` in both `.claude-plugin/marketplace.json` and `plugins/rto-rebrander/.claude-plugin/plugin.json`
3. Commit and push
4. Staff run `/plugin update rto-rebrander` next time they start Claude Code

Versions follow semver. Major bumps for breaking changes (e.g. brand-spec format changes); minor for new behavior; patch for fixes.
