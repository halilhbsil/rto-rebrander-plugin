# Safety carveouts

This rebrander operates on RTO documents, where some text *looks like* the old brand but must never be changed because it's compliance metadata or historical record.

## What's protected

The carveouts live in `assets/brand_spec.yaml` under `preserve.text_patterns` and are applied as a sentinel-wrapping step *before* brand-name replacement, then restored *after*. This is what's currently protected:

| Category | Pattern | Examples |
|---|---|---|
| Unit codes | `\b[A-Z]{3,5}\d{2,3}[A-Z]?\b` | `BSBWHS411`, `FSKNUM017`, `CHCAGE001`, `TAEDEL411` |
| BSB qualification codes | `\bBSB\d{5,6}\b` | `BSB30120`, `BSB40520` |
| CHC qualification codes | `\bCHC\d{5,6}\b` | `CHC30121` |
| FSK qualification codes | `\bFSK\d{5,6}\b` | `FSK10219` |
| CPP qualification codes | `\bCPP\d{5,6}\b` | `CPP20218` |
| TAE qualification codes | `\bTAE\d{5,6}\b` | `TAE40122` |
| Training Package literal | `\bTraining Package\b` | "FSK Foundation Skills Training Package" |
| Regulatory bodies | `\bASQA\b`, `\btraining\.gov\.au\b`, `\btga\.gov\.au\b` | URL/citation references |
| Standards for RTOs | `Standards for RTOs`, `Clause \d+(\.\d+)*` | "Standards for RTOs 2015 Clause 3.1.1" |
| ABN | `ABN\s*\d{2}\s*\d{3}\s*\d{3}\s*\d{3}` | "ABN 20 106 311 086" |
| RTO code | `RTO\s*21371` | "RTO 21371" |

## How sentinel-wrapping works

In `rebrand_docx.py` (and the equivalent in `rebrand_pptx.py`), every text node is processed like this:

```python
def _protected_replace_in_t(m):
    body = m.group(2)
    # Wrap each preserve-pattern match in sentinels
    protected, preserved = _protect_preserved_spans(body, preserve_patterns)
    # Now do brand replacement — sentinels can't match because they
    # contain non-text characters that don't appear in any pattern
    for variant in [*variants, old_name]:
        protected = protected.replace(variant, new_name)
    # Restore the original preserved text
    return _restore_preserved_spans(protected, preserved)
```

The sentinel itself is `\x00\x01PRESERVE\x00\x01<index>\x00\x02PRESERVE\x00\x02` — non-printable boundary chars guarantee the sentinel can't accidentally match a brand pattern.

## What's *not* automatically protected — and what to do

### Version history rows

A typical RTO document has a Version History table at the front like:

```
Version | Date       | Author              | Description
1.0     | 20.07.2021 | BSI Learning Curric.| Spellchecked
2.0     | 23.04.2024 | Daryl Gover        | Formatting
4.0     | 04.04.2026 | Eustace D'Mello    | New aEX Template
```

The 2021 entry's "BSI Learning Curric." is a **historical record** that should not be rewritten — doing so would falsify the document's audit trail.

**This is now handled by two regex carveouts** in `brand_spec.yaml`:

- `BSI Learning Curriculum` — protects the historical role/department name
- `BSI Learning \([^)]+\)` — protects historical individual author attributions like "BSI Learning (C Potts)"

These cover the two common shapes of version-history author entries. If the user has *other* historical phrasings (for example "BSI Learning Compliance Team" or "BSI Learning — J Smith"), add patterns to `preserve.text_patterns` for those specific shapes too.

**Recommended workflow:**

1. **Pre-flight inspection.** Run `inspect_doc.py` on a representative document. The text-occurrence count should match what's visible in the body — if it's higher, audit before running the batch (something is in a place you didn't expect).
2. **Manual post-edit for unusual phrasings.** After the rebrand, spot-check the Version History tables on highest-stakes documents. If a historical author phrase wasn't covered by the carveouts, either revert that one cell manually or add the pattern to the spec and re-run.
3. **More sophisticated detection (deferred).** A more robust approach would parse `<w:tbl>` elements, identify ones with a "Version History" heading row, and protect entire rows whose Date column predates `preserve.version_history_cutoff_date`. This isn't implemented because:
   - Visual-detection of "Version History" tables is fragile (different templates use different headings).
   - Date parsing is fragile (DD.MM.YYYY vs DD/MM/YYYY vs other formats).
   - The current regex carveouts catch the common cases, and the simple "manual spot-check" workflow handles the rest reliably for a 50–200 document one-off rebrand.

   If a future rebrand really needs this, the entry point would be in `rebrand_docx._process_xml_part` — look for `<w:tbl>` elements containing a heading row with text "Version History", then within those tables protect the entire `<w:tr>` for any row whose date precedes the cutoff.

### Email addresses and URLs

Strings like `info@bsilearning.com.au` or `https://bsilearning.com.au/...` are not currently protected. The text replacer would turn "BSI Learning" into "aEX Institute" but **not** the lowercase "bsilearning" inside an email/URL — because the search is case-sensitive on the exact phrase "BSI Learning". So in practice these tend to come through unchanged, which is what we want (changing them might break working links).

If the user *does* want to update email/URL references, add explicit replacement entries — but be aware that this can break live links until the new domain resolves.

### Cover-page rasterized text

If a document has a cover page where "BSI Learning" is a rasterized image (rather than editable text), the rebrander cannot change it. The image-swap logic only catches images that perceptually match the bundled `old_bsi_logo.png`. Stylised cover pages often won't match.

**Mitigation:** spot-check the first page of the first document in each family (learner guide, assessment tool, policy, handbook). If you find rasterized text, replace those images manually using the new aEX assets in `assets/`.
