# pst-to-markdown

A [Claude Code](https://claude.ai/code) skill that converts Outlook `.pst` email archives into an organized Markdown knowledge base — with threaded conversations, extracted attachments, and automatic noise filtering.

## What it does

Drop a `.pst` file into a project and ask Claude to extract the emails. The skill runs a two-stage pipeline:

1. **`readpst`** (from `libpst`) extracts the PST into per-folder `mbox` files, preserving the original folder hierarchy
2. **`pst_merge.py`** reads all `mbox` files globally and produces clean Markdown thread files

Each email thread becomes its own folder:

```
mail_markdown/
  Personal Folders/
    Work/Analyses/
      Project Alpha status__a1b2c3d4/
        conversation.md        ← full thread, chronological
        attachments/
          report_rev3.pdf
          data_analysis.xlsx
    Inbox/
      Re_ Project kickoff__e5f6a7b8/
        conversation.md
```

Each `conversation.md` is self-contained — metadata header, full message bodies in order, and relative links to every attachment:

```markdown
# Project Alpha status

*13 message(s) · merged from 3 folders: `Inbox` · `Sent messages` · `Work/Analyses`*

**From:** Alice Smith <alice@example.com>
**To:** Bob Jones <bob@example.com>
**Date:** Thu, 29 Oct 2015 11:08:09 +0100
**Attachments:**
- [report_rev3.pdf](attachments/report_rev3.pdf)

---

Hi Bob, please find the updated report attached...
```

## Key features

### Global thread merging
Standard PST exporters produce separate files for the Inbox copy and the Sent copy of the same thread. This skill reads all folders at once and unifies them using `Message-ID` / `In-Reply-To` / `References` chains (union-find). One thread = one folder, regardless of how many PST folders it lived in.

### Signature image filtering
Corporate HTML emails embed logo and icon images as inline CID attachments — these are not useful content. The skill detects the email signature block ("Best regards", "Mvh", "С уважением", and others) and skips any image referenced only after it. Images in the body (charts, screenshots, photos) are kept.

### Artifact cleanup
`readpst` emits several artifacts that clutter the output. The optional `cleanup_noise.py` script removes them:

| Artifact | What it is |
|----------|------------|
| `rtf-body.rtf` | RTF copy of the email body — already captured as plain text in `conversation.md` |
| `oledata.mso` | Outlook OLE metadata container |
| `ATT*` ≤ 4608 bytes | Outlook UI icon blobs with no filename |

Large `ATT*` files are kept — they may be real documents that lost their filename during export.

## Installation

```bash
gh repo clone GlebTsy/skill-pst-to-markdown ~/.claude/skills/pst-to-markdown
```

Or manually:
```bash
git clone https://github.com/GlebTsy/skill-pst-to-markdown.git ~/.claude/skills/pst-to-markdown
```

The skill activates automatically — no configuration needed. Claude Code loads skills from `~/.claude/skills/` at startup.

**Dependency:** `libpst` must be available on the system.
```bash
brew install libpst   # macOS
apt install libpst4   # Debian/Ubuntu
```

## Usage

Once installed, just describe what you want in natural language:

> "У меня есть файл Mail_Storage.pst, вытащи из него все письма"

> "Extract emails from my Outlook archive and convert to markdown"

> "I have a .pst file, convert it to a searchable email base"

Claude will detect the skill, confirm the paths, and run the full pipeline.

### What the skill asks
- Path to the `.pst` file (or detects it if there's only one in the project)
- Output directory (defaults to `mail_markdown/` next to the `.pst`)

### Pipeline steps Claude runs
```bash
# 1. Extract PST → mbox
readpst -r -o pst_raw/ Mail_Storage.pst

# 2. Convert mbox → Markdown threads
python3 pst_merge.py pst_raw/ mail_markdown/

# 3. Remove readpst artifacts (optional)
python3 cleanup_noise.py mail_markdown/
```

## Scale

| PST size | Extraction | Conversion |
|----------|------------|------------|
| ~1 GB    | 1–2 min    | 1–2 min    |
| ~4 GB    | 5–8 min    | 3–5 min    |
| ~10 GB   | 15–20 min  | 8–12 min   |

A 4.5 GB archive with ~7 300 messages produces ~5 700 thread folders and ~7 800 attachment files.

## What is and isn't filtered

| Kept | Filtered |
|------|----------|
| All `Content-Disposition: attachment` files | Corporate signature images (CID inline, signature-only) |
| Inline images referenced in the email body | `rtf-body.rtf` (cleanup step) |
| Every thread, including bulk/automated mail | `oledata.mso`, Outlook icon blobs (cleanup step) |

Automated emails (LinkedIn digests, IT maintenance notices, HR announcements) are kept — sorting and classification is a separate concern.

## Files

| File | Purpose |
|------|---------|
| `SKILL.md` | Skill instructions loaded by Claude Code |
| `scripts/pst_merge.py` | Main conversion script — global threading, signature filtering, attachment extraction |
| `scripts/cleanup_noise.py` | Optional post-processing — removes readpst artifacts |

## Requirements

- macOS or Linux
- `libpst` ≥ 0.6 (provides `readpst`)
- Python 3.8+ (standard library only, no pip packages)
- [Claude Code](https://claude.ai/code)

## License

MIT
