---
name: pst-to-markdown
description: >
  Convert Outlook .pst email archives into an organized Markdown knowledge base
  with threaded conversations and extracted attachments. Use when user has a .pst
  file and wants to extract emails, convert an Outlook archive to markdown, build
  a searchable email base, or process mail_storage.pst. Triggers on: "у меня есть
  pst", "extract emails from pst", "convert outlook archive", "вытащи письма из
  архива". Do NOT use for .eml or .msg files, IMAP/Gmail API access, or importing
  into another email client.
---

# PST → Markdown

Converts an Outlook `.pst` archive into a folder of Markdown thread files with
extracted attachments. Threads are merged globally across Inbox/Sent/work folders;
corporate signature images are filtered out automatically.

Two scripts are bundled in `scripts/`:
- `pst_merge.py` — main conversion (use this)
- `cleanup_noise.py` — optional post-processing to remove readpst artifacts

Note: if you see `pst_to_markdown.py` in a project folder, it is an older
per-folder version superseded by `pst_merge.py`. Do not use it.

## Prerequisites

Check and install `libpst` (provides `readpst`):

```bash
which readpst || brew install libpst
```

Python 3 standard library only — no extra packages needed.

## Step 1 — Locate the PST and set paths

Ask the user if not provided:
- Path to the `.pst` file
- Desired output directory (default: same folder as the `.pst`)

Set variables for the rest of the steps:
```
SKILL_DIR  = ~/.claude/skills/pst-to-markdown
PST_FILE   = /path/to/file.pst
PROJECT    = /path/to/project            # directory containing the .pst
PST_RAW    = $PROJECT/pst_raw            # intermediate mbox extraction
OUTPUT_DIR = $PROJECT/mail_markdown      # final Markdown output
```

## Step 2 — Extract PST → mbox

```bash
mkdir -p "$PST_RAW"
readpst -r -o "$PST_RAW" "$PST_FILE"
```

Flag `-r` preserves the PST folder hierarchy. Each PST folder becomes a directory
containing an `mbox` file. Extraction of a 4 GB PST takes ~5 minutes.

If `pst_raw/` already exists from a previous run, delete it first to avoid
stale data mixing with a fresh extraction:
```bash
rm -rf "$PST_RAW" && mkdir -p "$PST_RAW"
```

Verify success:
```bash
find "$PST_RAW" -name mbox | wc -l   # should be > 0
```

## Step 3 — Convert mbox → Markdown threads

Copy the script from the skill and run it:

```bash
cp "$SKILL_DIR/scripts/pst_merge.py" "$PROJECT/"
python3 "$PROJECT/pst_merge.py" "$PST_RAW" "$OUTPUT_DIR"
```

`pst_merge.py` does the following in one pass:
- Loads all `mbox` files globally (not per-folder)
- Groups messages into threads via `Message-ID` / `In-Reply-To` / `References`
  using a union-find structure across all folders
- Deduplicates messages: same email in Inbox + Sent becomes one copy
- Places each thread into the highest-priority folder
  (deep work folders > Inbox > Sent messages)
- Marks merged threads: `*N message(s) · merged from K folders*`
- Filters corporate signature images: CID-referenced images that appear only
  after the signature block ("Best regards", "Mvh", "С уважением", etc.) are
  skipped; content images in the body are kept
- Deduplicates attachments within each thread by MD5

## Step 4 — Remove readpst artifacts (optional but recommended)

`readpst` produces several artifact files that clutter the output:

| File | Count (typical) | What it is |
|------|-----------------|------------|
| `rtf-body.rtf` | ~400 | RTF copy of email body — already in conversation.md |
| `oledata.mso` | ~5 | Outlook OLE metadata container |
| `ATT*` (≤4608 bytes) | ~35 | Outlook UI icon blobs with no filename |

Run the cleanup script:
```bash
cp "$SKILL_DIR/scripts/cleanup_noise.py" "$PROJECT/"
python3 "$PROJECT/cleanup_noise.py" "$OUTPUT_DIR"
```

Large `ATT*` files (>4608 bytes) are kept — they may be real documents that
lost their filename during export. Review them manually.

## Step 5 — Report results

```bash
echo "Threads:     $(find "$OUTPUT_DIR" -name conversation.md | wc -l)"
echo "Attachments: $(find "$OUTPUT_DIR" -path '*/attachments/*' -type f | wc -l)"
du -sh "$OUTPUT_DIR"
```

## Output structure

```
mail_markdown/
  Personal Folders/
    Work/Analyses/Global/Project Alpha/
      Alpha status report__a1b2c3d4/
        conversation.md       ← full thread, chronological
        attachments/
          report_rev3.pdf
          data_analysis.xlsx
    Inbox/
      Re_ Project kickoff__e5f6a7b8/
        conversation.md
```

Each `conversation.md` starts with:
```
# <subject>

*N message(s) · merged from K folders: `folder1` · `folder2`*

**From:** Name <email>
**To:** ...
**Date:** ...
**Attachments:**
- [filename.pdf](attachments/filename.pdf)

---

<body text>
```

## What pst_merge.py filters

| Kept | Filtered |
|------|----------|
| All attachment-disposition files (PDF, DOCX, XLS, etc.) | CID images appearing **only** in the email signature block |
| Inline images referenced in the body before the signature | — |
| All threads regardless of folder | — |

Automated bulk emails (LinkedIn digests, IT maintenance, End User Digests) are
**kept** — sorting and classification is a separate step.

## Troubleshooting

| Error | Cause | Fix |
|-------|-------|-----|
| `readpst: command not found` | libpst not installed | `brew install libpst` |
| `0 mbox files found` | wrong readpst flag | Use `-r`, not `-S` or `-M` |
| `Cannot open mbox` | path has spaces or non-ASCII | Quote all paths in shell |
| Thread count lower than expected | Unicode PST format issue | `brew upgrade libpst` |
| Garbled subject/body text | Non-UTF-8 encoding | Script falls back to latin-1 automatically |
| Output and pst_raw are same dir | Path derivation error | Ensure PST_RAW ≠ OUTPUT_DIR |

## Notes on scale

| PST size | Extraction (readpst) | Conversion (pst_merge.py) |
|----------|----------------------|---------------------------|
| ~1 GB    | 1–2 min              | 1–2 min                   |
| ~4 GB    | 5–8 min              | 3–5 min                   |
| ~10 GB   | 15–20 min            | 8–12 min                  |

For PSTs larger than 8 GB, run extraction in background:
```bash
nohup readpst -r -o "$PST_RAW" "$PST_FILE" &
```
