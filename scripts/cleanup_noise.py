#!/usr/bin/env python3
"""
Remove known readpst/Outlook artifacts from a mail_markdown output directory.

Removed:
  - rtf-body.rtf  : readpst encodes RTF email bodies as attachments; body text
                    is already captured in conversation.md
  - oledata.mso   : Outlook OLE container metadata, not useful content
  - ATT* (≤4608b) : Outlook UI icon blobs with no filename (4 KB OLE objects)

Usage: python3 cleanup_noise.py <mail_markdown_dir>
"""

import sys
from pathlib import Path

EXACT_NAMES = {"rtf-body.rtf", "oledata.mso"}
ATT_MAX_JUNK_SIZE = 4608  # bytes — Outlook icon blobs are always exactly this size


def main():
    if len(sys.argv) != 2:
        print("Usage: python3 cleanup_noise.py <mail_markdown_dir>")
        sys.exit(1)

    root = Path(sys.argv[1]).resolve()
    if not root.exists():
        print(f"Directory not found: {root}")
        sys.exit(1)

    removed = []
    kept_att = []

    for f in root.rglob("*"):
        if not f.is_file() or "attachments" not in f.parts:
            continue

        name = f.name.lower()
        size = f.stat().st_size

        if name in EXACT_NAMES:
            f.unlink()
            removed.append(f)
        elif f.name.upper().startswith("ATT") and not f.suffix and size <= ATT_MAX_JUNK_SIZE:
            f.unlink()
            removed.append(f)
        elif f.name.upper().startswith("ATT") and not f.suffix:
            kept_att.append((size, f))  # large ATT* kept — may be real doc

    total_size = sum(0 for _ in removed)  # recount after removal

    print(f"Removed {len(removed)} noise files")
    if kept_att:
        print(f"Kept {len(kept_att)} large ATT* files (possible documents without filename):")
        for size, p in sorted(kept_att, reverse=True)[:10]:
            print(f"  {size // 1024:4} KB  {p.relative_to(root)}")

    # Remove now-empty attachments/ directories
    emptied = 0
    for att_dir in root.rglob("attachments"):
        if att_dir.is_dir() and not any(att_dir.iterdir()):
            att_dir.rmdir()
            emptied += 1

    if emptied:
        print(f"Removed {emptied} empty attachments/ directories")


if __name__ == "__main__":
    main()
