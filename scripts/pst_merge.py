#!/usr/bin/env python3
"""
Global thread merger: reads all mbox files, groups messages by Message-ID chain
across ALL folders, deduplicates attachments by MD5, writes clean Markdown output.

Usage: python3 pst_merge.py <pst_raw_dir> <output_dir>
"""

import sys, os, re, email, email.header, hashlib, html
from pathlib import Path
from collections import defaultdict
from datetime import datetime
from email.utils import parsedate_to_datetime, getaddresses
import mailbox

# Markers that start an email signature block (case-insensitive).
# We look for the FIRST occurrence before any quoted-reply block.
_SIG_MARKERS = [
    r"best\s+regards", r"kind\s+regards", r"warm\s+regards",
    r"\bregards\b", r"sincerely", r"yours\s+(?:sincerely|faithfully|truly)",
    r"med\s+vennlig\s+hilsen", r"\bmvh\b",           # Norwegian
    r"с\s+уважением", r"с\s+наилучшими\s+пожеланиями",  # Russian
    r"vänliga\s+hälsningar",                           # Swedish
    r"mit\s+freundlichen\s+grüßen",                    # German
    r"cordialement",                                    # French
    r"(?:^|\n)--\s*\n",                                # RFC 3676 sig delimiter
]
_SIG_RE = re.compile("|".join(_SIG_MARKERS), re.IGNORECASE)

# Marks the start of a quoted previous message — we stop looking after this.
_QUOTE_RE = re.compile(
    r"(?:-----\s*Original Message\s*-----|"
    r"<blockquote|"
    r"On .{10,80} wrote:|"
    r"От:.{1,80}Дата:)",
    re.IGNORECASE | re.DOTALL,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def decode_header(value):
    if not value:
        return ""
    parts = email.header.decode_header(value)
    decoded = []
    for part, charset in parts:
        if isinstance(part, bytes):
            try:
                decoded.append(part.decode(charset or "utf-8", errors="replace"))
            except (LookupError, UnicodeDecodeError):
                decoded.append(part.decode("latin-1", errors="replace"))
        else:
            decoded.append(str(part))
    return "".join(decoded)


def safe_filename(name, maxlen=80):
    name = re.sub(r"[^\w\s\-\.]", "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name[:maxlen] if name else "unnamed"


def html_to_text(raw_html):
    text = re.sub(r"<br\s*/?>", "\n", raw_html, flags=re.IGNORECASE)
    text = re.sub(r"<p[^>]*>", "\n\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</p>", "", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    text = html.unescape(text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def get_body(msg):
    plain = None
    html_body = None
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            cd = str(part.get("Content-Disposition", ""))
            if "attachment" in cd:
                continue
            charset = part.get_content_charset() or "utf-8"
            payload = part.get_payload(decode=True)
            if not payload:
                continue
            try:
                text = payload.decode(charset, errors="replace")
            except (LookupError, UnicodeDecodeError):
                text = payload.decode("latin-1", errors="replace")
            if ct == "text/plain" and plain is None:
                plain = text
            elif ct == "text/html" and html_body is None:
                html_body = html_to_text(text)
    else:
        charset = msg.get_content_charset() or "utf-8"
        ct = msg.get_content_type()
        payload = msg.get_payload(decode=True)
        if payload:
            try:
                text = payload.decode(charset, errors="replace")
            except (LookupError, UnicodeDecodeError):
                text = payload.decode("latin-1", errors="replace")
            if ct == "text/plain":
                plain = text
            elif ct == "text/html":
                html_body = html_to_text(text)
    return plain or html_body or ""


def _signature_cids(msg) -> set:
    """
    Return the set of Content-IDs referenced ONLY inside the signature block
    of the HTML body. Images used before the signature (charts, screenshots)
    are NOT in this set and will be kept.
    """
    html_body = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/html":
                cs = part.get_content_charset() or "utf-8"
                raw = part.get_payload(decode=True)
                if raw:
                    try:
                        html_body = raw.decode(cs, errors="replace")
                    except (LookupError, UnicodeDecodeError):
                        html_body = raw.decode("latin-1", errors="replace")
                break
    if not html_body:
        return set()

    # Restrict search to the new message only (before any quoted block)
    quote_m = _QUOTE_RE.search(html_body)
    search_space = html_body[: quote_m.start()] if quote_m else html_body

    sig_m = _SIG_RE.search(search_space)
    if not sig_m:
        return set()

    before_sig = search_space[: sig_m.start()]
    after_sig  = search_space[sig_m.start():]

    body_cids = set(re.findall(r"cid:([^\"'>\s]+)", before_sig, re.IGNORECASE))
    sig_cids  = set(re.findall(r"cid:([^\"'>\s]+)", after_sig,  re.IGNORECASE))

    return sig_cids - body_cids  # exclude CIDs also used in the body


def get_attachments(msg):
    """
    Extract attached and inline files.
    Parts that have a Content-ID referenced ONLY inside the email signature
    block (logos, contact icons) are skipped. Content-ID images used in the
    body (charts, screenshots) and all regular attachments are kept.
    Outlook labels CID-referenced signature images as 'attachment', so we
    filter by CID presence, not by Content-Disposition value.
    """
    sig_cids = _signature_cids(msg)

    result = []
    if msg.is_multipart():
        for part in msg.walk():
            cd = str(part.get("Content-Disposition", ""))
            if "attachment" not in cd and "inline" not in cd:
                continue
            filename = decode_header(part.get_filename() or "")
            if not filename:
                continue
            # If the part has a Content-ID, check whether it's signature-only
            raw_cid = part.get("Content-ID", "").strip().strip("<>")
            if raw_cid and raw_cid in sig_cids:
                continue  # signature decoration — skip
            payload = part.get_payload(decode=True)
            if payload:
                result.append((filename, payload))
    return result


def parse_date(msg_dict):
    try:
        return parsedate_to_datetime(msg_dict["date_raw"])
    except Exception:
        return datetime(1970, 1, 1)


def format_addresses(raw):
    if not raw:
        return ""
    try:
        pairs = getaddresses([decode_header(raw)])
    except Exception:
        return decode_header(raw)
    parts = []
    for name, addr in pairs:
        parts.append(f"{name} <{addr}>" if name else addr)
    return ", ".join(parts)


# ---------------------------------------------------------------------------
# Union-Find for thread grouping
# ---------------------------------------------------------------------------

class UnionFind:
    def __init__(self):
        self.parent = {}

    def find(self, x):
        self.parent.setdefault(x, x)
        if self.parent[x] != x:
            self.parent[x] = self.find(self.parent[x])
        return self.parent[x]

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.parent[rb] = ra


# ---------------------------------------------------------------------------
# Folder priority: the deeper / more specific the work folder, the higher
# ---------------------------------------------------------------------------

FOLDER_PRIORITY = {
    "sent messages": 0,
    "inbox": 1,
    "copies from printer": 1,
    "courses notifications": 1,
    "quality": 2,
    "hr": 2,
    "registration info": 2,
    "travelling": 2,
    "personal": 2,
    "courses": 2,
    "new knowledge": 3,
    "discussions": 3,
    "work": 4,
}

def folder_score(folder_str):
    """Higher = more preferred as the canonical folder for a merged thread."""
    parts = [p.lower() for p in Path(folder_str).parts]
    # deeper work folders score higher
    base = 0
    for p in parts:
        base = max(base, FOLDER_PRIORITY.get(p, 5 if p not in ("personal folders",) else 0))
    return base + len(parts) * 0.1


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def load_all_messages(pst_raw: Path):
    """Load every message from every mbox, return list of dicts."""
    messages = []
    mboxes = sorted(pst_raw.rglob("mbox"))
    print(f"Loading {len(mboxes)} mbox files...")

    for mbox_path in mboxes:
        folder_rel = str(mbox_path.parent.relative_to(pst_raw))
        try:
            mb = mailbox.mbox(str(mbox_path))
        except Exception as e:
            print(f"  [SKIP] {mbox_path}: {e}")
            continue

        for key in mb.keys():
            try:
                raw = mb.get_bytes(key)
                msg = email.message_from_bytes(raw)
            except Exception:
                continue

            mid = (msg.get("Message-ID") or "").strip()
            irt = (msg.get("In-Reply-To") or "").strip()
            refs = (msg.get("References") or "").strip()
            subject = decode_header(msg.get("Subject", ""))
            subject_norm = re.sub(r"^(RE_?:?\s*|FW_?:?\s*|Fwd_?:?\s*)+", "", subject, flags=re.IGNORECASE).strip()

            messages.append({
                "mid": mid,
                "irt": irt,
                "refs": refs.split() if refs else [],
                "subject": subject,
                "subject_norm": re.sub(r"\s+", " ", subject_norm).strip(),
                "date_raw": msg.get("Date", ""),
                "from_raw": msg.get("From", ""),
                "to_raw": msg.get("To", ""),
                "cc_raw": msg.get("Cc", ""),
                "folder": folder_rel,
                "msg_obj": msg,
            })

        mb.close()

    print(f"Loaded {len(messages):,} messages total")
    return messages


def build_threads(messages):
    """Group messages into threads using union-find on Message-ID chains."""
    uf = UnionFind()
    mid_to_idx = {}

    # Register all Message-IDs
    for i, m in enumerate(messages):
        if m["mid"]:
            mid_to_idx.setdefault(m["mid"], i)

    # Union via In-Reply-To and References
    for i, m in enumerate(messages):
        rep_id = m["mid"] if m["mid"] else f"__no_mid_{i}__"
        uf.find(rep_id)

        if m["irt"] and m["irt"] in mid_to_idx:
            uf.union(rep_id, m["irt"])
        for ref in m["refs"]:
            if ref in mid_to_idx:
                uf.union(rep_id, ref)

    # Group by root
    groups = defaultdict(list)
    for i, m in enumerate(messages):
        rep_id = m["mid"] if m["mid"] else f"__no_mid_{i}__"
        root = uf.find(rep_id)
        groups[root].append(i)

    return groups


def write_thread(root_id: str, indices, messages, output_root: Path, att_registry: dict):
    """
    Write one merged thread as a conversation.md + attachments/.
    att_registry: md5 → canonical Path (shared across all threads, dedup)
    """
    msgs = [messages[i] for i in indices]
    msgs.sort(key=parse_date)

    # Deduplicate by Message-ID (same email in Inbox + Sent = one copy)
    seen_mids = set()
    unique_msgs = []
    for m in msgs:
        key = m["mid"] if m["mid"] else id(m)
        if key not in seen_mids:
            seen_mids.add(key)
            unique_msgs.append(m)

    if not unique_msgs:
        return

    # Pick canonical folder: highest scoring folder
    best_folder = max(
        (m["folder"] for m in unique_msgs),
        key=folder_score
    )

    subject = unique_msgs[0]["subject_norm"] or unique_msgs[0]["subject"] or "no_subject"
    subject_file = safe_filename(subject, maxlen=60)
    # Use root_id (guaranteed unique per thread group) for the hash
    mid_hash = hashlib.md5(root_id.encode("utf-8")).hexdigest()[:8]

    base_dir = output_root / best_folder / f"{subject_file}__{mid_hash}"
    # Guard against path-length collisions on filesystems (rare but possible)
    thread_dir = base_dir
    counter = 1
    while thread_dir.exists() and (thread_dir / "conversation.md").exists():
        thread_dir = Path(str(base_dir) + f"_{counter}")
        counter += 1
    att_dir = thread_dir / "attachments"
    thread_dir.mkdir(parents=True, exist_ok=True)

    # Collect sources for header note
    folders_used = sorted({m["folder"] for m in unique_msgs})

    blocks = []
    for m in unique_msgs:
        msg_obj = m["msg_obj"]
        lines = []
        lines.append(f"**From:** {format_addresses(m['from_raw'])}")
        lines.append(f"**To:** {format_addresses(m['to_raw'])}")
        cc = format_addresses(m["cc_raw"])
        if cc:
            lines.append(f"**Cc:** {cc}")
        lines.append(f"**Date:** {m['date_raw']}")
        lines.append(f"**Subject:** {decode_header(msg_obj.get('Subject', ''))}")
        lines.append(f"**Folder:** `{m['folder']}`")

        # Attachments
        atts = get_attachments(msg_obj)
        if atts:
            lines.append("")
            lines.append("**Attachments:**")
            for orig_name, data in atts:
                md5 = hashlib.md5(data).hexdigest()
                if md5 in att_registry:
                    # Reuse existing canonical file — compute relative path from thread_dir
                    canonical = att_registry[md5]
                    try:
                        rel = canonical.relative_to(thread_dir)
                    except ValueError:
                        # different thread's attachments dir — copy locally anyway
                        safe = safe_filename(orig_name)
                        dest = att_dir / safe
                        att_dir.mkdir(parents=True, exist_ok=True)
                        counter = 1
                        stem, ext = os.path.splitext(safe)
                        while dest.exists() and dest.read_bytes() != data:
                            safe = f"{stem}_{counter}{ext}"
                            dest = att_dir / safe
                            counter += 1
                        if not dest.exists():
                            dest.write_bytes(data)
                        att_registry[md5] = dest
                        rel = Path("attachments") / safe
                    lines.append(f"- [{orig_name}]({rel})")
                else:
                    safe = safe_filename(orig_name)
                    dest = att_dir / safe
                    att_dir.mkdir(parents=True, exist_ok=True)
                    counter = 1
                    stem, ext = os.path.splitext(safe)
                    while dest.exists() and dest.read_bytes() != data:
                        safe = f"{stem}_{counter}{ext}"
                        dest = att_dir / safe
                        counter += 1
                    if not dest.exists():
                        dest.write_bytes(data)
                    att_registry[md5] = dest
                    lines.append(f"- [{orig_name}](attachments/{safe})")

        lines.append("")
        lines.append("---")
        lines.append("")
        body = get_body(msg_obj)
        lines.append(body if body else "*(no body)*")
        blocks.append("\n".join(lines))

    # Header
    sources_str = " · ".join(f"`{f}`" for f in folders_used) if len(folders_used) > 1 else f"`{folders_used[0]}`"
    md = f"# {subject}\n\n"
    md += f"*{len(unique_msgs)} message(s)*"
    if len(folders_used) > 1:
        md += f" · merged from {len(folders_used)} folders: {sources_str}"
    md += "\n\n"
    md += "\n\n---\n\n".join(blocks)

    (thread_dir / "conversation.md").write_text(md, encoding="utf-8")


def main():
    if len(sys.argv) != 3:
        print("Usage: python3 pst_merge.py <pst_raw_dir> <output_dir>")
        sys.exit(1)

    pst_raw = Path(sys.argv[1]).resolve()
    output = Path(sys.argv[2]).resolve()
    output.mkdir(parents=True, exist_ok=True)

    messages = load_all_messages(pst_raw)
    print("Building thread groups...")
    groups = build_threads(messages)
    print(f"Thread groups: {len(groups):,}")

    att_registry: dict = {}  # md5 → canonical Path
    written = 0
    skipped = 0

    for root_id, indices in groups.items():
        try:
            write_thread(root_id, indices, messages, output, att_registry)
            written += 1
        except Exception as e:
            skipped += 1
            if skipped <= 5:
                print(f"  [ERR] thread {root_id[:30]}: {e}")

    att_unique = len(att_registry)
    att_total = sum(
        1 for p in output.rglob("*")
        if p.is_file() and "attachments" in p.parts
    )
    print(f"\nDone.")
    print(f"  Threads written : {written:,}")
    print(f"  Threads skipped : {skipped}")
    print(f"  Unique attachments: {att_unique:,}")
    print(f"  Output: {output}")


if __name__ == "__main__":
    main()
