"""Extract attachments from a specific Outlook email to a folder.

Works for normal mail AND meeting invites (Class 53) / appointments (Class 26),
which `email_search` / the indexer skip (they only index olMail, Class 43). Uses raw
win32com (GetItemFromID + item.Attachments) so it never depends on _parse_email, which
assumes a MailItem and fails on invites.

Usage:
    # By EntryID (back-compatible positional form still works):
    python extract_one_email_attachments.py <entry_id> <out_dir>
    python extract_one_email_attachments.py --entry-id <id> --out <dir>

    # Or find the email by subject / sender substring (case-insensitive), scanning
    # Inbox + Sent Items including meeting invites:
    python extract_one_email_attachments.py --subject "Cal Poly - Building C" --out <dir>
    python extract_one_email_attachments.py --sender "Edmund" --subject "Cal Poly" --out <dir>

    # Skip tiny inline signature logos (e.g. image001.png):
    python extract_one_email_attachments.py --subject "..." --out <dir> --min-size 15000
"""
import argparse
import os
import sys

import win32com.client

# Outlook default-folder constants
OL_FOLDER_INBOX = 6
OL_FOLDER_SENT = 5


def _find_emails(ns, subject=None, sender=None, scan_limit=400):
    """Scan Inbox + Sent Items (all item classes, incl. meeting invites) for items whose
    subject/sender contain the given substrings. Returns [(entry_id, subject, sender, n_att)]."""
    subj_q = (subject or "").lower()
    send_q = (sender or "").lower()
    matches = []
    for folder_id in (OL_FOLDER_INBOX, OL_FOLDER_SENT):
        try:
            folder = ns.GetDefaultFolder(folder_id)
            items = folder.Items
            items.Sort("[ReceivedTime]", True)  # newest first
        except Exception:
            continue
        for i, item in enumerate(items):
            if i >= scan_limit:
                break
            try:
                subj = getattr(item, "Subject", "") or ""
                # SenderName for mail; Organizer for meeting invites
                snd = getattr(item, "SenderName", "") or getattr(item, "Organizer", "") or ""
                if subj_q and subj_q not in subj.lower():
                    continue
                if send_q and send_q not in snd.lower():
                    continue
                try:
                    n_att = item.Attachments.Count
                except Exception:
                    n_att = 0
                matches.append((item.EntryID, subj, snd, n_att))
            except Exception:
                continue
    return matches


def _save_attachments(ns, entry_id, out_dir, min_size=0):
    """Save all attachments of the item (any class) to out_dir. Returns saved paths."""
    item = ns.GetItemFromID(entry_id)  # works for mail, meeting invites, appointments
    os.makedirs(out_dir, exist_ok=True)
    saved = []
    count = item.Attachments.Count
    print(f"Found {count} attachment(s)")
    for ai in range(1, count + 1):  # Outlook Attachments are 1-based
        att = item.Attachments.Item(ai)
        filename = att.FileName
        try:
            size = att.Size
        except Exception:
            size = -1
        if min_size and 0 <= size < min_size:
            print(f"  [skip ] {filename} ({size} bytes < min-size {min_size})")
            continue
        dest = os.path.join(out_dir, filename)
        base, ext = os.path.splitext(filename)
        c = 1
        while os.path.exists(dest):
            dest = os.path.join(out_dir, f"{base}_{c}{ext}")
            c += 1
        try:
            att.SaveAsFile(dest)
            saved.append(dest)
            print(f"  [saved] {filename} ({size} bytes) -> {dest}")
        except Exception as e:
            print(f"  [FAIL ] {filename}: {e}")
    return saved


def main():
    parser = argparse.ArgumentParser(
        description="Extract Outlook email attachments (handles meeting invites)."
    )
    parser.add_argument("--entry-id", help="Outlook EntryID of the email")
    parser.add_argument("--subject", help="Subject substring to search for")
    parser.add_argument("--sender", help="Sender/Organizer substring to search for")
    parser.add_argument("--out", help="Output directory")
    parser.add_argument("--min-size", type=int, default=0,
                        help="Skip attachments smaller than N bytes (filters inline signature logos)")
    parser.add_argument("pos", nargs="*", help="Back-compat positional form: <entry_id> <out_dir>")
    args = parser.parse_args()

    # Back-compat: positional <entry_id> <out_dir>
    entry_id = args.entry_id
    out_dir = args.out
    if args.pos:
        if not entry_id and len(args.pos) >= 1:
            entry_id = args.pos[0]
        if not out_dir and len(args.pos) >= 2:
            out_dir = args.pos[1]

    if not out_dir:
        sys.exit("ERROR: output directory required (--out <dir> or positional <out_dir>)")

    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    if not entry_id:
        if not (args.subject or args.sender):
            sys.exit("ERROR: provide --entry-id, or --subject/--sender to search")
        matches = _find_emails(ns, subject=args.subject, sender=args.sender)
        with_att = [m for m in matches if m[3] > 0]
        chosen = with_att or matches
        if not chosen:
            sys.exit(f"No emails matched subject={args.subject!r} sender={args.sender!r}")
        print(f"Matched {len(matches)} email(s), {len(with_att)} with attachments:")
        for eid, subj, snd, n in chosen[:10]:
            print(f"  - {subj!r} from {snd!r} ({n} attachments)")
        entry_id, subj, snd, n = chosen[0]
        print(f"\nExtracting from: {subj!r} from {snd!r} ({n} attachments)")

    saved = _save_attachments(ns, entry_id, out_dir, min_size=args.min_size)
    print(f"\nSaved {len(saved)} file(s) to {out_dir}")


if __name__ == "__main__":
    main()
