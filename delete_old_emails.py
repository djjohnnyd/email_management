"""
delete_old_emails.py
--------------------
Connects to a GoDaddy (Outlook/Microsoft 365) mailbox via IMAP and deletes
emails that meet ALL of the following criteria:
  1. Have at least one attachment
  2. Are older than 10 years

SAFETY: Run with DRY_RUN = True first to preview what would be deleted.
        Set DRY_RUN = False only when you're confident in the results.

Requirements:
    pip install python-dotenv

Setup:
    Create a file named .env in the same folder as this script with:
        EMAIL=you@yourdomain.com
        PASSWORD=your_app_password_or_email_password

    GoDaddy IMAP settings:
        Host: imap.secureserver.net (or outlook.office365.com for M365)
        Port: 993
        SSL:  Yes
"""

import imaplib
import email
from email.header import decode_header
import datetime
import os
import sys

# ---------------------------------------------------------------------------
# CONFIGURATION — edit these or use a .env file
# ---------------------------------------------------------------------------

EMAIL_ADDRESS = os.environ.get("EMAIL", "you@yourdomain.com")
EMAIL_PASSWORD = os.environ.get("PASSWORD", "your_password_here")

# GoDaddy legacy hosting uses imap.secureserver.net
# GoDaddy Microsoft 365 accounts use outlook.office365.com
IMAP_HOST = "imap.secureserver.net"   # Change to outlook.office365.com if needed
IMAP_PORT = 993

# Folders to scan. Use "INBOX" for inbox only, or add more like "Sent"
FOLDERS_TO_SCAN = ["INBOX", "Sent"]

# How many years old an email must be to qualify for deletion
YEARS_OLD = 10

# ⚠️  SAFETY SWITCH: Set to False only when ready to actually delete
DRY_RUN = True

# How many emails to process per folder (None = all). Useful for testing.
LIMIT_PER_FOLDER = None   # e.g. 50 for a test run

# ---------------------------------------------------------------------------


def decode_str(s):
    """Decode an email header string to plain text."""
    if s is None:
        return ""
    parts = decode_header(s)
    decoded = []
    for part, charset in parts:
        if isinstance(part, bytes):
            decoded.append(part.decode(charset or "utf-8", errors="replace"))
        else:
            decoded.append(part)
    return " ".join(decoded)


def has_attachment(msg):
    """Return True if the email message has at least one attachment."""
    if msg.is_multipart():
        for part in msg.walk():
            disposition = part.get("Content-Disposition", "")
            if "attachment" in disposition.lower():
                return True
            # Some attachments have a filename but no explicit disposition
            if part.get_filename():
                return True
    return False


def get_cutoff_date(years=10):
    """Return the cutoff date (today minus N years) formatted for IMAP SEARCH."""
    today = datetime.date.today()
    try:
        cutoff = today.replace(year=today.year - years)
    except ValueError:
        # Feb 29 edge case
        cutoff = today.replace(year=today.year - years, day=28)
    # IMAP BEFORE date format: DD-Mon-YYYY (e.g. 27-May-2015)
    return cutoff.strftime("%d-%b-%Y")


def process_folder(mail, folder, cutoff_date_str, dry_run, limit):
    """Search a folder and delete qualifying emails."""

    status, _ = mail.select(folder, readonly=dry_run)
    if status != "OK":
        print(f"  [!] Could not open folder: {folder}")
        return 0, 0

    # Search for all emails BEFORE the cutoff date
    status, data = mail.search(None, f'(BEFORE "{cutoff_date_str}")')
    if status != "OK":
        print(f"  [!] Search failed in {folder}")
        return 0, 0

    all_ids = data[0].split()
    print(f"  Found {len(all_ids)} emails older than {YEARS_OLD} years in '{folder}'")

    if limit:
        all_ids = all_ids[:limit]
        print(f"  (Limited to first {limit} for this run)")

    checked = 0
    deleted = 0

    for msg_id in all_ids:
        # Fetch the full message to inspect for attachments
        status, msg_data = mail.fetch(msg_id, "(RFC822)")
        if status != "OK":
            continue

        raw = msg_data[0][1]
        msg = email.message_from_bytes(raw)
        checked += 1

        if not has_attachment(msg):
            continue

        # This email matches both criteria
        subject = decode_str(msg.get("Subject", "(no subject)"))
        date_str = msg.get("Date", "(unknown date)")
        sender = decode_str(msg.get("From", "(unknown sender)"))

        deleted += 1
        if dry_run:
            print(f"  [DRY RUN] Would delete: | {date_str[:22]} | From: {sender[:40]} | {subject[:60]}")
        else:
            mail.store(msg_id, "+FLAGS", "\\Deleted")
            print(f"  [DELETED] | {date_str[:22]} | From: {sender[:40]} | {subject[:60]}")

    # Permanently expunge deleted messages
    if not dry_run and deleted > 0:
        mail.expunge()

    return checked, deleted


def main():
    # Try to load .env if python-dotenv is available
    try:
        from dotenv import load_dotenv
        load_dotenv()
        global EMAIL_ADDRESS, EMAIL_PASSWORD
        EMAIL_ADDRESS = os.environ.get("EMAIL", EMAIL_ADDRESS)
        EMAIL_PASSWORD = os.environ.get("PASSWORD", EMAIL_PASSWORD)
    except ImportError:
        pass  # dotenv not installed; fall back to hardcoded values

    if "your_password_here" in EMAIL_PASSWORD or "yourdomain.com" in EMAIL_ADDRESS:
        print("ERROR: Please set your EMAIL and PASSWORD in this script or in a .env file.")
        sys.exit(1)

    cutoff_date_str = get_cutoff_date(YEARS_OLD)
    mode = "DRY RUN (no emails will be deleted)" if DRY_RUN else "⚠️  LIVE RUN — emails WILL be permanently deleted"

    print("=" * 70)
    print(f"  Email:    {EMAIL_ADDRESS}")
    print(f"  Server:   {IMAP_HOST}:{IMAP_PORT}")
    print(f"  Cutoff:   Emails before {cutoff_date_str} ({YEARS_OLD} years ago)")
    print(f"  Criteria: Has attachment AND older than {YEARS_OLD} years")
    print(f"  Mode:     {mode}")
    print("=" * 70)

    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        print(f"\nLogged in successfully.\n")
    except imaplib.IMAP4.error as e:
        print(f"Login failed: {e}")
        print("\nTroubleshooting tips:")
        print("  - For GoDaddy legacy hosting: host should be imap.secureserver.net")
        print("  - For GoDaddy M365:           host should be outlook.office365.com")
        print("  - If using M365, you may need an App Password (see README below)")
        sys.exit(1)

    total_checked = 0
    total_deleted = 0

    for folder in FOLDERS_TO_SCAN:
        print(f"\nScanning folder: {folder}")
        checked, deleted = process_folder(
            mail, folder, cutoff_date_str, DRY_RUN, LIMIT_PER_FOLDER
        )
        total_checked += checked
        total_deleted += deleted
        print(f"  → Checked: {checked} | {'Would delete' if DRY_RUN else 'Deleted'}: {deleted}")

    mail.logout()

    print("\n" + "=" * 70)
    print(f"  SUMMARY")
    print(f"  Total emails checked:  {total_checked}")
    action = "Would be deleted" if DRY_RUN else "Deleted"
    print(f"  {action}: {total_deleted}")
    if DRY_RUN:
        print("\n  ✓ This was a DRY RUN. Set DRY_RUN = False to perform actual deletion.")
    else:
        print("\n  ✓ Deletion complete. Emails have been permanently removed.")
    print("=" * 70)


if __name__ == "__main__":
    main()
