"""
delete_old_emails.py  (OAuth2 version for Microsoft 365 / GoDaddy M365)
------------------------------------------------------------------------
Connects via IMAP using OAuth2 (no plain password needed) and deletes
emails that meet ALL of the following criteria:
  1. Have at least one attachment
  2. Are older than 10 years (configurable)

SAFETY: Run with DRY_RUN = True first to preview what would be deleted.

Requirements:
    pip install msal python-dotenv

Setup:
    1. Register an Azure App (see README in comments below).
    2. Create a .env file with CLIENT_ID, TENANT_ID, and EMAIL.
       Or just fill in the config section below directly.
"""

import imaplib
import email
import base64
import datetime
import os
import sys
from email.header import decode_header

# ---------------------------------------------------------------------------
# CONFIGURATION — fill these in or put them in a .env file
# ---------------------------------------------------------------------------

EMAIL_ADDRESS = os.environ.get("EMAIL", "you@yourdomain.com")

# From your Azure App Registration:
CLIENT_ID  = os.environ.get("CLIENT_ID",  "your-client-id-here")
TENANT_ID  = os.environ.get("TENANT_ID",  "your-tenant-id-here")

IMAP_HOST = "outlook.office365.com"
IMAP_PORT = 993

# Folders to scan (use exact names; common ones listed)
FOLDERS_TO_SCAN = ["INBOX", "Sent Items"]

# Delete emails older than this many years
YEARS_OLD = 10

# ⚠️  SAFETY: Set to False only when ready to actually delete
DRY_RUN = True

# Limit per folder for test runs (None = unlimited)
LIMIT_PER_FOLDER = None   # e.g. 50 for a test run

# ---------------------------------------------------------------------------


def get_oauth2_token(client_id, tenant_id, email_address):
    """Interactively authenticate with Microsoft and return an access token."""
    try:
        import msal
    except ImportError:
        print("ERROR: msal not installed. Run:  pip install msal")
        sys.exit(1)

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scopes = ["https://outlook.office.com/IMAP.AccessAsUser.All", "offline_access"]

    app = msal.PublicClientApplication(client_id, authority=authority)

    # Try silent (cached) token first
    accounts = app.get_accounts(username=email_address)
    result = None
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])

    if not result:
        # Device code flow — works over SSH with no display
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            print(f"Failed to create device flow: {flow}")
            sys.exit(1)
        print("\n" + "=" * 60)
        print("  ACTION REQUIRED — Sign in on your local machine:")
        print(f"  1. Open this URL in your browser: {flow['verification_uri']}")
        print(f"  2. Enter this code:               {flow['user_code']}")
        print("=" * 60 + "\n")
        # Waits here until you complete login in the browser (up to 15 min)
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        print(f"Authentication failed: {result.get('error_description', result)}")
        sys.exit(1)

    print("Authentication successful.\n")
    return result["access_token"]


def build_xoauth2_string(email_address, access_token):
    """Build the XOAUTH2 authentication string for IMAP."""
    auth_string = f"user={email_address}\x01auth=Bearer {access_token}\x01\x01"
    return base64.b64encode(auth_string.encode()).decode()


def decode_str(s):
    if s is None:
        return ""
    parts = decode_header(s)
    decoded = []
    for part, charset in parts:
        if isinstance(part, bytes):
            decoded.append(part.decode(charset or "utf-8", errors="replace"))
        else:
            decoded.append(str(part))
    return " ".join(decoded)


def has_attachment(msg):
    if msg.is_multipart():
        for part in msg.walk():
            disposition = str(part.get("Content-Disposition", ""))
            if "attachment" in disposition.lower():
                return True
            if part.get_filename():
                return True
    return False


def get_cutoff_date_str(years=10):
    today = datetime.date.today()
    try:
        cutoff = today.replace(year=today.year - years)
    except ValueError:
        cutoff = today.replace(year=today.year - years, day=28)
    return cutoff.strftime("%d-%b-%Y")


def process_folder(mail, folder, cutoff_date_str, dry_run, limit):
    try:
        status, _ = mail.select(f'"{folder}"', readonly=dry_run)
    except Exception:
        status, _ = mail.select(folder, readonly=dry_run)

    if status != "OK":
        print(f"  [!] Could not open folder: {folder} — skipping")
        return 0, 0

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
        status, msg_data = mail.fetch(msg_id, "(RFC822)")
        if status != "OK":
            continue

        raw = msg_data[0][1]
        msg = email.message_from_bytes(raw)
        checked += 1

        if not has_attachment(msg):
            continue

        subject  = decode_str(msg.get("Subject", "(no subject)"))
        date_str = msg.get("Date", "(unknown date)")
        sender   = decode_str(msg.get("From", "(unknown)"))
        deleted += 1

        if dry_run:
            print(f"  [DRY RUN] Would delete | {date_str[:22]} | {sender[:35]} | {subject[:55]}")
        else:
            mail.store(msg_id, "+FLAGS", "\\Deleted")
            print(f"  [DELETED]              | {date_str[:22]} | {sender[:35]} | {subject[:55]}")

    if not dry_run and deleted > 0:
        mail.expunge()

    return checked, deleted


def main():
    # Load .env if available
    try:
        from dotenv import load_dotenv
        load_dotenv()
        global EMAIL_ADDRESS, CLIENT_ID, TENANT_ID
        EMAIL_ADDRESS = os.environ.get("EMAIL", EMAIL_ADDRESS)
        CLIENT_ID     = os.environ.get("CLIENT_ID", CLIENT_ID)
        TENANT_ID     = os.environ.get("TENANT_ID", TENANT_ID)
    except ImportError:
        pass

    # Validate config
    missing = []
    if "yourdomain.com" in EMAIL_ADDRESS or not EMAIL_ADDRESS:
        missing.append("EMAIL")
    if "your-client-id" in CLIENT_ID or not CLIENT_ID:
        missing.append("CLIENT_ID")
    if "your-tenant-id" in TENANT_ID or not TENANT_ID:
        missing.append("TENANT_ID")
    if missing:
        print(f"ERROR: Please set the following in your .env or in the script: {', '.join(missing)}")
        sys.exit(1)

    cutoff_date_str = get_cutoff_date_str(YEARS_OLD)
    mode = "DRY RUN (nothing will be deleted)" if DRY_RUN else "⚠️  LIVE — emails WILL be permanently deleted"

    print("=" * 70)
    print(f"  Email:   {EMAIL_ADDRESS}")
    print(f"  Cutoff:  Before {cutoff_date_str} ({YEARS_OLD} years ago)")
    print(f"  Folders: {', '.join(FOLDERS_TO_SCAN)}")
    print(f"  Mode:    {mode}")
    print("=" * 70)

    # Get OAuth2 token (opens browser if no cached token)
    access_token = get_oauth2_token(CLIENT_ID, TENANT_ID, EMAIL_ADDRESS)
    xoauth2_str  = build_xoauth2_string(EMAIL_ADDRESS, access_token)

    # Connect via IMAP
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.authenticate("XOAUTH2", lambda x: xoauth2_str)
        print("Connected to IMAP server.\n")
    except imaplib.IMAP4.error as e:
        print(f"\nIMAP connection/auth failed: {e}")
        print("\nThings to check:")
        print("  1. IMAP is enabled for your mailbox in Outlook settings")
        print("  2. Your Azure app has the IMAP.AccessAsUser.All permission")
        print("  3. Admin consent was granted for that permission")
        sys.exit(1)

    total_checked = 0
    total_deleted = 0

    for folder in FOLDERS_TO_SCAN:
        print(f"Scanning: {folder}")
        checked, deleted = process_folder(
            mail, folder, cutoff_date_str, DRY_RUN, LIMIT_PER_FOLDER
        )
        total_checked += checked
        total_deleted += deleted
        action = "Would delete" if DRY_RUN else "Deleted"
        print(f"  → Checked: {checked}  |  {action}: {deleted}\n")

    mail.logout()

    print("=" * 70)
    print(f"  Total checked:  {total_checked}")
    action = "Would be deleted" if DRY_RUN else "Permanently deleted"
    print(f"  {action}: {total_deleted}")
    if DRY_RUN:
        print("\n  ✓ DRY RUN complete. Set DRY_RUN = False to actually delete.")
    print("=" * 70)


if __name__ == "__main__":
    main()

