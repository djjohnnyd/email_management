"""
debug_imap_auth.py
------------------
Tests IMAP OAuth2 authentication in isolation and prints diagnostic info.
Run this to diagnose connection issues before using the main script.

Usage:
    python3 debug_imap_auth.py
"""

import imaplib
import base64
import os
import sys
import json

EMAIL_ADDRESS = os.environ.get("EMAIL", "you@yourdomain.com")
CLIENT_ID     = os.environ.get("CLIENT_ID", "your-client-id-here")
TENANT_ID     = os.environ.get("TENANT_ID", "your-tenant-id-here")
IMAP_HOST     = "outlook.office365.com"
IMAP_PORT     = 993

def get_token():
    try:
        import msal
    except ImportError:
        print("ERROR: pip install msal")
        sys.exit(1)

    try:
        from dotenv import load_dotenv
        load_dotenv()
        global EMAIL_ADDRESS, CLIENT_ID, TENANT_ID
        EMAIL_ADDRESS = os.environ.get("EMAIL", EMAIL_ADDRESS)
        CLIENT_ID     = os.environ.get("CLIENT_ID", CLIENT_ID)
        TENANT_ID     = os.environ.get("TENANT_ID", TENANT_ID)
    except ImportError:
        pass

    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}"
    )
    scopes = ["https://outlook.office.com/IMAP.AccessAsUser.All"]

    accounts = app.get_accounts(username=EMAIL_ADDRESS)
    result = None
    if accounts:
        print("Found cached account, trying silent token...")
        result = app.acquire_token_silent(scopes, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=scopes)
        print(f"\n  1. Open: {flow['verification_uri']}")
        print(f"  2. Code: {flow['user_code']}\n")
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        print(f"Token error: {result.get('error_description')}")
        sys.exit(1)

    return result["access_token"]

def decode_token_claims(token):
    """Decode JWT payload (no verification) to inspect claims."""
    try:
        parts = token.split(".")
        if len(parts) != 3:
            return None
        padding = 4 - len(parts[1]) % 4
        payload = base64.urlsafe_b64decode(parts[1] + "=" * padding)
        return json.loads(payload)
    except Exception:
        return None

def main():
    print("\n=== STEP 1: Getting OAuth2 token ===")
    token = get_token()
    print("Token obtained successfully.")

    print("\n=== STEP 2: Inspecting token claims ===")
    claims = decode_token_claims(token)
    if claims:
        print(f"  Audience (aud): {claims.get('aud', 'NOT FOUND')}")
        print(f"  Scope   (scp):  {claims.get('scp', 'NOT FOUND')}")
        print(f"  UPN     (upn):  {claims.get('upn', 'NOT FOUND')}")
        print(f"  App ID  (appid):{claims.get('appid', 'NOT FOUND')}")

        aud = claims.get("aud", "")
        scp = claims.get("scp", "")
        if "outlook.office.com" not in aud and "outlook.office365.com" not in aud:
            print("\n  ⚠️  WARNING: Token audience is not outlook.office.com")
            print("     This token may be rejected by the IMAP server.")
        if "IMAP" not in scp:
            print("\n  ⚠️  WARNING: IMAP scope not found in token.")
            print("     Make sure IMAP.AccessAsUser.All permission is granted in Azure.")
        else:
            print("\n  ✓ Token looks correct for IMAP access.")
    else:
        print("  Could not decode token claims.")

    print("\n=== STEP 3: Testing IMAP connection ===")
    auth_string = f"user={EMAIL_ADDRESS}\x01auth=Bearer {token}\x01\x01"
    auth_bytes = auth_string.encode()

    try:
        print(f"  Connecting to {IMAP_HOST}:{IMAP_PORT}...")
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        print("  TCP connection OK.")

        print("  Sending XOAUTH2 authentication...")
        mail.authenticate("XOAUTH2", lambda x: auth_bytes)
        print("  ✓ IMAP authentication SUCCESSFUL!")

        print("\n=== STEP 4: Listing folders ===")
        status, folders = mail.list()
        if status == "OK":
            print("  Available folders:")
            for f in folders[:10]:
                print(f"    {f.decode()}")
        mail.logout()

    except imaplib.IMAP4.error as e:
        print(f"\n  ✗ IMAP auth failed: {e}")
        print("\n  Likely causes:")
        print("  1. IMAP not enabled — go to outlook.office.com → Settings → Mail → Sync email → turn IMAP ON")
        print("  2. Token audience wrong (check Step 2 output above)")
        print("  3. IMAP.AccessAsUser.All permission not granted in Azure")
        print("  4. GoDaddy M365 tenant has IMAP disabled at the admin level")

if __name__ == "__main__":
    main()
