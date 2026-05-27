## Approach Options

Since GoDaddy email uses Microsoft 365 (Outlook/Exchange), you have two solid paths:

| Approach | Best For | Complexity |
|---|---|---|
| **Python + IMAP** | Simple, works on any OS, no extra setup | Low |
| **Microsoft Graph API** | More robust, handles large mailboxes better | Medium |

The **Python + IMAP** approach is the easiest and works well for this use case. Here's a complete script with a safe **dry-run mode** so you can preview what will be deleted before actually deleting anything.## How to use it

**1. Install the one dependency (optional but recommended):**
```bash
pip install python-dotenv
```

**2. Create a `.env` file next to the script:**
```
EMAIL=you@yourdomain.com
PASSWORD=your_password
```

**3. Figure out which IMAP host you need** — edit `IMAP_HOST` in the script:

| GoDaddy Account Type | IMAP Host |
|---|---|
| Legacy GoDaddy hosting email | `imap.secureserver.net` |
| GoDaddy Microsoft 365 | `outlook.office365.com` |

**4. Do a dry run first:**
```bash
python delete_old_emails.py
```
This will print every email it *would* delete without touching anything.

**5. When you're confident, flip the switch:**
```python
DRY_RUN = False
```
and run it again — emails will be permanently deleted.

---

## Important notes

- **M365 / Modern Auth**: If your GoDaddy account is Microsoft 365 and you have MFA enabled, you'll need to generate an **App Password** in your Microsoft account security settings and use that instead of your regular password.
- **`FOLDERS_TO_SCAN`**: The script checks `INBOX` and `Sent` by default. You can add other folders like `"Archive"` or `"Junk Email"`.
- **`LIMIT_PER_FOLDER`**: Set this to something small (e.g. `50`) for your first live run as an extra safety check before removing everything.
- **Speed**: If you have thousands of qualifying emails, fetching full message bodies to check for attachments can be slow. It's normal — each email needs to be downloaded to inspect.
