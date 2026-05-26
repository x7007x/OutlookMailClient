# OutlookMailClient

> Headless Outlook Web scraper built with Playwright — handles full Microsoft login flow, MFA prompts, password changes, and inbox email extraction.

---

## Features

- Full Microsoft login flow automation (email → password → MFA redirects)
- Handles forced password change screens
- Skips MFA setup prompts (Authenticator, mysignins)
- Auto-retries with alternative passwords via env variable
- Extracts inbox emails: sender, subject, preview, time
- Reads full email details: from, to, date, body (text + HTML)
- Cookie-based session reuse (login once, fetch separately)
- Headless by default, easy to flip for debugging

---

## Requirements

- Python 3.8+
- [Playwright](https://playwright.dev/python/)

```bash
pip install playwright
playwright install chromium
```

---

## Usage

### Via environment variables (recommended)

```bash
export OUTLOOK_EMAIL="your@email.com"
export OUTLOOK_PASSWORD="yourpassword"
python main.py
```

### Optional env variables

| Variable | Description |
|---|---|
| `OUTLOOK_EMAIL` | Primary Outlook email |
| `OUTLOOK_PASSWORD` | Primary password |
| `OUTLOOK_NEW_PASSWORD` | New password if forced change is triggered (default: `{password}@1`) |
| `OUTLOOK_ALT_PASSWORDS` | Comma-separated list of fallback passwords to try |

---

## Output

Returns a JSON array of emails from the inbox:

```json
[
  {
    "sender": "John Doe",
    "subject": "Meeting Tomorrow",
    "preview": "Hi, just a reminder...",
    "time": "10:30 AM",
    "details": {
      "subject": "Meeting Tomorrow",
      "from": { "name": "John Doe", "email": "john@example.com" },
      "to": [{ "name": "You", "email": "you@example.com" }],
      "date": "Mon 5/26/2025 10:30 AM",
      "body_text": "Hi, just a reminder about the meeting tomorrow at 9 AM.",
      "body_html": "<div>...</div>"
    }
  }
]
```

---

## How It Works

1. **`login()`** — Launches a headless Chromium browser, navigates to `outlook.office.com/mail`, and walks through all Microsoft login screens automatically. Returns session cookies on success.

2. **`get_emails(cookies)`** — Reuses the session cookies to load the inbox and scrape email list rows. Opens the first email to extract full details.

---

## Notes

- Credentials in `__main__` are blank by default — always use environment variables.
- The scraper uses `--no-sandbox` and `--disable-dev-shm-usage` flags for compatibility with Linux/Docker environments.
- Body cleaning strips UI artifacts like "Reply", "Forward", "Summarize" action labels automatically.

---

## License

MIT
