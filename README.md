# AI Email Sender v4 — Draft → Edit → Send

The key improvement in v4: **the AI drafts once, you edit, every send respects your edits.**

---

## The correct workflow

```
Step 1 — Draft
  python send_emails.py --draft
  AI generates all emails → saved to drafts.json + drafts_preview.txt
  Open drafts.json and edit any subject or body you want

Step 2 — Test (send to yourself)
  python send_emails.py --test
  Reads YOUR edited drafts.json
  Sends everything to TEST_EMAIL so you can check your inbox

Step 3 — Send (real recipients)
  python send_emails.py --send
  Reads the SAME drafts.json (your edited version)
  Sends to real recipients with delays + updates Excel status

Optional — Check replies
  python send_emails.py --check-replies
  Scans Gmail for replies, marks repliers in Excel
```

---

## The two draft files and what they do

| File | Purpose | Edit it? |
|---|---|---|
| `drafts.json` | Machine-readable. Stores all email data. The send reads from HERE | ✅ Yes — edit subject and body fields |
| `drafts_preview.txt` | Human-readable version of the same data. For easy reading only | ❌ Edits here are ignored — always edit drafts.json |

**Important:** `drafts_preview.txt` is auto-regenerated every time you run `--draft`. Edits to it are lost. Always edit `drafts.json`.

---

## How to edit drafts.json

Open it in any text editor (Notepad, VS Code). It looks like this:

```json
[
  {
    "email": "priya@stripe.com",
    "name": "Priya Sharma",
    "attempt_type": "first",
    "matched_file": "resume_2024.pdf",
    "attachment_path": "attachments/resume_2024.pdf",
    "subject": "Backend Engineering role at Stripe",
    "body": "Hi Priya,\n\nI came across the Backend Engineer opening...",
    "row_idx": 2,
    "file_hint": "resume",
    "current_status": ""
  },
  ...
]
```

**Only change** the `subject` and `body` fields. Leave everything else untouched.

**Line breaks in body:** Use `\n` for line breaks inside the JSON string. Example:
```json
"body": "Hi Priya,\n\nI saw your opening...\n\nBest,\nYour Name"
```

---

## Setup (one time)

### 1. Install Ollama and pull Gemma 4
```bash
ollama pull gemma4:e4b
```

### 2. Install Python dependencies
```bash
python -m pip install pandas openpyxl ollama python-dotenv --user
```

### 3. Create .env file
```env
GMAIL_USER=your@gmail.com
GMAIL_APP_PASSWORD=your-16-char-app-password
SENDER_NAME=Your Full Name
TEST_EMAIL=your@gmail.com
OLLAMA_MODEL=gemma4:e4b

DELAY_MIN_SECONDS=30
DELAY_MAX_SECONDS=60
DAILY_SEND_CAP=200
MAX_RETRIES=3
DRAFT_WORKERS=3
```

### 4. Folder structure
```
email-sender/
├── send_emails.py
├── context.md
├── contacts.xlsx        ← name, email, description, file_hint, status
├── attachments/
│   └── resume.pdf
├── .env
├── drafts.json          ← auto-created by --draft, YOU EDIT THIS
├── drafts_preview.txt   ← auto-created, read-only
└── sent_log.csv         ← auto-created
```

---

## Status column reference

| Value | Meaning | Next run action |
|---|---|---|
| *(empty)* | Never contacted | ✅ Draft + send first email |
| `email sent` | Sent once | ✅ Draft + send follow-up |
| `sent twice` | Followed up | ⛔ Skip forever |
| `replied` | They replied | ⛔ Skip forever |
| `failed` | Last attempt failed | 🔁 Retry |

---

## Common issues

| Problem | Fix |
|---|---|
| `drafts.json not found` | Run `--draft` first |
| Edits not showing in sent email | Make sure you edited `drafts.json`, not `drafts_preview.txt` |
| JSON parse error after editing | Check for missing commas, unclosed quotes in drafts.json |
| Ollama not responding | Make sure Ollama is running (system tray on Windows) |
| Gmail App Passwords unavailable | Enable 2-Step Verification first |
| IMAP error on check-replies | Enable IMAP in Gmail Settings → Forwarding and POP/IMAP |
