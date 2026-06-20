# AI Email Sender v3 — Production-Grade Bulk Email Tool

Send personalised bulk emails using AI — for job applications, event invites, follow-ups, sales outreach, or any campaign where every email needs to feel personally written.

Runs entirely on your computer using **Phi-4 Mini** (a free local AI model) and **Gmail**. No API costs, no data leaving your machine.

---

## What's new in v3

| Feature | What it does |
|---|---|
| **Status column tracking** | Excel sheet tracks who's been emailed — never accidentally send twice |
| **Resume from crash** | Stop the script anytime — re-running picks up exactly where it left off |
| **Dry-run mode** | Preview every email before sending with `--dry-run` |
| **Test recipient override** | Send all emails to yourself with `--test` for full QA |
| **Email validation** | Catches typos like `john@gmial.com` before sending |
| **Duplicate detection** | Same email twice in the sheet = sent only once |
| **Daily send cap** | Auto-stops at 200/day (configurable) to protect your Gmail reputation |
| **Retry with backoff** | Transient SMTP errors retried 3 times with exponential backoff |
| **Reply tracking** | Checks Gmail inbox — skips contacts who already replied |
| **Parallel drafting** | All emails drafted at once (3x faster), then sent with delays |

---

## How the status column works

Add a `status` column to your `contacts.xlsx`. The script automatically updates it after each email:

| Status value | Meaning | Next run action |
|---|---|---|
| *(empty)* | Never contacted | ✅ Send first email → mark as `email sent` |
| `email sent` | Sent once | ✅ Send follow-up → mark as `sent twice` |
| `sent twice` | Already followed up | ⛔ Skip — do not contact again |
| `failed` | Last attempt failed | 🔁 Retry on next run |
| `replied` | They replied to a previous email | ⛔ Skip — auto-set by reply tracker |

If you don't add the status column yourself, the script adds it automatically on first run.

**This gives you a built-in 2-email-max rule.** You can never accidentally spam someone — once they're at `sent twice`, the script will refuse to send them anything ever again.

---

## Running modes

### Normal run (sends real emails)
```bash
python send_emails.py
```

### Dry-run — preview drafts, don't send
```bash
python send_emails.py --dry-run
```
Drafts every email and writes them to `drafts_preview.txt`. No emails are sent. Review the file, then run without the flag.

### Test mode — send all to yourself
```bash
python send_emails.py --test
```
Sends every email to your `TEST_EMAIL` instead of real recipients. Perfect for confirming attachments, subjects, and formatting before a real run. Status column is NOT updated in test mode.

### Check replies only
```bash
python send_emails.py --check-replies
```
Scans your Gmail inbox for replies from contacts in the sheet and updates their status to `replied`. Doesn't send anything.

---

## Setup (one time)

### 1. Install Ollama and pull Phi-4 Mini
```bash
# Download from ollama.com, then:
ollama pull phi4-mini
```

### 2. Install Python dependencies
```bash
python -m pip install pandas openpyxl ollama python-dotenv --user
```

### 3. Get a Gmail App Password
1. Go to **myaccount.google.com → Security**
2. Enable **2-Step Verification**
3. Search **"App Passwords"** → create one → copy the 16-char password

### 4. Enable IMAP in Gmail (for reply tracking)
1. Gmail → Settings → See all settings
2. Forwarding and POP/IMAP tab
3. Enable IMAP → Save changes

### 5. Set up your folder
```
email-sender/
├── send_emails.py        ← the script
├── context.md            ← your campaign context
├── contacts.xlsx         ← your spreadsheet (with status column)
├── attachments/          ← files to attach
│   ├── resume.pdf
│   └── portfolio.pdf
├── .env                  ← your credentials
├── sent_log.csv          ← auto-created
└── drafts_preview.txt    ← auto-created when using --dry-run
```

### 6. Create your .env file
Copy `.env.example` to `.env` and fill in your values:
```env
GMAIL_USER=your@gmail.com
GMAIL_APP_PASSWORD=your-16-char-password
SENDER_NAME=Your Name
TEST_EMAIL=your@gmail.com

DELAY_MIN_SECONDS=30
DELAY_MAX_SECONDS=60
DAILY_SEND_CAP=200
MAX_RETRIES=3
DRAFT_WORKERS=3
```

---

## Excel sheet format

| name | email | description | file_hint | status |
|---|---|---|---|---|
| Priya Sharma | priya@stripe.com | Backend Engineer role at Stripe | resume | |
| David Kim | david@startup.io | Senior SWE position at fintech startup | resume | email sent |
| Anna Mueller | anna@techcorp.com | LinkedIn follow-up | NA | sent twice |

- **name** — recipient's name
- **email** — their email address
- **description** — specifics for this person (what makes this email different)
- **file_hint** — keyword matching a file in `/attachments` (e.g. "resume" → resume.pdf). Use **NA** or leave empty for no attachment
- **status** — leave empty for first send. Auto-updated by the script

---

## Recommended workflow for job applications

```bash
# Step 1 — Add 10 contacts to your sheet, leave status empty
# Step 2 — Dry-run to preview drafts
python send_emails.py --dry-run
# Review drafts_preview.txt

# Step 3 — Test by sending all to yourself
python send_emails.py --test
# Check your inbox — does formatting/attachment look right?

# Step 4 — Real send
python send_emails.py
# Excel updates: all rows now say "email sent"

# Step 5 — Wait 7 days, check for replies
python send_emails.py --check-replies
# Repliers get marked "replied" — they won't get follow-ups

# Step 6 — Send follow-up to the rest
python send_emails.py
# Anyone still on "email sent" gets a polite follow-up
# After this run, they're marked "sent twice" — no further emails ever
```

---

## Why XLSX (not CSV / ODS / JSON)?

XLSX is the right choice for this tool:
- Universal — opens in Excel, Google Sheets, Numbers, LibreOffice
- Supports the status column workflow naturally (cell formatting, dropdowns)
- The format itself is an open ISO standard (Office Open XML)
- The library we use (`openpyxl`) is fully open source
- The actual bottleneck is AI drafting and send delays — not the file format

If you ever need a 100% open alternative, CSV works too — but you lose cell formatting and validation. Just rename to `.csv` and the script reads it the same way.

---

## Anti-spam features explained

| Feature | Why it matters |
|---|---|
| 30–60s randomised delay | Real humans don't send emails every 5 seconds |
| Unique AI-generated content | Spam filters detect identical templates instantly |
| Real Gmail SMTP | Sending from `gmail.com` has best deliverability |
| Plain text only | HTML emails get more scrutiny from spam filters |
| Daily cap (200) | Gmail's hard limit is ~500/day — staying under buffers safely |
| App Password (not OAuth) | More stable for scripted sending |
| Personal sign-off | AI always signs off with your real name |

**Realistic throughput**: At 30-60s delays, expect ~60-90 emails per hour. A 200-email daily cap means about 3 hours of script running.

---

## Common issues

| Problem | Fix |
|---|---|
| `Module not found` | `python -m pip install <name> --user` |
| Emails going to spam | Increase `DELAY_MIN_SECONDS` to 60+, send <100/day |
| IMAP reply check fails | Enable IMAP in Gmail settings (Forwarding and POP/IMAP) |
| Phi-4 Mini hangs | Make sure Ollama is running (system tray) — restart it if stuck |
| Status column not updating | Close `contacts.xlsx` before running — file lock issue |
| `Gmail throttling` warning | You hit Gmail's limit — stop and wait 24 hours |

---

## What to put in context.md

Be specific. The more context, the better the emails. Examples:

**For job applications:**
- Your name, current role, years of experience
- The kind of roles you're targeting
- 2-3 standout achievements
- Your portfolio/LinkedIn URLs
- Tone you want (professional, friendly, technical)

**For event invites:**
- Event details (date, time, venue, dress code)
- Why you're inviting people
- What you want them to do (RSVP, share, bring something)

**For sales follow-ups:**
- Your product and value prop
- Common objections to address
- What action you want them to take

---

## License
Free to use, modify, and share. No warranty — use responsibly and don't spam people.
