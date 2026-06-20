# AI Email Sender — Local, Free, Context-Driven Bulk Email Tool

Send personalised bulk emails using AI — for job applications, event invites, follow-ups, sales outreach, or any campaign where every email needs to feel personally written.

Runs entirely on your computer using **Phi-4 Mini** (a free local AI model) and **Gmail**. No API costs, no data leaving your machine.

---

## What this does

You provide:
1. A shared **context.md** file — what the campaign is about, your tone, what to include in every email
2. A **contacts.xlsx** spreadsheet — one row per recipient, with name + email + specific details
3. Optional **attachment files** — resumes, brochures, invoices, etc.

The tool:
- Reads each row from your spreadsheet
- Combines your shared context with the per-row description
- Drafts a unique, personalised email using AI
- Attaches the right file based on a keyword hint (or skips attachment if marked "NA")
- Sends via your Gmail with a random delay between sends to avoid spam filters
- Logs every result to a CSV so you can track what was sent

---

## Why use this instead of a normal mail merge?

A normal mail merge replaces placeholders like `{name}` and `{company}` with values from a sheet — every email has the same wording, just different names. That's exactly what spam filters and hiring managers learn to spot.

This tool actually **writes** each email from scratch using your context and a short description. Two emails from the same campaign will read like two genuinely different messages — different opening sentences, different phrasing, different structures — because the AI generates each one independently.

---

## Perfect for

- **Job applications** — send personalised outreach to recruiters with your resume attached
- **Event invites** — tailor invites for each guest based on what you know about them
- **Wishing emails** — birthdays, anniversaries, holidays, congratulations
- **Sales outreach** — follow up on leads with context-aware messages
- **Networking** — reach out to multiple people with genuinely different emails
- **Course / community announcements** — announce launches to mailing lists

---

## Requirements

- **Python 3.9 or higher**
- **Ollama** installed and running ([ollama.com](https://ollama.com))
- **Phi-4 Mini** model pulled (~2.5 GB, runs on most laptops with 8GB RAM)
- A **Gmail account** with 2-Step Verification enabled

---

## Setup (one time)

### 1. Install Ollama and pull the model
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
3. Search **"App Passwords"** → create one named "Email Sender"
4. Copy the 16-character password

### 4. Set up your project folder
```
email-sender/
├── send_emails.py        ← the main script
├── context.md            ← your campaign context (see below)
├── contacts.xlsx         ← your spreadsheet
├── attachments/          ← any files to attach
│   ├── resume.pdf
│   ├── portfolio.pdf
│   └── brochure.pdf
├── .env                  ← your Gmail credentials
└── sent_log.csv          ← auto-created on first run
```

### 5. Create your .env file
```env
GMAIL_USER=your@gmail.com
GMAIL_APP_PASSWORD=your-16-char-app-password
SENDER_NAME=Your Full Name

# Optional — control how slow/fast emails send (in seconds)
DELAY_MIN_SECONDS=30
DELAY_MAX_SECONDS=60
```

---

## How to use it

### Step 1 — Edit `context.md`
Tell the AI what your campaign is about. Be specific about tone, what every email should include, and what to avoid.

See the included `context.md` for a job application example. Adapt it for your purpose.

### Step 2 — Fill in `contacts.xlsx`
Required columns (lowercase, exactly these names):

| name | email | description | file_hint |
|---|---|---|---|
| Priya Sharma | priya@stripe.com | Backend Engineer role at Stripe — values distributed systems experience | resume |
| Anna Mueller | anna@techcorp.com | Following up on our LinkedIn chat about their platform team | NA |

- **name** — recipient's name
- **email** — their email address
- **description** — what makes THIS email different from the others (specific details, the role, the context)
- **file_hint** — keyword to match a file in `/attachments` (e.g. "resume" → resume.pdf). Use **NA** or leave blank if no file needed.

### Step 3 — Add your attachments
Drop files into the `attachments/` folder. The AI matches them to each contact via the file_hint keyword.

### Step 4 — Run
```bash
python send_emails.py
```

The script will:
- Validate your setup
- Process each contact one by one
- Show real-time progress with a countdown between emails
- Save a full log to `sent_log.csv`

---

## Anti-spam features built in

| Feature | What it does |
|---|---|
| **Random delays** | Waits 30–60 seconds between emails (configurable) so it doesn't look bot-like |
| **Unique content per email** | AI generates fresh wording each time — no two emails are identical |
| **Real Gmail SMTP** | Uses your actual Gmail server, not a third-party relay |
| **Plain text body** | Avoids HTML which triggers more spam filters for cold outreach |
| **Natural language** | AI writes like a person, no "Dear Sir/Madam" templates |

**Gmail free account limits**: ~500 emails/day. The tool will happily run through that, but stagger large batches across multiple days to stay safe.

---

## Example use cases

### Job applications
**context.md** — explains your background, what kind of roles you want, your portfolio
**description column** — the specific role and what to emphasize for that company
**file_hint** — `resume` (matches resume.pdf in attachments)

### Birthday wishes
**context.md** — warm, personal tone, mention something genuine
**description column** — what you remember about them, shared memory, or recent news
**file_hint** — `NA` (no attachment needed)

### Event invitations
**context.md** — the event details, date, venue, dress code
**description column** — why you want THIS person there, personal touch
**file_hint** — `invite` (matches invitation.pdf)

### Sales follow-ups
**context.md** — your product, value proposition, what to mention
**description column** — what was discussed in the last call, their pain point
**file_hint** — `proposal` or `case_study` or `NA`

---

## Reading the log file

After every run, `sent_log.csv` shows:

| column | meaning |
|---|---|
| timestamp | when the email was processed |
| name, email | recipient |
| file_hint | what the row asked for |
| matched_file | what file was actually attached (if any) |
| subject | the AI-generated subject line |
| status | SENT / SKIPPED / FAILED |
| note | reason if skipped or failed |

---

## Common issues

| Problem | Fix |
|---|---|
| `Module not found: ollama` | Run `python -m pip install ollama --user` |
| `App Passwords not available` | Enable 2-Step Verification first |
| Emails going to spam | Increase `DELAY_MIN_SECONDS` to 60+, send fewer per day |
| Email skipped — "no file matched" | Check filename in attachments matches the file_hint keyword |
| Ollama connection refused | Make sure Ollama app is running (check system tray) |
| Generic-sounding emails | Make your `description` column more specific |

---

## Tips for best results

1. **Be specific in descriptions** — "Backend role at Stripe, values Python and distributed systems" beats "apply for backend role"
2. **Test on yourself first** — put your own email in row 1 and check the output looks good
3. **Iterate on context.md** — if emails feel off, refine the context with examples and tone notes
4. **Use real filenames** — name your files clearly (`resume_2024.pdf`, not `doc1.pdf`) so the AI matches them correctly
5. **Run small batches first** — try 3–5 contacts before sending to 50
6. **Don't skip the delay** — even 30 seconds between sends makes a huge difference for deliverability

---

## What's NOT included (you'd need to add yourself)

- Email reply tracking — this only sends, doesn't monitor responses
- A/B testing of subject lines
- Web UI — this is a CLI tool
- Multiple email account rotation — uses one Gmail account
- Email warmup — if your Gmail is brand new, send a few personal emails first

---

## License
Free to use, modify, and share. No warranty — use responsibly and don't spam people.
