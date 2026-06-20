"""
AI Email Sender v3 — Production-Grade Bulk Email Tool
=====================================================
Features:
- Status column tracking in Excel (empty / email sent / sent twice)
- Resume from crash — skips already-sent rows automatically
- Dry-run mode — preview drafts without sending
- Test recipient override — send entire batch to yourself for QA
- Email validation + duplicate detection
- Daily send cap with auto-pause
- Retry with exponential backoff on transient failures
- Reply tracking via Gmail IMAP (optional)
- Parallel drafting — pre-drafts all emails, then sends with delays

Usage:
  python send_emails.py                  # normal run
  python send_emails.py --dry-run        # draft only, no sending
  python send_emails.py --test           # send all to TEST_EMAIL
  python send_emails.py --check-replies  # only check for replies
"""

import os
import csv
import sys
import time
import random
import smtplib
import imaplib
import email as email_lib
import logging
import json
import re
import argparse
from pathlib import Path
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import ollama
from openpyxl import load_workbook
from dotenv import load_dotenv

# ── Config ─────────────────────────────────────────────────────────────────────

load_dotenv()

EXCEL_FILE      = "contacts.xlsx"
CONTEXT_FILE    = "context.md"
ATTACHMENTS_DIR = "attachments"
LOG_FILE        = "sent_log.csv"
DRAFTS_FILE     = "drafts_preview.txt"
MODEL           = "phi4-mini"

GMAIL_USER     = os.getenv("GMAIL_USER")
GMAIL_APP_PASS = os.getenv("GMAIL_APP_PASSWORD")
SENDER_NAME    = os.getenv("SENDER_NAME", "Your Name")
TEST_EMAIL     = os.getenv("TEST_EMAIL", "")   # used in --test mode

DELAY_MIN      = int(os.getenv("DELAY_MIN_SECONDS", "30"))
DELAY_MAX      = int(os.getenv("DELAY_MAX_SECONDS", "60"))
DAILY_CAP      = int(os.getenv("DAILY_SEND_CAP", "200"))   # Gmail free allows ~500/day
MAX_RETRIES    = int(os.getenv("MAX_RETRIES", "3"))
DRAFT_WORKERS  = int(os.getenv("DRAFT_WORKERS", "3"))      # parallel drafting threads

# Status column values
STATUS_EMPTY      = ""
STATUS_SENT_ONCE  = "email sent"
STATUS_SENT_TWICE = "sent twice"
STATUS_FAILED     = "failed"
STATUS_SKIPPED    = "skipped"
STATUS_REPLIED    = "replied"

NO_ATTACHMENT_VALUES = {"", "na", "n/a", "none", "no", "nan", "-"}
EMAIL_REGEX = re.compile(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$")

# ── Logging ────────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S"
)
log = logging.getLogger(__name__)

# ── Helpers ────────────────────────────────────────────────────────────────────

def load_context():
    p = Path(CONTEXT_FILE)
    if not p.exists():
        log.error(f"Context file '{CONTEXT_FILE}' not found. Create one (see README).")
        return None
    return p.read_text(encoding="utf-8").strip()


def list_attachments(folder):
    p = Path(folder)
    return [f.name for f in p.iterdir() if f.is_file()] if p.exists() else []


def is_valid_email(email):
    return bool(EMAIL_REGEX.match(email.strip()))


def needs_no_attachment(file_hint):
    return file_hint.lower().strip() in NO_ATTACHMENT_VALUES


def next_status(current):
    """Decide what status to set on success based on current value."""
    current = (current or "").lower().strip()
    if current == "":
        return STATUS_SENT_ONCE
    if current == STATUS_SENT_ONCE:
        return STATUS_SENT_TWICE
    return current  # shouldn't happen — we skip these earlier


def should_skip(current_status):
    """Returns (skip: bool, reason: str)."""
    s = (current_status or "").lower().strip()
    if s == STATUS_SENT_TWICE:
        return True, "Already sent twice — skipping"
    if s == STATUS_REPLIED:
        return True, "Recipient already replied — skipping"
    return False, ""


# ── AI calls ───────────────────────────────────────────────────────────────────

def match_file(hint, filenames):
    if not hint or not filenames:
        return None
    file_list = "\n".join(f"- {f}" for f in filenames)
    prompt = f"""You are a file-matching assistant.

Available files:
{file_list}

The contact needs: "{hint}"

Reply with ONLY the exact filename from the list, or 'null' if nothing matches."""
    try:
        r = ollama.chat(
            model=MODEL,
            messages=[{"role": "user", "content": prompt}],
            options={"temperature": 0}
        )
        result = r["message"]["content"].strip().strip('"').strip("'")
        return result if result in filenames else None
    except Exception as e:
        log.error(f"File match error: {e}")
        return None


def draft_email(context, name, description, attached_file, attempt_type="first"):
    """attempt_type: 'first' = initial outreach, 'followup' = second attempt."""
    attachment_note = (
        f"\nAttachment: '{attached_file}' — reference it naturally."
        if attached_file else
        "\nAttachment: None — do not mention any attachment."
    )

    followup_note = ""
    if attempt_type == "followup":
        followup_note = """
=== THIS IS A FOLLOW-UP ===
You previously sent an email to this person that didn't get a reply.
Write a polite, brief follow-up email — NOT a copy of the first message.
Acknowledge that you reached out before, keep it short (2-3 sentences),
and gently re-state the ask. Do not be pushy."""

    prompt = f"""You are writing an email on behalf of {SENDER_NAME}.

=== CAMPAIGN CONTEXT ===
{context}

=== RECIPIENT ===
Name: {name}
Specific details: {description}
{attachment_note}
{followup_note}

Sign off with: {SENDER_NAME}

Reply with JSON only, no markdown:
{{"subject": "...", "body": "...with \\n for line breaks"}}"""

    r = ollama.chat(
        model=MODEL,
        messages=[{"role": "user", "content": prompt}],
        options={"temperature": 0.5}
    )
    text = re.sub(r"```json|```", "", r["message"]["content"]).strip()

    try:
        parsed = json.loads(text)
        return parsed.get("subject", "Following up"), parsed.get("body", text)
    except json.JSONDecodeError:
        return "Following up", text


# ── Send + retry ──────────────────────────────────────────────────────────────

def send_email_with_retry(to_email, subject, body, attachment_path):
    last_err = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            msg = MIMEMultipart()
            msg["From"]    = f"{SENDER_NAME} <{GMAIL_USER}>"
            msg["To"]      = to_email
            msg["Subject"] = subject
            msg.attach(MIMEText(body, "plain", "utf-8"))

            if attachment_path:
                path = Path(attachment_path)
                with open(path, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f'attachment; filename="{path.name}"')
                msg.attach(part)

            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
                server.login(GMAIL_USER, GMAIL_APP_PASS)
                server.sendmail(GMAIL_USER, to_email, msg.as_string())
            return True, None
        except smtplib.SMTPException as e:
            last_err = str(e)
            # Throttling — pause longer and stop retrying further
            if any(code in last_err for code in ["421", "454", "550"]):
                log.warning(f"   ⏸  Gmail throttling detected: {last_err}")
                return False, f"Gmail throttling: {last_err}"
            backoff = 2 ** attempt
            log.warning(f"   ↻  Retry {attempt}/{MAX_RETRIES} in {backoff}s — {last_err}")
            time.sleep(backoff)
        except Exception as e:
            last_err = str(e)
            backoff = 2 ** attempt
            log.warning(f"   ↻  Retry {attempt}/{MAX_RETRIES} in {backoff}s — {last_err}")
            time.sleep(backoff)
    return False, last_err or "Unknown error"


# ── Reply tracking ────────────────────────────────────────────────────────────

def check_replies(contact_emails):
    """Check Gmail inbox for replies from anyone in our contact list."""
    log.info("Checking Gmail inbox for replies…")
    replied = set()
    try:
        with imaplib.IMAP4_SSL("imap.gmail.com") as mail:
            mail.login(GMAIL_USER, GMAIL_APP_PASS)
            mail.select("INBOX")

            for email_addr in contact_emails:
                # Search for emails FROM this contact in the last 30 days
                date_since = (datetime.now() - timedelta(days=30)).strftime("%d-%b-%Y")
                status, data = mail.search(None, f'(FROM "{email_addr}" SINCE {date_since})')
                if status == "OK" and data[0]:
                    replied.add(email_addr.lower())
    except Exception as e:
        log.warning(f"Could not check replies via IMAP: {e}")
        log.warning("Make sure IMAP is enabled in Gmail settings.")
        return set()
    log.info(f"Found {len(replied)} replies")
    return replied


# ── Excel I/O ────────────────────────────────────────────────────────────────

def read_contacts():
    """Read xlsx with openpyxl so we can also write back to it."""
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    headers = [str(c.value).lower().strip() if c.value else "" for c in ws[1]]

    contacts = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=False), start=2):
        data = {}
        for i, cell in enumerate(row):
            if i < len(headers) and headers[i]:
                data[headers[i]] = str(cell.value).strip() if cell.value is not None else ""
        data["_row_idx"] = row_idx
        contacts.append(data)

    return wb, ws, headers, contacts


def ensure_status_column(wb, ws, headers):
    if "status" not in headers:
        col_idx = len(headers) + 1
        ws.cell(row=1, column=col_idx, value="status")
        headers.append("status")
        wb.save(EXCEL_FILE)
        log.info("Added 'status' column to spreadsheet")
    return headers.index("status") + 1


def update_status(wb, ws, row_idx, status_col, value):
    ws.cell(row=row_idx, column=status_col, value=value)
    wb.save(EXCEL_FILE)


# ── CSV log ───────────────────────────────────────────────────────────────────

def init_log():
    if not Path(LOG_FILE).exists():
        with open(LOG_FILE, "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow(["timestamp", "name", "email", "attempt_type", "file_hint", "matched_file", "subject", "status", "note"])


def write_log(row):
    with open(LOG_FILE, "a", newline="", encoding="utf-8") as f:
        csv.DictWriter(f, fieldnames=["timestamp", "name", "email", "attempt_type", "file_hint", "matched_file", "subject", "status", "note"]).writerow(row)


# ── Pretty terminal helpers ───────────────────────────────────────────────────

def countdown_delay(seconds):
    for r in range(seconds, 0, -1):
        print(f"\r   ⏳ Waiting {r}s before next send... ", end="", flush=True)
        time.sleep(1)
    print("\r" + " " * 60 + "\r", end="", flush=True)


def write_drafts_preview(drafts):
    with open(DRAFTS_FILE, "w", encoding="utf-8") as f:
        f.write(f"DRAFTS PREVIEW — generated {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("=" * 70 + "\n\n")
        for d in drafts:
            f.write(f"To:      {d['name']} <{d['email']}>\n")
            f.write(f"Type:    {d['attempt_type']}\n")
            f.write(f"Attach:  {d.get('matched_file') or '(none)'}\n")
            f.write(f"Subject: {d['subject']}\n")
            f.write("-" * 70 + "\n")
            f.write(d["body"] + "\n")
            f.write("=" * 70 + "\n\n")
    log.info(f"✓ Wrote {len(drafts)} drafts to '{DRAFTS_FILE}' for your review")


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run",       action="store_true", help="Draft emails but don't send them")
    parser.add_argument("--test",          action="store_true", help="Send all emails to TEST_EMAIL instead")
    parser.add_argument("--check-replies", action="store_true", help="Only check for replies, don't send")
    args = parser.parse_args()

    log.info("=" * 60)
    log.info("AI Email Sender v3 — Production Edition")
    if args.dry_run:       log.info("MODE: DRY RUN (no emails will be sent)")
    if args.test:          log.info(f"MODE: TEST (all emails redirected to {TEST_EMAIL})")
    if args.check_replies: log.info("MODE: REPLY CHECK ONLY")
    log.info("=" * 60)

    # Validate
    if not GMAIL_USER or not GMAIL_APP_PASS:
        log.error("Missing GMAIL_USER or GMAIL_APP_PASSWORD in .env")
        return
    if args.test and not TEST_EMAIL:
        log.error("--test mode needs TEST_EMAIL in .env")
        return

    context = load_context()
    if not context:
        return
    log.info(f"✓ Loaded context.md ({len(context)} chars)")

    # Read spreadsheet
    try:
        wb, ws, headers, all_contacts = read_contacts()
    except FileNotFoundError:
        log.error(f"Excel file '{EXCEL_FILE}' not found")
        return

    status_col = ensure_status_column(wb, ws, headers)
    log.info(f"✓ Loaded {len(all_contacts)} rows from '{EXCEL_FILE}'")

    # Reply check mode — just update status of repliers and exit
    if args.check_replies:
        all_emails = [c.get("email", "") for c in all_contacts if c.get("email")]
        replied = check_replies(all_emails)
        if replied:
            for c in all_contacts:
                if c.get("email", "").lower() in replied:
                    update_status(wb, ws, c["_row_idx"], status_col, STATUS_REPLIED)
                    log.info(f"   ↩ {c.get('name')} replied")
        log.info("=" * 60)
        return

    # Dedup + validation
    seen = set()
    contacts = []
    for c in all_contacts:
        email = c.get("email", "").strip().lower()
        if not email or email == "nan":
            continue
        if not is_valid_email(email):
            log.warning(f"⚠ Invalid email skipped: {c.get('email')}")
            continue
        if email in seen:
            log.warning(f"⚠ Duplicate email skipped: {email}")
            continue
        seen.add(email)
        contacts.append(c)

    # Filter to only contacts that need sending
    to_process = []
    for c in contacts:
        status = c.get("status", "")
        skip, reason = should_skip(status)
        if skip:
            log.info(f"   ⊘ {c['email']} — {reason}")
            continue
        to_process.append(c)

    if not to_process:
        log.info("No contacts need processing. Done.")
        return

    log.info(f"✓ {len(to_process)} contacts to process this run")

    # Optional: check for replies first so we don't email people who already replied
    if not args.dry_run and not args.test:
        try:
            replied = check_replies([c["email"] for c in to_process])
            new_to_process = []
            for c in to_process:
                if c["email"].lower() in replied:
                    update_status(wb, ws, c["_row_idx"], status_col, STATUS_REPLIED)
                    log.info(f"   ↩ {c['name']} already replied — skipping")
                else:
                    new_to_process.append(c)
            to_process = new_to_process
        except Exception as e:
            log.warning(f"Reply check skipped: {e}")

    available_files = list_attachments(ATTACHMENTS_DIR)
    log.info(f"✓ {len(available_files)} files in '{ATTACHMENTS_DIR}/'")
    log.info("=" * 60)

    # ── STAGE 1: Pre-draft ALL emails in parallel ───────────────────────────────
    log.info("Stage 1: Pre-drafting all emails in parallel…")

    def make_draft(c):
        try:
            status = c.get("status", "")
            attempt_type = "followup" if status.lower() == STATUS_SENT_ONCE else "first"
            file_hint = c.get("file_hint", "") or c.get("file", "")

            matched_file    = None
            attachment_path = None
            if not needs_no_attachment(file_hint):
                if available_files:
                    matched_file = match_file(file_hint, available_files)
                    if matched_file:
                        attachment_path = str(Path(ATTACHMENTS_DIR) / matched_file)

            subject, body = draft_email(
                context, c.get("name", "there"),
                c.get("description", ""), matched_file, attempt_type
            )
            return {
                **c,
                "attempt_type":    attempt_type,
                "matched_file":    matched_file,
                "attachment_path": attachment_path,
                "subject":         subject,
                "body":            body,
                "draft_ok":        True,
            }
        except Exception as e:
            return {**c, "draft_ok": False, "error": str(e)}

    drafts = []
    init_log()
    with ThreadPoolExecutor(max_workers=DRAFT_WORKERS) as executor:
        futures = {executor.submit(make_draft, c): c for c in to_process}
        for i, fut in enumerate(as_completed(futures), start=1):
            d = fut.result()
            drafts.append(d)
            if d["draft_ok"]:
                tag = "📎" if d["matched_file"] else "📭"
                log.info(f"   [{i}/{len(to_process)}] {tag} Drafted for {d['name']}: \"{d['subject']}\"")
            else:
                log.error(f"   [{i}/{len(to_process)}] ✗ Draft failed for {d['name']}: {d['error']}")

    # Restore original order from to_process
    order = {c["email"]: idx for idx, c in enumerate(to_process)}
    drafts.sort(key=lambda d: order.get(d.get("email", ""), 999999))

    successful_drafts = [d for d in drafts if d.get("draft_ok")]
    write_drafts_preview(successful_drafts)

    # ── Dry run — stop here ─────────────────────────────────────────────────────
    if args.dry_run:
        log.info("=" * 60)
        log.info(f"DRY RUN COMPLETE — {len(successful_drafts)} drafts in '{DRAFTS_FILE}'")
        log.info("Review the file, then run without --dry-run to send.")
        log.info("=" * 60)
        return

    # ── STAGE 2: Send sequentially with delays ──────────────────────────────────
    log.info("=" * 60)
    log.info(f"Stage 2: Sending {len(successful_drafts)} emails with {DELAY_MIN}-{DELAY_MAX}s delays…")
    log.info(f"Daily cap: {DAILY_CAP} emails. Will pause when reached.")
    log.info("=" * 60)

    stats = {"sent": 0, "skipped": 0, "failed": 0, "replied": 0}
    sent_today = 0

    for idx, d in enumerate(successful_drafts):
        if sent_today >= DAILY_CAP:
            log.warning(f"⏸  Daily cap of {DAILY_CAP} reached. Stopping for today.")
            log.warning("    Re-run tomorrow to continue (status column tracks progress).")
            break

        target_email = TEST_EMAIL if args.test else d["email"]
        log.info(f"[{idx+1}/{len(successful_drafts)}] → {d['name']} <{target_email}>")
        if args.test:
            log.info(f"   (originally: {d['email']})")

        log_entry = {
            "timestamp":    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "name":         d["name"],
            "email":        d["email"],
            "attempt_type": d["attempt_type"],
            "file_hint":    d.get("file_hint", ""),
            "matched_file": d.get("matched_file") or "",
            "subject":      d["subject"],
            "status":       "",
            "note":         "",
        }

        success, err = send_email_with_retry(
            target_email, d["subject"], d["body"], d.get("attachment_path")
        )

        if success:
            log.info(f"   ✓ Sent — {d['subject']}")
            log_entry["status"] = "SENT"
            stats["sent"] += 1
            sent_today += 1
            if not args.test:
                # Update Excel status column only on real sends
                new_status = next_status(d.get("status", ""))
                update_status(wb, ws, d["_row_idx"], status_col, new_status)
        else:
            log.error(f"   ✗ Failed: {err}")
            log_entry["status"] = "FAILED"
            log_entry["note"]   = err
            stats["failed"] += 1
            if not args.test:
                update_status(wb, ws, d["_row_idx"], status_col, STATUS_FAILED)
            # If Gmail is throttling, stop the whole run
            if err and any(x in err for x in ["throttling", "421", "454"]):
                log.error("Gmail is rate-limiting. Stopping run to protect your account.")
                break

        write_log(log_entry)

        if idx < len(successful_drafts) - 1 and sent_today < DAILY_CAP:
            countdown_delay(random.randint(DELAY_MIN, DELAY_MAX))

    # Summary
    log.info("=" * 60)
    log.info(f"Done. Sent: {stats['sent']}  Failed: {stats['failed']}  Today's total: {sent_today}/{DAILY_CAP}")
    log.info(f"Log:    '{LOG_FILE}'")
    log.info(f"Drafts: '{DRAFTS_FILE}'")
    log.info(f"Excel status column updated.")
    log.info("=" * 60)


if __name__ == "__main__":
    main()
