"""
AI Email Sender v2 — Context-Driven Bulk Email Tool
====================================================
- Uses a shared context.md file to set the purpose and tone of all emails
- Per-row description adds the specific details for each recipient
- Configurable delay between sends to avoid Gmail spam filters
- File hint "NA" or empty = send without attachment (no flag)
- Optimized for job applications, event invites, follow-ups, outreach
"""

import os
import csv
import time
import random
import smtplib
import logging
import json
import re
from pathlib import Path
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import pandas as pd
import ollama
from dotenv import load_dotenv

# ── Config ─────────────────────────────────────────────────────────────────────

load_dotenv()

EXCEL_FILE      = "contacts.xlsx"
CONTEXT_FILE    = "context.md"
ATTACHMENTS_DIR = "attachments"
LOG_FILE        = "sent_log.csv"
MODEL           = "phi4-mini"

GMAIL_USER     = os.getenv("GMAIL_USER")
GMAIL_APP_PASS = os.getenv("GMAIL_APP_PASSWORD")
SENDER_NAME    = os.getenv("SENDER_NAME", "Your Name")

# Anti-spam delay settings (in seconds)
# Gmail allows ~500 emails/day for free accounts. Going too fast triggers spam filters.
DELAY_MIN = int(os.getenv("DELAY_MIN_SECONDS", "30"))   # minimum wait between emails
DELAY_MAX = int(os.getenv("DELAY_MAX_SECONDS", "60"))   # maximum wait between emails
# A random delay between these is used → looks more human, less bot-like

# Values in file_hint that mean "no attachment needed"
NO_ATTACHMENT_VALUES = {"", "na", "n/a", "none", "no", "nan", "-"}

# ── Logging ────────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S"
)
log = logging.getLogger(__name__)

# ── Load shared context ────────────────────────────────────────────────────────

def load_context():
    """Read context.md — the shared purpose/tone for all emails in this batch."""
    p = Path(CONTEXT_FILE)
    if not p.exists():
        log.error(f"Context file '{CONTEXT_FILE}' not found.")
        log.error("Create one — it tells the AI what these emails are about.")
        log.error("See the sample context.md or the README for examples.")
        return None
    with open(p, "r", encoding="utf-8") as f:
        return f.read().strip()


# ── Helpers ────────────────────────────────────────────────────────────────────

def list_attachments(folder):
    p = Path(folder)
    if not p.exists():
        return []
    return [f.name for f in p.iterdir() if f.is_file()]


def needs_no_attachment(file_hint):
    """Returns True if file_hint says 'no attachment'."""
    return file_hint.lower().strip() in NO_ATTACHMENT_VALUES


def match_file(hint, filenames):
    """Ask Phi-4 Mini to pick the best matching file."""
    if not hint or not filenames:
        return None

    file_list = "\n".join(f"- {f}" for f in filenames)
    prompt = f"""You are a file-matching assistant.

Available files:
{file_list}

The contact needs: "{hint}"

Which single filename is the best match? Reply with ONLY the exact filename.
If nothing matches, reply with exactly: null"""

    try:
        response = ollama.chat(
            model=MODEL,
            messages=[{"role": "user", "content": prompt}],
            options={"temperature": 0}
        )
        result = response["message"]["content"].strip().strip('"').strip("'")
        if result.lower() == "null" or result not in filenames:
            return None
        return result
    except Exception as e:
        log.error(f"File matching failed: {e}")
        return None


def draft_email(context, name, description, attached_file):
    """
    Generate subject + body using:
    - shared context (the campaign purpose)
    - per-row description (specifics for this recipient)
    - optional attached file reference
    """
    attachment_note = (
        f"\nAttachment: You are attaching '{attached_file}' — reference it naturally in the email."
        if attached_file else
        "\nAttachment: None — do not mention any attachment."
    )

    prompt = f"""You are writing a professional email on behalf of {SENDER_NAME}.

=== SHARED CONTEXT (applies to every email in this campaign) ===
{context}

=== THIS SPECIFIC EMAIL ===
Recipient name: {name}
Specific details for this person: {description}
{attachment_note}

=== INSTRUCTIONS ===
- Combine the shared context and the specific details to write the email
- Sign off with: {SENDER_NAME}
- Keep it natural, professional, and not obviously AI-generated
- Length: appropriate for the context (job applications can be longer, brief follow-ups shorter)

Reply with JSON only, no markdown, no extra text:
{{"subject": "the subject line", "body": "the full email body with proper line breaks using \\n"}}"""

    response = ollama.chat(
        model=MODEL,
        messages=[{"role": "user", "content": prompt}],
        options={"temperature": 0.5}
    )
    text = response["message"]["content"].strip()
    text = re.sub(r"```json|```", "", text).strip()

    try:
        parsed = json.loads(text)
        return parsed.get("subject", "Following up"), parsed.get("body", text)
    except json.JSONDecodeError:
        # Fallback: try to extract subject and body manually
        subject_match = re.search(r'"subject"\s*:\s*"([^"]+)"', text)
        body_match    = re.search(r'"body"\s*:\s*"(.+?)"\s*\}', text, re.DOTALL)
        subject = subject_match.group(1) if subject_match else "Following up"
        body    = body_match.group(1).replace("\\n", "\n") if body_match else text
        return subject, body


def send_email(to_email, subject, body, attachment_path):
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


def init_log():
    if not Path(LOG_FILE).exists():
        with open(LOG_FILE, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["timestamp", "name", "email", "file_hint", "matched_file", "subject", "status", "note"])


def write_log(row):
    with open(LOG_FILE, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["timestamp", "name", "email", "file_hint", "matched_file", "subject", "status", "note"])
        writer.writerow(row)


def countdown_delay(seconds):
    """Visual countdown so you know it's not frozen."""
    for remaining in range(seconds, 0, -1):
        print(f"\r   ⏳ Waiting {remaining}s before next email... ", end="", flush=True)
        time.sleep(1)
    print("\r" + " " * 60 + "\r", end="", flush=True)   # clear the line


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    log.info("=" * 60)
    log.info("AI Email Sender v2 — Context-Driven Edition")
    log.info("=" * 60)

    # 1. Validate credentials
    if not GMAIL_USER or not GMAIL_APP_PASS:
        log.error("Missing GMAIL_USER or GMAIL_APP_PASSWORD in .env")
        return

    # 2. Load context
    context = load_context()
    if not context:
        return
    log.info(f"✓ Loaded context.md ({len(context)} chars)")

    # 3. Load contacts
    try:
        df = pd.read_excel(EXCEL_FILE)
        df.columns = df.columns.str.lower().str.strip()
        log.info(f"✓ Loaded {len(df)} contacts from '{EXCEL_FILE}'")
    except FileNotFoundError:
        log.error(f"Excel file '{EXCEL_FILE}' not found")
        return

    # 4. Map column names
    col_map = {
        "email":       ["email", "e-mail", "email address"],
        "name":        ["name", "full name", "contact name"],
        "description": ["description", "desc", "message", "note", "notes", "short description"],
        "file_hint":   ["file_hint", "file", "attachment", "file hint"],
    }
    cols = {}
    for key, candidates in col_map.items():
        for c in candidates:
            if c in df.columns:
                cols[key] = c
                break

    if "email" not in cols:
        log.error("No 'email' column found in spreadsheet")
        return

    # 5. Scan attachments
    available_files = list_attachments(ATTACHMENTS_DIR)
    log.info(f"✓ Found {len(available_files)} files in '{ATTACHMENTS_DIR}/'")
    log.info(f"✓ Delay between sends: {DELAY_MIN}-{DELAY_MAX} seconds (randomised)")
    log.info("=" * 60)

    init_log()
    stats = {"sent": 0, "skipped": 0, "failed": 0}

    rows_to_process = []
    for i, row in df.iterrows():
        email = str(row.get(cols.get("email", ""), "")).strip()
        if email and email.lower() != "nan":
            rows_to_process.append((i, row))

    total = len(rows_to_process)

    for idx, (i, row) in enumerate(rows_to_process):
        email       = str(row.get(cols.get("email", ""), "")).strip()
        name        = str(row.get(cols.get("name", ""), "there")).strip()
        description = str(row.get(cols.get("description", ""), "")).strip()
        file_hint   = str(row.get(cols.get("file_hint", ""), "")).strip()

        log.info(f"[{idx+1}/{total}] Processing: {name} <{email}>")

        log_entry = {
            "timestamp":    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "name":         name,
            "email":        email,
            "file_hint":    file_hint,
            "matched_file": "",
            "subject":      "",
            "status":       "",
            "note":         "",
        }

        try:
            # ── File matching (only if file_hint is set and not "NA") ────────────
            matched_file    = None
            attachment_path = None

            if not needs_no_attachment(file_hint):
                if not available_files:
                    log.warning(f"   ⚠ Skipped: file '{file_hint}' requested but attachments folder is empty")
                    log_entry.update({"status": "SKIPPED", "note": "Attachments folder empty"})
                    stats["skipped"] += 1
                    write_log(log_entry)
                    continue

                matched_file = match_file(file_hint, available_files)
                if not matched_file:
                    log.warning(f"   ⚠ Skipped: no file matched '{file_hint}'")
                    log_entry.update({"status": "SKIPPED", "note": f"No file matched: '{file_hint}'"})
                    stats["skipped"] += 1
                    write_log(log_entry)
                    continue

                attachment_path = str(Path(ATTACHMENTS_DIR) / matched_file)
                log.info(f"   📎 Attaching: {matched_file}")
            else:
                log.info(f"   📭 No attachment (file_hint = '{file_hint or 'empty'}')")

            # ── Draft email ──────────────────────────────────────────────────────
            subject, body = draft_email(context, name, description, matched_file)
            log.info(f"   ✉  Subject: \"{subject}\"")

            # ── Send ─────────────────────────────────────────────────────────────
            send_email(email, subject, body, attachment_path)
            log.info(f"   ✓  Sent to {email}")
            log_entry.update({"matched_file": matched_file or "", "subject": subject, "status": "SENT"})
            stats["sent"] += 1

        except Exception as e:
            log.error(f"   ✗  Failed: {e}")
            log_entry.update({"status": "FAILED", "note": str(e)})
            stats["failed"] += 1

        write_log(log_entry)

        # ── Anti-spam delay (skip after the last email) ──────────────────────────
        if idx < total - 1:
            delay = random.randint(DELAY_MIN, DELAY_MAX)
            countdown_delay(delay)

    log.info("=" * 60)
    log.info(f"Done. Sent: {stats['sent']}  Skipped: {stats['skipped']}  Failed: {stats['failed']}")
    log.info(f"Full log: '{LOG_FILE}'")
    log.info("=" * 60)


if __name__ == "__main__":
    main()
