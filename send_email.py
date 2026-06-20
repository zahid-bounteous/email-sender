"""
AI Email Sender — Local version
Uses: Phi-4 Mini via Ollama · Gmail SMTP · Excel/CSV input · File attachments
"""

import os
import csv
import smtplib
import logging
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
ATTACHMENTS_DIR = "attachments"
LOG_FILE        = "sent_log.csv"
MODEL           = "phi4-mini"

GMAIL_USER      = os.getenv("GMAIL_USER")
GMAIL_APP_PASS  = os.getenv("GMAIL_APP_PASSWORD")
SENDER_NAME     = os.getenv("SENDER_NAME", "The Team")

EMAIL_CONTEXT = f"""
You are a professional email assistant. Write clear, concise, friendly emails.
Keep emails to 3-5 sentences. Always sign off with the sender's name: {SENDER_NAME}.
Never mention that the email was AI-generated.
""".strip()

# ── Logging ────────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S"
)
log = logging.getLogger(__name__)

# ── Helpers ────────────────────────────────────────────────────────────────────

def list_attachments(folder):
    p = Path(folder)
    if not p.exists():
        log.warning(f"Attachments folder '{folder}' not found.")
        return []
    return [f.name for f in p.iterdir() if f.is_file()]


def match_file(hint, filenames):
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


def is_description_clear(description):
    if not description or len(description.strip()) < 5:
        return False, "Description is empty or too short"

    prompt = f"""Is this instruction clear enough to write a professional email?

Instruction: "{description}"

Reply with JSON only, no extra text:
{{"clear": true}} if it is clear enough.
{{"clear": false, "reason": "explain why briefly"}} if it is too vague."""

    try:
        response = ollama.chat(
            model=MODEL,
            messages=[{"role": "user", "content": prompt}],
            options={"temperature": 0}
        )
        text = response["message"]["content"].strip()
        if '"clear": true' in text or "'clear': true" in text:
            return True, ""
        import re
        match = re.search(r'"reason":\s*"([^"]+)"', text)
        reason = match.group(1) if match else "Description too vague"
        return False, reason
    except Exception as e:
        log.error(f"Clarity check failed: {e}")
        return False, f"Clarity check error: {e}"


def draft_email(name, description, attached_file):
    attachment_note = f"\nNote: You are attaching the file '{attached_file}' — mention it naturally." if attached_file else ""

    prompt = f"""Write a professional email to {name}.

Context: {EMAIL_CONTEXT}

Task: {description}{attachment_note}

Reply with JSON only, no extra text:
{{"subject": "the subject line", "body": "the full email body"}}"""

    import json, re
    response = ollama.chat(
        model=MODEL,
        messages=[{"role": "user", "content": prompt}],
        options={"temperature": 0.4}
    )
    text = response["message"]["content"].strip()
    text = re.sub(r"```json|```", "", text).strip()
    try:
        parsed = json.loads(text)
        return parsed.get("subject", "Hello"), parsed.get("body", text)
    except json.JSONDecodeError:
        return "Following up", text


def send_email(to_email, subject, body, attachment_path):
    msg = MIMEMultipart()
    msg["From"]    = f"{SENDER_NAME} <{GMAIL_USER}>"
    msg["To"]      = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

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
        with open(LOG_FILE, "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["timestamp", "name", "email", "file_hint", "matched_file", "subject", "status", "note"])

def write_log(row):
    with open(LOG_FILE, "a", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["timestamp", "name", "email", "file_hint", "matched_file", "subject", "status", "note"])
        writer.writerow(row)


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    log.info("=" * 55)
    log.info("AI Email Sender — Phi-4 Mini + Gmail SMTP")
    log.info("=" * 55)

    if not GMAIL_USER or not GMAIL_APP_PASS:
        log.error("Missing GMAIL_USER or GMAIL_APP_PASSWORD in .env — aborting")
        return

    try:
        df = pd.read_excel(EXCEL_FILE)
        df.columns = df.columns.str.lower().str.strip()
        log.info(f"Loaded {len(df)} contacts from '{EXCEL_FILE}'")
    except FileNotFoundError:
        log.error(f"Excel file '{EXCEL_FILE}' not found")
        return

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

    available_files = list_attachments(ATTACHMENTS_DIR)
    log.info(f"Found {len(available_files)} files in '{ATTACHMENTS_DIR}/'")

    init_log()
    stats = {"sent": 0, "flagged": 0, "failed": 0}

    for i, row in df.iterrows():
        email       = str(row.get(cols.get("email", ""), "")).strip()
        name        = str(row.get(cols.get("name", ""), "Unknown")).strip()
        description = str(row.get(cols.get("description", ""), "")).strip()
        file_hint   = str(row.get(cols.get("file_hint", ""), "")).strip()

        if not email or email == "nan":
            continue

        log.info(f"[{i+1}/{len(df)}] Processing: {name} <{email}>")

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
            # Step 1 — Check description clarity
            clear, reason = is_description_clear(description)
            if not clear:
                log.warning(f"  ⚠ Flagged: {reason}")
                log_entry.update({"status": "FLAGGED", "note": reason})
                stats["flagged"] += 1
                write_log(log_entry)
                continue

            # Step 2 — Match file
            matched_file    = None
            attachment_path = None

            if file_hint and available_files:
                matched_file = match_file(file_hint, available_files)
                if matched_file:
                    attachment_path = str(Path(ATTACHMENTS_DIR) / matched_file)
                    log.info(f"  📎 Matched file: {matched_file}")
                else:
                    log.warning(f"  ⚠ Flagged: No file matched hint '{file_hint}'")
                    log_entry.update({"status": "FLAGGED", "note": f"No file matched: '{file_hint}'"})
                    stats["flagged"] += 1
                    write_log(log_entry)
                    continue
            elif file_hint and not available_files:
                log.warning(f"  ⚠ Flagged: Attachments folder empty")
                log_entry.update({"status": "FLAGGED", "note": "Attachments folder empty"})
                stats["flagged"] += 1
                write_log(log_entry)
                continue

            # Step 3 — Draft email
            subject, body = draft_email(name, description, matched_file)
            log.info(f"  ✉  Drafted: \"{subject}\"")

            # Step 4 — Send
            send_email(email, subject, body, attachment_path)
            log.info(f"  ✓  Sent to {email}")
            log_entry.update({"matched_file": matched_file or "", "subject": subject, "status": "SENT"})
            stats["sent"] += 1

        except Exception as e:
            log.error(f"  ✗  Failed: {e}")
            log_entry.update({"status": "FAILED", "note": str(e)})
            stats["failed"] += 1

        write_log(log_entry)

    log.info("=" * 55)
    log.info(f"Done. Sent: {stats['sent']}  Flagged: {stats['flagged']}  Failed: {stats['failed']}")
    log.info(f"Log saved to '{LOG_FILE}'")
    log.info("=" * 55)


if __name__ == "__main__":
    main()
