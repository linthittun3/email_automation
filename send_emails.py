#!/usr/bin/env python3
"""Bulk email sender: reads recipients from an Excel file and sends via Gmail SMTP."""

import argparse
import csv
import json
import random
import re
import smtplib
import ssl
import sys
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr, parseaddr
from html.parser import HTMLParser
from pathlib import Path

from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"
TEMPLATE_PATH = BASE_DIR / "template.html"
SENT_LOG_PATH = BASE_DIR / "sent_log.csv"
LOG_PATH = BASE_DIR / "send_log.txt"

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


class _HTMLTextExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self._parts = []
        self._skip = 0

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            self._skip += 1
        elif tag in ("p", "br", "div", "tr", "li", "h1", "h2", "h3", "h4"):
            self._parts.append("\n")

    def handle_endtag(self, tag):
        if tag in ("script", "style") and self._skip:
            self._skip -= 1
        elif tag in ("p", "div", "tr", "li", "h1", "h2", "h3", "h4"):
            self._parts.append("\n")

    def handle_data(self, data):
        if not self._skip:
            self._parts.append(data)

    def text(self):
        raw = "".join(self._parts)
        lines = [line.strip() for line in raw.splitlines()]
        return "\n".join(line for line in lines if line)


def log(msg):
    stamp = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{stamp}] {msg}"
    print(line, flush=True)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def load_config():
    if not CONFIG_PATH.exists():
        sys.exit("Missing config.json — fill it in with your Gmail address and app password.")
    with open(CONFIG_PATH, encoding="utf-8") as f:
        cfg = json.load(f)
    for key in ("gmail_address", "app_password", "subject"):
        if not cfg.get(key) or str(cfg[key]).startswith(("your_", "xxxx")):
            sys.exit(f'config.json: set a real value for "{key}".')
    return cfg


def load_template():
    if not TEMPLATE_PATH.exists():
        sys.exit("Missing template.html — create your email HTML body there.")
    return TEMPLATE_PATH.read_text(encoding="utf-8")


def pick_email_column(headers):
    for i, h in enumerate(headers):
        if h and "mail" in str(h).lower():
            return i
    return None


def load_recipients(path):
    if not path.exists():
        sys.exit(f"Excel file not found: {path}")

    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    try:
        headers = [str(c).strip() if c is not None else "" for c in next(rows)]
    except StopIteration:
        sys.exit("Excel file is empty.")

    email_idx = pick_email_column(headers)
    if email_idx is None:
        first = next(rows, None)
        if first and EMAIL_RE.match(str(first[0] or "").strip()):
            headers = ["email"] + [h for h in headers[1:]]
            email_idx = 0
        else:
            sys.exit(
                "Could not find an email column. "
                f"Headers found: {headers}. Rename a column to contain 'email'."
            )

    recipients = []
    seen = set()
    skipped = 0
    for row in rows:
        if row is None or email_idx >= len(row):
            continue
        email = str(row[email_idx] or "").strip().lower()
        if not EMAIL_RE.match(email):
            if any(row):
                skipped += 1
            continue
        if email in seen:
            skipped += 1
            continue
        seen.add(email)
        record = {"email": email}
        for i, h in enumerate(headers):
            if i == email_idx or not h:
                continue
            val = row[i] if i < len(row) else None
            record[h] = "" if val is None else str(val).strip()
        recipients.append(record)

    wb.close()
    if not recipients:
        sys.exit("No valid email addresses found in the Excel file.")
    log(f"Loaded {len(recipients)} unique recipients from {path.name} ({skipped} skipped).")
    return recipients


def load_sent():
    if not SENT_LOG_PATH.exists():
        return set()
    with open(SENT_LOG_PATH, newline="", encoding="utf-8") as f:
        return {row["email"].strip().lower() for row in csv.DictReader(f) if row.get("email")}


def mark_sent(email, ok, error=""):
    new_file = not SENT_LOG_PATH.exists()
    with open(SENT_LOG_PATH, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new_file:
            w.writerow(["email", "status", "error", "timestamp"])
        w.writerow([email, "sent" if ok else "failed", error, time.strftime("%Y-%m-%d %H:%M:%S")])


def render(template, record):
    class _SafeDict(dict):
        def __missing__(self, key):
            return "{" + key + "}"

    try:
        return template.format_map(_SafeDict(record))
    except (ValueError, IndexError):
        # Template may contain CSS braces; treat as literal text.
        body = template
        for k, v in record.items():
            body = body.replace("{" + k + "}", v)
        return body


def build_message(cfg, template, record):
    html = render(template, record)
    text = _HTMLTextExtractor()
    text.feed(html)

    msg = MIMEMultipart("alternative")
    msg["Subject"] = render(cfg["subject"], record)
    msg["From"] = formataddr((cfg.get("from_name", ""), cfg["gmail_address"]))
    msg["To"] = record["email"]
    msg.attach(MIMEText(text.text(), "plain", "utf-8"))
    msg.attach(MIMEText(html, "html", "utf-8"))
    return msg


def main():
    parser = argparse.ArgumentParser(description="Send bulk emails from an Excel file via Gmail.")
    parser.add_argument("excel", nargs="?", default="recipients.xlsx", help="Path to Excel file (default: recipients.xlsx)")
    parser.add_argument("--dry-run", action="store_true", help="Build and validate emails without sending.")
    parser.add_argument("--limit", type=int, default=0, help="Send at most N emails this run (0 = all).")
    parser.add_argument("--fresh", action="store_true", help="Ignore sent_log.csv and resend everyone.")
    args = parser.parse_args()

    cfg = load_config()
    template = load_template()
    recipients = load_recipients(BASE_DIR / args.excel)

    sent = set() if args.fresh else load_sent()
    todo = [r for r in recipients if r["email"] not in sent]
    if args.limit > 0:
        todo = todo[: args.limit]

    if sent:
        log(f"{len(sent)} already processed previously; {len(todo)} left to send.")
    if not todo:
        log("Nothing to send.")
        return
    if args.dry_run:
        log(f"DRY RUN: would send {len(todo)} emails. No mail will be sent.")

    context = ssl.create_default_context()
    server = None
    if not args.dry_run:
        try:
            server = smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context, timeout=30)
            server.login(cfg["gmail_address"], cfg["app_password"])
        except smtplib.SMTPAuthenticationError:
            sys.exit(
                "Gmail login failed. Use an App Password (myaccount.google.com > "
                "Security > 2-Step Verification > App passwords), not your normal password."
            )
        except Exception as e:
            sys.exit(f"Could not connect to Gmail: {e}")

    ok_count = fail_count = 0
    try:
        for i, record in enumerate(todo, 1):
            email = record["email"]
            try:
                msg = build_message(cfg, template, record)
                if args.dry_run:
                    log(f"[{i}/{len(todo)}] OK (dry run) -> {email}")
                    ok_count += 1
                    continue
                server.sendmail(cfg["gmail_address"], [email], msg.as_string())
                mark_sent(email, True)
                ok_count += 1
                log(f"[{i}/{len(todo)}] Sent -> {email}")
            except Exception as e:
                fail_count += 1
                mark_sent(email, False, str(e))
                log(f"[{i}/{len(todo)}] FAILED -> {email}: {e}")

            if i < len(todo) and not args.dry_run:
                time.sleep(random.uniform(cfg.get("delay_min_seconds", 1), cfg.get("delay_max_seconds", 2)))
                every = int(cfg.get("pause_every", 100))
                if every > 0 and i % every == 0:
                    pause = int(cfg.get("pause_seconds", 60))
                    log(f"Pause {pause}s after {i} emails...")
                    time.sleep(pause)
    finally:
        if server is not None:
            try:
                server.quit()
            except Exception:
                pass

    log(f"Done. Sent: {ok_count}, Failed: {fail_count}.")


if __name__ == "__main__":
    main()
