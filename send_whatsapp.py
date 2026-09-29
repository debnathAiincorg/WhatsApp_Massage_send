# ============ EDIT THIS VALUE TO CHANGE THE MESSAGE ============
OFFER_LINE = "enjoy 20% discount on any service"   # {{2}} - offer line ({{1}} is the customer's name from Excel)
# ===================================================================

# Send birthday and anniversary WhatsApp templates via the Meta WhatsApp Cloud API
# to everyone in the Excel file whose birthday / anniversary is today.
#
# Usage:
#     python send_whatsapp.py                      send today's messages
#     python send_whatsapp.py --dry-run            show who would get what, send nothing
#     python send_whatsapp.py --dry-run --date YYYY-MM-DD
#                                                  same, pretending today is that date
#
# Credentials come from .env (WHATSAPP_ACCESS_TOKEN, WHATSAPP_PHONE_NUMBER_ID).

import argparse
import calendar
import json
import os
import re
import sys
from datetime import date, datetime
from pathlib import Path

import requests
from dotenv import load_dotenv

GRAPH_API_VERSION = "v20.0"
REQUEST_TIMEOUT = 30  # seconds
LANGUAGE_CODE = "en"  # must match the templates' language in WhatsApp Manager

TEMPLATES = {
    "birthday": "birthday_wish_discount",
    "anniversary": "anniversary_wish_discount",
}

HERE = Path(__file__).resolve().parent
EXCEL_FILE = HERE / "Salon Customers Database.xlsx"  # every sheet is read
SENT_LOG = HERE / "sent_log.json"                    # today's sends, to avoid repeats
SENT_LOG_JS = HERE / "sent_log_data.js"              # same data for dashboard.html (works from file://)
LOCK_FILE = HERE / "send_whatsapp.lock"              # stops two runs sending at once
DEFAULT_COUNTRY_CODE = "91"                          # Excel numbers are stored without it

# Meta error codes worth explaining in plain language.
KNOWN_ERRORS = {
    190: "Access token is invalid or expired. Generate a new token and update WHATSAPP_ACCESS_TOKEN in .env.",
    100: "Invalid parameter. Check WHATSAPP_PHONE_NUMBER_ID in .env and the phone number in the Excel file.",
    131030: "Recipient number is not in the allowed list. In test mode, add it under 'To' in the Meta app's WhatsApp > API Setup page.",
    131026: "Message undeliverable. The recipient may not have WhatsApp or has not accepted the latest terms.",
    131049: "This recipient has hit Meta's marketing message limit (shared across all businesses, not just yours). Wait a day or more before retrying - this isn't a code or account problem.",
    131047: "More than 24 hours have passed since the recipient last replied. Only approved templates can be sent.",
    131031: "The WhatsApp Business account is locked or restricted.",
    133010: "The sender phone number is not registered with the Cloud API.",
    130429: "Rate limit hit. Wait a moment and try again.",
    80007: "Rate limit hit for this WhatsApp Business account. Wait and try again.",
    132000: "Number of template variables doesn't match the template (expected {{1}} name and {{2}} offer line).",
    132001: "Template not found or not approved yet. Check its status in WhatsApp Manager (it must be Active).",
    132005: "Template text is too long after filling in the variables. Shorten OFFER_LINE or the customer's name.",
    132007: "Template content violates WhatsApp policy.",
    132012: "Template variable format doesn't match what the template expects.",
    132015: "Template is paused due to low quality. Check its status in WhatsApp Manager.",
    132016: "Template is disabled. Check its status in WhatsApp Manager.",
}


class RequestFailed(Exception):
    """The request never got an HTTP response (timeout, no connection, ...).

    maybe_sent is False only when no connection was ever made, so WhatsApp certainly didn't get it.
    """

    def __init__(self, message, maybe_sent):
        super().__init__(message)
        self.maybe_sent = maybe_sent


def never_connected(exc):
    """True if requests failed before any data reached the server (DNS failure, refused, connect timeout)."""
    from urllib3.exceptions import ConnectTimeoutError, NewConnectionError

    if isinstance(exc, requests.exceptions.ConnectTimeout):
        return True
    reason = getattr(exc.args[0], "reason", None) if exc.args else None
    return isinstance(reason, (NewConnectionError, ConnectTimeoutError))


def load_config():
    """Load credentials from the .env file next to this script."""
    load_dotenv(HERE / ".env")

    config = {
        "token": os.getenv("WHATSAPP_ACCESS_TOKEN", "").strip(),
        "phone_number_id": os.getenv("WHATSAPP_PHONE_NUMBER_ID", "").strip(),
    }
    missing = [
        name
        for name, key in (("WHATSAPP_ACCESS_TOKEN", "token"), ("WHATSAPP_PHONE_NUMBER_ID", "phone_number_id"))
        if not config[key]
    ]
    if missing:
        sys.exit(f"Error: missing required value(s) in .env: {', '.join(missing)}. See .env.example.")

    return config


def build_payload(to, template_name, variables):
    return {
        "messaging_product": "whatsapp",
        "to": to,
        "type": "template",
        "template": {
            "name": template_name,
            "language": {"code": LANGUAGE_CODE},
            "components": [
                {
                    "type": "body",
                    "parameters": [{"type": "text", "text": str(v)} for v in variables],
                }
            ],
        },
    }


def send_message(token, phone_number_id, payload):
    """POST the payload and return the requests.Response."""
    url = f"https://graph.facebook.com/{GRAPH_API_VERSION}/{phone_number_id}/messages"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    return requests.post(url, headers=headers, json=payload, timeout=REQUEST_TIMEOUT)


def send_template(config, recipient, template_name, variables):
    """Send one template message. Returns (response, parsed body); raises RequestFailed on network errors."""
    payload = build_payload(recipient, template_name, variables)
    try:
        response = send_message(config["token"], config["phone_number_id"], payload)
    except requests.exceptions.Timeout as exc:
        raise RequestFailed(f"request timed out after {REQUEST_TIMEOUT}s. Check your connection and try again.",
                            maybe_sent=not never_connected(exc))
    except requests.exceptions.ConnectionError as exc:
        raise RequestFailed("could not reach graph.facebook.com. Check your internet connection.",
                            maybe_sent=not never_connected(exc))
    except requests.exceptions.RequestException as exc:
        raise RequestFailed(f"request failed: {exc}", maybe_sent=not never_connected(exc))

    try:
        body = response.json()
    except ValueError:
        body = {"raw": response.text}
    return response, body


def message_id_of(body):
    return (body.get("messages") or [{}])[0].get("id", "n/a")


def explain_error(body):
    """Build a readable message from a Graph API error response."""
    error = body.get("error", {}) if isinstance(body, dict) else {}
    code = error.get("code")
    details = error.get("error_data", {}).get("details") or error.get("message", "Unknown error")
    hint = KNOWN_ERRORS.get(code)
    lines = [f"API error {code}: {details}"]
    if hint:
        lines.append(f"Hint: {hint}")
    return "\n".join(lines)


# ------------------------------------------------------------------ Excel file

MONTHS = {name.lower(): i for i, name in enumerate(calendar.month_name) if name}
MONTHS.update({name.lower(): i for i, name in enumerate(calendar.month_abbr) if name})
MONTHS["sept"] = 9

REQUIRED_COLUMNS = ("First_Name", "Mobile_1", "Birth_Day", "Birth_Month", "Anniversary_Day", "Anniversary_Month")


def cell_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def parse_day_month(day, month):
    """Return (day, month_number), or None if either cell is blank or unreadable."""
    day_text, month_text = cell_text(day), cell_text(month).lower().rstrip(".")
    if not day_text or not month_text:
        return None
    if not day_text.isdigit() or month_text not in MONTHS:
        raise ValueError(f"unreadable date {day_text!r} {cell_text(month)!r}")
    return int(day_text), MONTHS[month_text]


def excel_phone(value):
    """Turn an Excel phone cell into API digits (country code added), or '' if blank/invalid."""
    digits = re.sub(r"\D", "", cell_text(value))
    if len(digits) == 11 and digits.startswith("0"):
        digits = digits[1:]
    if len(digits) == 10:
        digits = DEFAULT_COUNTRY_CODE + digits
    return digits if 11 <= len(digits) <= 15 else ""


def load_customers(path):
    """Yield one dict per data row from every sheet, looking columns up by header name."""
    from openpyxl import load_workbook

    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        for sheet in workbook.worksheets:
            rows = sheet.iter_rows(values_only=True)
            header = [cell_text(h) for h in next(rows, [])]
            missing = [c for c in REQUIRED_COLUMNS if c not in header]
            if missing:
                print(f"WARNING: sheet '{sheet.title}' skipped - missing column(s): {', '.join(missing)}")
                continue
            col = {name: header.index(name) for name in header if name}

            def get(row, name):
                i = col.get(name)
                return row[i] if i is not None and i < len(row) else None

            for row_number, row in enumerate(rows, start=2):
                if not any(cell_text(v) for v in row):
                    continue
                yield {
                    "where": f"{sheet.title} row {row_number}",
                    "name": " ".join(p for p in (cell_text(get(row, "First_Name")), cell_text(get(row, "Last_Name"))) if p),
                    "mobile_1": get(row, "Mobile_1"),
                    "mobile_2": get(row, "Mobile_2"),
                    "birthday": (get(row, "Birth_Day"), get(row, "Birth_Month")),
                    "anniversary": (get(row, "Anniversary_Day"), get(row, "Anniversary_Month")),
                }
    finally:
        workbook.close()


def find_todays_sends(customers, today):
    """Return one entry per (person, occasion) matching today; a person can appear twice."""
    sends, warnings = [], []
    for person in customers:
        for occasion in TEMPLATES:  # birthday, anniversary - checked independently
            try:
                day_month = parse_day_month(*person[occasion])
            except ValueError as exc:
                warnings.append(f"{person['where']} ({person['name']}): {occasion} {exc}, ignored")
                continue
            if day_month == (today.day, today.month):
                phone = excel_phone(person["mobile_1"]) or excel_phone(person["mobile_2"])
                raw = [cell_text(person[k]) for k in ("mobile_1", "mobile_2") if cell_text(person[k])]
                no_phone = f"invalid phone number ({', '.join(raw)})" if raw else "no phone number"
                sends.append({**person, "occasion": occasion, "phone": phone, "no_phone": no_phone})
    return sends, warnings


# ------------------------------------------------------- duplicate protection
#
# sent_log.json holds one entry per send attempt today that may have reached WhatsApp
# (entries from earlier dates are removed at the start of each run).
# An entry is written as "pending" BEFORE the API call and becomes "sent" after
# it succeeds (or is removed if the API refuses it). Any entry for the same
# date + phone + occasion blocks a repeat, so even a crash mid-send can't cause
# a second message. The lock file stops two runs from sending at the same time.


def load_sent_log():
    if not SENT_LOG.exists():
        return []
    try:
        entries = json.loads(SENT_LOG.read_text(encoding="utf-8"))
    except (ValueError, OSError) as exc:
        sys.exit(f"Error: cannot read {SENT_LOG.name} ({exc}). Fix or move it; refusing to run without the duplicate check.")
    if not isinstance(entries, list):
        sys.exit(f"Error: {SENT_LOG.name} should contain a JSON list. Fix or move it.")
    return entries


def write_atomically(path, text):
    tmp = path.with_suffix(".tmp")
    tmp.write_text(text, encoding="utf-8")
    tmp.replace(path)


def save_sent_log(entries):
    data = json.dumps(entries, indent=2)
    write_atomically(SENT_LOG, data)
    # A copy for dashboard.html, loaded with <script src> because browsers block fetch() on file:// pages.
    # "window." rather than "const" so the dashboard can load it again on Refresh.
    write_atomically(SENT_LOG_JS, f"window.sentLogData = {data};\n")


def acquire_lock():
    try:
        fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        started = LOCK_FILE.read_text(encoding="utf-8", errors="replace").strip() or "unknown time"
        sys.exit(
            f"Error: another run is already sending (started {started}).\n"
            f"If no other run is active (e.g. the last one was killed), delete {LOCK_FILE.name} and try again."
        )
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        f.write(f"{datetime.now().isoformat(timespec='seconds')} by process {os.getpid()}")


def release_lock():
    try:
        LOCK_FILE.unlink()
    except FileNotFoundError:
        pass


# ------------------------------------------------------------------------ run


def run(today, dry_run, excel_path):
    if not excel_path.exists():
        sys.exit(f"Error: Excel file not found: {excel_path}")
    if not OFFER_LINE.strip():
        sys.exit("Error: OFFER_LINE cannot be empty.")
    config = None if dry_run else load_config()

    sends, warnings = find_todays_sends(load_customers(excel_path), today)
    today_str = today.isoformat()
    # The log only ever holds today's entries: earlier dates are dropped (and saved, on a real run)
    # before anything is sent. Today's own entries stay, so re-runs today still skip them.
    full_log = load_sent_log()
    sent_log = [e for e in full_log if isinstance(e, dict) and e.get("date") == today_str]
    removed_old = len(full_log) - len(sent_log)
    if removed_old and not dry_run:
        save_sent_log(sent_log)
    logged = {(e.get("date"), e.get("phone"), e.get("occasion")): e for e in sent_log}

    mode = "DRY RUN - nothing will be sent" if dry_run else "sending"
    print(f"Birthday/anniversary messages for {today.day} {calendar.month_name[today.month]} ({today_str}), {mode}")
    print(f"Excel: {excel_path.name}  ->  {len(sends)} message(s) due today")
    if removed_old:
        print(f"{SENT_LOG.name}: {removed_old} entr{'y' if removed_old == 1 else 'ies'} from earlier dates "
              + ("would be removed on a real run" if dry_run else "removed"))
    print()
    for w in warnings:
        print(f"WARNING: {w}")
    if warnings:
        print()

    counts = {"birthday": 0, "anniversary": 0}
    skipped_no_phone, skipped_duplicate, failures = [], [], []

    for s in sends:
        template_name = TEMPLATES[s["occasion"]]
        label = f"[{s['occasion'].upper():11}] {s['name'] or '(no name)'} | " \
                f"{'+' + s['phone'] if s['phone'] else 'no phone'} | {template_name} | {s['where']}"

        if not s["phone"]:
            print(f"{label} -> SKIPPED: {s['no_phone']}")
            skipped_no_phone.append(s)
            continue
        if not s["name"]:
            print(f"{label} -> FAILED: no name in First_Name")
            failures.append((s, "no name in First_Name"))
            continue
        key = (today_str, s["phone"], s["occasion"])
        if key in logged:
            if logged[key].get("status") == "pending":
                print(f"{label} -> SKIPPED: an earlier run was interrupted while sending this one; "
                      f"it may have gone out (see {SENT_LOG.name})")
            else:
                print(f"{label} -> SKIPPED: already sent today")
            skipped_duplicate.append(s)
            continue
        if dry_run:
            print(f"{label} -> WOULD SEND  {{{{1}}}}={s['name']!r}  {{{{2}}}}={OFFER_LINE!r}")
            counts[s["occasion"]] += 1
            logged[key] = {"status": "planned"}
            continue

        entry = {"date": today_str, "phone": s["phone"], "occasion": s["occasion"], "name": s["name"],
                 "status": "pending", "sent_at": datetime.now().isoformat(timespec="seconds")}
        sent_log.append(entry)
        logged[key] = entry
        save_sent_log(sent_log)  # recorded before sending, so a crash can't cause a repeat

        try:
            response, body = send_template(config, s["phone"], template_name, [s["name"], OFFER_LINE])
        except RequestFailed as exc:
            if exc.maybe_sent:
                # The request may have reached WhatsApp without us seeing the reply: the entry
                # stays "pending" and blocks a retry today rather than risk a duplicate.
                reason = f"{exc} It may have been delivered, so it stays 'pending' in {SENT_LOG.name} " \
                         f"and won't be retried today (delete that entry to allow a retry)."
            else:
                sent_log.remove(entry)  # never left this computer: safe to retry on the next run
                del logged[key]
                save_sent_log(sent_log)
                reason = str(exc)
        else:
            if response.ok:
                entry.update(status="sent", message_id=message_id_of(body))
                save_sent_log(sent_log)
                print(f"{label} -> SENT (id: {entry['message_id']})")
                counts[s["occasion"]] += 1
                continue
            # WhatsApp refused it, so nothing went out: forget it so a later run can retry.
            sent_log.remove(entry)
            del logged[key]
            save_sent_log(sent_log)
            reason = explain_error(body).replace("\n", " | ")
        print(f"{label} -> FAILED: {reason}")
        failures.append((s, reason))

    verb = "would be sent" if dry_run else "sent"
    print("\n=========== SUMMARY ===========")
    print(f"Birthday messages {verb}:     {counts['birthday']}")
    print(f"Anniversary messages {verb}:  {counts['anniversary']}")
    print(f"Skipped (no phone number):  {len(skipped_no_phone)}")
    for s in skipped_no_phone:
        print(f"    - {s['name']} ({s['occasion']}, {s['where']}): {s['no_phone']}")
    print(f"Skipped (already sent today): {len(skipped_duplicate)}")
    print(f"Failed:                     {len(failures)}")
    for s, reason in failures:
        print(f"    - {s['name']} +{s['phone']} ({s['occasion']}): {reason}")
    if not dry_run and sum(counts.values()):
        print("\nNote: 'sent' means WhatsApp accepted the message. Meta can still drop it later "
              "(e.g. error 131049); that only shows up in webhook delivery statuses.")
    return 1 if failures else 0


def main():
    parser = argparse.ArgumentParser(
        description="Send the birthday / anniversary WhatsApp template to everyone in the Excel file "
                    "whose day is today. Safe to re-run: anyone already sent today is skipped "
                    f"(tracked in {SENT_LOG.name}).",
    )
    parser.add_argument("--dry-run", action="store_true", help="list who would get which message; send nothing")
    parser.add_argument("--date", metavar="YYYY-MM-DD", help="with --dry-run: pretend today is this date")
    parser.add_argument("--excel", metavar="PATH", help=f"Excel file to read (default: {EXCEL_FILE.name})")
    args = parser.parse_args()

    if args.date and not args.dry_run:
        parser.error("--date only works with --dry-run, so real messages always go out on the real date")
    try:
        today = date.fromisoformat(args.date) if args.date else date.today()
    except ValueError:
        parser.error(f"--date {args.date!r} is not a valid YYYY-MM-DD date")
    excel_path = Path(args.excel).resolve() if args.excel else EXCEL_FILE

    if args.dry_run:
        return run(today, True, excel_path)
    acquire_lock()
    try:
        return run(today, False, excel_path)
    finally:
        release_lock()


if __name__ == "__main__":
    sys.exit(main())
