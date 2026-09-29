# WhatsApp Birthday / Anniversary Sender

Sends approved WhatsApp templates through the Meta WhatsApp Cloud API to every customer in the Excel file whose birthday or anniversary is today:

| Occasion | Template sent |
|---|---|
| Birthday | `birthday_wish_discount` |
| Anniversary | `anniversary_wish_discount` |

## Setup

```bash
pip install -r requirements.txt
```

Copy `.env.example` to `.env` and fill in:

```
WHATSAPP_ACCESS_TOKEN=...          # permanent (System User) token
WHATSAPP_PHONE_NUMBER_ID=...       # WhatsApp > API Setup
```

Keep `.env` out of version control.

## Daily use

```bash
python send_whatsapp.py
```

This sends to everyone who is due today. Running it again the same day is safe: anyone who already got today's message is skipped (see [No repeat messages](#no-repeat-messages)).

To preview without sending anything:

```bash
python send_whatsapp.py --dry-run                     # today
python send_whatsapp.py --dry-run --date 2026-10-05   # pretend it's another date
```

`--date` only works with `--dry-run`, so real messages always go out on the real date.

## The Excel file

Put `Salon Customers Database.xlsx` in the same folder as `send_whatsapp.py`. Every sheet is read. Each sheet needs a header row with these column names, in any order. Extra columns such as `Middle_Name` or `Notes` are ignored, and a sheet without these columns is skipped with a warning.

| Column | Example | Notes |
|---|---|---|
| `First_Name` | Riya | required for a row to be sent |
| `Last_Name` | Sen | optional; `{{1}}` is `First_Name Last_Name` |
| `Mobile_1` | 9876543210 | 10 digits; `+91` is added |
| `Mobile_2` | | used if `Mobile_1` is blank or invalid |
| `Birth_Day`, `Birth_Month` | 29, September | month as a name (`September` or `Sep`) |
| `Anniversary_Day`, `Anniversary_Month` | 1, October | may be blank |

Birthdays and anniversaries are checked separately. Someone who has both today gets both templates, as two messages.

The script reads the last saved version of the file, so save it in Excel before running.

## What it prints

One line per person: name, phone, template, and the result (sent, skipped, or failed with the reason). Then a summary: birthday and anniversary messages sent, how many were skipped (no or invalid phone, already sent today), and each failure with its reason.

"Sent" means WhatsApp accepted the message. Meta can still drop it afterwards; error 131049, for example, means the recipient hit Meta's marketing message limit. Those drops only show up in webhook delivery statuses, not in this output.

## No repeat messages

- Every message is recorded in `sent_log.json` (date, phone, occasion) *before* it's sent. A later run the same day skips any person and occasion already in the log, so running twice never sends twice.
- If WhatsApp refuses a message, or the internet is down and it never left the computer, the record is removed. The next run tries again.
- If the run dies or times out while a message is in flight, the record stays as `"status": "pending"` and that person isn't retried today, because the message may already have gone out. To send it anyway, delete that entry from `sent_log.json`.
- Only one run can send at a time. A second run started meanwhile stops with an error. If a run was killed and left `send_whatsapp.lock` behind, delete that file.

The log only ever holds today's messages. Each run first removes entries from earlier dates, then adds today's; `sent_log_data.js` (the copy `dashboard.html` reads) always matches it. A `--dry-run` never changes either file.

Keep `sent_log.json` during the day: it is the duplicate check. It holds customer numbers, so keep it out of version control.

## Dashboard

Double-click `dashboard.html` to open it in a browser; no server is needed. It shows today's messages (name, occasion, phone, status, time) from `sent_log_data.js` and re-reads that file every 5 seconds, so it can stay open on a screen. "Last updated" shows when it last checked; the Refresh button checks immediately.

## Change what is sent

- **Offer text:** edit `OFFER_LINE` ({{2}}) at the top of `send_whatsapp.py`. {{1}} is always the customer's name from the Excel file.
- **Who receives it:** edit the Excel file.

Both templates must show **Active** in WhatsApp Manager; a template that is in review or rejected can't be sent.
