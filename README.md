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
| `First_Name` | FirstName | required for a row to be sent |
| `Last_Name` | SecondName | optional; `{{1}}` is `First_Name Last_Name` |
| `Mobile_1` | 1234567890 | 10 digits; `+91` is added |
| `Mobile_2` | | used if `Mobile_1` is blank or invalid |
| `Birth_Day`, `Birth_Month` | 29, September | month as a name (`September` or `Sep`) |
| `Anniversary_Day`, `Anniversary_Month` | 1, October | may be blank |

Birthdays and anniversaries are checked separately. Someone who has both today gets both templates, as two messages.

The script reads the last saved version of the file, so save it in Excel before running.

## What it prints

One line per person: name, phone, template, and the result (sent, skipped, or failed with the reason). Then a summary: birthday and anniversary messages sent, how many were skipped (no or invalid phone, already sent today), and each failure with its reason.

"Sent" means WhatsApp accepted the message. Meta can still drop it afterwards; error 131049, for example, means the recipient hit Meta's marketing message limit. Those drops don't show in this output; the delivery reports at the end of the run show them (see [Delivery reports](#delivery-reports)).

## No repeat messages

- Every message is recorded in `sent_log.json` (date, phone, occasion) *before* it's sent. What a later run the same day does depends on that entry's status:

  | Status in the log | Next run |
  |---|---|
  | `delivered` / `read` | skipped: done |
  | `sent`, last try less than 2 hours ago | skipped: waiting for its delivery report |
  | `sent`, 2 hours or more, 1 try so far | **sent once more** (try 2 of 2) |
  | `sent`, 2 hours or more, 2 tries made | marked **`unconfirmed`**, never tried again |
  | `pending` (a run stopped mid-send; it may have gone out) | never resent; flagged for a manual check |
  | `failed` (Meta dropped it, e.g. error 131049) | never resent; flagged for a manual check |
  | `unconfirmed` | never resent; flagged for a manual check |

  The limits are `MAX_ATTEMPTS` and `RETRY_AFTER` near the top of `send_whatsapp.py`. A retry overwrites the same entry (no duplicate entries): `attempts` counts the tries, and `earlier_message_ids` keeps the old message ids, so a late "delivered" report for an earlier try still marks it delivered. A retry can mean a customer gets the message twice if the first one did arrive but no report came back; the cap keeps that to at most one extra message.
- To clear a manual check: if you've confirmed the customer got it, set its `status` to `delivered`; to send it again, delete that entry. Either way, edit `sent_log.json` only while no run is going.
- If WhatsApp refuses a message, or the internet is down and it never left the computer, the entry goes back to how it was before that try. The next run tries again.
- A person listed twice in the Excel file still gets one message per run.
- Only one run can send at a time. A second run started meanwhile stops with an error. If a run was killed and left `send_whatsapp.lock` behind, delete that file.

The log only ever holds today's messages. Each run first removes entries from earlier dates, then adds today's; `sent_log_data.js` (the copy `dashboard.html` reads) always matches it. A `--dry-run` never changes either file.

Every run first reads today's list from the Excel file, then checks `sent_log.json` for entries dated today:

- **None:** a new day, so it runs normally.
- **Some:** it compares everyone due today with them, names first (it prints any name that isn't there), then mobile numbers. Anyone whose mobile number for that occasion isn't in the log (a new person, or a changed number) gets a message; everyone else is handled by the table above. If nobody needs a message or a status change, it prints "Nothing to do" and stops.

For example, if A and B were delivered this morning and the Excel file now shows A, B and C due today, only C gets a message.

`--force` skips the "Nothing to do" stop, but each person is still handled by the table above.

Keep `sent_log.json` during the day: it is the duplicate check. It holds customer numbers, so keep it out of version control.

## Dashboard

Double-click `dashboard.html` to open it in a browser; no server is needed. It shows today's messages (name, occasion, phone, status, time) from `sent_log_data.js` and re-reads that file every 5 seconds, so it can stay open on a screen. "Last updated" shows when it last checked; the Refresh button checks immediately.

Above the table, **Last run summary** shows the same counts as the script's SUMMARY (birthday/anniversary sent, skipped, failed), with who was skipped or failed and why. It comes from `run_summary_data.js`, which each real run (not `--dry-run`) writes when it finishes, and it describes the most recent run today that actually ran (a run that stops with "Nothing to do" doesn't replace it).

The Status column shows what is known about each message:

| Status | Colour | Meaning |
|---|---|---|
| `sent` | grey | WhatsApp accepted it; no delivery report yet |
| `delivered` / `read` | green | confirmed on the customer's phone (with the time) |
| `Failed · Error 131049` | red | Meta dropped it, with the error code; what the code means is in the key below the table. Check manually |
| `unconfirmed` | amber | no delivery report after 2 tries. Check manually |
| `pending` | amber | a run stopped while sending it; it may or may not have gone out. Check manually |

The key below the table gets one line per error code that appears in today's messages, e.g. "Error 131049 — Meta is limiting marketing messages to that customer's number…". The text comes from the `ERROR_CODES` table near the top of the script in `dashboard.html`; add a line there for a new code. A code not in the table gets a line saying no explanation has been added yet.

**Last run summary** also lists who is waiting for a delivery report and who needs a manual check.

Delivered, read and failed come from the delivery reports the script collects at the end of each run (below). Without them, messages stay `sent`.

## Delivery reports

After sending, the script waits up to 90 seconds for Meta's reports on each message, prints them, and writes them into `sent_log.json`, so the dashboard turns each row green or red:

```
Waiting up to 90 s for delivery reports (Ctrl+C to stop waiting)...
    DELIVERED FirstName SecondName +911234567890 (birthday)
    FAILED    FirstName SecondName +910987654321 (anniversary): error 131049 ...
Delivery reports: 1 delivered, 1 failed, 0 with no report yet
```

It stops waiting as soon as every message has a report. `--wait 30` changes the limit for one run (`--wait 0` skips it); `DELIVERY_WAIT_SECONDS` at the top of the script changes the default.

**How:** Meta sends reports to a web address (a *webhook*). While the script runs, it starts a small web server on this PC and an ngrok tunnel that gives it a fixed public address, `https://<your ngrok domain>/webhook`. When the script ends, both stop. Nothing runs in the background, and nothing is installed as a service. Only requests signed with your app secret are accepted, so nobody else can fake a report.

### One-time setup

1. **ngrok account.** Sign up free at ngrok.com. From its dashboard copy your **authtoken** (Getting Started > Your Authtoken) and your **dev domain** (Domains, e.g. `something.ngrok-free.dev`; it's yours permanently).
2. **Install ngrok:** `winget install ngrok.ngrok`, then close and reopen the terminal.
3. **Fill in `.env`:**
   - `NGROK_AUTHTOKEN=` and `NGROK_DOMAIN=` from step 1.
   - `WHATSAPP_APP_SECRET=`: Meta App Dashboard > your app > **App settings > Basic** > App secret > **Show**.
   - `WHATSAPP_WABA_ID` and `WHATSAPP_WEBHOOK_VERIFY_TOKEN` are already filled in (see `.env.example`).
4. **Register the address with Meta.** Run:
   ```bash
   python send_whatsapp.py --setup-webhook
   ```
   It opens the webhook and prints the **Callback URL** and **Verify token**. While it waits, go to the Meta App Dashboard > your app > **WhatsApp > Configuration** > Webhook > **Edit**, paste both, and click **Verify and save**. The script confirms and closes. Then, on the same page under **Webhook fields**, click **Subscribe** next to `messages`.

That's all. The address never changes, so this is never repeated. Every run from then on collects reports automatically.

### Limits

- **Reports only arrive while the script is waiting.** A report that comes later (e.g. the customer's phone was off) is missed, and that row stays `sent` (grey). Meta retries missed reports for a while, but by the next run the log holds only the new day, so they aren't recorded. Failures like 131049 usually come within seconds.
- **Messages sent before setup, or on a run where ngrok didn't start, stay `sent`.** Meta sends each report once, and there's no way to ask for it later.
- If ngrok can't start, the script says why and **still sends**, just without reports.
- A delivery failure doesn't cause a resend. The message still counts as sent today, which is usually right (131049 means "stop messaging this person for now").

### If reports don't show up

- `Delivery reports: off (...)` at the start: a setting is missing from `.env`, and the message names it.
- `ngrok stopped: ... ERR_NGROK_334`: an earlier ngrok is still running (e.g. the window was killed mid-run). End `ngrok.exe` in Task Manager and run again.
- Otherwise check `webhook.log` (every report recorded and every request rejected; "wrong signature" means `WHATSAPP_APP_SECRET` is wrong), `ngrok.log` (the tunnel), and `webhook_events.jsonl` (every report exactly as Meta sent it).

## Change what is sent

- **Offer text:** edit `OFFER_LINE` ({{2}}) at the top of `send_whatsapp.py`. {{1}} is always the customer's name from the Excel file.
- **Who receives it:** edit the Excel file.

Both templates must show **Active** in WhatsApp Manager; a template that is in review or rejected can't be sent.
