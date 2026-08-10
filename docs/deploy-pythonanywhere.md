# Deploying the Tinat bot to PythonAnywhere (Free Tier)

This guide migrates the Node.js bot to Python/Flask running in webhook mode.
PythonAnywhere's Free Tier has no background tasks and no long-polling, so the
bot lives entirely inside Flask request handlers: Telegram pings
`https://<username>.pythonanywhere.com/telegram-webhook` on every message and
callback, and the app answers in the same request.

## 1. Files added by the migration

| File | Purpose |
| --- | --- |
| `flask_app.py` | Full Flask webhook bot: `/start`, package flow, admin approval, voucher reservation, Excel export, redemption API |
| `wsgi.py` | PythonAnywhere WSGI entry point that loads `flask_app.application` |
| `requirements.txt` | `flask`, `requests`, `openpyxl` |
| `docs/deploy-pythonanywhere.md` | This guide |

## 2. Folder structure on PythonAnywhere

Create a project folder under `/home/<username>/` and keep the existing
structure. On the **Files** tab create/upload:

```
/home/<username>/tinat-bot/
├── flask_app.py
├── wsgi.py
├── requirements.txt
├── .env                      (optional, only for local testing)
├── phrases/                  # one .txt per pool, one phrase per line
│   ├── euee-preo.txt
│   ├── freshman.txt
│   ├── uat.txt
│   ├── university-department.txt
│   └── exit-exam.txt
├── data/                     # created automatically on first run
│   ├── users.json
│   ├── vouchers.json
│   ├── processed_updates.json
│   └── ... .lock files
└── exports/                  # created automatically; approved .xlsx files
```

`data/` and `exports/` are created on first request; you only need to upload
`phrases/` and the code.

## 3. WSGI configuration

On the **Web** tab, add a new web app:

1. **Manual configuration**, Python 3.10+.
2. Set the paths:
   - Source code and working directory: `/home/<username>/tinat-bot`
   - WSGI configuration file: `/home/<username>/tinat-bot/wsgi.py`
3. `wsgi.py` imports `application` from `flask_app.py`, so no other setup is needed there.

The file:

```python
import os
import sys

PROJECT_HOME = os.path.expanduser('~/tinat-bot')
if PROJECT_HOME not in sys.path:
    sys.path.insert(0, PROJECT_HOME)

from flask_app import application  # noqa: E402, F401
```

## 4. Environment variables

Set these on the **Web tab -> Environment variables** (one line per variable):

| Variable | Example | Purpose |
| --- | --- | --- |
| `TELEGRAM_BOT_TOKEN` | `123456:ABC-...` | Bot token from BotFather |
| `ADMIN_CHAT_ID` | `1905159972` | Numeric chat ID that receives approval requests and Excel files |

Optional:

| Variable | Default | Purpose |
| --- | --- | --- |
| `PACKAGES_JSON` | the 5 default packages (EUEE Preo, Freshman, UAT, University Department, Exit Exam) | JSON array to override the package list |
| `PHRASES_DIR` | `./phrases` | Phrase pool folder |
| `DATA_DIR` | `./data` | JSON stores |
| `EXPORTS_DIR` | `./exports` | Excel output |
| `USERS_FILE` / `VOUCHERS_FILE` | `data/users.json` / `data/vouchers.json` | Store file paths |

`flask_app.py` also reads a local `.env` file if present (real environment
variables always win), which is handy for running the bot locally.

## 5. Deployment checklist

1. **Upload code** — Files tab -> upload `flask_app.py`, `wsgi.py`,
   `requirements.txt`, and the `phrases/` folder into `/home/<username>/tinat-bot/`.
2. **Install dependencies** — open a **Bash console** and run:

   ```bash
   cd ~/tinat-bot
   pip install --user -r requirements.txt
   ```

   (Equivalent: `pip install --user flask requests openpyxl`.)
3. **Set environment variables** on the Web tab (see section 4), then click
   **Reload**.
4. **Register the webhook** — from the Bash console (or locally):

   ```bash
   curl -X POST "https://api.telegram.org/bot<TELEGRAM_BOT_TOKEN>/setWebhook" \
     -d url="https://<username>.pythonanywhere.com/telegram-webhook" \
     -d allowed_updates='["message","callback_query","edited_message"]'
   ```

   Or run the helper:

   ```bash
   python -c "from flask_app import register_webhook; register_webhook('https://<username>.pythonanywhere.com/telegram-webhook')"
   ```

   Verify with:

   ```bash
   curl "https://api.telegram.org/bot<TELEGRAM_BOT_TOKEN>/getWebhookInfo"
   ```

   You should see `"pending_update_count": 0` and `"last_error_message": ""`.
5. **Reload the web app** again after setting env vars, then test (section 6).

## 6. Testing that the bot is live

- Open `https://<username>.pythonanywhere.com/health` in a browser — expect `{"ok": true}`.
- Open `https://<username>.pythonanywhere.com/` — expect the "live" message.
- Send `/start` to the bot in Telegram; the package chooser should appear.
- Follow the flow: pick a package -> name -> transaction link -> transaction
  ID. The admin chat should receive the review message with Approve/Reject
  buttons.
- Click **Approve**: the user gets a voucher phrase, and the admin chat
  receives the regenerated `approved-*.xlsx`.
- Test redemption (from anywhere):

  ```bash
  curl -X POST "https://<username>.pythonanywhere.com/api/redeem" \
    -H "Content-Type: application/json" \
    -d '{"phrase":"<VOUCHER_PHRASE>","userId":"user_123"}'
  # -> {"status":"success", "packageKey":"...", "packageLabel":"..."}
  ```

  Redeeming the same phrase again returns `{"status":"invalid","error":"ALREADY_REDEEMED"}`.
  The legacy contract endpoint `POST /api/v1/vouchers/redeem` is also served.

## 7. Error handling & edge cases

- **Empty phrase pool**: approval refuses with an admin message ("No voucher
  phrases left for package ... Refill phrases/<pool>.txt") and the request
  stays `pending`; no phrase is double-issued.
- **Duplicate callback**: re-delivered Approve/Reject callbacks are idempotent
  — the request is checked for `status == 'pending'` before reserving again.
- **Webhook retries**: Telegram webhook delivery is at-least-once. Every update
  is deduplicated by `update_id` in `data/processed_updates.json` (24h TTL,
  capped at 500 entries), so a retried POST can never advance a draft two steps
  or double-submit a request.
- **Non-admin callback**: anyone other than `ADMIN_CHAT_ID` pressing a button
  gets "You are not allowed to review requests."
- **Invalid inputs**: blank messages are ignored; the flow simply waits for a
  valid reply. Unknown commands (`/foo`) are ignored silently.
- **Duplicate buyer**: a user whose request was approved cannot submit a second
  request (`/buy` blocks them and points to `/myaccess`).
- **File safety**: every JSON read-modify-write (drafts, requests, voucher
  reservation, redemption) happens under an atomic file lock
  (`data/*.json.lock`, `fcntl` on Linux, `msvcrt` fallback on Windows) so
  concurrent webhook hits can't double-issue or corrupt the stores.

## 8. Debugging

- **Web tab -> Error log** shows `print` output, `logging`, and tracebacks from
  `flask_app.py`. Look for `tinat-bot` lines.
- **Web tab -> Server log** shows WSGI startup messages.
- Raise verbosity by editing `logging.basicConfig(level=...)` in
  `flask_app.py` (default is `INFO`).
- Telegram's `getWebhookInfo` reports webhook errors (`last_error_message`)
  when Telegram could not reach your app.
- On the Free Tier the Free PythonAnywhere clock runs the app, but requests
  still terminate the process after idle; nothing is needed — the webhook
  wakes the worker per request.

## 9. Running locally (optional)

```bash
pip install -r requirements.txt
python flask_app.py          # serves on http://127.0.0.1:5000
```

With a local `.env` set to your real `BOT_TOKEN`, you can forward Telegram
updates with a tunnel (e.g. `ngrok http 5000`) and point the webhook at the
tunnel URL while testing.
