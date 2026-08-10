# Tinat Telegram Bot

This bot collects a user's selected package, verified phone number, name, and payment details, sends the purchase to an admin for payment verification, and then lets a support operator assign exactly one unused voucher from the correct package.

## What it does

- Shows an intro and package chooser from `/start`
- Never requires a Telegram username — the Telegram numeric user ID is the primary identity
- Collects the phone number via Telegram's official contact-sharing button (typed numbers are marked unverified)
- Validates the payment method (CBE / Telebirr), receipt link, and transaction ID before submitting
- Sends an approval request to the admin with Approve / Reject buttons
- Approving a payment immediately reserves one unused voucher from the correct package pool and sends the phrase to the buyer in chat
- Rejected payments are marked rejected and never issue a voucher
- Support can still assign a voucher to already-approved purchases that missed one (`/pending` or the Assign buttons)
- Exports approved purchases to an `.xlsx` file (one sheet per package) with phone number, voucher code, voucher status, assigned time, and redeemed time
- Includes the existing local redemption API (`redeem-server.js`) that the Tinat Mini App calls to redeem vouchers once, using Telegram Mini App authentication

## Flow

1. User sends `/start` and picks a package.
2. Bot asks the user to share their phone number via Telegram's contact-sharing button.
3. Bot asks for payment method, full name, transaction link, and transaction ID (validated per method).
4. The purchase (Telegram user ID, phone, username if present, package, payment details) is sent to the admin.
5. Admin verifies the payment and presses Approve or Reject.
6. Rejected → the payment is marked rejected, no voucher is issued.
7. Approved → the bot immediately takes one unused code from that package's pool, records ownership, and sends the voucher phrase to the buyer in chat.
8. The user opens the study app (a Telegram Mini App), which sends `initData` + the voucher code to `redeem-server.js`.
9. The server verifies the Telegram identity, ownership, payment approval, and voucher state, atomically marks the voucher `REDEEMED`, stores the entitlement, and unlocks the package.

## Setup

1. Install Node.js 18 or newer.
2. Run `npm install`.
3. Copy `.env.example` to `.env` and fill in the values.
4. Start the bot with `npm start`.
5. Start the local redemption API with `npm run redeem-api`.

## Environment variables

- `BOT_TOKEN`: your Telegram bot token from BotFather
- `ADMIN_CHAT_ID`: your Telegram numeric user ID or group chat ID that should receive approval requests
- `SUPPORT_CHAT_ID`: Telegram numeric user ID of the support operator who assigns vouchers. Leave empty to let the admin act as support too. The support account must press `/start` on the bot first.
- `PACKAGES_JSON`: JSON array that defines the packages shown to users and which phrase pool each one uses
- `PHRASES_DIR`: folder that contains text files with one secret phrase per line, named after each package pool
- `VOUCHERS_FILE`: optional JSON file for voucher assignments, defaults to `data/vouchers.json`
- `USERS_FILE`: optional JSON file for purchases and user identity, defaults to `data/users.json`
- `ENTITLEMENTS_FILE`: optional JSON file for entitlements created on redemption, defaults to `data/entitlements.json`
- `REDEEM_API_PORT`: port for the local redemption API, defaults to `8787`
- `REDEEM_RATE_LIMIT_MAX` / `REDEEM_RATE_LIMIT_WINDOW_MS`: anti-guessing limits on the redeem endpoint

Example `PACKAGES_JSON`:

```json
[
	{"key":"euee-preo","label":"EUEE Preo","priceCents":25000,"currency":"ETB","phrasePool":"euee-preo"},
	{"key":"freshman","label":"Freshman","priceCents":25000,"currency":"ETB","phrasePool":"freshman"},
	{"key":"uat","label":"UAT","priceCents":25000,"currency":"ETB","phrasePool":"uat"},
	{"key":"university-department","label":"University Department","priceCents":25000,"currency":"ETB","phrasePool":"university-department"},
	{"key":"exit-exam","label":"Exit Exam","priceCents":25000,"currency":"ETB","phrasePool":"exit-exam"}
]
```

## Roles

- **Admin** (from `ADMIN_CHAT_ID`): verifies payments and approves/rejects purchases.
- **Support** (from `SUPPORT_CHAT_ID`): sees approved purchases waiting for vouchers (`/pending`) and assigns one unused code from the correct package.
- **Users**: buyers. They are identified by their Telegram numeric user ID, never by username.

## Notes

- Put one phrase per line in `phrases/<phrasePool>.txt`.
- The Excel report is for admin/support reporting only — it is never used as an authentication database.
- See [docs/voucher-redemption-contract.md](docs/voucher-redemption-contract.md) for the redemption API contract and voucher lifecycle.
- To test redemption locally, POST to `http://127.0.0.1:8787/api/v1/vouchers/redeem` with a valid Telegram Mini App `initData` and a voucher code.
