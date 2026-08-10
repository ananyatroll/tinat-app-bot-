# Tinat Telegram Bot

This bot collects a user's chosen package, name, transaction link, and transaction ID, then sends the request to an admin for approval. After approval, it delivers one voucher phrase for the selected package.

## What it does

- Shows an intro and package chooser from `/start`
- Lets the user choose a package before submitting the request
- Collects the user's name, transaction link, and transaction ID
- Pings the admin with approve and reject buttons
- Reserves one unused phrase from the correct package pool after approval
- Exports approved requests to an `.xlsx` file with one sheet per package and sends it to the admin chat
- Stores approved voucher deliveries so `/myaccess` can resend them later
- Includes a local redemption API the Tinat app can call to redeem voucher phrases

## Setup

1. Install Node.js 18 or newer.
2. Run `npm install`.
3. Copy `.env.example` to `.env` and fill in the values.
4. Start the bot with `npm start`.
5. Start the local redemption API with `npm run redeem-api`.

## GitHub

This project includes a GitHub Actions workflow in [.github/workflows/ci.yml](.github/workflows/ci.yml) that validates the bot and redemption server on every push or pull request.

The workflow also supports manual runs and a daily scheduled validation, but GitHub Actions still cannot keep the Telegram bot running 24/7.

GitHub cannot keep a Telegram bot running continuously by itself. To keep the bot live, run it on a server, VM, or container host, and use GitHub for source control and validation.

## Environment variables

- `BOT_TOKEN`: your Telegram bot token from BotFather
- `ADMIN_CHAT_ID`: your Telegram numeric user ID or group chat ID that should receive approval requests
- `PACKAGES_JSON`: JSON array that defines the packages shown to users and which phrase pool each one uses
- `PHRASES_DIR`: folder that contains text files with one secret phrase per line, named after each package pool
- `VOUCHERS_FILE`: optional JSON file for voucher reservations, defaults to `data/vouchers.json`
- `REDEEM_API_PORT`: port for the local redemption API, defaults to `8787`

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

## Flow

- The user sends `/start`.
- The bot shows a short intro and package options.
- The bot asks the user to choose a package, then the user's full name, then the transaction link, then the transaction ID.
- The bot sends the details to the admin chat for manual review.
- When the admin approves the request, the bot reserves one unused phrase from the matching package pool and sends it to the user.
- When the admin approves the request, the bot also creates an Excel file in `exports/` with all approved requests so far (one sheet per package) and sends that file to the admin chat.

## Notes

- Put one phrase per line in `phrases/<phrasePool>.txt`.
- The app backend should verify the phrase, mark it redeemed, and unlock the matching package only once. See [docs/voucher-redemption-contract.md](docs/voucher-redemption-contract.md).
- If a pool runs out of phrases, approvals for that package should stop until you refill the file.
- To test redemption locally, POST to `http://127.0.0.1:8787/api/v1/vouchers/redeem` with the JSON body described in the contract.
