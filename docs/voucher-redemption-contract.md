# Voucher Redemption Contract

This document defines the backend contract that the Tinat study app uses to redeem voucher codes. The live implementation is the Flask API (`flask_app.py`, deployed on Railway); `redeem-server.js` is the older Node implementation and is kept for reference.

## Purpose

- The Telegram bot collects a purchase (package, verified phone number, payment details).
- The admin verifies and approves the payment.
- Support assigns exactly one unused voucher from the correct package to the buyer.
- The study app sends a redeem request to this backend. The backend verifies Telegram identity, voucher ownership, and voucher state, then atomically marks the voucher `REDEEMED` and returns the granted package.

## Voucher states

A voucher moves through these states and can never go backwards:

- `UNUSED` – still in the package pool, assigned to no one.
- `ASSIGNED` – bound to one Telegram user (and one purchase) by Support.
- `REDEEMED` – consumed by the owner; cannot be reused.
- `REVOKED` – blocked; can never be redeemed.

## Voucher record fields

Stored in `data/vouchers.json` under the matching package pool. Each assigned voucher carries:

- `phrase` – the secret code the user enters
- `packageKey`, `packageLabel`, `packagePool` – which package it unlocks
- `status` – one of the states above
- `requestId` – the Telegram approval request it belongs to
- `assignedToUserId` – the Telegram numeric user ID of the owner
- `assignedToUsername` – Telegram username if the owner has one
- `assignedToPhone` – the verified phone number from the purchase
- `phoneVerified` – whether the phone came from Telegram's contact share
- `transactionId` – the payment/transaction ID of the approved purchase
- `assignedAt`, `assignedBy` – when and by whom it was assigned
- `redeemedAt`, `redeemedByUserId` – redemption audit fields

## Redeem endpoint

`POST /api/v1/vouchers/redeem`

The study app must authenticate as a Telegram user. Two auth modes are supported:

- **Mini App** (web/webview): send the raw `initData` (from `window.Telegram.WebApp.initData`) — never a hand-typed Telegram ID.
- **Native app** (Android/iOS): the app can't run Telegram's JS layer, so it uses the official *Login with Telegram* SDK (e.g. `telegram-login-android`), which returns a signed payload. The app exchanges it for a short-lived `authToken` via `/api/v1/auth/login` and then sends that token to `/redeem` or `/status`.

Request body:

```json
{
  "initData": "query_id=...&user=%7B...%7D&auth_date=...&hash=...",
  "code": "G5LE37QI45M2",
  "deviceId": "optional-device-id"
}
```

Native apps instead send `authToken` (from `/api/v1/auth/login`) instead of `initData`:

```json
{
  "authToken": "RcpU2FKvlea57n5cgSvSW58rMYxHQ83XUq85Rji_WCc",
  "code": "G5LE37QI45M2",
  "deviceId": "optional-device-id"
}
```

`code` is accepted as an alias too (legacy `phrase` field also works).

Response on success:

```json
{
  "ok": true,
  "packageKey": "euee-preo",
  "packageLabel": "EUEE Preo",
  "redeemedAt": "2026-08-09T12:00:00.000Z",
  "access": {
    "packageKey": "euee-preo",
    "enabled": true
  }
}
```

## Verification order (server-side)

1. Authenticate the Telegram user: `initData` is verified via HMAC-SHA256 with the bot token (`WebAppData` secret) plus `auth_date` freshness; `authToken` is looked up in the issued-token store and checked against its expiry. The authenticated Telegram numeric user ID comes only from the verified source — never from the request body.
2. Rate-limit attempts per user and per IP to slow down code guessing.
3. Find the voucher by `code`.
4. `INVALID_CODE` if it does not exist.
5. `VOUCHER_REVOKED` if the voucher was revoked.
6. `ALREADY_REDEEMED` if already redeemed.
7. `NOT_ASSIGNED` if the code is still unused or not assigned to a user.
8. `NOT_OWNER` if `assignedToUserId` does not match the authenticated Telegram user.
9. `PAYMENT_NOT_APPROVED` if the backing request is not approved.
10. Atomically set `status = redeemed`, record `redeemedAt`/`redeemedByUserId`, write the user's package entitlement to `data/entitlements.json`, and return the package.

All redemption read-modify-write steps run inside a single serialized queue so two simultaneous requests can never both redeem the same voucher (`ASSIGNED -> REDEEMED` exactly once).

## Error codes (safe, no sensitive data)

| HTTP | code                | meaning                                  |
|------|---------------------|------------------------------------------|
| 400  | `INVALID_CODE`      | unknown code                             |
| 400  | `NOT_ASSIGNED`      | code not assigned to anyone              |
| 400  | `ALREADY_REDEEMED`  | code already consumed                    |
| 400  | `VOUCHER_REVOKED`   | code blocked                             |
| 400  | `NOT_OWNER`         | code belongs to a different user         |
| 400  | `PAYMENT_NOT_APPROVED` | backing payment not approved          |
| 401  | `UNAUTHORIZED`      | initData/authToken missing, stale, or invalid |
| 429  | `RATE_LIMITED`      | too many attempts                        |
| 400  | `BAD_REQUEST`       | malformed request body                   |

## Data files

- `data/vouchers.json` – package pools and assigned/redeemed voucher records
- `data/users.json` – purchases (requests), drafts, and user identity (Telegram ID, phone)
- `data/entitlements.json` – package entitlements created on successful redemption

## Running

```bash
python run_local.py   # Flask app: long-polling bot + redeem API on $PORT (Railway)
```

Health check: `GET /health` returns `{ "ok": true }`.

## Endpoints

- `POST /api/v1/vouchers/redeem` — consume a code and grant access
- `POST /api/v1/vouchers/status` — check a code's validity without consuming it (same auth and error format as `/redeem`)
- `POST /api/v1/auth/login` — native-app login: verify a *Login with Telegram* payload and receive a short-lived `authToken`
- `POST /api/v1/android/activate` — standalone app: phone + code → redeem once, create entitlement, issue app access token
- `GET /api/v1/android/entitlement` — app access token → authorized package

Both redeem/status endpoints accept `initData` **or** `authToken`, plus `code` (aliases: `phrase`, `voucherPhrase`).

## Storage / persistence on Railway

All state is JSON under `DATA_DIR` (`data/` locally; `/data` on Railway): `users.json`, `vouchers.json`, `entitlements.json`, `auth_tokens.json`, `access_tokens.json`, `processed_updates.json`.

These files live on the Railway volume mounted at `/data`. Without that volume the filesystem is ephemeral and resets on every redeploy — so the volume must be created (`Ctrl+K` → *Create Volume* → mount path `/data`) and `DATA_DIR=/data` must be set. This is the persistence layer; nothing is lost across restarts when the volume is in place.

## Native app login (`/api/v1/auth/login`)

The app uses the official *Login with Telegram* SDK (`telegram-login-android`), which shows a Telegram-branded button that opens Telegram to confirm login. The SDK hands the app back a signed payload:

```json
{
  "id": 42424242,
  "first_name": "Test",
  "last_name": "Buyer",
  "username": "testbuyer",
  "photo_url": "https://t.me/i/userpic/...",
  "auth_date": 1754800000,
  "hash": "..."
}
```

`POST /api/v1/auth/login` with that JSON:

```json
{
  "id": 42424242,
  "first_name": "Test",
  "last_name": "Buyer",
  "username": "testbuyer",
  "auth_date": 1754800000,
  "hash": "..."
}
```

Success (`200`):

```json
{
  "ok": true,
  "userId": "42424242",
  "authToken": "RcpU2FKvlea57n5cgSvSW58rMYxHQ83XUq85Rji_WCc",
  "expiresAt": "2026-08-11T12:00:00.000Z"
}
```

Failure (`401`): `{ "ok": false, "error": "UNAUTHORIZED" }` — tampered payload or stale `auth_date`.

The server verifies the `hash` as HMAC-SHA256 over the sorted `key=value` fields (excluding `hash`), keyed with the SHA-256 of the bot token — the same scheme as the web Login Widget — and rejects payloads older than `INIT_DATA_MAX_AGE_SECONDS`. No new secrets are shared with the app; the app never needs the bot token.

`authToken` is valid for `AUTH_TOKEN_TTL_SECONDS` (default 24h) and is stored in `data/auth_tokens.json` (`AUTH_TOKENS_FILE`). The app should call `/auth/login` again when it gets a `401` and treat the token as opaque.

## Android activation (standalone app, phone + code)

The standalone Android Kotlin app has no Telegram session, so it authenticates with the **verified phone number** the bot captured via Telegram's official Share Contact button, plus the redeem code assigned by Support.

### Activate: `POST /api/v1/android/activate`

```json
{
  "phone": "+2519XXXXXXXX",
  "code": "BIO-82KF-91XP"
}
```

Success (`200`):

```json
{
  "success": true,
  "accessToken": "opaque-app-token",
  "package": {
    "id": "biology",
    "name": "Biology"
  }
}
```

The backend trusts only its own records. It verifies, in order, inside one serialized write lock on `data/vouchers.json`:

1. The phone number is normalized (digits only) and plausible.
2. A voucher with that `code` exists.
3. The voucher is `ASSIGNED` (not `UNUSED`, not `REDEEMED`, not `REVOKED`).
4. The voucher's stored owner phone matches the submitted phone.
5. The phone was Telegram-verified (`phone.verified: true` on the backing purchase request; `phoneVerified` on the voucher is the assignment-time record). A manually typed number never qualifies.
6. The backing purchase was approved (`request.status` is `approved` or `delivered`).
7. The package still exists in `PACKAGES`.

Only then is the voucher marked `REDEEMED` atomically, an entitlement recorded in `data/entitlements.json`, and an app `accessToken` issued. The app never supplies Telegram ID, username, package, payment, or entitlement state.

### Check entitlement: `GET /api/v1/android/entitlement`

```
Authorization: Bearer <accessToken>
```

Success (`200`):

```json
{
  "success": true,
  "package": { "id": "biology", "name": "Biology" },
  "entitlement": {
    "entitlementId": "...",
    "packageKey": "biology",
    "packageLabel": "Biology",
    "activatedAt": "2026-08-11T00:00:00.000Z"
  }
}
```

The app stores the `accessToken` in Android secure storage (Keystore) and uses it on launch instead of asking for the code again.

### Android error codes

| HTTP | code                | meaning                                        |
|------|---------------------|------------------------------------------------|
| 400  | `INVALID_CODE`      | unknown code                                   |
| 400  | `ALREADY_REDEEMED`  | code already consumed                          |
| 400  | `VOUCHER_REVOKED`   | code blocked                                   |
| 400  | `NOT_ASSIGNED`      | code not assigned to a purchase                |
| 400  | `PHONE_MISMATCH`    | phone does not own this code                   |
| 400  | `PHONE_NOT_VERIFIED`| phone was not shared via Telegram's contact    |
| 400  | `PAYMENT_NOT_APPROVED` | backing purchase not approved              |
| 400  | `PACKAGE_INVALID`   | package no longer offered                      |
| 400  | `INVALID_PHONE` / `BAD_REQUEST` | malformed input             |
| 429  | `RATE_LIMITED`      | too many attempts (per phone, code, or IP)     |
| 401  | `UNAUTHORIZED`      | missing/invalid access token                   |
| 401  | `TOKEN_EXPIRED`     | access token past its TTL                      |
| 401  | `TOKEN_REVOKED`     | access token (or its entitlement) revoked      |
| 404  | `ENTITLEMENT_NOT_FOUND` | token valid but entitlement gone           |

### Access tokens & revocation

- Access tokens are opaque random strings stored in `data/access_tokens.json` (`ACCESS_TOKENS_FILE`) with a default TTL of 365 days (`ACCESS_TOKEN_TTL_SECONDS`).
- Rate limiting reuses the existing in-memory limiter (`REDEEM_RATE_LIMIT_*`, overridable via `ANDROID_RATE_LIMIT_MAX` / `ANDROID_RATE_LIMIT_WINDOW_MS`) with separate buckets for phone, code, and client IP.
- Admin revocation (Telegram): `/revokeaccess <token>` invalidates one token; `/revokeentitlement <entitlementId>` revokes an entitlement (all its tokens stop opening the package).
- Privacy: activation attempts are logged with masked phone (`****0000`) and code length only; responses never leak other users' phones, Telegram IDs, payments, or unused codes.
