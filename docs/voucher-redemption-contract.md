# Voucher Redemption Contract

This document defines the backend contract that the Tinat study app uses to redeem voucher codes against the existing `redeem-server.js`.

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

The study app must run as a Telegram Mini App so the backend can authenticate the user. The frontend sends the raw `initData` (from `window.Telegram.WebApp.initData`) — never a hand-typed Telegram ID.

Request body:

```json
{
  "initData": "query_id=...&user=%7B...%7D&auth_date=...&hash=...",
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

1. Validate the Telegram `initData` signature (HMAC-SHA256 with the bot token, `WebAppData` secret) and `auth_date` freshness. The authenticated Telegram numeric user ID comes from the signed `user` field only — never from the request body.
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
| 401  | `UNAUTHORIZED`      | initData missing, stale, or invalid      |
| 429  | `RATE_LIMITED`      | too many attempts                        |
| 400  | `BAD_REQUEST`       | malformed request body                   |

## Data files

- `data/vouchers.json` – package pools and assigned/redeemed voucher records
- `data/users.json` – purchases (requests), drafts, and user identity (Telegram ID, phone)
- `data/entitlements.json` – package entitlements created on successful redemption

## Running

```bash
npm run redeem-api   # starts redeem-server.js on REDEEM_API_PORT (default 8787)
```

Health check: `GET /health` returns `{ "ok": true }`.
