# Voucher Redemption Contract

This document defines the local backend contract for Tinat voucher phrases.

## Purpose

The Telegram bot issues exactly one phrase after admin approval. The Tinat app sends that phrase to your backend. The backend must validate it, redeem it exactly once, and return the granted package.

## Data model

Store each voucher phrase with at least these fields:

- `phrase`: the secret phrase shown to the user
- `packageKey`: the package this phrase unlocks
- `status`: `available`, `reserved`, or `redeemed`
- `reservedAt`: when the bot assigned it to a user
- `redeemedAt`: when the app redeemed it
- `redeemedByUserId`: the Tinat user ID that redeemed it
- `requestId`: the Telegram approval request ID

## Redeem endpoint

`POST /api/v1/vouchers/redeem`

Request body:

```json
{
  "phrase": "G5LE37QI45M2",
  "userId": "user_123",
  "deviceId": "optional-device-id"
}
```

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

Response when already redeemed:

```json
{
  "ok": false,
  "error": "ALREADY_REDEEMED"
}
```

Response when invalid:

```json
{
  "ok": false,
  "error": "INVALID_PHRASE"
}
```

## Redemption rules

1. Trim the submitted phrase before checking it.
2. Match the phrase against your voucher store exactly.
3. Reject it if the phrase does not exist.
4. Reject it if it is already redeemed.
5. If it is reserved but not redeemed yet, mark it redeemed atomically.
6. Return only the package that phrase unlocks.
7. Never allow the same phrase to unlock more than one account.

## Recommended backend behavior

- Keep the voucher store server-side only.
- Use a transaction or equivalent lock around lookup and redemption.
- Record which user redeemed the phrase for audit history.
- If the bot and app share the same store, treat the bot as the only issuer and the app as the only redeemer.

## Example verification logic

```text
find voucher by phrase
if not found -> INVALID_PHRASE
if status == redeemed -> ALREADY_REDEEMED
if status == available or reserved -> set status = redeemed, set redeemedAt, set redeemedByUserId
return packageKey
```

## Suggested file format for phrase pools

The bot reads one phrase per line from `phrases/<phrasePool>.txt`. After approval it reserves one phrase from that pool and stores the reservation locally in `data/vouchers.json`.