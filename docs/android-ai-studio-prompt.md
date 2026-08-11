# Tinat Android App — AI Studio Build Prompt

Paste the block below into Google AI Studio to generate the Android app that
connects to the Tinat backend. Everything in this prompt is verified against
the running server code (endpoints, schemas, error codes, phone rules).

> IMPORTANT: replace `<RAILWAY_APP_DOMAIN>` with the real deployed Railway URL
> (e.g. `https://tinat-bot.up.railway.app`) before generating. The backend
> serves HTTPS only and never permits cleartext HTTP.

The live deployment base URL is: **https://web-production-16f9b.up.railway.app**

---

```
You are building the Android app for "Tinat", an Ethiopian study-prep product.
Users buy a package in the Telegram bot, get approved, and receive a voucher
code in Telegram. The Android app lets them activate that code with their
phone number and then unlocks the package's content.

## Backend contract (base URL: https://<RAILWAY_APP_DOMAIN>)

All requests/responses are JSON. No other auth for activation.

### 1) Activate a package
POST {BASE}/api/v1/android/activate
Content-Type: application/json

Request body:
{
  "phone": "+251912345678",
  "code": "BIO-82KF-91XP"
}
- phone: the phone number the user shared with the Telegram bot. The server
  strips all non-digit characters and requires 9-15 digits total.
- code: the voucher code the user received in Telegram (format like
  BIO-82KF-91XP; treat it as opaque text).

Success (HTTP 200):
{
  "success": true,
  "accessToken": "<43-char random token>",
  "package": { "id": "euee-prep", "name": "EUEE Prep" }
}

Errors come as JSON with HTTP status per this table:
| error                | HTTP | meaning                                          |
| BAD_REQUEST          | 400  | malformed body or empty phone/code               |
| INVALID_PHONE        | 400  | phone not 9-15 digits                            |
| INVALID_CODE         | 400  | no voucher matches the code                      |
| ALREADY_REDEEMED     | 400  | code already used by anyone                      |
| VOUCHER_REVOKED      | 400  | code was revoked                                 |
| NOT_ASSIGNED         | 400  | code not yet assigned to a purchase              |
| PHONE_MISMATCH       | 400  | phone does not match the voucher owner           |
| PHONE_NOT_VERIFIED   | 400  | phone was not verified through Telegram          |
| PAYMENT_NOT_APPROVED | 400  | the purchase was not approved yet                |
| PACKAGE_INVALID      | 400  | package no longer exists                         |
| RATE_LIMITED         | 429  | too many attempts; wait and retry with backoff   |
Error body shape: { "success": false, "error": "<CODE>" }

### 2) Check entitlement (call on app launch)
GET {BASE}/api/v1/android/entitlement
Authorization: Bearer <accessToken>

Success (HTTP 200):
{
  "success": true,
  "package": { "id": "euee-prep", "name": "EUEE Prep" },
  "entitlement": {
    "entitlementId": "<id>",
    "packageKey": "euee-prep",
    "packageLabel": "EUEE Prep",
    "activatedAt": "2026-08-11T00:00:00Z"
  }
}

Errors:
| error                | HTTP | meaning                                        |
| UNAUTHORIZED         | 401  | missing/invalid Authorization header           |
| TOKEN_EXPIRED        | 401  | access token past its expiry                   |
| TOKEN_REVOKED        | 401  | access token or entitlement revoked            |
| ENTITLEMENT_NOT_FOUND| 404  | token valid but no entitlement record          |

## Packages (current catalog, all 300.00 ETB)
key=id                   label
euee-prep                EUEE Prep
freshman                 Freshman
uat                      UAT
university-department    University Department
exit-exam                Exit Exam

## User flow to implement
1. Splash screen: if a saved accessToken exists, call GET entitlement.
   - 200 -> unlock content for the returned package.
   - any 401/404 -> clear the token and show the activation screen.
2. Activation screen: two inputs (phone with Ethiopian formatting, e.g.
   starting +251, and voucher code) + an "Activate" button.
   - POST activate; on success store accessToken securely and go to content.
   - on error, map each code to a clear user-facing message, e.g.
     INVALID_CODE -> "This code is not recognised. Check the code you got in
     Telegram."; PHONE_MISMATCH -> "This code belongs to a different phone
     number."; PAYMENT_NOT_APPROVED -> "Your purchase has not been approved
     yet."; RATE_LIMITED -> "Too many attempts. Try again in a few minutes."
3. Content screen: shows the unlocked package name; replace the "content"
   section with a placeholder that says it is ready for the study material.

## Technical requirements
- Kotlin, MVVM, single Activity + Compose or Fragments (choose the modern
  default), Retrofit + OkHttp + kotlinx.serialization or Moshi.
- Custom Retrofit error converter to read the "error" field from non-2xx
  bodies so errors map to the table above.
- Store the accessToken with EncryptedSharedPreferences (or Android Keystore);
  never log the token, phone, or code.
- Disable cleartext traffic (HTTPS only). Base URL constant in one place.
- Phone field: allow +, digits, spaces; validate 9-15 digits after stripping
  non-digits before submitting.
- Network errors (no internet, timeout, 5xx) -> generic retry message.
- 429 -> exponential backoff with a short pause before re-enabling the button.
- Producing a small Retrofit interface like:

interface TinatApi {
  @POST("api/v1/android/activate")
  suspend fun activate(@Body body: ActivateRequest): ActivationResponse

  @GET("api/v1/android/entitlement")
  suspend fun entitlement(@Header("Authorization") bearer: String): EntitlementResponse
}

Generate the complete project (Gradle files, manifest, build config, all
screens, DI or manual wiring, network layer, models, token store, and a
README with run instructions).
```

---

## Notes
- The voucher code is delivered in Telegram by the support/admin flow; the app
  only collects it from the user.
- The backend never trusts identity from the app: it validates the code,
  the owner phone, Telegram phone verification, and payment approval itself.
- `accessToken` is valid for 1 year by default and is bound to one entitlement.
