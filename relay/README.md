# Meta webhook verification relay (DEV)

The public Apps Script web app cannot read Meta's `X-Hub-Signature-256` header. This Cloudflare Worker verifies the signature over the exact request bytes before forwarding a short-lived HMAC envelope to Apps Script.

## Setup

1. Deploy `worker.js` as a Cloudflare Worker using the Cloudflare dashboard.
2. Set encrypted Worker secrets: `META_APP_SECRET` (from the Meta app), `WEBHOOK_VERIFY_TOKEN` (the current Meta callback verify token), `APPS_SCRIPT_URL` (the DEV web app URL), and a newly generated random `RELAY_SHARED_SECRET` (at least 32 random bytes).
3. Set the same `RELAY_SHARED_SECRET` in the DEV Apps Script project Script Properties. This turns on rejection of direct, unsigned WhatsApp events.
4. Update the Meta WhatsApp webhook callback URL to the Worker URL, using the existing verify token, and complete Meta's verification challenge.
5. Send a real DEV test message and confirm one incoming event and reply. Test that an unsigned POST to the old Apps Script URL returns `EVENT_REJECTED`. Leave `WA_SEND_MODE=TEST` and the test recipient allowlist in place.

Do not commit any secret. The relay responds with a generic error and does not log webhook bodies or credentials. A direct scheduler `system_action` still uses the existing token in the GitHub Actions workflow; rotate it into a dedicated secret before production. Until steps 1–4 are complete, the new Apps Script code accepts the existing direct Meta callback to avoid interrupting DEV.
