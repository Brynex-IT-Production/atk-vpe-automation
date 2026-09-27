// Deploy this Worker as the Meta callback after setting all four secrets.
// Required secrets: META_APP_SECRET, WEBHOOK_VERIFY_TOKEN,
// APPS_SCRIPT_URL, RELAY_SHARED_SECRET.
const encoder = new TextEncoder();

function hex(bytes) {
  return [...new Uint8Array(bytes)].map(byte => byte.toString(16).padStart(2, "0")).join("");
}

function equalHex(a, b) {
  if (typeof a !== "string" || typeof b !== "string" || a.length !== b.length) return false;
  let difference = 0;
  for (let i = 0; i < a.length; i++) difference |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return difference === 0;
}

async function hmac(secret, bytes) {
  const key = await crypto.subtle.importKey(
    "raw", encoder.encode(secret), { name: "HMAC", hash: "SHA-256" }, false, ["sign"]
  );
  return hex(await crypto.subtle.sign("HMAC", key, bytes));
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (request.method === "GET") {
      const valid = url.searchParams.get("hub.mode") === "subscribe" &&
        url.searchParams.get("hub.verify_token") === env.WEBHOOK_VERIFY_TOKEN;
      return valid
        ? new Response(url.searchParams.get("hub.challenge") || "", { status: 200 })
        : new Response("Verification failed", { status: 403 });
    }
    if (request.method !== "POST") return new Response("Method not allowed", { status: 405 });
    if (!env.META_APP_SECRET || !env.RELAY_SHARED_SECRET || !env.APPS_SCRIPT_URL) {
      return new Response("Relay not configured", { status: 503 });
    }

    const raw = new Uint8Array(await request.arrayBuffer());
    const signature = request.headers.get("X-Hub-Signature-256") || "";
    const expected = "sha256=" + await hmac(env.META_APP_SECRET, raw);
    if (!equalHex(signature, expected)) return new Response("Invalid signature", { status: 403 });

    const timestamp = String(Date.now());
    const payload = new TextDecoder().decode(raw);
    const mac = await hmac(env.RELAY_SHARED_SECRET, encoder.encode(timestamp + "." + payload));
    const response = await fetch(env.APPS_SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ relay_timestamp: timestamp, relay_payload: payload, relay_mac: mac }),
      redirect: "follow"
    });
    if (!response.ok) return new Response("Upstream unavailable", { status: 502 });
    const result = await response.text();
    return result === "EVENT_RECEIVED"
      ? new Response("EVENT_RECEIVED", { status: 200 })
      : new Response("Upstream rejected event", { status: 502 });
  }
};
