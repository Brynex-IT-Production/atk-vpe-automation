# Web API V1 — Flutter handoff

Branch: `feature/web-api-v1`

## Principle
The API is the permanent boundary between Flutter and the automation engine. Flutter must never read/write Google Sheets directly and must never hold Meta credentials.

V1 uses a single POST gateway with an `action` field because Apps Script web apps expose `doPost`. A future backend can preserve the same logical operations as REST endpoints.

## Current safe operations
- `health`
- `members.list`
- `meeting_template.get`
- `meeting.create`
- `meeting.get`
- `meeting.save`
- `meeting.sync` — writes supported assignments into the legacy planner; does not send WhatsApp
- `meeting.activate`

`meeting.sync` is explicit and separate from saving. `meeting.activate` is intentionally safe in this branch: it changes status to `READY_FOR_SYNC` but does **not** trigger WhatsApp. We will wire messaging only after the adapter to `VPE Weekly Planner` passes tests.

## Request examples

### Members
```json
{"action":"members.list","club_id":"ATK"}
```

### Default meeting structure
```json
{"action":"meeting_template.get","club_id":"ATK"}
```

### Create meeting
```json
{"action":"meeting.create","club_id":"ATK","meeting_date":"2026-10-12"}
```

### Save/reorder meeting
```json
{
  "action":"meeting.save",
  "club_id":"ATK",
  "meeting_id":"MTG_...",
  "items":[
    {"item_id":"ITM_1","type":"role","code":"TMOD","label":"TMOD","member_id":"M001","member_name":"Anoop Paul"},
    {"item_id":"ITM_2","type":"speech","code":"S1","label":"Speaker 1","member_id":"M002","member_name":"Priyesh"},
    {"item_id":"ITM_3","type":"session","code":"","label":"Education Session","member_id":"M003","member_name":"Arjun Sundar Raj"}
  ]
}
```

Array order is authoritative. The backend writes positions 1..N, so Flutter drag-and-drop only needs to submit the reordered array.

### Activate
```json
{"action":"meeting.activate","club_id":"ATK","meeting_id":"MTG_..."}
```

## Flutter rules
1. Store IDs, not spreadsheet row numbers.
2. Display member names, but submit `member_id` plus name in V1.
3. Never call Meta directly.
4. Never trigger messaging when saving a draft.
5. Custom items use `type:"session"`; standard roles/speeches retain their codes.
6. The same member may be assigned twice only after an explicit UI warning/override (server validation will be added before production).

## Next backend steps
1. Add authenticated API gateway / CORS-capable relay.
2. Add Club Master and club-scoped Member Master.
3. Add adapter: ordered API meeting -> legacy `VPE Weekly Planner`.
4. Validate role/speaker/evaluator constraints.
5. Add API error codes instead of generic internal errors.
6. Add readiness/exceptions endpoints.
7. Test ATK end-to-end with messaging disabled.
8. Enable messaging only after explicit activation and successful sync.


## Gateway
All requests are POST JSON to the Apps Script web-app URL and must include:
```json
{"api_version":"v1","api_key":"<server-issued key>","action":"health"}
```
The `health` action does not require the key. All other actions do.

### Sync to existing automation planner
```json
{"api_version":"v1","api_key":"...","action":"meeting.sync","club_id":"ATK","meeting_id":"MTG_..."}
```
This maps standard roles/speeches to the existing VPE Weekly Planner but sends no WhatsApp messages.

## Security note for Flutter Web
Do not permanently embed `WEB_API_KEY` in a public Flutter Web bundle. For the internal Dev handoff it can be supplied through a controlled relay/test environment. Before public launch, authentication must move to a user/session token or a server/BFF layer. Meta credentials always remain server-side.
