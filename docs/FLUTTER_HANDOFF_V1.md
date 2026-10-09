# Flutter handoff — VPE Assistant API V1

Status: backend contract frozen for frontend integration.

## Endpoint
Use the Apps Script DEV web-app deployment URL supplied by the backend owner. All requests are POST JSON.

## Request envelope
Every frontend API request uses:
```json
{
  "api_version": "v1",
  "api_key": "<DEV ONLY>",
  "action": "club.bootstrap",
  "club_id": "ATK"
}
```

Do not ship the DEV API key in a public production Flutter Web bundle.

## Startup
Call `club.bootstrap` for club settings, ExCom, the default ordered meeting template, and capability flags. Load active members separately with `members.list`.

## Meeting planner mapping
- Choose date -> `meeting.create` with `meeting_date: "YYYY-MM-DD"`
- Member dropdown -> use `bootstrap.members`
- Drag/drop -> reorder the local `items` array
- Add custom session -> add an item with `type: "session"`, blank `code`, label, and optional member assignment
- Delete -> remove item locally
- Save as Draft -> `meeting.save` with `save_mode: "DRAFT"`
- Save Meeting Plan -> `meeting.save` with `save_mode: "CONFIRMED"`
- Reload detail -> `meeting.get`
- Meetings screen -> `meetings.list`

Standard codes are TMOD, GE, TTM, S1, S2, S3, EV1, EV2, EV3, Timer, Ah, GRM. Custom sessions MUST have an empty code.

## Dashboard
Use `dashboard.get`. Render `readiness` and `pending_actions`. Pending actions already contain UI fields: `title`, `description`, `severity`, `state`, `due_at`.

## Onboarding
- `club.settings.get` / `club.settings.save`
- `members.list` / `member.save`
- `excom.list` / `excom.save`

## Automation safety
Do NOT call `meeting.sync` from the UI. It is an internal migration endpoint and is server-disabled by default.
`meeting.activate` currently validates and marks readiness only; WhatsApp activation remains disabled during integration.

## Errors
Read `ok`, `status`, and `error.code`.
Important codes: UNAUTHORIZED, MISSING_ACTION, UNSUPPORTED_API_VERSION, BAD_REQUEST, NOT_FOUND, MEETING_NOT_READY, LEGACY_SYNC_DISABLED.

## Acceptance rule
The VPE must not see or depend on Google Sheets, Apps Script internals, Meta template names, task rows, or planner column numbers.


## Save semantics
`meeting.save` accepts `save_mode`:
- `DRAFT`: persist an incomplete working plan.
- `CONFIRMED`: mark the plan as `PLANNED`. Vacant roles are allowed because the automation engine is expected to fill them through role-call workflows.

Confirming a meeting plan does not itself expose a separate VPE-facing "Start Automation" control in V1.


## Confirmed-plan automation bridge
The backend now has a disabled-by-default bridge from a confirmed web meeting plan into the existing ATK automation stack.

Server properties:
- `WEB_API_AUTO_SYNC_PLANNED` — when TRUE, a `meeting.save` with `save_mode:"CONFIRMED"` attempts the legacy bridge.
- `WEB_API_ALLOW_LEGACY_SYNC` — separate safety permission required for any planner write.

Both remain FALSE/absent during frontend integration. The Flutter client must not expose or control these flags.

When deliberately enabled later, a confirmed plan can update the existing planner, refresh programme slots, and rebuild automation tasks. This is how manual web assignments become unavailable in the WhatsApp role-call flow without changing the existing role-selection engine.
