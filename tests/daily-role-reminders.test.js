'use strict';
const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');

const source = fs.readFileSync('Tasks.js', 'utf8');
const cron = fs.readFileSync('.github/workflows/phase1-scheduler.yml', 'utf8');
const marker = 'function dueVacantRoleCallSlot_';
const start = source.indexOf(marker);
const end = source.indexOf('\n}\n',start);
assert.ok(start >= 0 && end > start, 'Due-slot evaluator must exist');
const context = {};
vm.createContext(context);
vm.runInContext(source.slice(start,end+3), context);
const due = context.dueVacantRoleCallSlot_;

const times = '08:15,20:15';
const day = '2026-10-22';
const meeting = '2026-10-26';
const at = (h,m) => h*60+m;
function expected(date,h,m,slot) {
  const result = due(date,at(h,m),meeting,'2026-10-22',times);
  assert.equal(result && result.slot,slot);
}
expected('2026-10-21',20,23,null); // no extra calls on activation day
expected(day,8,14,null);
expected(day,8,15,'AM');
expected(day,8,19,'AM'); // scheduler a few minutes late
expected(day,8,42,'AM'); // half-hour fallback
expected(day,10,14,'AM'); // delayed runner within catch-up window
expected(day,10,15,null); // stale notifications expire
expected(day,12,0,null);
expected(day,20,14,null);
expected(day,20,15,'PM');
expected(day,20,19,'PM'); // scheduler delayed
expected(day,20,42,'PM'); // fallback
expected(day,22,14,'PM');
expected(day,22,15,null);
expected('2026-10-25',20,42,'PM');
expected('2026-10-26',8,42,null); // meeting day: no role blasts
expected('2026-10-27',8,42,null); // expired meeting
assert.throws(() => due(day,at(8,42),meeting,day,'bad,20:15'));
assert.ok(source.includes("!meetingAutomationEligible_(meetingNo)"),
  'Every run must recheck active meeting');
assert.ok(source.includes('availableRoleDefs_(row)'),
  'Every run must recheck live vacancies');
assert.ok(source.includes('previouslyAttempted.has(key)'),
  'Use persistent outbox deduplication');
assert.ok(source.includes('LockService.getScriptLock()'),
  'Avoid concurrent duplicate broadcasts');
assert.ok(source.includes('sendApprovedRoleCall_(member.phone, meetingNo, date)'),
  'Only approved Utility V3 may be used');
assert.ok(source.includes("const key = [\n          'ROLE_DAILY', meetingNo, due.date, due.slot, member.name"),
  'Dedup key is meeting + date + slot + recipient');
assert.ok(cron.includes("cron: '46 2,14 * * *'"),
  'Dedicated approximately 08:16/20:16 IST runner required');
assert.ok(cron.includes("cron: '12,42 * * * *'"),
  'Keep 30-minute fallback runner');
console.log('Twice-daily due-based role reminders: boundary, late-tick, lock, no-duplicate checks PASS');
