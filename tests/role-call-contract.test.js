'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const source = fs.readFileSync('Messaging.js', 'utf8');
const tasks = fs.readFileSync('Tasks.js', 'utf8');
const webApi = fs.readFileSync('WebApiV1.js', 'utf8');
const templates = fs.readFileSync('RoleReminderV2.js', 'utf8');
const setup = fs.readFileSync('TemplateSetup.js', 'utf8');
const context = {
  console,
  normalizePhone_: String,
  configValue_: (_name, fallback) => fallback,
  Utilities: {formatDate: (_date, _zone, _format) => '26 October 2026'}
};
vm.createContext(context);
vm.runInContext(source, context);
let approved = true;
let approvalStatus = 'APPROVED · UTILITY';
context.SPREADSHEET_ID = 'unit-test-only';
context.SpreadsheetApp = {
  openById: () => ({
    getSheetByName: () => ({
      getLastRow: () => 48,
      getRange: () => ({
        getValues: () => [['atk_role_reminder_v3', '', approvalStatus]]
      })
    })
  })
};
const sends = [];
context.isApprovedTemplate_ = name => approved && name === 'atk_role_reminder_v3';
context.sendWhatsAppTemplate_ = (...args) => {
  sends.push(args);
  return {ok: true, messageId: 'mock-meta-message'};
};
// An open service window must not change role calls into freeform sends.
context.hasOpenWhatsAppServiceWindow_ = () => true;
context.sendWhatsAppPayload_ = () => {
  throw new Error('Role call must never use freeform payload');
};
context.sendWhatsAppText_ = () => {
  throw new Error('Role call must never use text-only payload');
};
assert.equal(context.templateNameForKind_('ROLE_REMINDER'), 'atk_role_reminder_v3');
let response = context.sendApprovedRoleCall_('919746462423', 454, '2026-10-26');
assert.equal(response.ok, true);
assert.equal(sends.length, 1);
assert.deepEqual(Array.from(sends[0][2]), [454, '26 October 2026']);
assert.equal(sends[0][1], 'atk_role_reminder_v3');
assert.equal(sends[0][3], 'ROLE_MENU|454');

response = context.sendAutomatedQuickReply_(
  '919746462423', 'ROLE_REMINDER', {}, [455, '2026-11-02'], 'ROLE_MENU|455'
);
assert.equal(response.ok, true);
assert.equal(sends[1][1], 'atk_role_reminder_v3');
assert.deepEqual(Array.from(sends[1][2]), [455, '2 November 2026']);
response = context.sendAutomatedText_(
  '919746462423', 'ROLE_REMINDER', 'Reply ROLE', [456, '2026-11-09']
);
assert.equal(response.ok, true);
assert.equal(sends[2][1], 'atk_role_reminder_v3');
assert.deepEqual(Array.from(sends[2][2]), [456, '9 November 2026']);

const sentCount = sends.length;
approvalStatus = 'REJECTED · UTILITY';
response = context.sendApprovedRoleCall_('919746462423', 454, '2026-10-26');
assert.equal(response.ok, false);
assert.equal(response.deferred, true);
assert.equal(sends.length, sentCount);
approvalStatus = 'APPROVED · MARKETING';
response = context.sendApprovedRoleCall_('919746462423', 454, '2026-10-26');
assert.equal(response.ok, false);
assert.equal(response.deferred, true);
assert.equal(sends.length, sentCount);
approvalStatus = 'PENDING · UTILITY';
response = context.sendApprovedRoleCall_('919746462423', 454, '2026-10-26');
assert.equal(response.ok, false);
assert.equal(sends.length, sentCount);
approvalStatus = 'APPROVED · UTILITY';
response = context.sendApprovedRoleCall_('919746462423', 454, '2026-02-30');
assert.equal(response.ok, false);
assert.equal(sends.length, sentCount);

assert.ok(tasks.includes('sendApprovedRoleCall_(member.phone, meetingNo, meetingDate)'));
assert.ok(tasks.includes('sendApprovedRoleCall_(member.phone, meetingNo, row[1])'));
assert.ok(!tasks.includes("'atk_role_reminder_v2'"));
assert.ok(!webApi.includes("'atk_role_reminder_v2'"));
assert.ok(webApi.includes('sendApprovedRoleCall_(phone,state.meetingNo,dateText)'));
assert.ok(templates.includes("name: 'atk_role_reminder_v3'"));
assert.ok(templates.includes("category: 'UTILITY'"));
assert.ok(templates.includes("text: 'View roles'"));
assert.ok(setup.includes('roleReminderV3Spec_(),'));
console.log('Role call contract: Utility V3, meeting/date parameters, View roles and no fallback PASS');
