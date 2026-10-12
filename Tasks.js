/**
 * Central Phase 1 task engine.
 * Tasks are deterministic and rebuilt from the canonical meeting plan.
 * Existing reminder/send history is preserved across synchronizations.
 */

const TASK_SHEET_ = 'Automation Tasks';
const SUBMISSION_SHEET_ = 'WA Submissions';

// ATK_AGENDA_TEMPLATE_SUBMIT_RETRY_20261002
function phase1ScheduledTick() {
  if (!configBool_('TASK_ENGINE_ENABLED')) return;

  // Member introductions are intentionally opt-in. Keep this workflow off unless
  // the VPE explicitly enables INTRO_WORKFLOW_ENABLED for a future use case.
  const introWorkflowEnabled = configBool_('INTRO_WORKFLOW_ENABLED');

  try {
    ensureProgrammePlanningSetup_();
    syncProgrammeSlots_();
    ensureProgrammePlanningTemplates_();
  } catch (error) {
    logAutomation_('Programme Planning', '', '', 'Programme bootstrap', 'Error', String(error));
  }

  try {
    const agendaTemplateStatus = String(configValue_('WA_TEMPLATE_AGENDA_APPROVAL_STATUS', '')).toUpperCase();
    if (agendaTemplateStatus !== 'SUBMITTED' && agendaTemplateStatus !== 'APPROVED') {
      const templateResult = createAtkWhatsAppTemplates_();
      logAutomation_('WhatsApp Templates', '', '', 'Agenda approval template submission',
        templateResult && templateResult.ok ? 'Success' : 'Error',
        JSON.stringify(templateResult || {}));
    }
    syncAtkTemplateApprovalStatus_();
  } catch (error) {
    logAutomation_('WhatsApp Templates', '', '', 'Approval status sync', 'Error', String(error));
  }

  try {
    sendApprovedRoleReminderV2DevTestOnce_();
  } catch (error) {
    logAutomation_('WhatsApp Templates', '', '', 'Role reminder v2 DEV test',
      'Error', String(error));
  }

  syncPhase1Tasks_();
  processVacantRoleReminders_();
  processSpeakerDetailPrompts_();
  processSpeakerDetailReminders_();
  processAssignedRolePrompts_();
  processGrammarFlowTestStartOnce_();
  processVpprThemeHandoffs_();
  // Run scheduled reminders first, then catch up any pending Grammarian whose
  // original assignment prompt was never sent (for example, manual planner entry).
  // This order prevents a first-time catch-up and a reminder firing in the same tick.
  processGrammarAssignmentReminder_();
  processMissingGrammarAssignmentPrompts_();
  if (introWorkflowEnabled) processIntroChoiceTimeouts_();
  processThemeRequestReminder_();
  if (introWorkflowEnabled) processReadyIntroductionBundles_();
  processReadySpeakerDetailsToTmod_();
  if (introWorkflowEnabled) processMemberIntroOnboarding_();
  processEvaluatorNotifications_();
  processMonthlyProgrammePlanning_();
  processSpeakingSlotTargeting_();
  updateControlCenterReadiness_();
  processAgendaApprovalAutomation_();
  processVpeExceptionDigest_();
}

function installEvaluatorEditTrigger() {
  const handler = 'handlePlannerEvaluatorEdit_';

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger(handler)
    .forSpreadsheet(SPREADSHEET_ID)
    .onEdit()
    .create();

  logAutomation_(
    'Evaluation Coordination',
    '',
    configValue_('VPE_MEMBER_NAME', 'Riyas AF'),
    'Installed evaluator assignment edit trigger',
    'Success',
    handler
  );

  return 'Evaluator assignment edit trigger installed ✅';
}

function installPhase1Triggers_() {
  const hourlyHandler = 'phase1ScheduledTick';
  const evaluatorEditHandler = 'handlePlannerEvaluatorEdit_';
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    const handler = trigger.getHandlerFunction();
    if (handler === hourlyHandler || handler === evaluatorEditHandler) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger(hourlyHandler)
    .timeBased()
    .everyHours(1)
    .create();

  ScriptApp.newTrigger(evaluatorEditHandler)
    .forSpreadsheet(SPREADSHEET_ID)
    .onEdit()
    .create();

  logAutomation_(
    'Task Engine',
    '',
    configValue_('VPE_MEMBER_NAME', 'Riyas AF'),
    'Installed hourly scheduler + evaluator edit trigger',
    'Success',
    hourlyHandler + ' | ' + evaluatorEditHandler
  );

  return 'ATK Phase 1 setup is complete ✅\nHourly scheduler and evaluator-assignment edit trigger are active.';
}

function syncPhase1Tasks_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  const taskSheet = ss.getSheetByName(TASK_SHEET_);
  if (!planner || !taskSheet) throw new Error('Task engine sheets are missing.');

  const introWorkflowEnabled = configBool_('INTRO_WORKFLOW_ENABLED');
  const existing = readExistingTasks_(taskSheet);
  const submissions = readSubmissionIndex_(ss);
  const plannerRowCount = Math.max(0, planner.getLastRow() - 2);
  const plannerRows = plannerRowCount
    ? planner.getRange(3, 1, plannerRowCount, 28).getValues()
    : [];

  const today = startOfDay_(new Date());
  const scanDays = Math.max(7, Number(configValue_('REMINDER_WINDOW_DAYS', '30')) || 30);
  const scanEnd = new Date(today.getTime() + scanDays * 86400000);
  const generated = {};

  plannerRows.forEach((row, idx) => {
    const meetingNo = Number(row[0]);
    const meetingDate = row[1];
    const meetingStatus = String(row[2] || '');

    if (!meetingNo || !(meetingDate instanceof Date)) return;
    if (meetingStatus === 'Completed' || meetingStatus === 'Cancelled') return;

    const meetingDay = startOfDay_(meetingDate);
    if (meetingDay < today || meetingDay > scanEnd) return;

    // Evaluator slots are tracked by speaker-specific assignment tasks below.
    ROLE_DEFS_.filter(def => !def.evaluatorSlot).forEach(def => {
      const member = String(row[def.column - 1] || '').trim();
      const key = taskId_(meetingNo, 'VACANT_ROLE', def.code, '');

      generated[key] = buildTask_(existing[key], {
        taskId: key,
        meetingNo,
        taskType: 'VACANT_ROLE',
        member: member,
        related: def.code,
        state: member ? 'Completed' : 'Pending',
        dueAt: daysBefore_(meetingDate, Number(configValue_('ROLE_VACANCY_DUE_DAYS', '5')) || 5),
        reference: 'VPE Weekly Planner!' + columnLetter_(def.column) + (idx + 3),
        payload: JSON.stringify({ roleCode: def.code, roleName: def.name })
      });

      // Role-player introduction task once the role is filled. TMOD does not need
      // an introduction for themselves.
      if (introWorkflowEnabled && member && def.code !== 'TMOD') {
        const introKey = taskId_(meetingNo, 'INTRO', def.code, member);
        const introDone = submissions.has(submissionKey_(meetingNo, member, 'INTRO', def.code));

        generated[introKey] = buildTask_(existing[introKey], {
          taskId: introKey,
          meetingNo,
          taskType: 'INTRO',
          member,
          related: def.code,
          state: introDone ? 'Completed' : 'Pending',
          dueAt: daysBefore_(meetingDate, Number(configValue_('INTRO_DUE_DAYS', '3')) || 3),
          reference: def.name,
          payload: JSON.stringify({ roleCode: def.code, roleName: def.name })
        });
      }

      if (member && def.code === 'TMOD') {
        const themeKey = taskId_(meetingNo, 'THEME', def.code, member);
        const themeDone = Boolean(String(row[3] || '').trim());

        generated[themeKey] = buildTask_(existing[themeKey], {
          taskId: themeKey,
          meetingNo,
          taskType: 'THEME',
          member,
          related: def.code,
          state: themeDone ? 'Completed' : 'Pending',
          dueAt: daysBefore_(meetingDate, Number(configValue_('THEME_DUE_DAYS', '6')) || 6),
          reference: 'VPE Weekly Planner!D' + (idx + 3),
          payload: JSON.stringify({})
        });
      }

      if (member && def.code === 'GRM') {
        const grammarKey = taskId_(meetingNo, 'GRAMMAR', def.code, member);
        const legacyGrammarDone =
          submissions.has(submissionKey_(meetingNo, member, 'GRAMMAR', def.code));
        const grammarWordDone =
          submissions.has(submissionKey_(meetingNo, member, 'GRAMMAR_WORD', def.code));
        const grammarIdiomDone =
          submissions.has(submissionKey_(meetingNo, member, 'GRAMMAR_IDIOM', def.code));
        const grammarDone = legacyGrammarDone || (grammarWordDone && grammarIdiomDone);

        generated[grammarKey] = buildTask_(existing[grammarKey], {
          taskId: grammarKey,
          meetingNo,
          taskType: 'GRAMMAR',
          member,
          related: def.code,
          state: grammarDone ? 'Completed' : 'Pending',
          dueAt: daysBefore_(meetingDate, Number(configValue_('GRAMMARIAN_DUE_DAYS', '3')) || 3),
          reference: 'Word + Meaning + Usage / Idiom + Meaning + Usage',
          payload: JSON.stringify({
            flow: legacyGrammarDone ? 'LEGACY_COMPLETE' : 'SPLIT_WORD_IDIOM',
            wordDone: grammarWordDone,
            idiomDone: grammarIdiomDone
          })
        });
      }
    });

    const speakers = [
      { slot: 1, member: String(row[6] || '').trim(), evaluator: String(row[25] || '').trim(), title: String(row[6] || '').trim() },
      { slot: 2, member: String(row[10] || '').trim(), evaluator: String(row[26] || '').trim(), title: String(row[10] || '').trim() },
      { slot: 3, member: String(row[14] || '').trim(), evaluator: String(row[27] || '').trim(), title: String(row[14] || '').trim() }
    ];

    speakers.forEach(sp => {
      const slotRef = 'S' + sp.slot;
      const programmeState = programmeSlotState_(meetingNo, sp.slot);
      const programmeBlocked =
        programmeState === 'VPE_HOLD' || programmeState === 'SESSION_CONFIRMED';

      const openKey = taskId_(meetingNo, 'SPEAKING_SLOT_OPEN', slotRef, '');
      generated[openKey] = buildTask_(existing[openKey], {
        taskId: openKey,
        meetingNo,
        taskType: 'SPEAKING_SLOT_OPEN',
        member: sp.member,
        related: slotRef,
        state: (sp.member || programmeBlocked) ? 'Completed' : 'Pending',
        dueAt: daysBefore_(meetingDate, Number(configValue_('SPEAKING_SLOT_DUE_DAYS', '10')) || 10),
        reference: 'Speaker slot ' + sp.slot,
        payload: JSON.stringify({ slot: sp.slot, programmeState: programmeState })
      });

      if (!sp.member) return;

      const detailsKey = taskId_(meetingNo, 'SPEAKER_DETAILS', slotRef, sp.member);
      const detailsDone = submissions.has(submissionKey_(meetingNo, sp.member, 'SPEAKER_DETAILS', slotRef));

      generated[detailsKey] = buildTask_(existing[detailsKey], {
        taskId: detailsKey,
        meetingNo,
        taskType: 'SPEAKER_DETAILS',
        member: sp.member,
        related: slotRef,
        state: detailsDone ? 'Completed' : 'Pending',
        dueAt: daysBefore_(meetingDate, Number(configValue_('SPEAKER_DETAILS_DUE_DAYS', '6')) || 6),
        reference: 'Speaker slot ' + sp.slot,
        payload: JSON.stringify({ slot: sp.slot })
      });

      if (introWorkflowEnabled) {
      const speakerIntroKey = taskId_(meetingNo, 'INTRO', slotRef, sp.member);
      const speakerIntroDone = submissions.has(submissionKey_(meetingNo, sp.member, 'INTRO', slotRef));
      generated[speakerIntroKey] = buildTask_(existing[speakerIntroKey], {
        taskId: speakerIntroKey,
        meetingNo,
        taskType: 'INTRO',
        member: sp.member,
        related: slotRef,
        state: speakerIntroDone ? 'Completed' : 'Pending',
        dueAt: daysBefore_(meetingDate, Number(configValue_('INTRO_DUE_DAYS', '7')) || 7),
        reference: 'Speaker ' + sp.slot,
        payload: JSON.stringify({ slot: sp.slot })
      });

      }

      const evaluatorKey = taskId_(meetingNo, 'EVALUATOR_ASSIGNMENT', slotRef, sp.member);
      generated[evaluatorKey] = buildTask_(existing[evaluatorKey], {
        taskId: evaluatorKey,
        meetingNo,
        taskType: 'EVALUATOR_ASSIGNMENT',
        member: configValue_('VPE_MEMBER_NAME', 'Riyas AF'),
        related: slotRef,
        state: sp.evaluator ? 'Completed' : 'Pending',
        dueAt: daysBefore_(meetingDate, Number(configValue_('EVALUATOR_ASSIGNMENT_DUE_DAYS', '7')) || 7),
        reference: sp.member,
        payload: JSON.stringify({ speaker: sp.member, evaluator: sp.evaluator, slot: sp.slot })
      });

      if (sp.evaluator) {
        const evaluatorRole = 'EV' + sp.slot;
        if (introWorkflowEnabled) {
        const evaluatorIntroKey = taskId_(meetingNo, 'INTRO', evaluatorRole, sp.evaluator);
        const evaluatorIntroDone = submissions.has(
          submissionKey_(meetingNo, sp.evaluator, 'INTRO', evaluatorRole));
        generated[evaluatorIntroKey] = buildTask_(existing[evaluatorIntroKey], {
          taskId: evaluatorIntroKey,
          meetingNo,
          taskType: 'INTRO',
          member: sp.evaluator,
          related: evaluatorRole,
          state: evaluatorIntroDone ? 'Completed' : 'Pending',
          dueAt: daysBefore_(meetingDate, Number(configValue_('INTRO_DUE_DAYS', '7')) || 7),
          reference: 'Evaluator for ' + sp.member,
          payload: JSON.stringify({roleCode: evaluatorRole, roleName: 'Evaluator ' + sp.slot})
        });
        }
        const coordinationKey = taskId_(meetingNo, 'EVALUATOR_COORDINATION', slotRef, sp.member);
        const speakerPairKey = ['EVALUATOR_PAIR_SPEAKER', meetingNo, slotRef, sp.member, sp.evaluator].join('|');
        const evaluatorPairKey = ['EVALUATOR_PAIR_EVALUATOR', meetingNo, slotRef, sp.evaluator, sp.member].join('|');
        const pairInformed = hasReminderBeenSent_(speakerPairKey) && hasReminderBeenSent_(evaluatorPairKey);

        generated[coordinationKey] = buildTask_(existing[coordinationKey], {
          taskId: coordinationKey,
          meetingNo,
          taskType: 'EVALUATOR_COORDINATION',
          member: sp.member,
          related: slotRef,
          state: pairInformed ? 'Completed' : 'Pending',
          dueAt: daysBefore_(meetingDate, Number(configValue_('EVALUATOR_COORDINATION_DUE_DAYS', '5')) || 5),
          reference: sp.evaluator,
          payload: JSON.stringify({
            speaker: sp.member,
            evaluator: sp.evaluator,
            slot: sp.slot,
            mode: 'DIRECT_COMMUNICATION_REMINDER_ONLY'
          })
        });
      }
    });
  });

  // Preserve historical/finished records and cancel only generated tasks that
  // are no longer applicable.
  Object.keys(existing).forEach(key => {
    if (!generated[key]) {
      const row = existing[key];
      if (row[5] === 'Pending' || row[5] === 'Sent' || row[5] === 'Escalated') {
        row[5] = 'Cancelled';
        row[13] = new Date();
      }
      generated[key] = row;
    }
  });

  const rows = Object.keys(generated)
    .sort()
    .map(key => generated[key]);

  const maxRows = Math.max(0, taskSheet.getMaxRows() - 1);
  if (rows.length > maxRows) {
    throw new Error('Automation Tasks capacity exceeded: ' + rows.length);
  }

  taskSheet.getRange(2, 1, maxRows, 18).clearContent();
  if (rows.length) {
    taskSheet.getRange(2, 1, rows.length, 18).setValues(rows);
  }

  logAutomation_('Task Engine', '', '', 'Task synchronization', 'Success', rows.length + ' task rows');
  return rows.length;
}

function readExistingTasks_(sheet) {
  const out = {};
  if (sheet.getLastRow() < 2) return out;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 18).getValues();
  rows.forEach(row => {
    const key = String(row[0] || '').trim();
    if (key) out[key] = row;
  });
  return out;
}

function readSubmissionIndex_(ss) {
  const map = new Set();
  const sheet = ss.getSheetByName(SUBMISSION_SHEET_);
  if (!sheet || sheet.getLastRow() < 2) return map;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
  rows.forEach(row => {
    const meetingNo = Number(row[1]);
    const member = String(row[2] || '').trim();
    const type = String(row[3] || '').trim();
    const related = String(row[4] || '').trim();
    if (meetingNo && member && type) {
      map.add(submissionKey_(meetingNo, member, type, related));
    }
  });
  return map;
}

function buildTask_(existing, data) {
  const row = existing ? existing.slice() : new Array(18).fill('');
  const now = new Date();
  const previousState = String(row[5] || '');

  row[0] = data.taskId;
  row[1] = data.meetingNo;
  row[2] = data.taskType;
  row[3] = data.member || '';
  row[4] = data.related || '';
  row[5] = data.state || 'Pending';
  row[6] = data.dueAt || '';
  row[10] = data.reference || '';
  row[11] = data.payload || '';
  row[12] = row[12] || now;
  row[13] = now;

  if (row[5] === 'Completed' && previousState !== 'Completed') {
    row[14] = now;
  }
  if (row[5] !== 'Completed') {
    row[14] = row[14] || '';
  }

  return row;
}

function taskId_(meetingNo, type, related, member) {
  return [
    'M' + Number(meetingNo),
    String(type || '').replace(/[^A-Z0-9_]/gi, '_'),
    String(related || '').replace(/[^A-Z0-9_]/gi, '_'),
    String(member || '').replace(/[^A-Z0-9_]/gi, '_')
  ].join('|');
}

function submissionKey_(meetingNo, member, type, related) {
  return [
    Number(meetingNo),
    String(member || '').trim().toLowerCase(),
    String(type || '').trim().toUpperCase(),
    String(related || '').trim().toUpperCase()
  ].join('|');
}

function daysBefore_(date, days) {
  const d = startOfDay_(date);
  return new Date(d.getTime() - Math.max(0, Number(days) || 0) * 86400000);
}

function columnLetter_(column) {
  let n = Number(column);
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function buildPhase1StatusText_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(TASK_SHEET_);
  if (!sheet) return 'Phase 1 task sheet is missing.';

  if (configBool_('TASK_ENGINE_ENABLED')) {
    syncPhase1Tasks_();
  }

  const rows = sheet.getLastRow() >= 2
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 18).getValues()
    : [];

  const pending = rows.filter(r => r[5] === 'Pending').length;
  const completed = rows.filter(r => r[5] === 'Completed').length;
  const escalated = rows.filter(r => r[5] === 'Escalated').length;

  return [
    'ATK Phase 1 status',
    'Pending: ' + pending,
    'Completed: ' + completed,
    'Escalated: ' + escalated,
    'Role selector: ' + (configBool_('ROLE_SELECTION_ENABLED') ? 'ON' : 'OFF'),
    'Vacancy reminders: ' + (configBool_('VACANT_ROLE_REMINDERS_ENABLED') ? 'ON' : 'OFF')
  ].join('\n');
}

/**
 * Reminder logic is intentionally gated. It creates no outbound messages unless
 * VACANT_ROLE_REMINDERS_ENABLED is explicitly TRUE.
 */
function processVacantRoleReminders_() {
  if (!configBool_('VACANT_ROLE_REMINDERS_ENABLED')) return;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  const members = ss.getSheetByName('Member Master');
  if (!planner || !members) return;

  const today = startOfDay_(new Date());
  const configuredDays = String(configValue_('ROLE_REMINDER_DAYS', '7,3'))
    .split(',')
    .map(v => Number(String(v).trim()))
    .filter(v => Number.isFinite(v) && v >= 0);

  const plannerRows = planner.getRange(3, 1, Math.max(0, planner.getLastRow() - 2), 28).getValues();
  const memberRows = members.getRange(2, 1, Math.max(0, members.getLastRow() - 1), 4).getValues();

  plannerRows.forEach(row => {
    const meetingNo = Number(row[0]);
    const meetingDate = row[1];
    const status = String(row[2] || '');
    if (!meetingNo || !(meetingDate instanceof Date) || status === 'Completed' || status === 'Cancelled') return;

    const daysLeft = Math.round((startOfDay_(meetingDate) - today) / 86400000);
    if (configuredDays.indexOf(daysLeft) === -1) return;

    const openRoles = availableRoleDefs_(row);
    if (!openRoles.length) return;

    const speakers = new Set([row[6], row[10], row[14]].map(v => String(v || '').trim()).filter(Boolean));
    const assigned = new Set(ROLE_DEFS_.map(def => String(row[def.column - 1] || '').trim()).filter(Boolean));

    const eligible = memberRows
      .filter(mr => String(mr[3] || '').trim() === 'Active')
      .map(mr => ({
        name: String(mr[1] || '').trim(),
        phone: normalizePhone_(mr[2] || '')
      }))
      .filter(m => m.name && m.phone && !speakers.has(m.name) && !assigned.has(m.name));

    eligible.forEach(member => {
      // In DEV, do not create failed outbox entries for members outside the test allowlist.
      if (!isWhatsAppRecipientAllowed_(member.phone)) return;
      const idempotency = ['ROLE_REMINDER', meetingNo, daysLeft, member.name].join('|');
      if (hasReminderBeenSent_(idempotency)) return;

      const roleNames = openRoles.map(r => r.name).join(', ');
      const revised = templateNameForKind_('ROLE_REMINDER') === 'atk_role_reminder_v2';
      const text = revised
        ? 'Hi,\n\nHope you are attending the meeting this week.\n\n' +
          'Please tap below to see the available roles for ATK Meeting No. ' +
          meetingNo + '.\n\nThank you.'
        : 'ATK Meeting ' + meetingNo + ' still has open roles: ' + roleNames +
          '.\n\nIf you would like to take one, reply ROLE to see the live available list.';
      const result = revised
        ? sendAutomatedQuickReply_(member.phone, 'ROLE_REMINDER', {
            messaging_product: 'whatsapp',
            recipient_type: 'individual',
            to: member.phone,
            type: 'interactive',
            interactive: {
              type: 'button',
              body: {text: text},
              action: {buttons: [{type: 'reply', reply: {id: 'ROLE_MENU', title: 'View roles'}}]}
            }
          }, [meetingNo], 'ROLE_MENU|' + meetingNo)
        : sendAutomatedText_(member.phone, 'ROLE_REMINDER', text,
            [member.name, meetingNo, roleNames]);
      recordReminderSend_(idempotency, member, meetingNo, result, text);
    });
  });
}

function hasReminderBeenSent_(idempotencyKey) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('WA Outbox');
  if (!sheet || sheet.getLastRow() < 2) return false;

  const rows = sheet.getRange(2, 7, sheet.getLastRow() - 1, 10).getValues();
  const successfulStates = new Set(['SENT', 'DELIVERED', 'READ', 'ACCEPTED']);

  return rows.some(row =>
    String(row[9] || '') === idempotencyKey &&
    successfulStates.has(String(row[0] || '').toUpperCase())
  );
}

function recordReminderSend_(idempotencyKey, member, meetingNo, result, text) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('WA Outbox');
  if (!sheet) return;

  const now = new Date();
  sheet.appendRow([
    Utilities.getUuid(),
    member.name,
    member.phone,
    'ROLE_REMINDER',
    'ROLE',
    text,
    result.ok ? 'SENT' : 'FAILED',
    1,
    '',
    result.messageId || '',
    meetingNo,
    '',
    now,
    result.ok ? now : '',
    result.error || '',
    idempotencyKey
  ]);
}


function processReadyIntroductionBundles_() {
  if (!configBool_('INTRO_WORKFLOW_ENABLED')) return;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  if (!planner) return;

  const today = startOfDay_(new Date());
  const dueDays = Math.max(
    0,
    Number(configValue_('INTRO_BUNDLE_DUE_DAYS', '2')) || 2
  );

  const rows = planner.getRange(
    3,
    1,
    Math.max(0, planner.getLastRow() - 2),
    24
  ).getValues();

  rows.forEach(row => {
    const meetingNo = Number(row[0]);
    const meetingDate = row[1];
    const status = String(row[2] || '');

    if (!meetingNo || !(meetingDate instanceof Date)) return;
    if (status === 'Completed' || status === 'Cancelled') return;

    const daysLeft = Math.round(
      (startOfDay_(meetingDate) - today) / 86400000
    );
    if (daysLeft !== dueDays) return;

    sendConsolidatedIntroductionsIfReady_(meetingNo);
  });
}

/**
 * Sends one approved role reminder to the sole DEV test recipient for the
 * next meeting where that member can take a vacant role.
 */
function sendDevRoleReminderForNextVacancy_() {
  if (configValue_('SYSTEM_MODE','') !== 'DEV' ||
      configValue_('WA_SEND_MODE','') !== 'TEST') {
    throw new Error('DEV role reminder test requires DEV TEST mode.');
  }

  const member = configValue_('VPE_MEMBER_NAME','Riyas AF');
  const phone = normalizePhone_(phoneForMember_(member));
  const allowlist = String(configValue_('WA_TEST_RECIPIENTS',''))
    .split(',').map(normalizePhone_).filter(Boolean);

  if (!phone || allowlist.length !== 1 || allowlist[0] !== phone ||
      !isWhatsAppRecipientAllowed_(phone)) {
    throw new Error('DEV role reminder test requires the VPE as the sole allowlisted recipient.');
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  if (!planner) throw new Error('VPE Weekly Planner not found.');

  const today = startOfDay_(new Date());
  const windowDays = Math.max(
    1,
    Number(configValue_('ROLE_SELECTION_WINDOW_DAYS','14')) || 14
  );
  const latest = new Date(today.getTime() + windowDays * 86400000);
  const rowCount = Math.max(0, planner.getLastRow() - 2);
  const rows = rowCount ? planner.getRange(3,1,rowCount,28).getValues() : [];

  let meeting = null;
  for (let i=0; i<rows.length; i++) {
    const row = rows[i];
    const meetingNo = Number(row[0]);
    const meetingDate = row[1];
    const status = String(row[2] || '');
    if (!meetingNo || !(meetingDate instanceof Date)) continue;

    const d = startOfDay_(meetingDate);
    if (d <= today || d > latest) continue;
    if (status === 'Completed' || status === 'Cancelled') continue;

    const speakers = [row[6],row[10],row[14]]
      .map(v => String(v || '').trim()).filter(Boolean);
    if (speakers.indexOf(member) !== -1) continue;

    const assigned = ROLE_DEFS_.some(def =>
      String(row[def.column - 1] || '').trim() === member
    );
    if (assigned) continue;

    const available = availableRoleDefs_(row);
    if (!available.length) continue;

    meeting = {
      meetingNo:meetingNo,
      meetingDate:meetingDate,
      available:available,
      rowNumber:i+3,
      reason:'AVAILABLE'
    };
    break;
  }

  if (!meeting) {
    return {ok:false,sent:false,reason:'NO_AVAILABLE_FUTURE_MEETING'};
  }

  const templateName = templateNameForKind_('ROLE_REMINDER') || 'atk_role_reminder_v2';
  if (!isApprovedTemplate_(templateName)) {
    throw new Error('Approved role reminder template is unavailable: ' + templateName);
  }

  const result = sendWhatsAppTemplate_(
    phone, templateName, [meeting.meetingNo], 'ROLE_MENU|' + meeting.meetingNo);

  logAutomation_('Role Selection', meeting.meetingNo, member,
    'DEV role reminder test sent',
    result && result.ok ? 'Sent' : 'Error',
    result && (result.messageId || result.error) || '');

  return {
    ok:Boolean(result && result.ok),
    sent:Boolean(result && result.ok),
    meeting_no:meeting.meetingNo,
    message_id:result && result.messageId || '',
    error:result && result.error || ''
  };
}


/**
 * One controlled test of the approved View roles template. Runs through the
 * DEV scheduler, never through the member reminder pilot.
 */
function sendApprovedRoleReminderV2DevTestOnce_() {
  if (configValue_('SYSTEM_MODE', '') !== 'DEV' ||
      configValue_('WA_SEND_MODE', '') !== 'TEST' ||
      configBool_('VACANT_ROLE_REMINDERS_ENABLED')) return;

  const templateName = 'atk_role_reminder_v2';
  const testKey = 'DEV_ROLE_V2_TEST|945075534849918';
  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (props.getProperty(testKey) || hasReminderBeenSent_(testKey)) return;
    if (!isApprovedTemplate_(templateName)) return;

    const member = configValue_('VPE_MEMBER_NAME', 'Riyas AF');
    const phone = normalizePhone_(phoneForMember_(member));
    const allowlist = String(configValue_('WA_TEST_RECIPIENTS', ''))
      .split(',').map(normalizePhone_).filter(Boolean);
    if (member !== 'Riyas AF' || allowlist.length !== 1 ||
        allowlist[0] !== phone || !isWhatsAppRecipientAllowed_(phone)) {
      throw new Error('DEV role test requires Riyas as the sole allowlisted recipient.');
    }

    const meetingNo = 451;
    const result = sendWhatsAppTemplate_(
      phone, templateName, [meetingNo], 'ROLE_MENU');
    const body = 'View roles DEV test for ATK Meeting No. ' + meetingNo;
    if (result && result.ok) {
      // Write the guard before logging so a later logging error cannot resend.
      props.setProperty(testKey, result.messageId || 'META_ACCEPTED');
    }
    recordReminderSend_(testKey, {name: member, phone: phone},
      meetingNo, result || {ok: false, error: 'No Meta response'}, body);
    logAutomation_('WhatsApp Templates', meetingNo, member,
      'Role reminder v2 DEV test',
      result && result.ok ? 'Sent' : 'Error',
      result && (result.messageId || result.error) || 'No Meta response');
  } finally {
    lock.releaseLock();
  }
}


/* ATK_AGENDA_APPROVAL_SCHEDULER_V1 */
function processAgendaApprovalAutomation_() {
  if (!configBool_('AGENDA_APPROVAL_WORKFLOW_ENABLED')) return;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  const preview = ss.getSheetByName('Agenda Preview');
  if (!planner || !preview) return;

  const previewMeetingNo = Number(preview.getRange('B4').getValue());
  if (!previewMeetingNo) return;

  const ids = planner.getRange(3,1,Math.max(0, planner.getLastRow() - 2),1)
    .getValues().flat();
  const idx = ids.findIndex(v => Number(v) === previewMeetingNo);
  if (idx < 0) return;

  const meetingDate = planner.getRange(idx+3,2).getValue();
  if (!(meetingDate instanceof Date)) return;

  // The preview itself is the current canonical readiness surface. If any
  // explicit pending markers remain, readiness is false; the hard Saturday
  // deadline may still force publication for approval.
  const display = preview.getDataRange().getDisplayValues();
  const pending = display.some(row => row.some(v => /^\s*(\[Pending\]|\[Pending time\]|Pending)\s*$/i.test(String(v||''))));
  const decision = agendaPublicationDecision_(meetingDate, {ready:!pending});
  if (!decision.publishNow) return;

  const current = latestAgendaApprovalRecord_(previewMeetingNo);
  const fingerprint = agendaPreviewFingerprint_();
  if (current && String(current.values[4]||'') === fingerprint &&
      ['AWAITING_APPROVAL','APPROVED'].indexOf(String(current.values[5]||'')) >= 0) {
    return;
  }

  sendAgendaForApproval_(previewMeetingNo);
}
