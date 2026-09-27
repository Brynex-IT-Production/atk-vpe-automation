/**
 * TMOD theme, Grammarian and role-player introduction workflows.
 * These flows use explicit reply prefixes so submissions are unambiguous.
 */

function handleStructuredSubmission_(member, phone, rawText, messageId) {
  const text = String(rawText || '').trim();
  if (!text) return false;

  const themeMatch = text.match(/^THEME(?:\s+(\d+))?\s*:\s*(.+)$/i);
  if (themeMatch) {
    handleThemeSubmission_(member, phone, Number(themeMatch[1] || 0), themeMatch[2], messageId);
    return true;
  }

  const grammarMatch = text.match(/^GRAMMAR(?:\s+(\d+))?\s*:\s*(.+)$/i);
  if (grammarMatch) {
    handleGrammarSubmission_(member, phone, Number(grammarMatch[1] || 0), grammarMatch[2], messageId);
    return true;
  }

  const introMatch = text.match(/^INTRO(?:\s+(\d+))?\s*:\s*(.+)$/i);
  if (introMatch) {
    handleIntroSubmission_(member, phone, Number(introMatch[1] || 0), introMatch[2], messageId);
    return true;
  }


  const speechMatch = text.match(/^SPEECH(?:\s+(\d+))?\s*:\s*(.+)$/i);
  if (speechMatch) {
    handleSpeakerDetailsSubmission_(member, phone, Number(speechMatch[1] || 0), speechMatch[2], messageId);
    return true;
  }

  // A plain-text reply can be an introduction when exactly one recent
  // introduction request is pending for this member. Explicit commands above
  // always take priority, and ambiguous replies remain unhandled.
  const pendingIntro = recentUnambiguousIntroTask_(member, phone, text);
  if (pendingIntro) {
    handleIntroSubmission_(member, phone, pendingIntro.meetingNo, text, messageId);
    return true;
  }

  return false;
}

function recentUnambiguousIntroTask_(member, phone, text) {
  if (!configBool_('INTRO_WORKFLOW_ENABLED') || text.length < 10) return null;
  if (/^(ROLE|ROLES|ADMIN|THEME|GRAMMAR|INTRO|SPEECH)\b/i.test(text)) return null;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tasks = ss.getSheetByName('Automation Tasks');
  const outbox = ss.getSheetByName('WA Outbox');
  if (!tasks || !outbox || tasks.getLastRow() < 2 || outbox.getLastRow() < 2) return null;

  const pending = tasks.getRange(2, 1, tasks.getLastRow() - 1, 18).getValues()
    .filter(r => String(r[2] || '') === 'INTRO' &&
      String(r[3] || '').trim() === member &&
      String(r[5] || '') === 'Pending');
  if (pending.length !== 1) return null;

  const meetingNo = Number(pending[0][1]);
  const rows = outbox.getRange(2, 1, outbox.getLastRow() - 1, 16).getValues();
  const now = Date.now();
  const recentRequest = rows.some(r =>
    String(r[2] || '').replace(/\D/g, '') === String(phone || '').replace(/\D/g, '') &&
    String(r[3] || '') === 'INTRO_REQUEST' &&
    Number(r[10]) === meetingNo &&
    ['SENT', 'DELIVERED', 'READ'].indexOf(String(r[6] || '')) !== -1 &&
    r[12] instanceof Date && now - r[12].getTime() >= 0 &&
    now - r[12].getTime() <= 24 * 60 * 60 * 1000
  );

  return recentRequest ? { meetingNo: meetingNo } : null;
}

function sendImmediateAssignmentPrompts_(meetingNo, roleDef, member, phone) {
  const prompts = [];

  if (roleDef.code === 'TMOD' && configBool_('TMOD_WORKFLOW_ENABLED')) {
    prompts.push(
      'Hi ' + member + ',\n\n' +
      'Please share the theme at the earliest, so that the role players can align with the same.\n\n' +
      'Thank you.\n\n' +
      'Reply in this format:\nTHEME ' + meetingNo + ': Your theme'
    );
  }

  if (roleDef.code === 'GRM' && configBool_('GRAMMARIAN_WORKFLOW_ENABLED')) {
    prompts.push(
      'As Grammarian for ATK Meeting ' + meetingNo +
      ', please send your Word and Idiom in this format:\n\nGRAMMAR ' + meetingNo +
      ': Word | Meaning | Idiom | Meaning'
    );
  }

  if (roleDef.code !== 'TMOD' && configBool_('INTRO_WORKFLOW_ENABLED')) {
    prompts.push(
      'Please also send a short introduction for the Toastmaster to use:\n\nINTRO ' +
      meetingNo + ': Your short introduction'
    );
  }

  if (!prompts.length) return;

  const body = prompts.join('\n\n');
  const result = sendWhatsAppText_(phone, body);
  recordWorkflowOutbound_(
    member,
    phone,
    'ASSIGNMENT_PROMPT',
    body,
    meetingNo,
    result,
    ['ASSIGNMENT_PROMPT', meetingNo, roleDef.code, member].join('|')
  );
}

function handleThemeSubmission_(member, phone, requestedMeetingNo, theme, messageId) {
  if (!configBool_('TMOD_WORKFLOW_ENABLED')) {
    sendWhatsAppText_(phone, 'Theme collection is not active yet.');
    return;
  }

  const task = findPendingTaskForMember_('THEME', member, requestedMeetingNo);
  if (!task) {
    sendWhatsAppText_(phone, 'I could not find a pending Toastmaster theme request for you.');
    return;
  }

  const cleanTheme = String(theme || '').trim();
  if (cleanTheme.length < 3) {
    sendWhatsAppText_(phone, 'Please send a clear theme after THEME ' + task.meetingNo + ':');
    return;
  }

  const assignment = verifyRoleAssignment_(task.meetingNo, 'TMOD', member);
  if (!assignment.ok) {
    sendWhatsAppText_(phone, 'Your Toastmaster assignment for Meeting ' + task.meetingNo + ' has changed, so I did not save this theme.');
    syncPhase1Tasks_();
    return;
  }

  assignment.sheet.getRange(assignment.rowNumber, 4).setValue(cleanTheme);
  appendSubmission_({
    meetingNo: task.meetingNo,
    member,
    type: 'THEME',
    related: 'TMOD',
    values: [cleanTheme, '', '', ''],
    sourceMessageId: messageId
  });

  syncPhase1Tasks_();
  sendWhatsAppText_(phone, 'Theme received ✅\nMeeting ' + task.meetingNo + ': ' + cleanTheme);
  logAutomation_('TMOD Theme', task.meetingNo, member, 'Theme received', 'Success', cleanTheme);
}

function handleGrammarSubmission_(member, phone, requestedMeetingNo, body, messageId) {
  if (!configBool_('GRAMMARIAN_WORKFLOW_ENABLED')) {
    sendWhatsAppText_(phone, 'Grammarian collection is not active yet.');
    return;
  }

  const task = findPendingTaskForMember_('GRAMMAR', member, requestedMeetingNo);
  if (!task) {
    sendWhatsAppText_(phone, 'I could not find a pending Grammarian submission for you.');
    return;
  }

  const parts = String(body || '').split('|').map(v => v.trim());
  if (parts.length < 4 || parts.slice(0, 4).some(v => !v)) {
    sendWhatsAppText_(phone,
      'Please use this format:\nGRAMMAR ' + task.meetingNo +
      ': Word | Meaning | Idiom | Meaning');
    return;
  }

  const assignment = verifyRoleAssignment_(task.meetingNo, 'GRM', member);
  if (!assignment.ok) {
    sendWhatsAppText_(phone, 'Your Grammarian assignment for Meeting ' + task.meetingNo + ' has changed, so I did not save this.');
    syncPhase1Tasks_();
    return;
  }

  appendSubmission_({
    meetingNo: task.meetingNo,
    member,
    type: 'GRAMMAR',
    related: 'GRM',
    values: parts.slice(0, 4),
    sourceMessageId: messageId
  });

  syncPhase1Tasks_();
  sendWhatsAppText_(phone,
    'Grammarian details received ✅\nWord: ' + parts[0] + '\nIdiom: ' + parts[2]);
  logAutomation_('Grammarian', task.meetingNo, member, 'Word and idiom received', 'Success', parts[0] + ' / ' + parts[2]);
}

function handleIntroSubmission_(member, phone, requestedMeetingNo, intro, messageId) {
  if (!configBool_('INTRO_WORKFLOW_ENABLED')) {
    sendWhatsAppText_(phone, 'Introduction collection is not active yet.');
    return;
  }

  const task = findPendingTaskForMember_('INTRO', member, requestedMeetingNo);
  if (!task) {
    sendWhatsAppText_(phone, 'I could not find a pending role-player introduction for you.');
    return;
  }

  const cleanIntro = String(intro || '').trim();
  if (cleanIntro.length < 10) {
    sendWhatsAppText_(phone,
      'Please send a slightly fuller introduction after INTRO ' + task.meetingNo + ':');
    return;
  }

  const assignment = verifyRoleAssignment_(task.meetingNo, task.related, member);
  if (!assignment.ok) {
    sendWhatsAppText_(phone, 'Your role assignment for Meeting ' + task.meetingNo + ' has changed, so I did not save this introduction.');
    syncPhase1Tasks_();
    return;
  }

  appendSubmission_({
    meetingNo: task.meetingNo,
    member,
    type: 'INTRO',
    related: task.related,
    values: [cleanIntro, '', '', ''],
    sourceMessageId: messageId
  });

  syncPhase1Tasks_();
  sendWhatsAppText_(phone, 'Introduction received ✅ for Meeting ' + task.meetingNo + '.');
  logAutomation_('Introduction', task.meetingNo, member, 'Role-player introduction received', 'Success', task.related);

  sendConsolidatedIntroductionsIfReady_(task.meetingNo);
}

function findPendingTaskForMember_(type, member, meetingNo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Automation Tasks');
  if (!sheet || sheet.getLastRow() < 2) return null;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 18).getValues();
  const candidates = rows
    .filter(r =>
      String(r[2] || '') === type &&
      String(r[3] || '').trim() === member &&
      ['Pending', 'Sent', 'Escalated'].indexOf(String(r[5] || '')) !== -1 &&
      (!meetingNo || Number(r[1]) === Number(meetingNo))
    )
    .sort((a, b) => Number(a[1]) - Number(b[1]));

  if (!candidates.length) return null;
  const r = candidates[0];
  return {
    taskId: String(r[0] || ''),
    meetingNo: Number(r[1]),
    type: String(r[2] || ''),
    member: String(r[3] || ''),
    related: String(r[4] || ''),
    state: String(r[5] || '')
  };
}

function verifyRoleAssignment_(meetingNo, roleCode, member) {
  const def = ROLE_DEFS_.find(r => r.code === roleCode);
  if (!def) return { ok: false };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('VPE Weekly Planner');
  if (!sheet) return { ok: false };

  const lastRow = Math.min(62, sheet.getLastRow());
  const ids = sheet.getRange(3, 1, Math.max(0, lastRow - 2), 1).getValues().flat();
  const idx = ids.findIndex(v => Number(v) === Number(meetingNo));
  if (idx < 0) return { ok: false };

  const rowNumber = idx + 3;
  const assigned = String(sheet.getRange(rowNumber, def.column).getValue() || '').trim();
  return { ok: assigned === member, sheet, rowNumber, assigned };
}

function appendSubmission_(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('WA Submissions');
  if (!sheet) throw new Error('WA Submissions sheet not found.');

  const sourceId = String(data.sourceMessageId || '');
  if (sourceId && sheet.getLastRow() >= 2) {
    const existingIds = sheet.getRange(2, 11, sheet.getLastRow() - 1, 1).getValues().flat();
    if (existingIds.some(v => String(v || '') === sourceId)) return false;
  }

  const now = new Date();
  const values = data.values || [];
  sheet.appendRow([
    Utilities.getUuid(),
    data.meetingNo,
    data.member,
    data.type,
    data.related || '',
    values[0] || '',
    values[1] || '',
    values[2] || '',
    values[3] || '',
    data.fileOrLink || '',
    sourceId,
    now,
    now,
    ''
  ]);
  return true;
}

let MEMBER_PHONE_CACHE_ = null;

function getMemberPhoneMap_() {
  if (MEMBER_PHONE_CACHE_) return MEMBER_PHONE_CACHE_;

  const map = {};
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Member Master');

  if (sheet && sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    rows.forEach(r => {
      const name = String(r[1] || '').trim();
      const status = String(r[3] || '').trim();
      if (name && status === 'Active') {
        map[name] = normalizePhone_(r[2] || '');
      }
    });
  }

  MEMBER_PHONE_CACHE_ = map;
  return MEMBER_PHONE_CACHE_;
}

function phoneForMember_(member) {
  return getMemberPhoneMap_()[String(member || '').trim()] || '';
}

function recordWorkflowOutbound_(member, phone, kind, text, meetingNo, result, idempotencyKey) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('WA Outbox');
  if (!sheet) return;

  if (idempotencyKey && hasReminderBeenSent_(idempotencyKey)) return;

  const now = new Date();
  sheet.appendRow([
    Utilities.getUuid(),
    member || '',
    phone || '',
    kind || '',
    '',
    text || '',
    result && result.ok ? 'SENT' : 'FAILED',
    1,
    '',
    result && result.messageId ? result.messageId : '',
    meetingNo || '',
    '',
    now,
    result && result.ok ? now : '',
    result && result.error ? result.error : '',
    idempotencyKey || ''
  ]);
}

function sendConsolidatedIntroductionsIfReady_(meetingNo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  const submissions = ss.getSheetByName('WA Submissions');
  if (!planner || !submissions) return false;

  const lastRow = Math.min(62, planner.getLastRow());
  const ids = planner.getRange(3, 1, Math.max(0, lastRow - 2), 1).getValues().flat();
  const idx = ids.findIndex(v => Number(v) === Number(meetingNo));
  if (idx < 0) return false;

  const rowNumber = idx + 3;
  const row = planner.getRange(rowNumber, 1, 1, 24).getValues()[0];
  const tmod = String(row[18] || '').trim();
  if (!tmod) return false;

  const assigned = ROLE_DEFS_
    .filter(def => def.code !== 'TMOD')
    .map(def => ({ def, member: String(row[def.column - 1] || '').trim() }))
    .filter(x => x.member);

  if (!assigned.length) return false;

  const submissionRows = submissions.getLastRow() >= 2
    ? submissions.getRange(2, 1, submissions.getLastRow() - 1, 14).getValues()
    : [];

  const intros = [];
  for (let i = 0; i < assigned.length; i++) {
    const item = assigned[i];
    const found = submissionRows.find(r =>
      Number(r[1]) === Number(meetingNo) &&
      String(r[2] || '').trim() === item.member &&
      String(r[3] || '').trim() === 'INTRO' &&
      String(r[4] || '').trim() === item.def.code
    );
    if (!found) return false;
    intros.push({
      role: item.def.name,
      member: item.member,
      text: String(found[5] || '').trim()
    });
  }

  const phone = phoneForMember_(tmod);
  if (!phone) return false;

  const idempotencyKey = 'INTRO_BUNDLE|' + meetingNo + '|' + tmod;
  if (hasReminderBeenSent_(idempotencyKey)) return true;

  const lines = intros.map(x => x.role + ' — ' + x.member + '\n' + x.text);
  const body =
    'ATK Meeting ' + meetingNo + ' · Role-player introductions\n\n' +
    lines.join('\n\n');

  const result = sendAutomatedText_(
    phone,
    'INTRO_BUNDLE',
    body,
    [tmod, meetingNo, lines.join('\n\n')]
  );
  recordWorkflowOutbound_(tmod, phone, 'INTRO_BUNDLE', body, meetingNo, result, idempotencyKey);
  logAutomation_('Introduction', meetingNo, tmod, 'Consolidated introductions sent to TMOD', result.ok ? 'Sent' : 'Error', result.error || '');

  return Boolean(result.ok);
}


function handleSpeakerDetailsSubmission_(member, phone, requestedMeetingNo, body, messageId) {
  if (!configBool_('SPEAKER_DETAILS_WORKFLOW_ENABLED')) {
    sendWhatsAppText_(phone, 'Speaker-detail collection is not active yet.');
    return;
  }

  const task = findPendingTaskForMember_('SPEAKER_DETAILS', member, requestedMeetingNo);
  if (!task) {
    sendWhatsAppText_(phone, 'I could not find a pending speech-detail request for you.');
    return;
  }

  const parts = String(body || '').split('|').map(v => v.trim());
  if (parts.length < 4 || parts.slice(0, 4).some(v => !v)) {
    sendWhatsAppText_(phone,
      'Please use this format:\nSPEECH ' + task.meetingNo +
      ': Speech Title | Path | Level | Project');
    return;
  }

  const assignment = verifySpeakerAssignment_(task.meetingNo, task.related, member);
  if (!assignment.ok) {
    sendWhatsAppText_(phone,
      'Your speaking assignment for Meeting ' + task.meetingNo +
      ' has changed, so I did not save these details.');
    syncPhase1Tasks_();
    return;
  }

  appendSubmission_({
    meetingNo: task.meetingNo,
    member,
    type: 'SPEAKER_DETAILS',
    related: task.related,
    values: parts.slice(0, 4),
    sourceMessageId: messageId
  });

  syncPhase1Tasks_();

  sendWhatsAppText_(phone,
    'Speech details received ✅\n' +
    'Title: ' + parts[0] + '\n' +
    'Path: ' + parts[1] + '\n' +
    'Level: ' + parts[2] + '\n' +
    'Project: ' + parts[3]);

  logAutomation_(
    'Speaker Details',
    task.meetingNo,
    member,
    'Speech details received',
    'Success',
    task.related
  );

  processEvaluatorNotificationsForMeeting_(task.meetingNo);
}

function verifySpeakerAssignment_(meetingNo, slotRef, member) {
  const slot = Number(String(slotRef || '').replace(/^S/i, ''));
  const columnMap = {1: 7, 2: 11, 3: 15};
  const column = columnMap[slot];
  if (!column) return { ok: false };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('VPE Weekly Planner');
  if (!sheet) return { ok: false };

  const lastRow = Math.min(62, sheet.getLastRow());
  const ids = sheet.getRange(3, 1, Math.max(0, lastRow - 2), 1).getValues().flat();
  const idx = ids.findIndex(v => Number(v) === Number(meetingNo));
  if (idx < 0) return { ok: false };

  const rowNumber = idx + 3;
  const assigned = String(sheet.getRange(rowNumber, column).getValue() || '').trim();
  return { ok: assigned === member, sheet, rowNumber, assigned, slot, column };
}

function processSpeakerDetailPrompts_() {
  if (!configBool_('SPEAKER_DETAILS_WORKFLOW_ENABLED')) return;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tasks = ss.getSheetByName('Automation Tasks');
  if (!tasks || tasks.getLastRow() < 2) return;

  const rows = tasks.getRange(2, 1, tasks.getLastRow() - 1, 18).getValues();
  const today = new Date();

  rows.forEach(row => {
    const taskType = String(row[2] || '');
    const state = String(row[5] || '');
    const dueAt = row[6];
    if (taskType !== 'SPEAKER_DETAILS' || state !== 'Pending') return;
    if (!(dueAt instanceof Date) || dueAt > today) return;

    const meetingNo = Number(row[1]);
    const member = String(row[3] || '').trim();
    const slotRef = String(row[4] || '').trim();
    const phone = phoneForMember_(member);
    if (!meetingNo || !member || !slotRef || !phone) return;
    if (!isWhatsAppRecipientAllowed_(phone)) return;

    const key = ['SPEAKER_DETAILS_PROMPT', meetingNo, slotRef, member].join('|');
    if (hasReminderBeenSent_(key)) return;

    const text =
      'Hi ' + member + ',\n\n' +
      'Please share your speech details for ATK Meeting ' + meetingNo + '.\n\n' +
      'Reply in this format:\n' +
      'SPEECH ' + meetingNo + ': Speech Title | Path | Level | Project\n\n' +
      'Thank you.';

    const result = sendAutomatedText_(
      phone,
      'SPEAKER_DETAILS',
      text,
      [member, meetingNo, meetingNo]
    );
    recordWorkflowOutbound_(member, phone, 'SPEAKER_DETAILS_PROMPT', text, meetingNo, result, key);
    logAutomation_(
      'Speaker Details',
      meetingNo,
      member,
      'Speech-detail prompt',
      result.ok ? 'Sent' : (result.blocked ? 'Blocked' : 'Error'),
      result.error || slotRef
    );
  });
}

function processEvaluatorNotifications_() {
  if (!configBool_('EVALUATION_WORKFLOW_ENABLED')) return;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  if (!planner) return;

  const today = startOfDay_(new Date());
  const reminderDays = String(configValue_('EVALUATOR_COORDINATION_REMINDER_DAYS', '5,2'))
    .split(',')
    .map(v => Number(String(v).trim()))
    .filter(v => Number.isFinite(v) && v >= 0);

  const rows = planner.getRange(
    3,
    1,
    Math.min(60, Math.max(0, planner.getLastRow() - 2)),
    28
  ).getValues();

  rows.forEach(row => {
    const meetingNo = Number(row[0]);
    const meetingDate = row[1];
    const status = String(row[2] || '');
    if (!meetingNo || !(meetingDate instanceof Date)) return;
    if (status === 'Completed' || status === 'Cancelled') return;

    const daysLeft = Math.round((startOfDay_(meetingDate) - today) / 86400000);
    processEvaluatorCoordinationForMeeting_(meetingNo, daysLeft, reminderDays);
  });
}

function processEvaluatorNotificationsForMeeting_(meetingNo) {
  // Immediate pairing notification after speaker details or evaluator assignment changes.
  return processEvaluatorCoordinationForMeeting_(meetingNo, null, []);
}

function processEvaluatorCoordinationForMeeting_(meetingNo, daysLeft, reminderDays) {
  if (!configBool_('EVALUATION_WORKFLOW_ENABLED')) return false;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  if (!planner) return false;

  const lastRow = Math.min(62, planner.getLastRow());
  const ids = planner.getRange(3, 1, Math.max(0, lastRow - 2), 1).getValues().flat();
  const idx = ids.findIndex(v => Number(v) === Number(meetingNo));
  if (idx < 0) return false;

  const row = planner.getRange(idx + 3, 1, 1, 28).getValues()[0];
  const speakers = [
    {slotRef: 'S1', member: String(row[6] || '').trim(), evaluator: String(row[25] || '').trim()},
    {slotRef: 'S2', member: String(row[10] || '').trim(), evaluator: String(row[26] || '').trim()},
    {slotRef: 'S3', member: String(row[14] || '').trim(), evaluator: String(row[27] || '').trim()}
  ];

  speakers.forEach(sp => {
    if (!sp.member || !sp.evaluator) return;

    const speakerPhone = phoneForMember_(sp.member);
    const evaluatorPhone = phoneForMember_(sp.evaluator);

    // Always try the initial pairing notification once.
    const initialSpeakerKey = ['EVALUATOR_PAIR_SPEAKER', meetingNo, sp.slotRef, sp.member, sp.evaluator].join('|');
    if (speakerPhone && isWhatsAppRecipientAllowed_(speakerPhone) && !hasReminderBeenSent_(initialSpeakerKey)) {
      const text =
        'Hi ' + sp.member + ',\n\n' +
        sp.evaluator + ' is your evaluator for ATK Meeting ' + meetingNo + '.\n\n' +
        'Please connect with your evaluator directly and share the following at the earliest:\n' +
        '• Speech title\n' +
        '• Path, Level & Project\n' +
        '• Project / evaluation sheet\n' +
        '• Your project objective and any specific area where you would like feedback\n\n' +
        'Direct communication between you and your evaluator is important before the meeting.\n\n' +
        'Thank you.';

      const result = sendAutomatedText_(
        speakerPhone,
        'EVALUATOR_PAIRING',
        text,
        [sp.member, sp.evaluator, meetingNo]
      );
      recordWorkflowOutbound_(sp.member, speakerPhone, 'EVALUATOR_PAIRING', text, meetingNo, result, initialSpeakerKey);
      logAutomation_(
        'Evaluation Coordination',
        meetingNo,
        sp.member,
        'Speaker informed of evaluator',
        result.ok ? 'Sent' : (result.blocked ? 'Blocked' : 'Error'),
        sp.evaluator
      );
    }

    const initialEvaluatorKey = ['EVALUATOR_PAIR_EVALUATOR', meetingNo, sp.slotRef, sp.evaluator, sp.member].join('|');
    if (evaluatorPhone && isWhatsAppRecipientAllowed_(evaluatorPhone) && !hasReminderBeenSent_(initialEvaluatorKey)) {
      const text =
        'Hi ' + sp.evaluator + ',\n\n' +
        'You are the evaluator for ' + sp.member + ' in ATK Meeting ' + meetingNo + '.\n\n' +
        'Please connect with the speaker directly before the meeting and ensure you receive:\n' +
        '• Speech title\n' +
        '• Path, Level & Project\n' +
        '• Project / evaluation sheet\n' +
        '• Project objective and any specific feedback request from the speaker\n\n' +
        'Please use the direct conversation to understand the project before the meeting.\n\n' +
        'Thank you.';

      const result = sendAutomatedText_(
        evaluatorPhone,
        'EVALUATOR_PAIRING',
        text,
        [sp.evaluator, sp.member, meetingNo]
      );
      recordWorkflowOutbound_(sp.evaluator, evaluatorPhone, 'EVALUATOR_PAIRING', text, meetingNo, result, initialEvaluatorKey);
      logAutomation_(
        'Evaluation Coordination',
        meetingNo,
        sp.evaluator,
        'Evaluator informed of speaker',
        result.ok ? 'Sent' : (result.blocked ? 'Blocked' : 'Error'),
        sp.member
      );
    }

    if (daysLeft === null || reminderDays.indexOf(daysLeft) === -1) return;

    const reminderSuffix = 'D' + daysLeft;

    const speakerReminderKey = ['EVALUATOR_COORD_REMINDER_SPEAKER', meetingNo, sp.slotRef, reminderSuffix, sp.member].join('|');
    if (speakerPhone && isWhatsAppRecipientAllowed_(speakerPhone) && !hasReminderBeenSent_(speakerReminderKey)) {
      const text =
        'Hi ' + sp.member + ',\n\n' +
        'A quick reminder for ATK Meeting ' + meetingNo + ': please connect directly with ' +
        sp.evaluator + ' and make sure you have shared your speech title, Path, Level, Project, project/evaluation sheet, and any specific feedback area.\n\nThank you.';
      const result = sendAutomatedText_(
        speakerPhone,
        'EVALUATOR_COORD_REMINDER',
        text,
        [sp.member, meetingNo, sp.evaluator]
      );
      recordWorkflowOutbound_(sp.member, speakerPhone, 'EVALUATOR_COORD_REMINDER', text, meetingNo, result, speakerReminderKey);
    }

    const evaluatorReminderKey = ['EVALUATOR_COORD_REMINDER_EVALUATOR', meetingNo, sp.slotRef, reminderSuffix, sp.evaluator].join('|');
    if (evaluatorPhone && isWhatsAppRecipientAllowed_(evaluatorPhone) && !hasReminderBeenSent_(evaluatorReminderKey)) {
      const text =
        'Hi ' + sp.evaluator + ',\n\n' +
        'A quick reminder for ATK Meeting ' + meetingNo + ': please connect directly with ' +
        sp.member + ' and confirm you have the speech title, Path, Level, Project, project/evaluation sheet, and the speaker\'s specific feedback expectations.\n\nThank you.';
      const result = sendAutomatedText_(
        evaluatorPhone,
        'EVALUATOR_COORD_REMINDER',
        text,
        [sp.evaluator, meetingNo, sp.member]
      );
      recordWorkflowOutbound_(sp.evaluator, evaluatorPhone, 'EVALUATOR_COORD_REMINDER', text, meetingNo, result, evaluatorReminderKey);
    }
  });

  return true;
}


function processAssignedRolePrompts_() {
  if (
    !configBool_('TMOD_WORKFLOW_ENABLED') &&
    !configBool_('GRAMMARIAN_WORKFLOW_ENABLED') &&
    !configBool_('INTRO_WORKFLOW_ENABLED')
  ) return;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tasks = ss.getSheetByName('Automation Tasks');
  if (!tasks || tasks.getLastRow() < 2) return;

  const rows = tasks.getRange(2, 1, tasks.getLastRow() - 1, 18).getValues();
  const now = new Date();

  rows.forEach(row => {
    const taskId = String(row[0] || '');
    const meetingNo = Number(row[1]);
    const type = String(row[2] || '');
    const member = String(row[3] || '').trim();
    const related = String(row[4] || '').trim();
    const state = String(row[5] || '');
    const dueAt = row[6];

    if (!taskId || !meetingNo || !member || state !== 'Pending') return;
    if (!(dueAt instanceof Date) || dueAt > now) return;

    let kind = '';
    let text = '';
    let params = [];

    if (type === 'THEME' && configBool_('TMOD_WORKFLOW_ENABLED')) {
      kind = 'THEME_REQUEST';
      text =
        'Hi ' + member + ',\n\n' +
        'Please share the theme at the earliest, so that the role players can align with the same.\n\n' +
        'Thank you.\n\n' +
        'Reply in this format:\nTHEME ' + meetingNo + ': Your theme';
      params = [member, meetingNo, meetingNo];
    }

    if (type === 'GRAMMAR' && configBool_('GRAMMARIAN_WORKFLOW_ENABLED')) {
      kind = 'GRAMMAR_REQUEST';
      text =
        'Hi ' + member + ',\n\n' +
        'Please share the Word of the Day and Idiom of the Day for ATK Meeting ' + meetingNo + ' at the earliest.\n\n' +
        'Reply in this format:\nGRAMMAR ' + meetingNo +
        ': Word | Meaning | Idiom | Meaning\n\nThank you.';
      params = [member, meetingNo, meetingNo];
    }

    if (type === 'INTRO' && configBool_('INTRO_WORKFLOW_ENABLED')) {
      const role = ROLE_DEFS_.find(r => r.code === related);
      kind = 'INTRO_REQUEST';
      text =
        'Hi ' + member + ',\n\n' +
        'Please share a short introduction that the Toastmaster can use while introducing you as ' +
        (role ? role.name : related) + ' for ATK Meeting ' + meetingNo + '.\n\n' +
        'Reply in this format:\nINTRO ' + meetingNo +
        ': Your short introduction\n\nThank you.';
      params = [member, meetingNo, role ? role.name : related, meetingNo];
    }

    if (!kind || !text) return;

    const phone = phoneForMember_(member);
    if (!phone || !isWhatsAppRecipientAllowed_(phone)) return;

    const key = [kind, meetingNo, related, member].join('|');
    if (hasReminderBeenSent_(key)) return;

    const result = sendAutomatedText_(phone, kind, text, params);
    recordWorkflowOutbound_(member, phone, kind, text, meetingNo, result, key);

    logAutomation_(
      'Assignment Follow-up',
      meetingNo,
      member,
      kind,
      result.ok ? 'Sent' : (result.deferred ? 'Deferred' : (result.blocked ? 'Blocked' : 'Error')),
      result.error || taskId
    );
  });
}
