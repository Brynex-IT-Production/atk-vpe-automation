/**
 * TMOD theme, Grammarian and role-player introduction workflows.
 * These flows use explicit reply prefixes so submissions are unambiguous.
 */

function parseSpeakerDetailsBlock_(text) {
  const fields = {};
  let recognized = 0;

  String(text || '').split(/\r?\n/).forEach(line => {
    const match = String(line || '').match(/^\s*(Path|Level|Project|Title)\s*:\s*(.*?)\s*$/i);
    if (!match) return;
    const key = match[1].toLowerCase();
    if (!Object.prototype.hasOwnProperty.call(fields, key)) recognized++;
    fields[key] = String(match[2] || '').trim();
  });

  return {
    recognized: recognized,
    path: fields.path || '',
    level: fields.level || '',
    project: fields.project || '',
    title: fields.title || ''
  };
}

function parseGrammarDetailsBlock_(text) {
  if (text && typeof text === 'object') {
    const word = String(text.word || '').trim();
    const wordMeaning = String(text.wordMeaning || '').trim();
    const idiom = String(text.idiom || '').trim();
    const idiomMeaning = String(text.idiomMeaning || '').trim();
    return {
      recognized: [word, wordMeaning, idiom, idiomMeaning].filter(Boolean).length,
      word: word,
      wordMeaning: wordMeaning,
      idiom: idiom,
      idiomMeaning: idiomMeaning
    };
  }

  const raw = String(text || '').trim();
  const fields = {};
  let recognized = 0;

  raw.split(/\r?\n/).forEach(line => {
    const match = String(line || '').match(
      /^\s*(Word(?: of the Day)?|Word meaning|Meaning of (?:the )?word|Idiom(?: of the Day)?|Idiom meaning|Meaning of (?:the )?idiom)\s*:\s*(.*?)\s*$/i
    );
    if (!match) return;

    const label = String(match[1] || '').toLowerCase();
    const value = String(match[2] || '').trim();
    let key = '';

    if (/^word(?: of the day)?$/.test(label)) key = 'word';
    else if (/^(word meaning|meaning of (?:the )?word)$/.test(label)) key = 'wordMeaning';
    else if (/^idiom(?: of the day)?$/.test(label)) key = 'idiom';
    else if (/^(idiom meaning|meaning of (?:the )?idiom)$/.test(label)) key = 'idiomMeaning';

    if (!key) return;
    if (!Object.prototype.hasOwnProperty.call(fields, key)) recognized++;
    fields[key] = value;
  });

  if (recognized < 4 && raw.indexOf('|') !== -1) {
    const parts = raw.split('|').map(v => String(v || '').trim());
    if (parts.length >= 4 && parts.slice(0, 4).every(Boolean)) {
      return {
        recognized: 4,
        word: parts[0],
        wordMeaning: parts[1],
        idiom: parts[2],
        idiomMeaning: parts[3]
      };
    }
  }

  return {
    recognized: recognized,
    word: fields.word || '',
    wordMeaning: fields.wordMeaning || '',
    idiom: fields.idiom || '',
    idiomMeaning: fields.idiomMeaning || ''
  };
}


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

  const grammarBlock = parseGrammarDetailsBlock_(text);
  if (grammarBlock.recognized >= 4) {
    handleGrammarSubmission_(member, phone, 0, grammarBlock, messageId);
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

  const speechBlock = parseSpeakerDetailsBlock_(text);
  if (speechBlock.recognized >= 2) {
    handleSpeakerDetailsSubmission_(member, phone, 0, speechBlock, messageId);
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
  if (roleDef.code === 'TMOD' && !configBool_('TMOD_WORKFLOW_ENABLED')) return;
  if (roleDef.code === 'GRM' && !configBool_('GRAMMARIAN_WORKFLOW_ENABLED')) return;
  if (roleDef.code !== 'TMOD' && roleDef.code !== 'GRM' &&
      !configBool_('INTRO_WORKFLOW_ENABLED')) return;
  const kind = roleDef.code === 'TMOD' ? 'THEME' :
    roleDef.code === 'GRM' ? 'GRAMMAR' : 'INTRO';
  const request = startGuidedRequest_(member, phone, kind, meetingNo, roleDef.code, roleDef.name);
  if (!request) return;
  const result = sendWhatsAppText_(phone, request.text);
  if (result.ok) setConversation_(phone, request.state);
  recordWorkflowOutbound_(member, phone, 'ASSIGNMENT_PROMPT',
    request.text, meetingNo, result,
    ['ASSIGNMENT_PROMPT', meetingNo, roleDef.code, member].join('|'));
}

function meetingTheme_(meetingNo) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('VPE Weekly Planner');
  if (!sheet) return '';
  const lastRow = Math.min(62, sheet.getLastRow());
  if (lastRow < 3) return '';
  const ids = sheet.getRange(3, 1, lastRow - 2, 1).getValues().flat();
  const idx = ids.findIndex(v => Number(v) === Number(meetingNo));
  if (idx < 0) return '';
  return String(sheet.getRange(idx + 3, 4).getValue() || '').trim();
}

function sendMeetingThemeToRole_(meetingNo, member, roleName, phone) {
  const theme = meetingTheme_(meetingNo);
  const destination = phone || phoneForMember_(member);
  if (!theme || !member || !destination || !isWhatsAppRecipientAllowed_(destination)) return false;

  const key = ['THEME_TO_ROLE', meetingNo, member, roleName || 'Role', theme].join('|');
  if (hasReminderBeenSent_(key)) return false;

  const body =
    'Hi ' + member + ',\n\n' +
    'The theme for ATK Meeting ' + meetingNo + ' is:\n\n' +
    theme + '\n\n' +
    'Please keep the theme in mind while preparing for your ' + (roleName || 'role') + ' role.';

  const result = sendAutomatedText_(
    destination,
    'TMOD_UPDATE',
    body,
    [member, meetingNo, 'Theme: ' + theme]
  );

  recordWorkflowOutbound_(
    member,
    destination,
    'THEME_TO_ROLE',
    body,
    meetingNo,
    result,
    key
  );

  logAutomation_(
    'TMOD Theme',
    meetingNo,
    member,
    'Theme shared with role player',
    result.ok ? 'Sent' : (result.blocked ? 'Blocked' : 'Error'),
    (roleName || 'Role') + ' | ' + theme
  );

  return Boolean(result && result.ok);
}

function sendMeetingThemeToAssignedRoles_(meetingNo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('VPE Weekly Planner');
  if (!sheet) return false;

  const lastRow = Math.min(62, sheet.getLastRow());
  if (lastRow < 3) return false;
  const ids = sheet.getRange(3, 1, lastRow - 2, 1).getValues().flat();
  const idx = ids.findIndex(v => Number(v) === Number(meetingNo));
  if (idx < 0) return false;

  const row = sheet.getRange(idx + 3, 1, 1, 28).getValues()[0];
  let sent = false;
  const seen = {};

  ROLE_DEFS_.forEach(def => {
    if (def.code === 'TMOD') return;
    const assigned = String(row[def.column - 1] || '').trim();
    if (!assigned || seen[assigned]) return;
    seen[assigned] = true;
    if (sendMeetingThemeToRole_(meetingNo, assigned, def.name, phoneForMember_(assigned))) {
      sent = true;
    }
  });

  return sent;
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
  sendMeetingThemeToAssignedRoles_(task.meetingNo);
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

  const details = parseGrammarDetailsBlock_(body);
  if (details.recognized < 4) {
    sendWhatsAppText_(phone,
      'Please send all four details in one message using this format:\n\n' +
      'Word:\nWord meaning:\nIdiom:\nIdiom meaning:');
    return;
  }

  const parts = [
    details.word,
    details.wordMeaning,
    details.idiom,
    details.idiomMeaning
  ];

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

function handleIntroSubmission_(member, phone, requestedMeetingNo, intro, messageId, requestedRelated) {
  if (!configBool_('INTRO_WORKFLOW_ENABLED')) {
    sendWhatsAppText_(phone, 'Introduction collection is not active yet.');
    return;
  }

  const task = findPendingTaskForMember_('INTRO', member, requestedMeetingNo, requestedRelated);
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

  const assignment = /^S[123]$/.test(task.related)
    ? verifySpeakerAssignment_(task.meetingNo, task.related, member)
    : verifyRoleAssignment_(task.meetingNo, task.related, member);
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
}

function findPendingTaskForMember_(type, member, meetingNo, related) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Automation Tasks');
  if (!sheet || sheet.getLastRow() < 2) return null;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 18).getValues();
  const candidates = rows
    .filter(r =>
      String(r[2] || '') === type &&
      String(r[3] || '').trim() === member &&
      ['Pending', 'Sent', 'Escalated'].indexOf(String(r[5] || '')) !== -1 &&
      (!meetingNo || Number(r[1]) === Number(meetingNo)) &&
      (!related || String(r[4] || '') === related)
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

  const ids = planner.getRange(
    3, 1, Math.max(0, Math.min(60, planner.getLastRow() - 2)), 1
  ).getValues().flat();
  const index = ids.findIndex(v => Number(v) === Number(meetingNo));
  if (index < 0) return false;

  const row = planner.getRange(index + 3, 1, 1, 28).getValues()[0];
  const tmod = String(row[18] || '').trim();
  const phone = phoneForMember_(tmod);
  if (!tmod || !phone || !isWhatsAppRecipientAllowed_(phone)) return false;

  const key = ['INTRO_BUNDLE', meetingNo, tmod].join('|');
  if (hasReminderBeenSent_(key)) return false;

  const assigned = ROLE_DEFS_.filter(def => def.code !== 'TMOD')
    .map(def => ({
      member: String(row[def.column - 1] || '').trim(),
      related: def.code,
      role: def.name
    }))
    .concat([1, 2, 3].map(slot => ({
      member: String(row[slot === 1 ? 6 : slot === 2 ? 10 : 14] || '').trim(),
      related: 'S' + slot,
      role: 'Speaker ' + slot
    })))
    .filter(item => item.member);

  const submissionRows = submissions.getLastRow() >= 2
    ? submissions.getRange(2, 1, submissions.getLastRow() - 1, 14).getValues()
    : [];

  const introductions = [];
  const missing = [];
  let meetingSpecificCount = 0;
  let defaultCount = 0;

  assigned.forEach(item => {
    const matches = submissionRows.filter(r =>
      Number(r[1]) === Number(meetingNo) &&
      String(r[2] || '').trim() === item.member &&
      String(r[3] || '') === 'INTRO' &&
      String(r[4] || '') === item.related
    );
    const latest = matches[matches.length - 1];
    const meetingSpecific = latest ? String(latest[5] || '').trim() : '';
    const savedDefault = defaultIntro_(item.member);
    const intro = meetingSpecific || savedDefault;

    if (!intro) {
      missing.push(item.role + ' — ' + item.member);
      return;
    }

    if (meetingSpecific) meetingSpecificCount++;
    else defaultCount++;

    introductions.push(
      item.role + ' — ' + item.member + '\n' + intro
    );
  });

  let bundle = introductions.length
    ? introductions.join('\n\n')
    : 'No introductions are currently available.';

  if (missing.length) {
    bundle += '\n\nNo introduction available for:\n' + missing.join('\n');
  }

  const body =
    'Hi ' + tmod + ', here are the role-player introductions received for ATK Meeting ' +
    meetingNo + ':\n\n' + bundle +
    '\n\nPlease use these while introducing the role players.';

  const result = sendAutomatedText_(
    phone,
    'INTRO_BUNDLE',
    body,
    [tmod, meetingNo, bundle]
  );

  recordWorkflowOutbound_(
    tmod,
    phone,
    'INTRO_BUNDLE',
    body,
    meetingNo,
    result,
    key
  );

  logAutomation_(
    'Introduction',
    meetingNo,
    tmod,
    'Consolidated introduction bundle',
    result.ok ? 'Sent' : (result.blocked ? 'Blocked' : 'Error'),
    'Meeting-specific: ' + meetingSpecificCount +
      ' | Saved default: ' + defaultCount +
      ' | Missing: ' + missing.length
  );

  return Boolean(result && result.ok);
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

  let parts = [];
  if (body && typeof body === 'object' && Number(body.recognized || 0) > 0) {
    const missing = ['path', 'level', 'project', 'title']
      .filter(key => !String(body[key] || '').trim());

    if (missing.length) {
      sendWhatsAppText_(phone,
        'Please add the missing detail' + (missing.length > 1 ? 's' : '') +
        ': ' + missing.map(key => key.charAt(0).toUpperCase() + key.slice(1)).join(', ') +
        '\n\nPath:\nLevel:\nProject:\nTitle:');
      return;
    }

    parts = [
      String(body.title || '').trim(),
      String(body.path || '').trim(),
      String(body.level || '').trim(),
      String(body.project || '').trim()
    ];
  } else {
    parts = String(body || '').split('|').map(v => v.trim());
    if (parts.length < 4 || parts.slice(0, 4).some(v => !v)) {
      sendWhatsAppText_(phone,
        'Please use this format:\n\nPath:\nLevel:\nProject:\nTitle:');
      return;
    }
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

  const active = activeConversation_(phone);
  if (active && active.kind === 'SPEAKER_DETAILS') clearConversation_(phone);

  sendWhatsAppText_(phone,
    'Speech details received ✅\n' +
    'Path: ' + parts[1] + '\n' +
    'Level: ' + parts[2] + '\n' +
    'Project: ' + parts[3] + '\n' +
    'Title: ' + parts[0]);

  logAutomation_(
    'Speaker Details',
    task.meetingNo,
    member,
    'Speech details received',
    'Success',
    task.related
  );

  sendSpeakerDetailsToTmod_(task.meetingNo, member, task.related, parts);
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

  const activeTemplate = templateNameForKind_('SPEAKER_DETAILS');
  const blockFlow = activeTemplate === 'atk_speaker_guided_v2';
  const rows = tasks.getRange(2, 1, tasks.getLastRow() - 1, 18).getValues();

  rows.forEach(row => {
    const meetingNo = Number(row[1]);
    const member = String(row[3] || '').trim();
    const related = String(row[4] || '').trim();

    if (String(row[2]) !== 'SPEAKER_DETAILS' || String(row[5]) !== 'Pending' ||
        !(row[6] instanceof Date) || row[6] > new Date()) return;

    const phone = phoneForMember_(member);
    if (!phone || !isWhatsAppRecipientAllowed_(phone)) return;

    const key = ['SPEAKER_DETAILS_PROMPT', meetingNo, related, member].join('|');

    if (blockFlow) {
      let introSent = hasReminderBeenSent_(key);

      if (!introSent) {
        const body =
          'Hi ' + member + ',\n\n' +
          'You’re confirmed as a speaker for ATK Meeting ' + meetingNo + ' 🎤\n\n' +
          'Let’s get your speech details ready for the agenda.\n\n' +
          'Please fill the below details.\n\n' +
          'Thank you.';

        const result = sendAutomatedText_(
          phone,
          'SPEAKER_DETAILS',
          body,
          [member, meetingNo]
        );

        recordWorkflowOutbound_(
          member, phone, 'SPEAKER_DETAILS_PROMPT',
          body, meetingNo, result, key
        );

        if (!result.ok) return;
        introSent = true;
      }

      if (!introSent) return;

      const formatKey = ['SPEAKER_DETAILS_FORMAT', meetingNo, related, member].join('|');
      if (hasReminderBeenSent_(formatKey)) return;

      const formatBody = 'Path:\nLevel:\nProject:\nTitle:';
      const formatResult = sendAutomatedText_(
        phone,
        'SPEAKER_DETAILS_FORMAT',
        formatBody,
        []
      );

      recordWorkflowOutbound_(
        member, phone, 'SPEAKER_DETAILS_FORMAT',
        formatBody, meetingNo, formatResult, formatKey
      );
      return;
    }

    if (activeConversation_(phone) || hasReminderBeenSent_(key)) return;

    const prompt = startGuidedRequest_(
      member, phone, 'SPEAKER_DETAILS', meetingNo, related, 'speaker'
    );
    const result = sendAutomatedText_(
      phone,
      'SPEAKER_DETAILS',
      prompt.text,
      [member, meetingNo, meetingNo]
    );

    recordWorkflowOutbound_(
      member, phone, 'SPEAKER_DETAILS_PROMPT',
      prompt.text, meetingNo, result, key
    );

    if (result.ok) setConversation_(phone, prompt.state);
  });
}

function processEvaluatorNotifications_() {
  if (!configBool_('EVALUATION_WORKFLOW_ENABLED')) return;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  if (!planner) return;

  const today = startOfDay_(new Date());
  const reminderDays = String(configValue_('EVALUATOR_COORDINATION_REMINDER_DAYS', '2'))
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
        'you will be evaluated by ' + sp.evaluator +
        ' for ATK Meeting ' + meetingNo + '. Please connect directly and coordinate the project and evaluation requirements before the meeting.\n\n' +
        'Please ensure you send:\n' +
        'Your full Project and evaluation form.\n\n' +
        'Thank you.';

      const result = sendAutomatedText_(
        speakerPhone,
        'EVALUATOR_PAIRING_SPEAKER',
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
        'You will be evaluating ' + sp.member + ' for ATK Meeting ' + meetingNo + '.\n\n' +
        'Please connect directly with the speaker before the meeting and ensure you receive their complete project and evaluation form.\n\n' +
        'Thank you.';

      const result = sendAutomatedText_(
        evaluatorPhone,
        'EVALUATOR_ASSIGNMENT',
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

    const speakerReminderKey = ['EVALUATOR_COORD_CHECK_SPEAKER', meetingNo, sp.slotRef, reminderSuffix, sp.member].join('|');
    if (speakerPhone && isWhatsAppRecipientAllowed_(speakerPhone) && !hasReminderBeenSent_(speakerReminderKey)) {
      const text =
        'Hi ' + sp.member + ',\n\n' +
        'Hope you have coordinated with your evaluator, ' + sp.evaluator +
        ', for ATK Meeting ' + meetingNo + '.\n\n' +
        'Have you coordinated?';

      const payload = {
        messaging_product: 'whatsapp',
        recipient_type: 'individual',
        to: speakerPhone,
        type: 'interactive',
        interactive: {
          type: 'button',
          body: {text: text},
          action: {
            buttons: [
              {type: 'reply', reply: {id: 'EVAL_COORD|SPEAKER|' + meetingNo + '|YES', title: 'Yes'}},
              {type: 'reply', reply: {id: 'EVAL_COORD|SPEAKER|' + meetingNo + '|NO', title: 'No'}}
            ]
          }
        }
      };

      const result = sendAutomatedQuickReply_(
        speakerPhone,
        'EVALUATOR_COORD_CHECK_SPEAKER',
        payload,
        [sp.member, sp.evaluator, meetingNo],
        [
          'EVAL_COORD|SPEAKER|' + meetingNo + '|YES',
          'EVAL_COORD|SPEAKER|' + meetingNo + '|NO'
        ]
      );

      recordWorkflowOutbound_(
        sp.member, speakerPhone, 'EVALUATOR_COORD_CHECK_SPEAKER',
        text, meetingNo, result, speakerReminderKey
      );
    }

    const evaluatorReminderKey = ['EVALUATOR_COORD_CHECK_EVALUATOR', meetingNo, sp.slotRef, reminderSuffix, sp.evaluator].join('|');
    if (evaluatorPhone && isWhatsAppRecipientAllowed_(evaluatorPhone) && !hasReminderBeenSent_(evaluatorReminderKey)) {
      const text =
        'Hi ' + sp.evaluator + ',\n\n' +
        'Hope you have coordinated with your speaker, ' + sp.member +
        ', for ATK Meeting ' + meetingNo + '.\n\n' +
        'Have you coordinated?';

      const payload = {
        messaging_product: 'whatsapp',
        recipient_type: 'individual',
        to: evaluatorPhone,
        type: 'interactive',
        interactive: {
          type: 'button',
          body: {text: text},
          action: {
            buttons: [
              {type: 'reply', reply: {id: 'EVAL_COORD|EVALUATOR|' + meetingNo + '|YES', title: 'Yes'}},
              {type: 'reply', reply: {id: 'EVAL_COORD|EVALUATOR|' + meetingNo + '|NO', title: 'No'}}
            ]
          }
        }
      };

      const result = sendAutomatedQuickReply_(
        evaluatorPhone,
        'EVALUATOR_COORD_CHECK_EVALUATOR',
        payload,
        [sp.evaluator, sp.member, meetingNo],
        [
          'EVAL_COORD|EVALUATOR|' + meetingNo + '|YES',
          'EVAL_COORD|EVALUATOR|' + meetingNo + '|NO'
        ]
      );

      recordWorkflowOutbound_(
        sp.evaluator, evaluatorPhone, 'EVALUATOR_COORD_CHECK_EVALUATOR',
        text, meetingNo, result, evaluatorReminderKey
      );
    }
  });

  return true;
}

function handleEvaluatorCoordinationCheckReply_(member, phone, interaction) {
  const meetingNo = Number(interaction.meetingNo || 0);
  const role = String(interaction.role || '').toUpperCase();
  const answer = String(interaction.answer || '').toUpperCase();

  if (!meetingNo || ['SPEAKER', 'EVALUATOR'].indexOf(role) === -1 ||
      ['YES', 'NO'].indexOf(answer) === -1) {
    sendWhatsAppText_(phone, 'I could not understand that response.');
    return;
  }

  let counterpart = '';
  const planner = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('VPE Weekly Planner');
  if (planner) {
    const ids = planner.getRange(3, 1, Math.max(0, Math.min(60, planner.getLastRow() - 2)), 1)
      .getValues().flat();
    const idx = ids.findIndex(v => Number(v) === meetingNo);
    if (idx >= 0) {
      const row = planner.getRange(idx + 3, 1, 1, 28).getValues()[0];
      const pairs = [
        {speaker: String(row[6] || '').trim(), evaluator: String(row[25] || '').trim()},
        {speaker: String(row[10] || '').trim(), evaluator: String(row[26] || '').trim()},
        {speaker: String(row[14] || '').trim(), evaluator: String(row[27] || '').trim()}
      ];
      const pair = pairs.find(p => role === 'SPEAKER' ? p.speaker === member : p.evaluator === member);
      if (pair) counterpart = role === 'SPEAKER' ? pair.evaluator : pair.speaker;
    }
  }

  if (answer === 'YES') {
    sendWhatsAppText_(phone, 'Great, thank you ✅');
  } else {
    sendWhatsAppText_(
      phone,
      counterpart
        ? 'Please connect with ' + counterpart + ' today and complete the coordination before the meeting. Thank you.'
        : 'Please complete the speaker-evaluator coordination today. Thank you.'
    );
  }

  logAutomation_(
    'Evaluation Coordination',
    meetingNo,
    member,
    role + ' coordination check',
    answer === 'YES' ? 'Confirmed' : 'Not yet',
    counterpart
  );
}


function handlePlannerEvaluatorEdit_(e) {
  try {
    if (!e || !e.range) return;

    const range = e.range;
    const sheet = range.getSheet();
    if (!sheet || sheet.getName() !== 'VPE Weekly Planner') return;

    const firstRow = range.getRow();
    const lastRow = range.getLastRow();
    const firstCol = range.getColumn();
    const lastCol = range.getLastColumn();

    if (lastRow < 3 || lastCol < 26 || firstCol > 28) return;

    const startRow = Math.max(3, firstRow);
    for (let rowNo = startRow; rowNo <= lastRow; rowNo++) {
      const evaluatorValues = sheet.getRange(rowNo, 26, 1, 3).getValues()[0];
      if (!evaluatorValues.some(v => String(v || '').trim())) continue;

      const meetingNo = Number(sheet.getRange(rowNo, 1).getValue());
      if (!meetingNo) continue;

      processEvaluatorNotificationsForMeeting_(meetingNo);
    }
  } catch (error) {
    logAutomation_(
      'Evaluation Coordination',
      '',
      '',
      'Evaluator edit trigger failed',
      'Error',
      String(error)
    );
    console.error(error);
  }
}


function processAssignedRolePrompts_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tasks = ss.getSheetByName('Automation Tasks');
  if (!tasks || tasks.getLastRow() < 2) return;
  const rows = tasks.getRange(2, 1, tasks.getLastRow() - 1, 18).getValues();
  rows.forEach(row => {
    const type = String(row[2] || '');
    if (['THEME', 'INTRO'].indexOf(type) < 0 ||
        String(row[5] || '') !== 'Pending' ||
        !(row[6] instanceof Date) || row[6] > new Date()) return;
    const meetingNo = Number(row[1]);
    const member = String(row[3] || '').trim();
    const related = String(row[4] || '').trim();
    const phone = phoneForMember_(member);
    if (!phone || !isWhatsAppRecipientAllowed_(phone) || activeConversation_(phone)) return;
    const kind = type === 'THEME' ? 'THEME_REQUEST' :
      type === 'GRAMMAR' ? 'GRAMMAR_REQUEST' : 'INTRO_REQUEST';
    const key = [kind, meetingNo, related, member].join('|');
    if (hasReminderBeenSent_(key)) return;
    const role = ROLE_DEFS_.find(def => def.code === related);
    const prompt = startGuidedRequest_(member, phone, type, meetingNo,
      related, role ? role.name : 'speaker');
    const params = type === 'INTRO'
      ? [member, meetingNo, role ? role.name : 'speaker', meetingNo]
      : [member, meetingNo, meetingNo];
    const result = sendAutomatedText_(phone, kind, prompt.text, params);
    recordWorkflowOutbound_(member, phone, kind, prompt.text, meetingNo, result, key);
    if (result.ok) setConversation_(phone, prompt.state);
  });
}

function processGrammarAssignmentReminder_() {
  if (!configBool_('GRAMMARIAN_WORKFLOW_ENABLED')) return;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tasks = ss.getSheetByName('Automation Tasks');
  const outbox = ss.getSheetByName('WA Outbox');
  if (!tasks || !outbox || tasks.getLastRow() < 2 || outbox.getLastRow() < 2) return;

  const now = new Date();
  const timezone = configValue_('TIMEZONE', 'Asia/Kolkata');
  const localHour = Number(Utilities.formatDate(now, timezone, 'H'));

  const reminderHours = String(
    configValue_('GRAMMAR_REMINDER_HOURS', '7,20')
  ).split(',')
    .map(v => Number(String(v).trim()))
    .filter(v => Number.isFinite(v));

  if (reminderHours.indexOf(localHour) === -1) return;

  const successfulStates = new Set(['SENT', 'DELIVERED', 'READ', 'ACCEPTED']);
  const taskRows = tasks.getRange(2, 1, tasks.getLastRow() - 1, 18).getValues();
  const outboxRows = outbox.getRange(2, 1, outbox.getLastRow() - 1, 16).getValues();

  taskRows.forEach(row => {
    if (String(row[2] || '') !== 'GRAMMAR' ||
        String(row[5] || '') !== 'Pending') return;

    const meetingNo = Number(row[1]);
    const member = String(row[3] || '').trim();
    const related = String(row[4] || '').trim();
    if (!meetingNo || !member) return;

    const assignmentKey = ['ASSIGNMENT_PROMPT', meetingNo, related, member].join('|');
    const firstSend = outboxRows
      .filter(r =>
        String(r[15] || '') === assignmentKey &&
        successfulStates.has(String(r[6] || '').toUpperCase()) &&
        r[12] instanceof Date
      )
      .map(r => r[12])
      .sort((a, b) => a.getTime() - b.getTime())[0];

    if (!(firstSend instanceof Date) || now <= firstSend) return;

    const phone = phoneForMember_(member);
    if (!phone || !isWhatsAppRecipientAllowed_(phone)) return;

    // No reminder template is used here. Reminders are allowed only while
    // WhatsApp's customer-service window remains open from the member's inbound
    // role-selection interaction (or any later inbound reply).
    if (!hasOpenWhatsAppServiceWindow_(phone)) return;

    const slotDate = Utilities.formatDate(now, timezone, 'yyyyMMdd');
    const reminderKey = [
      'GRAMMAR_ASSIGNMENT_REMINDER',
      meetingNo,
      related,
      member,
      slotDate,
      'H' + localHour
    ].join('|');

    if (hasReminderBeenSent_(reminderKey)) return;

    const body =
      'Hi ' + member + ',\n\n' +
      'Just a reminder about the Grammarian details for ATK Meeting ' + meetingNo + '.\n\n' +
      'We haven’t received them yet.\n\n' +
      'Please reply in one message:\n\n' +
      'Word:\n' +
      'Word meaning:\n' +
      'Idiom:\n' +
      'Idiom meaning:\n\n' +
      'Thank you.';

    const result = sendWhatsAppText_(phone, body);

    recordWorkflowOutbound_(
      member,
      phone,
      'GRAMMAR_ASSIGNMENT_REMINDER',
      body,
      meetingNo,
      result,
      reminderKey
    );

    logAutomation_(
      'Grammarian',
      meetingNo,
      member,
      'Grammarian reminder at ' + (localHour === 20 ? '8 PM' : localHour + ' AM'),
      result.ok ? 'Sent' : (result.blocked ? 'Blocked' : 'Error'),
      'Pending; service window open'
    );
  });
}

function processThemeRequestReminder_() {
  if (!configBool_('TMOD_WORKFLOW_ENABLED')) return;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tasks = ss.getSheetByName('Automation Tasks');
  if (!tasks || tasks.getLastRow() < 2) return;

  const now = new Date();
  const timezone = configValue_('TIMEZONE', 'Asia/Kolkata');
  const localHour = Number(Utilities.formatDate(now, timezone, 'H'));

  const reminderHours = String(
    configValue_('THEME_REMINDER_HOURS', '7,20')
  ).split(',')
    .map(v => Number(String(v).trim()))
    .filter(v => Number.isFinite(v));

  if (reminderHours.indexOf(localHour) === -1) return;

  const reminderDaysBefore = String(
    configValue_('THEME_REMINDER_DAYS_BEFORE', '5,4')
  ).split(',')
    .map(v => Number(String(v).trim()))
    .filter(v => Number.isFinite(v));

  const firstRequestDaysBefore = Math.max(
    1,
    Number(configValue_('THEME_DUE_DAYS', '6')) || 6
  );

  const today = startOfDay_(now);
  const rows = tasks.getRange(2, 1, tasks.getLastRow() - 1, 18).getValues();

  rows.forEach(row => {
    if (String(row[2] || '') !== 'THEME' ||
        String(row[5] || '') !== 'Pending' ||
        !(row[6] instanceof Date)) return;

    const meetingNo = Number(row[1]);
    const member = String(row[3] || '').trim();
    const related = String(row[4] || '').trim();
    const firstRequestDay = startOfDay_(row[6]);

    if (!meetingNo || !member) return;

    const daysAfterFirstRequest = Math.round(
      (today.getTime() - firstRequestDay.getTime()) / 86400000
    );
    const daysBeforeMeeting = firstRequestDaysBefore - daysAfterFirstRequest;

    if (reminderDaysBefore.indexOf(daysBeforeMeeting) === -1) return;

    const firstKey = ['THEME_REQUEST', meetingNo, related, member].join('|');
    if (!hasReminderBeenSent_(firstKey)) return;

    const reminderKey = [
      'THEME_REQUEST_REMINDER',
      meetingNo,
      related,
      member,
      'D' + daysBeforeMeeting,
      'H' + localHour
    ].join('|');

    if (hasReminderBeenSent_(reminderKey)) return;

    const phone = phoneForMember_(member);
    if (!phone || !isWhatsAppRecipientAllowed_(phone)) return;

    const body =
      'Hi ' + member + ',\n\n' +
      'Just a reminder about the meeting theme for ATK Meeting ' + meetingNo + ' 🎤\n\n' +
      'We haven’t received the theme yet.\n\n' +
      'Please share it today so that all the role players can align with it.\n\n' +
      'Just reply with your theme.\n\n' +
      'Thank you.';

    const result = sendAutomatedText_(
      phone,
      'THEME_REMINDER',
      body,
      [member, meetingNo]
    );

    recordWorkflowOutbound_(
      member,
      phone,
      'THEME_REQUEST_REMINDER',
      body,
      meetingNo,
      result,
      reminderKey
    );

    logAutomation_(
      'TMOD Theme',
      meetingNo,
      member,
      'Theme reminder ' + daysBeforeMeeting + ' days before meeting at ' +
        (localHour === 20 ? '8 PM' : localHour + ' AM'),
      result.ok ? 'Sent' : (result.blocked ? 'Blocked' : 'Error'),
      'Theme still pending'
    );
  });
}

function sendSpeakerDetailsToTmod_(meetingNo, member, slotRef, parts) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  const ids = planner.getRange(3, 1, Math.max(0, Math.min(60, planner.getLastRow() - 2)), 1)
    .getValues().flat();
  const idx = ids.findIndex(v => Number(v) === Number(meetingNo));
  if (idx < 0) return;
  const tmod = String(planner.getRange(idx + 3, 19).getValue() || '').trim();
  const phone = phoneForMember_(tmod);
  if (!phone || !isWhatsAppRecipientAllowed_(phone)) return;
  const body = 'Speech details for ' + member + ' · ATK Meeting ' + meetingNo +
    '\nTitle: ' + parts[0] + '\nPath: ' + parts[1] +
    '\nLevel: ' + parts[2] + '\nProject: ' + parts[3];
  const key = ['SPEECH_TO_TMOD', meetingNo, slotRef, member, parts.join('|')].join('|');
  if (hasReminderBeenSent_(key)) return;
  const result = sendAutomatedText_(phone, 'TMOD_UPDATE', body,
    [tmod, meetingNo, member + ' — ' + parts[0] + ' | ' + parts[1] +
      ' | ' + parts[2] + ' | ' + parts[3]]);
  recordWorkflowOutbound_(tmod, phone, 'SPEECH_TO_TMOD', body, meetingNo, result, key);
}

function processReadySpeakerDetailsToTmod_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const subs = ss.getSheetByName('WA Submissions');
  if (!subs || subs.getLastRow() < 2) return;
  subs.getRange(2, 1, subs.getLastRow() - 1, 14).getValues().forEach(row => {
    if (String(row[3]) !== 'SPEAKER_DETAILS') return;
    sendSpeakerDetailsToTmod_(Number(row[1]), String(row[2] || ''),
      String(row[4] || ''), [row[5], row[6], row[7], row[8]].map(String));
  });
}
