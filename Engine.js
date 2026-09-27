/**
 * ATK VPE Phase 1 automation engine.
 * All logic is DEV-first and controlled by Automation Config switches.
 */

const ROLE_DEFS_ = [
  { code: 'TMOD', name: 'Toastmaster', column: 19 },
  { code: 'GE', name: 'General Evaluator', column: 20 },
  { code: 'TTM', name: 'Table Topics Master', column: 21 },
  { code: 'Timer', name: 'Timer', column: 22 },
  { code: 'Ah', name: 'Ah-Counter', column: 23 },
  { code: 'GRM', name: 'Grammarian', column: 24 }
];

function processWhatsAppAutomation_(payload) {
  const entries = payload.entry || [];

  entries.forEach(entry => {
    (entry.changes || []).forEach(change => {
      const value = change.value || {};

      (value.messages || []).forEach(message => {
        try {
          processInboundWhatsAppMessage_(message);
        } catch (error) {
          logAutomation_('WhatsApp', '', message.from || '', 'Inbound processing failed', 'Error', String(error));
          console.error(error);
        }
      });

      (value.statuses || []).forEach(status => {
        try {
          updateOutboxDeliveryStatus_(status);
        } catch (error) {
          console.error(error);
        }
      });
    });
  });
}

function processInboundWhatsAppMessage_(message) {
  const messageId = String(message.id || '');
  if (!messageId) return;

  // logWhatsAppWebhook_ runs first. Count > 1 means Meta retried this message.
  if (isDuplicateWebhookMessage_(messageId)) {
    logAutomation_('WhatsApp', '', '', 'Duplicate webhook ignored', 'Skipped', messageId);
    return;
  }

  const phone = normalizePhone_(message.from || '');
  const member = findMemberByWhatsApp_(phone);

  if (member === 'Unknown') {
    sendWhatsAppText_(phone,
      'Hi. I could not match this WhatsApp number with the ATK active-member list. Please contact the VPE so your number can be updated.');
    return;
  }

  if (configBool_('TASK_ENGINE_ENABLED')) {
    syncPhase1Tasks_();
  }

  const interaction = parseInteraction_(message);

  if (interaction.kind === 'ROLE_SELECTION') {
    handleRoleSelection_(member, phone, interaction);
    return;
  }

  if (interaction.kind === 'SPEECH_SLOT_ACCEPT') {
    handleSpeechSlotSelection_(member, phone, interaction);
    return;
  }

  if (interaction.kind === 'ROLE_MENU') {
    handleRoleMenuRequest_(member, phone);
    return;
  }

  if (interaction.kind === 'ADMIN_EXCEPTIONS') {
    const vpeMember = configValue_('VPE_MEMBER_NAME', 'Riyas AF');
    if (member === vpeMember) {
      const digest = buildVpeExceptionDigest_();
      sendWhatsAppText_(
        phone,
        digest || 'ATK VPE — No action needed in the configured exception window.'
      );
    }
    return;
  }

  // VPE-only administration through WhatsApp.
  const rawText = String(interaction.text || '').trim();
  const vpeMember = configValue_('VPE_MEMBER_NAME', 'Riyas AF');

  if (handleStructuredSubmission_(member, phone, rawText, messageId)) {
    return;
  }

  if (member === vpeMember && rawText.toUpperCase() === 'ADMIN SETUP') {
    if (configBool_('TASK_ENGINE_ENABLED')) {
      syncPhase1Tasks_();
    }
    sendWhatsAppText_(phone,
      'ATK Phase 1 setup is active ✅\nThe hourly scheduler is managed automatically through GitHub Actions, so no Apps Script trigger authorization is required.');
    return;
  }

  if (member === vpeMember && rawText.toUpperCase() === 'ADMIN STATUS') {
    sendWhatsAppText_(phone, buildPhase1StatusText_());
    return;
  }

  if (member === vpeMember && rawText.toUpperCase() === 'ADMIN EXCEPTIONS') {
    const digest = buildVpeExceptionDigest_();
    sendWhatsAppText_(
      phone,
      digest || 'ATK VPE — No action needed in the configured exception window.'
    );
    return;
  }

  // Natural-language shortcuts for the permanent group link or direct chat.
  const text = rawText.toLowerCase();
  if (['role', 'roles', 'choose role', 'choose my role', 'role selection'].includes(text)) {
    handleRoleMenuRequest_(member, phone);
    return;
  }
}

function parseInteraction_(message) {
  if (message.type === 'interactive') {
    const listReply = message.interactive && message.interactive.list_reply;
    const buttonReply = message.interactive && message.interactive.button_reply;
    const reply = listReply || buttonReply;

    if (reply) {
      const id = String(reply.id || '');
      if (id.indexOf('ROLE|') === 0) {
        const parts = id.split('|');
        return {
          kind: 'ROLE_SELECTION',
          meetingNo: Number(parts[1]),
          roleCode: parts[2],
          text: String(reply.title || '')
        };
      }
      if (id.indexOf('SPEAK|') === 0) {
        const parts = id.split('|');
        return {
          kind: 'SPEECH_SLOT_ACCEPT',
          meetingNo: Number(parts[1]),
          text: String(reply.title || '')
        };
      }
      if (id === 'ADMIN_EXCEPTIONS') {
        return { kind: 'ADMIN_EXCEPTIONS', text: String(reply.title || '') };
      }
      if (id === 'ROLE_MENU') {
        return { kind: 'ROLE_MENU', text: String(reply.title || '') };
      }
    }
  }

  if (message.type === 'button' && message.button) {
    const payload = String(message.button.payload || '');
    if (payload === 'ADMIN_EXCEPTIONS') {
      return {
        kind: 'ADMIN_EXCEPTIONS',
        text: String(message.button.text || '')
      };
    }
    if (payload.indexOf('SPEAK|') === 0) {
      const parts = payload.split('|');
      return {
        kind: 'SPEECH_SLOT_ACCEPT',
        meetingNo: Number(parts[1]),
        text: String(message.button.text || '')
      };
    }
    if (payload === 'ROLE_MENU') {
      return { kind: 'ROLE_MENU', text: String(message.button.text || '') };
    }
  }

  return {
    kind: 'TEXT',
    text: message.type === 'text' && message.text ? String(message.text.body || '') : ''
  };
}

function handleRoleMenuRequest_(member, phone) {
  if (!configBool_('ROLE_SELECTION_ENABLED')) {
    sendWhatsAppText_(phone, 'Role selection is being prepared. It will be available here shortly.');
    return;
  }

  const meeting = findNextRoleSelectionMeeting_(member);
  if (!meeting) {
    sendWhatsAppText_(phone, 'There are no role vacancies currently available for you in the upcoming ATK meetings.');
    return;
  }

  if (meeting.reason === 'SPEAKER') {
    sendWhatsAppText_(phone,
      'You are already listed as a speaker for Meeting ' + meeting.meetingNo +
      ', so I have kept you out of the role-selection list for that meeting.');
    return;
  }

  if (meeting.reason === 'ALREADY_ASSIGNED') {
    sendWhatsAppText_(phone,
      'You are already assigned as ' + meeting.assignedRole +
      ' for Meeting ' + meeting.meetingNo + '.');
    return;
  }

  sendRoleList_(phone, member, meeting);
}

function findNextRoleSelectionMeeting_(member) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('VPE Weekly Planner');
  if (!sheet) throw new Error('VPE Weekly Planner not found.');

  const windowDays = Math.max(1, Number(configValue_('ROLE_SELECTION_WINDOW_DAYS', '14')) || 14);
  const today = startOfDay_(new Date());
  const latest = new Date(today.getTime() + windowDays * 86400000);
  const values = sheet.getRange(3, 1, Math.min(60, Math.max(0, sheet.getLastRow() - 2)), 28).getValues();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const meetingNo = Number(row[0]);
    const meetingDate = row[1];
    const status = String(row[2] || '');

    if (!meetingNo || !(meetingDate instanceof Date)) continue;
    const d = startOfDay_(meetingDate);
    if (d < today || d > latest) continue;
    if (status === 'Completed' || status === 'Cancelled') continue;

    const speakers = [row[6], row[10], row[14]].map(v => String(v || '').trim()).filter(Boolean);
    if (speakers.indexOf(member) !== -1) {
      return { meetingNo, meetingDate, reason: 'SPEAKER' };
    }

    for (let r = 0; r < ROLE_DEFS_.length; r++) {
      const assigned = String(row[ROLE_DEFS_[r].column - 1] || '').trim();
      if (assigned === member) {
        return {
          meetingNo,
          meetingDate,
          reason: 'ALREADY_ASSIGNED',
          assignedRole: ROLE_DEFS_[r].name
        };
      }
    }

    const available = ROLE_DEFS_.filter(def => !String(row[def.column - 1] || '').trim());
    if (available.length) {
      return { meetingNo, meetingDate, available, rowNumber: i + 3, reason: 'AVAILABLE' };
    }
  }

  return null;
}

function sendRoleList_(phone, member, meeting) {
  const rows = meeting.available.map(def => ({
    id: 'ROLE|' + meeting.meetingNo + '|' + def.code,
    title: def.name,
    description: 'Meeting ' + meeting.meetingNo
  }));

  const dateText = Utilities.formatDate(meeting.meetingDate, configValue_('TIMEZONE', 'Asia/Kolkata'), 'EEE, d MMM');
  const body =
    'ATK Meeting ' + meeting.meetingNo + ' · ' + dateText +
    '\nChoose one of the roles still available below.';

  const payload = {
    messaging_product: 'whatsapp',
    recipient_type: 'individual',
    to: phone,
    type: 'interactive',
    interactive: {
      type: 'list',
      body: { text: body },
      action: {
        button: 'Choose my role',
        sections: [{
          title: 'Available roles',
          rows: rows
        }]
      }
    }
  };

  const result = sendWhatsAppPayload_(payload);
  logAutomation_('Role Selection', meeting.meetingNo, member, 'Role menu sent', result.ok ? 'Sent' : 'Error', result.messageId || result.error || '');
}

function handleRoleSelection_(member, phone, interaction) {
  if (!configBool_('ROLE_SELECTION_ENABLED')) {
    sendWhatsAppText_(phone, 'Role selection is currently unavailable.');
    return;
  }

  const roleDef = ROLE_DEFS_.find(r => r.code === interaction.roleCode);
  if (!roleDef || !interaction.meetingNo) {
    sendWhatsAppText_(phone, 'That role selection could not be understood. Please type ROLE to open the latest list.');
    return;
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    sendWhatsAppText_(phone, 'Someone else is choosing a role at the same moment. Please type ROLE again in a few seconds.');
    return;
  }

  try {
    const result = reserveRole_(interaction.meetingNo, roleDef, member);

    if (result.status === 'BOOKED') {
      sendWhatsAppText_(phone,
        'Confirmed ✅\n' + roleDef.name + ' · ATK Meeting ' + interaction.meetingNo +
        ' · ' + result.dateText + '\n\nThank you, ' + member + '.');

      logAutomation_('Role Selection', interaction.meetingNo, member,
        'Role booked: ' + roleDef.name, 'Success', 'Planner row ' + result.rowNumber);

      if (configBool_('TASK_ENGINE_ENABLED')) {
        syncPhase1Tasks_();
      }

      sendImmediateAssignmentPrompts_(
        interaction.meetingNo,
        roleDef,
        member,
        phone
      );
      return;
    }

    if (result.status === 'ALREADY_ASSIGNED') {
      sendWhatsAppText_(phone,
        'You are already assigned as ' + result.assignedRole +
        ' for Meeting ' + interaction.meetingNo + '.');
      return;
    }

    if (result.status === 'SPEAKER') {
      sendWhatsAppText_(phone,
        'You are already listed as a speaker for Meeting ' + interaction.meetingNo +
        ', so I have not added a role for the same meeting.');
      return;
    }

    if (result.status === 'TAKEN') {
      sendWhatsAppText_(phone,
        roleDef.name + ' was just taken by another member. I will show you the roles that are still open.');
      const meeting = findMeetingForRoleList_(interaction.meetingNo, member);
      if (meeting && meeting.available.length) sendRoleList_(phone, member, meeting);
      return;
    }

    sendWhatsAppText_(phone, 'That meeting is no longer available for role selection. Type ROLE to see the current options.');
  } finally {
    lock.releaseLock();
  }
}

function reserveRole_(meetingNo, roleDef, member) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('VPE Weekly Planner');
  const lastRow = Math.min(62, sheet.getLastRow());
  const meetingNos = sheet.getRange(3, 1, Math.max(0, lastRow - 2), 1).getValues().flat();

  const index = meetingNos.findIndex(v => Number(v) === Number(meetingNo));
  if (index < 0) return { status: 'NOT_FOUND' };

  const rowNumber = index + 3;
  const row = sheet.getRange(rowNumber, 1, 1, 28).getValues()[0];
  const status = String(row[2] || '');
  if (status === 'Completed' || status === 'Cancelled') return { status: 'NOT_AVAILABLE' };

  const speakers = [row[6], row[10], row[14]].map(v => String(v || '').trim());
  if (speakers.indexOf(member) !== -1) return { status: 'SPEAKER' };

  for (let i = 0; i < ROLE_DEFS_.length; i++) {
    const assigned = String(row[ROLE_DEFS_[i].column - 1] || '').trim();
    if (assigned === member) {
      return { status: 'ALREADY_ASSIGNED', assignedRole: ROLE_DEFS_[i].name };
    }
  }

  const cell = sheet.getRange(rowNumber, roleDef.column);
  const existing = String(cell.getValue() || '').trim();
  if (existing) return { status: 'TAKEN', existing };

  cell.setValue(member);
  SpreadsheetApp.flush();

  const dateText = row[1] instanceof Date
    ? Utilities.formatDate(row[1], configValue_('TIMEZONE', 'Asia/Kolkata'), 'EEE, d MMM')
    : String(row[1] || '');

  return { status: 'BOOKED', rowNumber, dateText };
}

function findMeetingForRoleList_(meetingNo, member) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('VPE Weekly Planner');
  const values = sheet.getRange(3, 1, Math.min(60, Math.max(0, sheet.getLastRow() - 2)), 28).getValues();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (Number(row[0]) !== Number(meetingNo)) continue;

    const speakers = [row[6], row[10], row[14]].map(v => String(v || '').trim());
    if (speakers.indexOf(member) !== -1) return null;

    const memberHasRole = ROLE_DEFS_.some(def => String(row[def.column - 1] || '').trim() === member);
    if (memberHasRole) return null;

    return {
      meetingNo: Number(row[0]),
      meetingDate: row[1],
      rowNumber: i + 3,
      available: ROLE_DEFS_.filter(def => !String(row[def.column - 1] || '').trim())
    };
  }
  return null;
}

function sendWhatsAppText_(phone, text) {
  if (!phone) return { ok: false, error: 'Missing destination phone.' };

  const payload = {
    messaging_product: 'whatsapp',
    recipient_type: 'individual',
    to: normalizePhone_(phone),
    type: 'text',
    text: {
      preview_url: false,
      body: String(text || '')
    }
  };

  return sendWhatsAppPayload_(payload);
}

function sendWhatsAppPayload_(payload) {
  const destination = normalizePhone_(payload && payload.to ? payload.to : '');
  const sendMode = String(configValue_('WA_SEND_MODE', 'TEST')).toUpperCase();

  if (sendMode !== 'LIVE') {
    const allowlist = String(configValue_('WA_TEST_RECIPIENTS', ''))
      .split(',')
      .map(v => normalizePhone_(v))
      .filter(Boolean);

    if (!destination || allowlist.indexOf(destination) === -1) {
      logAutomation_(
        'WhatsApp',
        '',
        destination,
        'DEV outbound blocked by test-recipient allowlist',
        'Blocked',
        sendMode
      );
      return {
        ok: false,
        blocked: true,
        error: 'DEV outbound blocked by WA_TEST_RECIPIENTS.'
      };
    }
  }

  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('META_ACCESS_TOKEN');
  const phoneNumberId = props.getProperty('WA_PHONE_NUMBER_ID');

  if (!token || !phoneNumberId) {
    throw new Error('WhatsApp Script Properties are incomplete.');
  }

  const url = 'https://graph.facebook.com/v26.0/' +
    encodeURIComponent(phoneNumberId) + '/messages';

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + token
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  const body = response.getContentText();
  let parsed = {};
  try { parsed = JSON.parse(body); } catch (ignore) {}

  const messageId = parsed.messages && parsed.messages[0] && parsed.messages[0].id
    ? parsed.messages[0].id : '';

  if (status >= 200 && status < 300) {
    return { ok: true, status, messageId, body };
  }

  const errorText = parsed.error && parsed.error.message ? parsed.error.message : body;
  logAutomation_('WhatsApp', '', normalizePhone_(payload.to || ''), 'Outbound send failed', 'Error', 'HTTP ' + status + ': ' + errorText);
  return { ok: false, status, error: errorText, body };
}

function updateOutboxDeliveryStatus_(status) {
  const id = String(status.id || '');
  if (!id) return;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('WA Outbox');
  if (!sheet || sheet.getLastRow() < 2) return;

  const ids = sheet.getRange(2, 10, sheet.getLastRow() - 1, 1).getValues().flat();
  const idx = ids.findIndex(v => String(v || '') === id);
  if (idx < 0) return;

  const row = idx + 2;
  const next = String(status.status || '').toUpperCase();
  if (!next) return;

  // Meta status webhooks can arrive out of order. Never downgrade a message
  // that was already delivered or read when a delayed SENT receipt arrives.
  const cell = sheet.getRange(row, 7);
  const current = String(cell.getValue() || '').toUpperCase();
  const rank = { SENT: 1, DELIVERED: 2, READ: 3 };
  if (rank[next] && rank[current] && rank[next] < rank[current]) return;
  if (next === 'FAILED' && (current === 'DELIVERED' || current === 'READ')) return;
  if (next === 'SENT' && current === 'FAILED') return;
  cell.setValue(next);
}

function isDuplicateWebhookMessage_(messageId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('WA Webhook Log');
  if (!sheet || sheet.getLastRow() < 2) return false;

  const values = sheet.getRange(2, 5, sheet.getLastRow() - 1, 1).getValues().flat();
  let count = 0;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i] || '') === messageId) count++;
    if (count > 1) return true;
  }
  return false;
}

let CONFIG_CACHE_ = null;

function getConfigMap_() {
  if (CONFIG_CACHE_) return CONFIG_CACHE_;

  const map = {};
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Automation Config');

  if (sheet && sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    rows.forEach(row => {
      const key = String(row[0] || '').trim();
      if (key) map[key] = row[1];
    });
  }

  CONFIG_CACHE_ = map;
  return CONFIG_CACHE_;
}

function configValue_(key, fallback) {
  const map = getConfigMap_();
  if (!Object.prototype.hasOwnProperty.call(map, key)) return fallback;

  const value = map[key];
  return value === '' || value === null ? fallback : String(value);
}

function configBool_(key) {
  return String(configValue_(key, 'FALSE')).toUpperCase() === 'TRUE';
}

function normalizePhone_(phone) {
  return String(phone || '').replace(/\D/g, '');
}

function startOfDay_(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

function logAutomation_(module, meetingNo, memberOrItem, action, result, notes) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Automation Log');
    if (!sheet) return;

    sheet.appendRow([
      new Date(),
      module || '',
      meetingNo || '',
      memberOrItem || '',
      action || '',
      result || '',
      notes || ''
    ]);
  } catch (error) {
    console.error(error);
  }
}

/**
 * DEV diagnostic only. Does not send messages or change assignments.
 */
function phase1Diagnostic() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const requiredSheets = [
    'VPE Weekly Planner',
    'Member Master',
    'Automation Config',
    'Automation Tasks',
    'WA Outbox',
    'WA Submissions',
    'WA Webhook Log',
    'Automation Log'
  ];

  const missingSheets = requiredSheets.filter(name => !ss.getSheetByName(name));
  const props = PropertiesService.getScriptProperties();

  const result = {
    spreadsheetId: SPREADSHEET_ID,
    systemMode: configValue_('SYSTEM_MODE', ''),
    roleSelectionEnabled: configBool_('ROLE_SELECTION_ENABLED'),
    missingSheets: missingSheets,
    metaTokenPresent: Boolean(props.getProperty('META_ACCESS_TOKEN')),
    phoneNumberIdPresent: Boolean(props.getProperty('WA_PHONE_NUMBER_ID')),
    wabaIdPresent: Boolean(props.getProperty('WA_WABA_ID')),
    nextRoleMeetingForVpe: findNextRoleSelectionMeeting_(configValue_('VPE_MEMBER_NAME', 'Riyas AF'))
  };

  console.log(JSON.stringify(result));
  return result;
}
