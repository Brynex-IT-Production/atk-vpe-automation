/**
 * Prepared-speech vacancy targeting and self-service reservation.
 * Uses actual ATK speaking history from legacy Meeting roles + posted Participation Log.
 */

function processSpeakingSlotTargeting_() {
  if (!configBool_('SPEAKING_SLOT_WORKFLOW_ENABLED')) return;
  if (templateNameForKind_('SPEECH_INVITE') === 'atk_speech_invite_v2') {
    processSpeechDateInvites_();
    return;
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  if (!planner) return;

  const today = startOfDay_(new Date());
  const windowDays = Math.max(1, Number(configValue_('SPEECH_TARGETING_WINDOW_DAYS', '14')) || 14);
  const broadenDays = Math.max(0, Number(configValue_('SPEECH_BROADEN_DAYS', '7')) || 7);
  const batchSize = Math.max(1, Number(configValue_('SPEECH_INVITE_BATCH_SIZE', '5')) || 5);

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
    if (daysLeft < 0 || daysLeft > windowDays) return;

    const openSlots = getOpenSpeechSlotsFromRow_(row);
    if (!openSlots.length) return;

    const stage = daysLeft <= broadenDays ? 'BROAD' : 'PRIORITY';
    const candidates = getSpeechPriorityCandidates_(meetingNo, stage);
    if (!candidates.length) return;

    const inviteLimit = Math.max(batchSize, openSlots.length * 2);
    const invitePrefix = ['SPEECH_INVITE', meetingNo, stage].join('|') + '|';
    const alreadyInvited = countSuccessfulOutboxByPrefix_(invitePrefix);
    const remainingQuota = Math.max(0, inviteLimit - alreadyInvited);
    if (!remainingQuota) return;

    let invited = 0;

    for (let i = 0; i < candidates.length && invited < remainingQuota; i++) {
      const candidate = candidates[i];
      if (!isWhatsAppRecipientAllowed_(candidate.phone)) continue;

      const key = ['SPEECH_INVITE', meetingNo, stage, candidate.name].join('|');
      if (hasReminderBeenSent_(key)) continue;

      const result = sendSpeechSlotInvite_(candidate, meetingNo, meetingDate, openSlots.length, stage);
      recordWorkflowOutbound_(
        candidate.name,
        candidate.phone,
        'SPEECH_INVITE',
        result.body || '',
        meetingNo,
        result,
        key
      );

      if (result.ok) invited++;
    }
  });
}

function sendSpeechSlotInvite_(candidate, meetingNo, meetingDate, openCount, stage) {
  const dateText = Utilities.formatDate(
    meetingDate,
    configValue_('TIMEZONE', 'Asia/Kolkata'),
    'EEE, d MMM'
  );

  const body =
    'Hi ' + candidate.name + ',\n\n' +
    'ATK Meeting ' + meetingNo + ' on ' + dateText +
    ' has ' + openCount + ' prepared-speech slot' + (openCount === 1 ? '' : 's') + ' open.\n\n' +
    (stage === 'PRIORITY'
      ? 'We are first offering the open slot to members who have not delivered a prepared speech recently.\n\n'
      : '') +
    'If you would like to take a slot, tap below. The first confirmed selection will be reserved.';

  const payload = {
    messaging_product: 'whatsapp',
    recipient_type: 'individual',
    to: candidate.phone,
    type: 'interactive',
    interactive: {
      type: 'button',
      body: { text: body },
      action: {
        buttons: [{
          type: 'reply',
          reply: {
            id: 'SPEAK|' + meetingNo,
            title: 'Reserve speech slot'
          }
        }]
      }
    }
  };

  const result = sendAutomatedQuickReply_(
    candidate.phone,
    'SPEECH_INVITE',
    payload,
    [candidate.name, meetingNo, dateText, openCount],
    'SPEAK|' + meetingNo
  );
  result.body = body;
  logAutomation_(
    'Speaking Slot',
    meetingNo,
    candidate.name,
    'Speech-slot invitation ' + stage,
    result.ok ? 'Sent' : (result.blocked ? 'Blocked' : 'Error'),
    candidate.lastSpeechMeeting ? 'Last speech: ' + candidate.lastSpeechMeeting : 'No recent speech found'
  );
  return result;
}

function handleSpeechSlotSelection_(member, phone, interaction) {
  if (!configBool_('SPEAKING_SLOT_WORKFLOW_ENABLED')) {
    sendWhatsAppText_(phone, 'Prepared-speech self-booking is currently unavailable.');
    return;
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    sendWhatsAppText_(phone, 'Another booking is being processed. Please try again in a few seconds.');
    return;
  }

  try {
    const result = reserveSpeechSlot_(interaction.meetingNo, member);

    if (result.status === 'BOOKED') {
      sendWhatsAppText_(
        phone,
        'Confirmed ✅\nYour prepared speech is reserved for ' +
        result.dateText + '.\n\nThank you, ' + member + '.'
      );

      logAutomation_(
        'Speaking Slot',
        interaction.meetingNo,
        member,
        'Speech slot booked: S' + result.slot,
        'Success',
        'Planner row ' + result.rowNumber
      );

      if (configBool_('TASK_ENGINE_ENABLED')) syncPhase1Tasks_();
      return;
    }

    if (result.status === 'ALREADY_SPEAKER') {
      sendWhatsAppText_(phone, 'You already have a prepared-speech slot in Meeting ' + interaction.meetingNo + '.');
      return;
    }

    if (result.status === 'HAS_ROLE') {
      sendWhatsAppText_(
        phone,
        'You are already assigned as ' + result.roleName +
        ' for Meeting ' + interaction.meetingNo + ', so I did not add a speech slot.'
      );
      return;
    }

    if (result.status === 'FULL') {
      sendWhatsAppText_(phone, 'The prepared-speech slots for Meeting ' + interaction.meetingNo + ' have just been filled.');
      return;
    }

    sendWhatsAppText_(phone, 'That meeting is no longer available for speech-slot booking.');
  } finally {
    lock.releaseLock();
  }
}

function reserveSpeechSlot_(meetingNo, member) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('VPE Weekly Planner');
  if (!sheet) return { status: 'NOT_FOUND' };

  const lastRow = Math.min(62, sheet.getLastRow());
  const ids = sheet.getRange(3, 1, Math.max(0, lastRow - 2), 1).getValues().flat();
  const idx = ids.findIndex(v => Number(v) === Number(meetingNo));
  if (idx < 0) return { status: 'NOT_FOUND' };

  const rowNumber = idx + 3;
  const row = sheet.getRange(rowNumber, 1, 1, 28).getValues()[0];
  const status = String(row[2] || '');
  if (status === 'Completed' || status === 'Cancelled' ||
      !(row[1] instanceof Date) || startOfDay_(row[1]) < startOfDay_(new Date())) {
    return { status: 'NOT_AVAILABLE' };
  }

  const speakerColumns = [7, 11, 15];
  const statusColumns = [9, 13, 17];
  const existingSpeakers = speakerColumns.map(c => String(row[c - 1] || '').trim());

  if (existingSpeakers.indexOf(member) !== -1) {
    return { status: 'ALREADY_SPEAKER' };
  }

  for (let i = 0; i < ROLE_DEFS_.length; i++) {
    const assigned = String(row[ROLE_DEFS_[i].column - 1] || '').trim();
    if (assigned === member) {
      return { status: 'HAS_ROLE', roleName: ROLE_DEFS_[i].name };
    }
  }

  const openIndex = existingSpeakers.findIndex(v => !v);
  if (openIndex < 0) return { status: 'FULL' };

  const speakerCol = speakerColumns[openIndex];
  const statusCol = statusColumns[openIndex];

  sheet.getRange(rowNumber, speakerCol).setValue(member);
  sheet.getRange(rowNumber, statusCol).setValue('Planned');
  SpreadsheetApp.flush();

  const dateText = row[1] instanceof Date
    ? Utilities.formatDate(row[1], configValue_('TIMEZONE', 'Asia/Kolkata'), 'EEE, d MMM')
    : String(row[1] || '');

  return {
    status: 'BOOKED',
    slot: openIndex + 1,
    rowNumber,
    dateText
  };
}

function sendSpeakerDetailsPrompt_(meetingNo, slot, member, phone) {
  if (!configBool_('SPEAKER_DETAILS_WORKFLOW_ENABLED')) return;

  const related = 'S' + slot;
  const key = ['SPEAKER_DETAILS_PROMPT', meetingNo, related, member].join('|');
  if (hasReminderBeenSent_(key) || activeConversation_(phone)) return;

  const prompt = startGuidedRequest_(member, phone, 'SPEAKER_DETAILS',
    meetingNo, related, 'speaker');
  if (!prompt) return;

  const result = sendWhatsAppText_(phone, prompt.text);
  recordWorkflowOutbound_(member, phone, 'SPEAKER_DETAILS_PROMPT',
    prompt.text, meetingNo, result, key);
  if (result.ok) setConversation_(phone, prompt.state);
}

function getOpenSpeechSlotsFromRow_(row) {
  const cols = [
    {slot: 1, index: 6},
    {slot: 2, index: 10},
    {slot: 3, index: 14}
  ];
  return cols.filter(x => !String(row[x.index] || '').trim()).map(x => x.slot);
}

function getSpeechPriorityCandidates_(targetMeetingNo, stage) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const members = ss.getSheetByName('Member Master');
  const planner = ss.getSheetByName('VPE Weekly Planner');
  if (!members || !planner) return [];

  const memberRows = members.getRange(2, 1, Math.max(0, members.getLastRow() - 1), 4).getValues();
  const active = memberRows
    .filter(r => String(r[3] || '').trim() === 'Active')
    .map(r => ({
      name: String(r[1] || '').trim(),
      phone: normalizePhone_(r[2] || '')
    }))
    .filter(x => x.name && x.phone);

  const plannerRows = planner.getRange(3, 1, Math.min(60, Math.max(0, planner.getLastRow() - 2)), 28).getValues();
  const targetRow = plannerRows.find(r => Number(r[0]) === Number(targetMeetingNo));
  if (!targetRow) return [];

  const sameMeetingRoles = new Set(
    ROLE_DEFS_.map(def => String(targetRow[def.column - 1] || '').trim()).filter(Boolean)
  );

  const futureSpeakers = new Set();
  plannerRows.forEach(r => {
    const meetingNo = Number(r[0]);
    if (!meetingNo || meetingNo < Number(targetMeetingNo)) return;
    [r[6], r[10], r[14]].forEach(v => {
      const name = String(v || '').trim();
      if (name) futureSpeakers.add(name);
    });
  });

  const history = getPreparedSpeechHistory_();
  const lookback = Math.max(1, Number(configValue_('SPEECH_PRIORITY_LOOKBACK_MEETINGS', '6')) || 6);
  const recentMeetings = history.meetings
    .filter(m => m < Number(targetMeetingNo))
    .sort((a, b) => b - a)
    .slice(0, lookback);

  const recentSet = new Set(recentMeetings);
  const recentlySpoke = new Set();
  Object.keys(history.byMember).forEach(member => {
    if (history.byMember[member].some(m => recentSet.has(m))) {
      recentlySpoke.add(member);
    }
  });

  const candidates = active
    .filter(x => !futureSpeakers.has(x.name))
    .filter(x => !sameMeetingRoles.has(x.name))
    .map(x => {
      const meetings = history.byMember[x.name] || [];
      const lastSpeechMeeting = meetings.length ? Math.max.apply(null, meetings) : 0;
      return Object.assign({}, x, {
        lastSpeechMeeting,
        spokeRecently: recentlySpoke.has(x.name)
      });
    })
    .filter(x => stage === 'BROAD' || !x.spokeRecently)
    .sort((a, b) => {
      if (a.lastSpeechMeeting === b.lastSpeechMeeting) {
        return a.name.localeCompare(b.name);
      }
      if (!a.lastSpeechMeeting) return -1;
      if (!b.lastSpeechMeeting) return 1;
      return a.lastSpeechMeeting - b.lastSpeechMeeting;
    });

  return candidates;
}

function getPreparedSpeechHistory_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const byMember = {};
  const meetings = new Set();

  // Legacy source: Meeting roles, columns are meeting numbers and first three rows are SPKR.
  const legacy = ss.getSheetByName('Meeting roles');
  if (legacy && legacy.getLastColumn() >= 2) {
    const width = Math.min(legacy.getLastColumn(), 60);
    const headers = legacy.getRange(1, 2, 1, width - 1).getValues()[0];
    const speakers = legacy.getRange(2, 2, 3, width - 1).getValues();

    headers.forEach((header, col) => {
      const meetingNo = Number(header);
      if (!meetingNo || meetingNo >= Number(configValue_('CUTOVER_MEETING', '450'))) return;
      meetings.add(meetingNo);

      for (let row = 0; row < speakers.length; row++) {
        const member = String(speakers[row][col] || '').trim();
        if (!member) continue;
        if (!byMember[member]) byMember[member] = [];
        byMember[member].push(meetingNo);
      }
    });
  }

  // New source: posted prepared speeches after cutover.
  const participation = ss.getSheetByName('Participation Log');
  if (participation && participation.getLastRow() >= 2) {
    const rows = participation.getRange(2, 1, participation.getLastRow() - 1, 8).getValues();
    rows.forEach(r => {
      const meetingNo = Number(r[0]);
      const member = String(r[2] || '').trim();
      const code = String(r[4] || '').trim();
      const status = String(r[6] || '').trim();
      if (!meetingNo || !member || code !== 'SPKR' || status !== 'Posted') return;
      meetings.add(meetingNo);
      if (!byMember[member]) byMember[member] = [];
      byMember[member].push(meetingNo);
    });
  }

  Object.keys(byMember).forEach(member => {
    byMember[member] = Array.from(new Set(byMember[member])).sort((a, b) => a - b);
  });

  return {
    byMember,
    meetings: Array.from(meetings).sort((a, b) => a - b)
  };
}

function isWhatsAppRecipientAllowed_(phone) {
  const destination = normalizePhone_(phone);
  const sendMode = String(configValue_('WA_SEND_MODE', 'TEST')).toUpperCase();
  if (sendMode === 'LIVE') return Boolean(destination);

  const allowlist = String(configValue_('WA_TEST_RECIPIENTS', ''))
    .split(',')
    .map(v => normalizePhone_(v))
    .filter(Boolean);

  return Boolean(destination) && allowlist.indexOf(destination) !== -1;
}


function countSuccessfulOutboxByPrefix_(prefix) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('WA Outbox');
  if (!sheet || sheet.getLastRow() < 2) return 0;

  const rows = sheet.getRange(2, 7, sheet.getLastRow() - 1, 10).getValues();
  const successfulStates = new Set(['SENT', 'DELIVERED', 'READ', 'ACCEPTED']);
  let count = 0;

  rows.forEach(row => {
    const state = String(row[0] || '').toUpperCase();
    const key = String(row[9] || '');
    if (successfulStates.has(state) && key.indexOf(prefix) === 0) {
      count++;
    }
  });

  return count;
}


/** Invite once for the nearest open meeting; show up to four dates at tap time. */
function processSpeechDateInvites_() {
  const meetings = availableSpeechMeetings_();
  if (!meetings.length) return;
  const first = meetings[0];
  const daysLeft = Math.round((startOfDay_(first.date) - startOfDay_(new Date())) / 86400000);
  const windowDays = Math.max(1, Number(configValue_('SPEECH_TARGETING_WINDOW_DAYS', '14')) || 14);
  if (daysLeft > windowDays) return;
  const broadenDays = Math.max(0, Number(configValue_('SPEECH_BROADEN_DAYS', '7')) || 7);
  const stage = daysLeft <= broadenDays ? 'BROAD' : 'PRIORITY';
  const candidates = getSpeechPriorityCandidates_(first.meetingNo, stage);
  const batchSize = Math.max(1, Number(configValue_('SPEECH_INVITE_BATCH_SIZE', '5')) || 5);
  const prefix = ['SPEECH_DATES_INVITE', first.meetingNo, stage].join('|') + '|';
  let remaining = Math.max(0, Math.max(batchSize, first.openCount * 2) -
    countSuccessfulOutboxByPrefix_(prefix));
  candidates.forEach(candidate => {
    if (!remaining || !isWhatsAppRecipientAllowed_(candidate.phone)) return;
    const key = prefix + candidate.name;
    if (hasReminderBeenSent_(key)) return;
    const body = speechDateInviteBody_(candidate.name);
    const payload = {
      messaging_product: 'whatsapp', recipient_type: 'individual',
      to: candidate.phone, type: 'interactive',
      interactive: {
        type: 'button', body: {text: body},
        action: {buttons: [{type: 'reply',
          reply: {id: 'SPEECH_DATES', title: 'Reserve speech slot'}}]}
      }
    };
    const result = sendAutomatedQuickReply_(candidate.phone, 'SPEECH_INVITE',
      payload, [candidate.name], 'SPEECH_DATES');
    recordWorkflowOutbound_(candidate.name, candidate.phone,
      'SPEECH_INVITE', body, first.meetingNo, result, key);
    if (result.ok) remaining--;
  });
}

function speechDateInviteBody_(name) {
  return 'Hi ' + name + ',\n\nWant to give a speech?\n\n' +
    'Select a slot before it’s gone! 🎤\n\n' +
    'If you’d like to take a speaking slot, choose a convenient meeting date using the button below.';
}

function availableSpeechMeetings_() {
  const planner = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('VPE Weekly Planner');
  if (!planner || planner.getLastRow() < 3) return [];
  const rows = planner.getRange(3, 1,
    Math.min(60, planner.getLastRow() - 2), 28).getValues();
  const today = startOfDay_(new Date());
  return rows.filter(row => Number(row[0]) && row[1] instanceof Date &&
      startOfDay_(row[1]) >= today &&
      !['Completed', 'Cancelled'].includes(String(row[2] || '')) &&
      getOpenSpeechSlotsFromRow_(row).length)
    .map(row => ({meetingNo: Number(row[0]), date: row[1],
      openCount: getOpenSpeechSlotsFromRow_(row).length, row: row}))
    .sort((a, b) => a.date - b.date || a.meetingNo - b.meetingNo);
}

function handleSpeechDateMenu_(member, phone) {
  if (!configBool_('SPEAKING_SLOT_WORKFLOW_ENABLED') ||
      templateNameForKind_('SPEECH_INVITE') !== 'atk_speech_invite_v2') {
    sendWhatsAppText_(phone, 'Speech booking is not available right now.');
    return;
  }
  const dates = availableSpeechMeetings_().filter(meeting => {
    const row = meeting.row;
    return ![row[6], row[10], row[14]].some(v => String(v || '').trim() === member) &&
      !ROLE_DEFS_.some(def => String(row[def.column - 1] || '').trim() === member);
  }).slice(0, 4);
  if (!dates.length) {
    sendWhatsAppText_(phone, 'There are no available meeting dates for you right now.');
    return;
  }
  const result = sendWhatsAppPayload_({
    messaging_product: 'whatsapp', recipient_type: 'individual', to: phone,
    type: 'interactive',
    interactive: {
      type: 'list',
      body: {text: 'Choose a convenient date for your prepared speech:'},
      action: {button: 'View dates', sections: [{
        title: 'Upcoming meetings',
        rows: dates.map(meeting => ({
          id: 'SPEAK|' + meeting.meetingNo,
          title: Utilities.formatDate(meeting.date,
            configValue_('TIMEZONE', 'Asia/Kolkata'), 'EEE, d MMM yyyy')
        }))
      }]}
    }
  });
  logAutomation_('Speaking Slot', '', member, 'Show speech meeting dates',
    result.ok ? 'Sent' : 'Error', result.error || '');
}
