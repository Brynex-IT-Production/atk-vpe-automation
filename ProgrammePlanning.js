/**
 * VPE programme-slot planning.
 *
 * Speech slots remain bookable by default. The VPE can place explicit holds,
 * request a confirmed speaker to move, and convert a hold into a confirmed
 * education/special session. Holds never auto-release.
 */

const PROGRAMME_SLOT_SHEET_ = 'Programme Slots';
const PROGRAMME_REVIEW_SHEET_ = 'Programme Reviews';

let PROGRAMME_SLOT_CACHE_ = null;

function ensureProgrammePlanningSetup_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let slots = ss.getSheetByName(PROGRAMME_SLOT_SHEET_);
  if (!slots) {
    slots = ss.insertSheet(PROGRAMME_SLOT_SHEET_);
    slots.getRange(1, 1, 1, 15).setValues([[
      'Meeting No','Meeting Date','Slot','State','Speaker',
      'Session Type','Presenter','Session Title','Time Allotted (Min)',
      'Move Requested At','Previous Speaker','Move Target Meeting',
      'Updated At','Updated By','Notes'
    ]]);
    slots.setFrozenRows(1);
    slots.hideSheet();
  }

  let reviews = ss.getSheetByName(PROGRAMME_REVIEW_SHEET_);
  if (!reviews) {
    reviews = ss.insertSheet(PROGRAMME_REVIEW_SHEET_);
    reviews.getRange(1, 1, 1, 10).setValues([[
      'Review Key','Month 1','Month 2','Status','First Sent At',
      'Last Reminder At','Completed At','Completed By','Outcome','Updated At'
    ]]);
    reviews.setFrozenRows(1);
    reviews.hideSheet();
  }

  const config = ss.getSheetByName('Automation Config');
  if (!config) throw new Error('Automation Config is missing.');

  const desired = [
    ['PROGRAMME_PLANNING_ENABLED','TRUE','Monthly two-month VPE programme review + slot holds','Pilot'],
    ['PROGRAMME_REVIEW_DAY','25','Monthly review day','Active'],
    ['PROGRAMME_REVIEW_HOUR','20','Monthly first prompt hour','Active'],
    ['PROGRAMME_REVIEW_REMINDER_DAYS','26,27','Follow-up days after monthly review prompt','Active'],
    ['PROGRAMME_REVIEW_REMINDER_HOURS','7,20','Follow-up hours','Active'],
    ['SPEECH_BOOKING_WINDOW_MONTHS','2','Rolling forward booking horizon in calendar months','Active'],
    ['PROGRAMME_REVIEW_URL_TTL_DAYS','75','Signed VPE planning-link lifetime','Active'],
    ['WA_TEMPLATE_VPE_PROGRAMME_REVIEW','atk_vpe_programme_review_v1','Monthly VPE programme review template','PENDING'],
    ['WA_TEMPLATE_SPEAKER_MOVE_REQUEST','atk_speaker_move_request_v1','Protected speaker move-request template','PENDING']
  ];

  const existing = config.getLastRow() >= 2
    ? config.getRange(2, 1, config.getLastRow() - 1, 4).getValues()
    : [];
  const rowByKey = {};
  existing.forEach((row, i) => {
    const key = String(row[0] || '').trim();
    if (key) rowByKey[key] = i + 2;
  });

  desired.forEach(row => {
    const rowNo = rowByKey[row[0]];
    if (rowNo) {
      // Keep live status/config values except for the feature enablement and
      // booking horizon explicitly approved for this DEV change.
      if (row[0] === 'PROGRAMME_PLANNING_ENABLED' ||
          row[0] === 'SPEECH_BOOKING_WINDOW_MONTHS') {
        config.getRange(rowNo, 2).setValue(row[1]);
      }
      return;
    }
    config.appendRow(row);
  });

  CONFIG_CACHE_ = null;
}

function ensureProgrammePlanningTemplates_() {
  if (!configBool_('PROGRAMME_PLANNING_ENABLED')) return;
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('META_ACCESS_TOKEN');
  const wabaId = props.getProperty('WA_WABA_ID');
  if (!token || !wabaId) return;

  const timezone = configValue_('TIMEZONE', 'Asia/Kolkata');
  const todayKey = Utilities.formatDate(new Date(), timezone, 'yyyyMMdd');
  if (props.getProperty('PROGRAMME_TEMPLATE_CHECK_DAY') === todayKey) return;

  const results = [
    createOrReadTemplate_(token, wabaId, vpeProgrammeReviewSpec_()),
    createOrReadTemplate_(token, wabaId, speakerMoveRequestSpec_())
  ];
  props.setProperty('PROGRAMME_TEMPLATE_CHECK_DAY', todayKey);

  logAutomation_(
    'Programme Planning',
    '',
    configValue_('VPE_MEMBER_NAME', 'Riyas AF'),
    'Programme templates ensured',
    results.every(r => r.ok) ? 'Success' : 'Error',
    results.map(r => r.name + ':' + (r.status || r.action || r.error || '')).join(' | ')
  );
}

function programmeSlotKey_(meetingNo, slot) {
  return Number(meetingNo) + '|S' + Number(slot);
}

function readProgrammeSlots_() {
  if (PROGRAMME_SLOT_CACHE_) return PROGRAMME_SLOT_CACHE_;
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(PROGRAMME_SLOT_SHEET_);
  const map = {};
  if (sheet && sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 15).getValues().forEach((row, i) => {
      const meetingNo = Number(row[0]);
      const slot = Number(row[2]);
      if (!meetingNo || !slot) return;
      map[programmeSlotKey_(meetingNo, slot)] = {
        rowNumber: i + 2,
        meetingNo: meetingNo,
        meetingDate: row[1],
        slot: slot,
        state: String(row[3] || 'OPEN'),
        speaker: String(row[4] || '').trim(),
        sessionType: String(row[5] || '').trim(),
        presenter: String(row[6] || '').trim(),
        sessionTitle: String(row[7] || '').trim(),
        minutes: Number(row[8] || 0),
        moveRequestedAt: row[9],
        previousSpeaker: String(row[10] || '').trim(),
        moveTargetMeeting: Number(row[11] || 0),
        updatedAt: row[12],
        updatedBy: String(row[13] || '').trim(),
        notes: String(row[14] || '').trim()
      };
    });
  }
  PROGRAMME_SLOT_CACHE_ = map;
  return map;
}

function programmeSlotState_(meetingNo, slot) {
  const item = readProgrammeSlots_()[programmeSlotKey_(meetingNo, slot)];
  return item ? item.state : 'OPEN';
}

function isProgrammeSlotBookable_(meetingNo, slot) {
  return programmeSlotState_(meetingNo, slot) === 'OPEN';
}

function syncProgrammeSlots_() {
  ensureProgrammePlanningSetup_();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  const sheet = ss.getSheetByName(PROGRAMME_SLOT_SHEET_);
  if (!planner || !sheet) return;

  const existingRows = sheet.getLastRow() >= 2
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 15).getValues()
    : [];
  const existing = {};
  existingRows.forEach(row => {
    const meetingNo = Number(row[0]);
    const slot = Number(row[2]);
    if (meetingNo && slot) existing[programmeSlotKey_(meetingNo, slot)] = row.slice();
  });

  const plannerRows = planner.getRange(
    3, 1, Math.max(0, planner.getLastRow() - 2), 28
  ).getValues();

  plannerRows.forEach(row => {
    const meetingNo = Number(row[0]);
    const meetingDate = row[1];
    if (!meetingNo || !(meetingDate instanceof Date)) return;

    [1,2,3].forEach(slot => {
      const key = programmeSlotKey_(meetingNo, slot);
      const speaker = String(row[[6,10,14][slot - 1]] || '').trim();
      const old = existing[key] || [
        meetingNo, meetingDate, slot, 'OPEN', '', '', '', '', '',
        '', '', '', '', '', ''
      ];
      old[0] = meetingNo;
      old[1] = meetingDate;
      old[2] = slot;
      old[4] = speaker;

      const state = String(old[3] || 'OPEN');
      if (state === 'VPE_HOLD' || state === 'SESSION_CONFIRMED') {
        if (speaker) {
          old[3] = 'SPEAKER_CONFIRMED';
          old[14] = 'Planner speaker assignment overrode previous hold/session state.';
          old[5] = ''; old[6] = ''; old[7] = ''; old[8] = '';
        }
      } else if (state === 'MOVE_REQUESTED') {
        if (!speaker) old[3] = 'VPE_HOLD';
      } else {
        old[3] = speaker ? 'SPEAKER_CONFIRMED' : 'OPEN';
      }
      existing[key] = old;
    });
  });

  const rows = Object.keys(existing)
    .map(key => existing[key])
    .sort((a,b) => Number(a[0]) - Number(b[0]) || Number(a[2]) - Number(b[2]));

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 15).clearContent();
  }
  if (rows.length) sheet.getRange(2, 1, rows.length, 15).setValues(rows);
  PROGRAMME_SLOT_CACHE_ = null;
}

function speechBookingWindowEnd_(now) {
  const d = now ? new Date(now) : new Date();
  const months = Math.max(1, Number(configValue_('SPEECH_BOOKING_WINDOW_MONTHS', '2')) || 2);
  return new Date(d.getFullYear(), d.getMonth() + months + 1, 0, 23, 59, 59, 999);
}

function programmeMonthKey_(date) {
  return Utilities.formatDate(date, configValue_('TIMEZONE', 'Asia/Kolkata'), 'yyyy-MM');
}

function programmeReviewPeriodFor_(date) {
  const d = new Date(date || new Date());
  const first = new Date(d.getFullYear(), d.getMonth() + 1, 1);
  const second = new Date(d.getFullYear(), d.getMonth() + 2, 1);
  const tz = configValue_('TIMEZONE', 'Asia/Kolkata');
  return {
    month1: Utilities.formatDate(first, tz, 'yyyy-MM'),
    month2: Utilities.formatDate(second, tz, 'yyyy-MM'),
    label1: Utilities.formatDate(first, tz, 'MMMM yyyy'),
    label2: Utilities.formatDate(second, tz, 'MMMM yyyy'),
    key: Utilities.formatDate(first, tz, 'yyyy-MM') + ',' +
      Utilities.formatDate(second, tz, 'yyyy-MM')
  };
}

function getProgrammeReviewRow_(period) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(PROGRAMME_REVIEW_SHEET_);
  if (!sheet) return null;
  if (sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
    const idx = rows.findIndex(r => String(r[0] || '') === period.key);
    if (idx >= 0) return {sheet: sheet, rowNumber: idx + 2, values: rows[idx]};
  }
  sheet.appendRow([
    period.key, period.label1, period.label2, 'PENDING', '', '', '', '', '', new Date()
  ]);
  return {sheet: sheet, rowNumber: sheet.getLastRow(),
    values: [period.key, period.label1, period.label2, 'PENDING','','','','','','']};
}

function markProgrammeReviewCompleted_(periodKey, member, outcome) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(PROGRAMME_REVIEW_SHEET_);
  if (!sheet || sheet.getLastRow() < 2) return;
  const keys = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
  const idx = keys.findIndex(v => String(v || '') === periodKey);
  if (idx < 0) return;
  const row = idx + 2;
  sheet.getRange(row, 4).setValue('COMPLETED');
  sheet.getRange(row, 7, 1, 4).setValues([[
    new Date(), member || '', outcome || 'REVIEWED', new Date()
  ]]);
}

function processMonthlyProgrammePlanning_() {
  if (!configBool_('PROGRAMME_PLANNING_ENABLED')) return;

  const now = new Date();
  const tz = configValue_('TIMEZONE', 'Asia/Kolkata');
  const day = Number(Utilities.formatDate(now, tz, 'd'));
  const hour = Number(Utilities.formatDate(now, tz, 'H'));
  const reviewDay = Number(configValue_('PROGRAMME_REVIEW_DAY', '25')) || 25;
  const reviewHour = Number(configValue_('PROGRAMME_REVIEW_HOUR', '20')) || 20;
  const reminderDays = String(configValue_('PROGRAMME_REVIEW_REMINDER_DAYS', '26,27'))
    .split(',').map(v => Number(v.trim())).filter(Number.isFinite);
  const reminderHours = String(configValue_('PROGRAMME_REVIEW_REMINDER_HOURS', '7,20'))
    .split(',').map(v => Number(v.trim())).filter(Number.isFinite);

  const isFirst = day === reviewDay && hour === reviewHour;
  const isReminder = reminderDays.indexOf(day) !== -1 && reminderHours.indexOf(hour) !== -1;
  if (!isFirst && !isReminder) return;

  syncProgrammeSlots_();
  const period = programmeReviewPeriodFor_(now);
  const review = getProgrammeReviewRow_(period);
  if (String(review.values[3] || '') === 'COMPLETED') return;

  const member = configValue_('VPE_MEMBER_NAME', 'Riyas AF');
  const phone = phoneForMember_(member);
  if (!phone || !isWhatsAppRecipientAllowed_(phone)) return;

  const body =
    'Hi ' + member + ',\n\n' +
    'Time to review ATK speech slots for ' + period.label1 + ' and ' + period.label2 + '.\n\n' +
    'Do you want to hold any slots for an Education Session or Special Meeting?\n\n' +
    'Confirmed speakers stay protected unless they agree to move.';

  const noHoldId = 'VPE_PLAN|NO_HOLD|' + period.key;
  const reviewId = 'VPE_PLAN|REVIEW|' + period.key;
  const payload = {
    messaging_product: 'whatsapp',
    recipient_type: 'individual',
    to: phone,
    type: 'interactive',
    interactive: {
      type: 'button',
      body: {text: body},
      action: {buttons: [
        {type:'reply', reply:{id:noHoldId, title:'No hold'}},
        {type:'reply', reply:{id:reviewId, title:'Review slots'}}
      ]}
    }
  };

  const sendKey = [
    'VPE_PROGRAMME_REVIEW', period.key, 'D' + day, 'H' + hour
  ].join('|');
  if (hasReminderBeenSent_(sendKey)) return;

  const result = sendAutomatedQuickReply_(
    phone, 'VPE_PROGRAMME_REVIEW', payload,
    [member, period.label1, period.label2], [noHoldId, reviewId]
  );
  recordWorkflowOutbound_(member, phone, 'VPE_PROGRAMME_REVIEW', body, '', result, sendKey);

  if (result.ok) {
    const targetCol = isFirst ? 5 : 6;
    review.sheet.getRange(review.rowNumber, targetCol).setValue(new Date());
    review.sheet.getRange(review.rowNumber, 10).setValue(new Date());
  }
}

function handleVpeProgrammeReviewAction_(member, phone, interaction) {
  if (member !== configValue_('VPE_MEMBER_NAME', 'Riyas AF')) return;
  const periodKey = String(interaction.periodKey || '');
  if (!/^\d{4}-\d{2},\d{4}-\d{2}$/.test(periodKey)) {
    sendWhatsAppText_(phone, 'That programme-review link is no longer valid.');
    return;
  }

  if (interaction.action === 'NO_HOLD') {
    markProgrammeReviewCompleted_(periodKey, member, 'NO_NEW_HOLDS');
    sendWhatsAppText_(phone,
      'Done ✅ No new holds added for this review. Existing holds, if any, remain unchanged.');
    return;
  }

  if (interaction.action === 'REVIEW') {
    sendWhatsAppText_(phone,
      'Open the VPE programme planner here:\n' + programmeReviewUrl_(periodKey) +
      '\n\nConfirmed speakers remain protected until they choose another available date.');
  }
}

function programmeReviewSecret_() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('PROGRAMME_REVIEW_SIGNING_SECRET');
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('PROGRAMME_REVIEW_SIGNING_SECRET', secret);
  }
  return secret;
}

function programmeReviewToken_(periodKey) {
  const days = Math.max(7, Number(configValue_('PROGRAMME_REVIEW_URL_TTL_DAYS', '75')) || 75);
  const expiry = Date.now() + days * 86400000;
  const raw = periodKey + '|' + expiry;
  const encoded = Utilities.base64EncodeWebSafe(raw, Utilities.Charset.UTF_8).replace(/=+$/,'');
  const sigBytes = Utilities.computeHmacSha256Signature(raw, programmeReviewSecret_(), Utilities.Charset.UTF_8);
  const sig = sigBytes.map(v => ('0' + (v & 255).toString(16)).slice(-2)).join('');
  return encoded + '.' + sig;
}

function validateProgrammeReviewToken_(token) {
  const parts = String(token || '').split('.');
  if (parts.length !== 2 || !/^[0-9a-f]{64}$/.test(parts[1])) return '';
  let raw = '';
  try {
    raw = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
  } catch (_) {
    return '';
  }
  const rawParts = raw.split('|');
  if (rawParts.length !== 2 || !/^\d{4}-\d{2},\d{4}-\d{2}$/.test(rawParts[0])) return '';
  if (Date.now() > Number(rawParts[1] || 0)) return '';
  const sigBytes = Utilities.computeHmacSha256Signature(raw, programmeReviewSecret_(), Utilities.Charset.UTF_8);
  const expected = sigBytes.map(v => ('0' + (v & 255).toString(16)).slice(-2)).join('');
  let diff = 0;
  for (let i = 0; i < expected.length; i++) diff |= expected.charCodeAt(i) ^ parts[1].charCodeAt(i);
  return diff === 0 ? rawParts[0] : '';
}

function programmeReviewUrl_(periodKey) {
  const base = ScriptApp.getService().getUrl();
  return base + '?view=programme&token=' + encodeURIComponent(programmeReviewToken_(periodKey));
}

function renderProgrammeReviewPage_(token) {
  if (!validateProgrammeReviewToken_(token)) {
    return HtmlService.createHtmlOutput(
      '<h3>ATK Programme Planner</h3><p>This link is invalid or has expired.</p>'
    ).setTitle('ATK Programme Planner');
  }
  const template = HtmlService.createTemplateFromFile('ProgrammeReview');
  template.token = token;
  return template.evaluate()
    .setTitle('ATK Programme Planner')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getProgrammeReviewData(token) {
  const periodKey = validateProgrammeReviewToken_(token);
  if (!periodKey) throw new Error('This programme-review link is invalid or expired.');
  syncProgrammeSlots_();

  const months = periodKey.split(',');
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  const rows = planner.getRange(
    3, 1, Math.max(0, planner.getLastRow() - 2), 28
  ).getValues();
  const slotMap = readProgrammeSlots_();
  const tz = configValue_('TIMEZONE', 'Asia/Kolkata');

  const meetings = rows.filter(row =>
    Number(row[0]) && row[1] instanceof Date &&
    months.indexOf(programmeMonthKey_(row[1])) !== -1 &&
    !['Completed','Cancelled'].includes(String(row[2] || ''))
  ).map(row => {
    const meetingNo = Number(row[0]);
    return {
      meetingNo: meetingNo,
      dateText: Utilities.formatDate(row[1], tz, 'EEE, d MMM yyyy'),
      slots: [1,2,3].map(slot => {
        const item = slotMap[programmeSlotKey_(meetingNo, slot)] || {};
        return {
          slot: slot,
          state: String(item.state || 'OPEN'),
          speaker: String(row[[6,10,14][slot - 1]] || '').trim(),
          sessionType: String(item.sessionType || ''),
          presenter: String(item.presenter || ''),
          sessionTitle: String(item.sessionTitle || ''),
          minutes: Number(item.minutes || 0),
          previousSpeaker: String(item.previousSpeaker || ''),
          moveTargetMeeting: Number(item.moveTargetMeeting || 0)
        };
      })
    };
  });

  return {periodKey: periodKey, months: months, meetings: meetings};
}

function updateProgrammeSlotRecord_(meetingNo, slot, changes) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(PROGRAMME_SLOT_SHEET_);
  PROGRAMME_SLOT_CACHE_ = null;
  const item = readProgrammeSlots_()[programmeSlotKey_(meetingNo, slot)];
  if (!item) throw new Error('Programme slot not found for Meeting ' + meetingNo + ' S' + slot + '.');
  const row = sheet.getRange(item.rowNumber, 1, 1, 15).getValues()[0];

  const index = {
    state:3, speaker:4, sessionType:5, presenter:6, sessionTitle:7, minutes:8,
    moveRequestedAt:9, previousSpeaker:10, moveTargetMeeting:11,
    updatedAt:12, updatedBy:13, notes:14
  };
  Object.keys(changes || {}).forEach(key => {
    if (Object.prototype.hasOwnProperty.call(index, key)) row[index[key]] = changes[key];
  });
  row[12] = new Date();
  row[13] = configValue_('VPE_MEMBER_NAME', 'Riyas AF');
  sheet.getRange(item.rowNumber, 1, 1, 15).setValues([row]);
  PROGRAMME_SLOT_CACHE_ = null;
}

function saveProgrammeReview(token, actions) {
  const periodKey = validateProgrammeReviewToken_(token);
  if (!periodKey) throw new Error('This programme-review link is invalid or expired.');
  if (!Array.isArray(actions)) throw new Error('No programme changes were supplied.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    syncProgrammeSlots_();
    const data = getProgrammeReviewData(token);
    const allowedMeetings = new Set(data.meetings.map(m => Number(m.meetingNo)));
    const messages = [];

    actions.forEach(action => {
      const meetingNo = Number(action.meetingNo);
      const slot = Number(action.slot);
      const verb = String(action.action || 'KEEP').toUpperCase();
      if (!allowedMeetings.has(meetingNo) || [1,2,3].indexOf(slot) === -1) return;

      const meeting = data.meetings.find(m => Number(m.meetingNo) === meetingNo);
      const current = meeting.slots.find(s => Number(s.slot) === slot);
      if (!current) return;

      if (verb === 'HOLD') {
        if (current.speaker) throw new Error('Meeting ' + meetingNo + ' S' + slot + ' already has a confirmed speaker.');
        updateProgrammeSlotRecord_(meetingNo, slot, {
          state:'VPE_HOLD', sessionType:'', presenter:'', sessionTitle:'', minutes:'',
          notes:'Held by VPE for possible education/special session.'
        });
        messages.push('Meeting ' + meetingNo + ' S' + slot + ' held.');
      } else if (verb === 'REQUEST_MOVE') {
        if (!current.speaker) throw new Error('Meeting ' + meetingNo + ' S' + slot + ' has no confirmed speaker to move.');
        const sent = requestSpeakerMove_(meetingNo, slot, current.speaker);
        if (!sent.ok) throw new Error(
          'Could not send move request to ' + current.speaker + ': ' + (sent.error || 'send failed')
        );
        messages.push('Move request sent to ' + current.speaker + '.');
      } else if (verb === 'RELEASE') {
        if (current.speaker) throw new Error('A confirmed speaker slot cannot be released directly.');
        updateProgrammeSlotRecord_(meetingNo, slot, {
          state:'OPEN', sessionType:'', presenter:'', sessionTitle:'', minutes:'',
          moveRequestedAt:'', previousSpeaker:'', moveTargetMeeting:'',
          notes:'Released explicitly by VPE.'
        });
        messages.push('Meeting ' + meetingNo + ' S' + slot + ' released.');
      } else if (verb === 'CONFIRM_SESSION') {
        if (current.speaker) throw new Error('Move the confirmed speaker before confirming a session in this slot.');
        const type = String(action.sessionType || '').trim();
        const presenter = String(action.presenter || '').trim();
        const title = String(action.sessionTitle || '').trim();
        const minutes = Number(action.minutes || 0);
        if (!type || !presenter || !title || !minutes || minutes < 1) {
          throw new Error('Meeting ' + meetingNo + ' S' + slot +
            ': session type, presenter, title and time allotted are required.');
        }
        updateProgrammeSlotRecord_(meetingNo, slot, {
          state:'SESSION_CONFIRMED', sessionType:type, presenter:presenter,
          sessionTitle:title, minutes:minutes,
          notes:'Session confirmed by VPE; no presenter confirmation required.'
        });
        messages.push('Meeting ' + meetingNo + ' S' + slot + ' session confirmed.');
      }
    });

    markProgrammeReviewCompleted_(
      periodKey, configValue_('VPE_MEMBER_NAME', 'Riyas AF'),
      messages.length ? 'REVIEWED_WITH_CHANGES' : 'REVIEWED_NO_CHANGES'
    );
    syncProgrammeSlots_();
    if (configBool_('TASK_ENGINE_ENABLED')) syncPhase1Tasks_();
    return {ok:true, messages:messages};
  } finally {
    lock.releaseLock();
  }
}

function requestSpeakerMove_(meetingNo, slot, speaker) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  const ids = planner.getRange(3, 1, Math.max(0, planner.getLastRow() - 2), 1)
    .getValues().flat();
  const idx = ids.findIndex(v => Number(v) === Number(meetingNo));
  if (idx < 0) return {ok:false,error:'Meeting not found.'};
  const row = planner.getRange(idx + 3, 1, 1, 28).getValues()[0];
  const current = String(row[[6,10,14][slot - 1]] || '').trim();
  if (current !== speaker) return {ok:false,error:'Speaker assignment changed.'};

  const phone = phoneForMember_(speaker);
  if (!phone || !isWhatsAppRecipientAllowed_(phone)) {
    return {ok:false,error:'Speaker is not an allowed WhatsApp recipient in current mode.'};
  }

  const dateText = Utilities.formatDate(row[1], configValue_('TIMEZONE','Asia/Kolkata'), 'EEE, d MMM yyyy');
  const buttonId = ['MOVE_SPEECH_DATES', meetingNo, 'S' + slot].join('|');
  const body =
    'Hi ' + speaker + ',\n\n' +
    'ATK is planning a programme adjustment for Meeting ' + meetingNo + ' on ' + dateText + '.\n\n' +
    'Would you be comfortable moving your prepared speech to another available meeting? ' +
    'Your current slot remains confirmed until you choose another date.\n\n' +
    'Tap below to see the next available dates.';

  const payload = {
    messaging_product:'whatsapp', recipient_type:'individual', to:phone,
    type:'interactive',
    interactive:{
      type:'button', body:{text:body},
      action:{buttons:[{type:'reply',reply:{id:buttonId,title:'View dates'}}]}
    }
  };
  const key = ['SPEAKER_MOVE_REQUEST', meetingNo, 'S' + slot, speaker].join('|');
  const result = sendAutomatedQuickReply_(
    phone, 'SPEAKER_MOVE_REQUEST', payload,
    [speaker, meetingNo, dateText], buttonId
  );
  recordWorkflowOutbound_(speaker, phone, 'SPEAKER_MOVE_REQUEST', body, meetingNo, result, key);
  if (result.ok) {
    updateProgrammeSlotRecord_(meetingNo, slot, {
      state:'MOVE_REQUESTED', moveRequestedAt:new Date(), previousSpeaker:speaker,
      notes:'Current speech slot remains protected until speaker selects a new date.'
    });
  }
  return result;
}

function handleSpeakerMoveDateMenu_(member, phone, interaction) {
  const oldMeeting = Number(interaction.meetingNo || 0);
  const slot = Number(String(interaction.slotRef || '').replace(/^S/i,''));
  syncProgrammeSlots_();
  const item = readProgrammeSlots_()[programmeSlotKey_(oldMeeting, slot)];
  if (!item || item.state !== 'MOVE_REQUESTED' || item.speaker !== member) {
    sendWhatsAppText_(phone, 'This move request is no longer active. Your current booking remains unchanged.');
    return;
  }

  const meetings = availableSpeechMeetings_().filter(meeting => {
    if (Number(meeting.meetingNo) === oldMeeting) return false;
    const row = meeting.row;
    return ![row[6],row[10],row[14]].some(v => String(v || '').trim() === member) &&
      !ROLE_DEFS_.some(def => String(row[def.column - 1] || '').trim() === member);
  }).slice(0, 10);

  if (!meetings.length) {
    sendWhatsAppText_(phone,
      'There are no suitable open speech dates in the current booking window. Your existing slot remains confirmed.');
    return;
  }

  const result = sendWhatsAppPayload_({
    messaging_product:'whatsapp', recipient_type:'individual', to:phone,
    type:'interactive',
    interactive:{
      type:'list',
      body:{text:'Choose a new date. Your current speech slot will remain confirmed until this move succeeds.'},
      action:{button:'View dates',sections:[{
        title:'Available speech dates',
        rows:meetings.map(meeting => ({
          id:['MOVE_SPEECH',oldMeeting,'S'+slot,meeting.meetingNo].join('|'),
          title:Utilities.formatDate(meeting.date,configValue_('TIMEZONE','Asia/Kolkata'),'EEE, d MMM yyyy')
        }))
      }]}
    }
  });
  logAutomation_('Programme Planning', oldMeeting, member, 'Show speaker move dates',
    result.ok ? 'Sent' : 'Error', result.error || '');
}

function handleSpeakerMoveSelection_(member, phone, interaction) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const result = moveProtectedSpeaker_(
      Number(interaction.oldMeetingNo || 0),
      Number(String(interaction.slotRef || '').replace(/^S/i,'')),
      Number(interaction.newMeetingNo || 0),
      member
    );
    if (!result.ok) {
      sendWhatsAppText_(phone, result.error || 'That move could not be completed. Your original slot remains confirmed.');
      return;
    }
    sendWhatsAppText_(phone,
      'Moved successfully ✅\nYour prepared speech is now reserved for ' + result.newDateText +
      '.\n\nYour previous slot has been released to the VPE for the planned session.');
  } finally {
    lock.releaseLock();
  }
}

function moveProtectedSpeaker_(oldMeetingNo, oldSlot, newMeetingNo, member) {
  syncProgrammeSlots_();
  const oldItem = readProgrammeSlots_()[programmeSlotKey_(oldMeetingNo, oldSlot)];
  if (!oldItem || oldItem.state !== 'MOVE_REQUESTED' || oldItem.speaker !== member) {
    return {ok:false,error:'The original move request is no longer active.'};
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  const rows = planner.getRange(3,1,Math.max(0, planner.getLastRow() - 2),28).getValues();
  const oldIdx = rows.findIndex(r => Number(r[0]) === oldMeetingNo);
  const newIdx = rows.findIndex(r => Number(r[0]) === newMeetingNo);
  if (oldIdx < 0 || newIdx < 0) return {ok:false,error:'One of the meeting dates is no longer available.'};

  const oldRow = rows[oldIdx];
  const newRow = rows[newIdx];
  if (String(oldRow[[6,10,14][oldSlot - 1]] || '').trim() !== member) {
    return {ok:false,error:'Your original speech assignment has changed.'};
  }

  if ([newRow[6],newRow[10],newRow[14]].some(v => String(v || '').trim() === member) ||
      ROLE_DEFS_.some(def => String(newRow[def.column - 1] || '').trim() === member)) {
    return {ok:false,error:'You already have an assignment on that meeting date.'};
  }

  const openSlots = getOpenSpeechSlotsFromRow_(newRow);
  if (!openSlots.length) return {ok:false,error:'That meeting has just filled. Please choose another date.'};
  const newSlot = openSlots[0];

  const oldSheetRow = oldIdx + 3;
  const newSheetRow = newIdx + 3;
  const speakerCols = [7,11,15];
  const speechStatusCols = [9,13,17];
  const evaluatorCols = [26,27,28];
  const oldEvaluator = String(oldRow[evaluatorCols[oldSlot - 1] - 1] || '').trim();

  planner.getRange(newSheetRow, speakerCols[newSlot - 1]).setValue(member);
  planner.getRange(newSheetRow, speechStatusCols[newSlot - 1]).setValue('Planned');
  planner.getRange(oldSheetRow, speakerCols[oldSlot - 1]).clearContent();
  planner.getRange(oldSheetRow, speechStatusCols[oldSlot - 1]).clearContent();
  if (oldEvaluator) planner.getRange(oldSheetRow, evaluatorCols[oldSlot - 1]).clearContent();
  SpreadsheetApp.flush();

  updateProgrammeSlotRecord_(oldMeetingNo, oldSlot, {
    state:'VPE_HOLD', speaker:'', moveTargetMeeting:newMeetingNo,
    notes:'Speaker moved to Meeting ' + newMeetingNo + '; slot now held for VPE session planning.'
  });
  syncProgrammeSlots_();

  const newItem = readProgrammeSlots_()[programmeSlotKey_(newMeetingNo, newSlot)];
  if (newItem) {
    updateProgrammeSlotRecord_(newMeetingNo, newSlot, {
      state:'SPEAKER_CONFIRMED', speaker:member,
      notes:'Speaker moved from Meeting ' + oldMeetingNo + ' S' + oldSlot + '.'
    });
  }

  if (configBool_('TASK_ENGINE_ENABLED')) syncPhase1Tasks_();

  logAutomation_(
    'Programme Planning', oldMeetingNo, member,
    'Protected speaker moved to Meeting ' + newMeetingNo + ' S' + newSlot,
    'Success', oldEvaluator ? 'Previous evaluator cleared: ' + oldEvaluator : ''
  );

  const newDateText = Utilities.formatDate(
    newRow[1], configValue_('TIMEZONE','Asia/Kolkata'), 'EEE, d MMM yyyy'
  );
  return {ok:true,newMeetingNo:newMeetingNo,newSlot:newSlot,newDateText:newDateText};
}

function programmeSlotIssuesForMeeting_(meetingNo) {
  const slotMap = readProgrammeSlots_();
  const issues = [];
  [1,2,3].forEach(slot => {
    const item = slotMap[programmeSlotKey_(meetingNo, slot)];
    if (!item) return;
    if (item.state === 'VPE_HOLD') issues.push('S' + slot + ' programme hold awaiting confirmation');
    if (item.state === 'MOVE_REQUESTED') issues.push('S' + slot + ' speaker move awaiting response');
  });
  return issues;
}
