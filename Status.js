/**
 * Exception-only VPE visibility and Control Center readiness.
 */

function updateControlCenterReadiness_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  const tasks = ss.getSheetByName('Automation Tasks');
  const control = ss.getSheetByName('VPE Control Center');
  if (!planner || !tasks || !control) return;

  const plannerRows = planner.getRange(
    3,
    1,
    Math.max(0, planner.getLastRow() - 2),
    28
  ).getValues();

  const taskRows = tasks.getLastRow() >= 2
    ? tasks.getRange(2, 1, tasks.getLastRow() - 1, 18).getValues()
    : [];

  const today = startOfDay_(new Date());
  const upcoming = plannerRows
    .filter(row => {
      const date = row[1];
      const status = String(row[2] || '');
      return Number(row[0]) &&
        date instanceof Date &&
        startOfDay_(date) >= today &&
        status !== 'Completed' &&
        status !== 'Cancelled';
    })
    .slice(0, 4);

  const rows = upcoming.map(row => buildReadinessRow_(row, taskRows));

  const target = control.getRange(55, 1, 4, 10);
  target.clearContent();

  // Keep ratio strings such as 1/3 and 3/3 as text. Without this,
  // Google Sheets can interpret them as dates.
  control.getRange(55, 2, 4, 8).setNumberFormat('@');

  if (rows.length) {
    control.getRange(55, 1, rows.length, 10).setValues(rows);
  }
}

function buildReadinessRow_(row, taskRows) {
  const meetingNo = Number(row[0]);
  const roles = ROLE_DEFS_.filter(def => String(row[def.column - 1] || '').trim());
  const speakers = [row[6], row[10], row[14]].map(v => String(v || '').trim()).filter(Boolean);
  const evaluators = [row[25], row[26], row[27]].map(v => String(v || '').trim()).filter(Boolean);
  const programmeItems = [1,2,3].map(slot =>
    readProgrammeSlots_()[programmeSlotKey_(meetingNo, slot)] || {state:'OPEN'}
  );
  const confirmedSessions = programmeItems.filter(item => item.state === 'SESSION_CONFIRMED').length;
  const pendingProgrammeHolds = programmeItems.filter(item =>
    item.state === 'VPE_HOLD' || item.state === 'MOVE_REQUESTED'
  ).length;
  const programmeFilled = speakers.length + confirmedSessions;

  const meetingTasks = taskRows.filter(r => Number(r[1]) === meetingNo);
  const countCompleted = type => meetingTasks.filter(r =>
    String(r[2] || '') === type && String(r[5] || '') === 'Completed'
  ).length;

  const tmod = String(row[18] || '').trim();
  const grammar = String(row[23] || '').trim();
  const themeStatus = !tmod ? 'Waiting TMOD' : (String(row[3] || '').trim() ? 'Received' : 'Pending');
  const grammarStatus = !grammar
    ? 'Waiting GRM'
    : (countCompleted('GRAMMAR') >= 1 ? 'Received' : 'Pending');

  const introAssigned = roles.filter(x => x.code !== 'TMOD').length;
  const introCompleted = countCompleted('INTRO');
  const speakerDetailsCompleted = countCompleted('SPEAKER_DETAILS');
  const coordinationCompleted = countCompleted('EVALUATOR_COORDINATION');

  const fullyReady =
    ROLE_DEFS_.filter(def => !def.evaluatorSlot)
      .every(def => String(row[def.column - 1] || '').trim()) &&
    programmeFilled === 3 &&
    pendingProgrammeHolds === 0 &&
    themeStatus === 'Received' &&
    grammarStatus === 'Received' &&
    introCompleted >= introAssigned &&
    speakerDetailsCompleted >= speakers.length &&
    evaluators.length >= speakers.length &&
    coordinationCompleted >= evaluators.length;

  return [
    meetingNo,
    roles.length + ' of 6',
    programmeFilled + ' of 3',
    themeStatus,
    grammarStatus,
    introCompleted + ' of ' + introAssigned,
    speakerDetailsCompleted + ' of ' + speakers.length,
    evaluators.length + ' of ' + speakers.length,
    coordinationCompleted + ' of ' + evaluators.length,
    fullyReady ? 'Ready' : 'Action needed'
  ];
}

function buildVpeExceptionDigest_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  const tasks = ss.getSheetByName('Automation Tasks');
  if (!planner || !tasks) return '';

  const today = startOfDay_(new Date());
  const windowDays = Math.max(1, Number(configValue_('VPE_EXCEPTION_WINDOW_DAYS', '7')) || 7);
  const plannerRows = planner.getRange(
    3,
    1,
    Math.max(0, planner.getLastRow() - 2),
    28
  ).getValues();
  const taskRows = tasks.getLastRow() >= 2
    ? tasks.getRange(2, 1, tasks.getLastRow() - 1, 18).getValues()
    : [];

  const sections = [];

  plannerRows.forEach(row => {
    const meetingNo = Number(row[0]);
    const date = row[1];
    const status = String(row[2] || '');

    if (!meetingNo || !(date instanceof Date)) return;
    if (status === 'Completed' || status === 'Cancelled') return;

    const daysLeft = Math.round((startOfDay_(date) - today) / 86400000);
    if (daysLeft < 0 || daysLeft > windowDays) return;

    const issues = [];
    const openRoles = ROLE_DEFS_
      .filter(def => {
        if (String(row[def.column - 1] || '').trim()) return false;
        if (!def.evaluatorSlot) return true;
        return Boolean(String(row[[6,10,14][def.evaluatorSlot - 1]] || '').trim());
      })
      .map(def => def.name);

    if (openRoles.length) {
      issues.push('Open roles: ' + openRoles.join(', '));
    }

    const speakers = [row[6], row[10], row[14]].map(v => String(v || '').trim());
    const openSpeechCount = getOpenSpeechSlotsFromRow_(row).length;
    if (openSpeechCount) issues.push(openSpeechCount + ' prepared-speech slot' + (openSpeechCount === 1 ? '' : 's') + ' open');

    programmeSlotIssuesForMeeting_(meetingNo).forEach(issue => issues.push(issue));

    const tmod = String(row[18] || '').trim();
    if (tmod && !String(row[3] || '').trim()) issues.push('Theme pending from ' + tmod);

    const meetingTasks = taskRows.filter(r => Number(r[1]) === meetingNo);
    const pendingByType = type => meetingTasks.filter(r =>
      String(r[2] || '') === type &&
      ['Pending', 'Sent', 'Escalated'].indexOf(String(r[5] || '')) !== -1
    );

    const grammarPending = pendingByType('GRAMMAR');
    if (grammarPending.length) issues.push('Grammarian Word/Idiom details pending');

    const introPending = pendingByType('INTRO');
    if (introPending.length) issues.push(introPending.length + ' role-player introduction' + (introPending.length === 1 ? '' : 's') + ' pending');

    const detailsPending = pendingByType('SPEAKER_DETAILS');
    if (detailsPending.length) issues.push(detailsPending.length + ' speaker-detail submission' + (detailsPending.length === 1 ? '' : 's') + ' pending');

    const evaluatorPending = pendingByType('EVALUATOR_ASSIGNMENT');
    if (evaluatorPending.length) {
      const slots = evaluatorPending.map(r => String(r[4] || '')).filter(Boolean);
      issues.push('Evaluator not assigned for ' + slots.join(', '));
    }

    const coordinationPending = pendingByType('EVALUATOR_COORDINATION');
    if (coordinationPending.length) {
      issues.push(coordinationPending.length + ' speaker-evaluator pairing notification' + (coordinationPending.length === 1 ? '' : 's') + ' pending');
    }

    if (!issues.length) return;

    const dateText = Utilities.formatDate(date, configValue_('TIMEZONE', 'Asia/Kolkata'), 'EEE, d MMM');
    sections.push(
      'Meeting ' + meetingNo + ' · ' + dateText + '\n' +
      issues.map(x => '• ' + x).join('\n')
    );
  });

  if (!sections.length) return '';
  return 'ATK VPE — Action needed\n\n' + sections.join('\n\n');
}

function processVpeExceptionDigest_() {
  if (!configBool_('VPE_EXCEPTION_DIGEST_ENABLED')) return;

  const digest = buildVpeExceptionDigest_();
  if (!digest) return;

  const member = configValue_('VPE_MEMBER_NAME', 'Riyas AF');
  const phone = phoneForMember_(member);
  if (!phone || !isWhatsAppRecipientAllowed_(phone)) return;

  const hashBytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    digest,
    Utilities.Charset.UTF_8
  );
  const hash = hashBytes.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('').slice(0, 16);
  const key = 'VPE_EXCEPTION_DIGEST|' + hash;

  if (hasReminderBeenSent_(key)) return;

  let result;
  if (hasOpenWhatsAppServiceWindow_(phone)) {
    result = sendWhatsAppText_(phone, digest);
  } else {
    const templateName = templateNameForKind_('VPE_EXCEPTION');
    if (!templateName) return;

    result = sendWhatsAppTemplate_(
      phone,
      templateName,
      [member],
      'ADMIN_EXCEPTIONS'
    );
  }
  recordWorkflowOutbound_(member, phone, 'VPE_EXCEPTION_DIGEST', digest, '', result, key);
  logAutomation_(
    'VPE Exception',
    '',
    member,
    'Exception digest',
    result.ok ? 'Sent' : (result.blocked ? 'Blocked' : 'Error'),
    hash
  );
}
