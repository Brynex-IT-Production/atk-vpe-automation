/**
 * One-shot DEV end-to-end verification.
 * Runs only in DEV + TEST, sends template smoke tests only to the existing
 * allowlisted VPE, suppresses later workflow outbound messages, restores the
 * temporary planner row/config, and writes one summary to Automation Log.
 */

var DEV_E2E_SUPPRESS_OUTBOUND_ = false;

function maybeRunDevE2eTest_() {
  if (configValue_('SYSTEM_MODE', '') !== 'DEV' ||
      configValue_('WA_SEND_MODE', '') !== 'TEST') return false;

  const runToken = 'ATK_E2E_20260930_02';
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('DEV_E2E_LAST_RUN') === runToken) return false;

  // Mark first so a scheduler retry cannot duplicate real template smoke sends.
  props.setProperty('DEV_E2E_LAST_RUN', runToken);

  const result = runDevE2eTest_(runToken);
  props.setProperty('DEV_E2E_LAST_RESULT', JSON.stringify(result));
  return true;
}

function runDevE2eTest_(runToken) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  const config = ss.getSheetByName('Automation Config');
  if (!planner || !config) throw new Error('DEV E2E required sheets are missing.');

  const member = configValue_('VPE_MEMBER_NAME', 'Riyas AF');
  const phone = normalizePhone_(phoneForMember_(member));
  if (!member || !phone || !isWhatsAppRecipientAllowed_(phone)) {
    throw new Error('DEV E2E requires the VPE to be the allowlisted test recipient.');
  }

  const rowNo = 62;
  const rowRange = planner.getRange(rowNo, 1, 1, 28);
  const originalValues = rowRange.getValues()[0];
  const originalFormulas = rowRange.getFormulas()[0];
  const configBackup = {};
  const checks = [];
  const templates = [];
  let fatal = '';

  function check(name, ok, detail) {
    checks.push({name: name, ok: Boolean(ok), detail: String(detail || '')});
  }

  try {
    devE2eCleanupArtifacts_();
    clearConversation_(phone);

    // 1) Real Meta template smoke test. This is the only part that actually
    // sends WhatsApp messages. All later workflow sends are suppressed.
    const templateCases = [
      {
        label: 'Role reminder v2',
        kind: 'ROLE_REMINDER',
        params: ['9901'],
        buttons: 'ROLE_MENU'
      },
      {
        label: 'Speech invite v2',
        kind: 'SPEECH_INVITE',
        params: [member],
        buttons: 'SPEECH_DATES'
      },
      {
        label: 'Theme request v2',
        kind: 'THEME_REQUEST',
        params: [member, '9902']
      },
      {
        label: 'Theme reminder v3',
        kind: 'THEME_REMINDER',
        params: [member, '9902']
      },
      {
        label: 'Speaker evaluator pairing v2',
        kind: 'EVALUATOR_PAIRING_SPEAKER',
        params: [member, member, '9904']
      },
      {
        label: 'Evaluator assignment v2',
        kind: 'EVALUATOR_ASSIGNMENT',
        params: [member, member, '9904']
      },
      {
        label: 'Speaker coordination check v2',
        kind: 'EVALUATOR_COORD_CHECK_SPEAKER',
        params: [member, member, '9904'],
        buttons: [
          'EVAL_COORD|SPEAKER|9904|YES',
          'EVAL_COORD|SPEAKER|9904|NO'
        ]
      },
      {
        label: 'Evaluator coordination check v2',
        kind: 'EVALUATOR_COORD_CHECK_EVALUATOR',
        params: [member, member, '9904'],
        buttons: [
          'EVAL_COORD|EVALUATOR|9904|YES',
          'EVAL_COORD|EVALUATOR|9904|NO'
        ]
      },
      {
        label: 'Speaker details current active template',
        kind: 'SPEAKER_DETAILS',
        params: [member, '9904']
      },
      {
        label: 'Role-player introduction',
        kind: 'INTRO_REQUEST',
        params: [member, '9901', 'Timer']
      },
      {
        label: 'Introduction bundle',
        kind: 'INTRO_BUNDLE',
        params: [member, '9901', 'Timer — ' + member + '\nDEV E2E introduction']
      },
      {
        label: 'TMOD update',
        kind: 'TMOD_UPDATE',
        params: [member, '9904', 'Speaker 1 — ' + member + '\nTitle: DEV E2E']
      }
    ];

    let previousResult = {};
    try {
      previousResult = JSON.parse(
        PropertiesService.getScriptProperties().getProperty('DEV_E2E_LAST_RESULT') || '{}'
      );
    } catch (_) {}
    const previousTemplates = previousResult.templates || [];

    templateCases.forEach(function(test, index) {
      const templateName = templateNameForKind_(test.kind);
      const approved = Boolean(templateName && isApprovedTemplate_(templateName));
      const previous = previousTemplates.find(function(t) {
        return t.kind === test.kind && t.template === templateName && t.ok;
      });
      let send = {ok: false, error: 'Not approved'};
      let reused = false;

      if (approved && previous) {
        send = {ok: true};
        reused = true;
      } else if (approved) {
        if (index) Utilities.sleep(450);
        send = sendWhatsAppTemplate_(
          phone,
          templateName,
          test.params,
          test.buttons || null
        );
      }

      templates.push({
        label: test.label,
        kind: test.kind,
        template: templateName,
        approved: approved,
        ok: Boolean(send && send.ok),
        reusedPreviousPass: reused,
        error: send && send.error ? String(send.error) : ''
      });
    });

    check(
      'Approved-template Meta smoke test',
      templates.every(function(t) { return t.approved && t.ok; }),
      templates.map(function(t) {
        return t.kind + '=' + t.template + ':' + (t.ok ? 'OK' : 'FAIL');
      }).join(', ')
    );

    // From here onward, exercise the workflow/state machinery without sending
    // more WhatsApp messages to the user.
    DEV_E2E_SUPPRESS_OUTBOUND_ = true;

    // 2) Role booking + meeting-specific intro + consolidated TMOD bundle.
    devE2ePrepareMeeting_(planner, rowNo, 9901, 5);
    syncPhase1Tasks_();
    devE2eInjectButton_(phone, 'ROLE|9901|Timer', 'Timer');
    check(
      'Role booking writes planner',
      String(planner.getRange(rowNo, 22).getValue() || '') === member,
      'Timer=' + planner.getRange(rowNo, 22).getValue()
    );

    if (defaultIntro_(member)) {
      devE2eInjectButton_(
        phone,
        'INTRO_CHOICE|9901|Timer|NEW',
        'I have something new'
      );
      devE2eInjectText_(
        phone,
        'DEV E2E meeting-specific introduction for ' + member + '.'
      );
    } else {
      handleIntroSubmission_(
        member,
        phone,
        9901,
        'DEV E2E meeting-specific introduction for ' + member + '.',
        'DEV_E2E_DIRECT_INTRO',
        'Timer'
      );
    }

    check(
      'Role-player intro saved',
      devE2eHasSubmission_(9901, member, 'INTRO', 'Timer'),
      'Timer intro submission'
    );

    planner.getRange(rowNo, 19).setValue(member);
    SpreadsheetApp.flush();
    const bundleSent = sendConsolidatedIntroductionsIfReady_(9901);
    const bundleText = devE2eOutboxTextByPrefix_('INTRO_BUNDLE|9901|' + member);
    check(
      'TMOD intro bundle labels role + member',
      bundleSent && bundleText.indexOf('Timer — ' + member) !== -1,
      bundleText
    );

    // 3) Six-day theme request -> inbound theme -> planner + distribution.
    devE2ePrepareMeeting_(planner, rowNo, 9902, 6);
    planner.getRange(rowNo, 19).setValue(member); // TMOD only while request starts
    SpreadsheetApp.flush();
    syncPhase1Tasks_();
    clearConversation_(phone);
    processAssignedRolePrompts_();

    check(
      'Six-day theme request task sends',
      devE2eHasOutboxPrefix_('THEME_REQUEST|9902|TMOD|' + member),
      'Theme request outbox'
    );

    // Add a role recipient only after the theme conversation is active, so the
    // same DEV member does not trigger the earlier 7-day intro workflow first.
    planner.getRange(rowNo, 22).setValue(member);
    SpreadsheetApp.flush();

    devE2eInjectText_(phone, 'The Power of Small Wins');
    check(
      'Theme saved to planner',
      String(planner.getRange(rowNo, 4).getValue() || '') === 'The Power of Small Wins',
      String(planner.getRange(rowNo, 4).getValue() || '')
    );
    check(
      'Theme submission recorded',
      devE2eHasSubmission_(9902, member, 'THEME', 'TMOD'),
      'THEME/TMOD'
    );
    check(
      'Theme distributed to assigned role',
      devE2eHasOutboxPrefix_('THEME_TO_ROLE|9902|' + member + '|Timer|'),
      'Timer received theme handoff'
    );

    // 4) Grammarian booking -> immediate request -> 7/8-style reminder ->
    // all-in-one reply -> completed task.
    devE2ePrepareMeeting_(planner, rowNo, 9903, 3);
    syncPhase1Tasks_();
    clearConversation_(phone);
    devE2eInjectButton_(phone, 'ROLE|9903|GRM', 'Grammarian');

    check(
      'Grammarian role booking',
      String(planner.getRange(rowNo, 24).getValue() || '') === member,
      'Grammarian=' + planner.getRange(rowNo, 24).getValue()
    );
    check(
      'Immediate Grammarian request',
      devE2eHasOutboxPrefix_('ASSIGNMENT_PROMPT|9903|GRM|' + member),
      'Immediate assignment prompt'
    );

    configBackup.GRAMMAR_REMINDER_HOURS =
      devE2eSetConfig_(config, 'GRAMMAR_REMINDER_HOURS', devE2eCurrentHour_());
    processGrammarAssignmentReminder_();
    check(
      'Grammarian repeated reminder',
      devE2eHasOutboxPrefix_('GRAMMAR_ASSIGNMENT_REMINDER|9903|GRM|' + member + '|'),
      'Current hour temporarily used for DEV test'
    );

    devE2eInjectText_(
      phone,
      'Word: Resilience\n' +
      'Word meaning: The ability to recover quickly\n' +
      'Idiom: Back on your feet\n' +
      'Idiom meaning: Recovered after a difficulty'
    );

    check(
      'Grammarian details saved',
      devE2eHasSubmission_(9903, member, 'GRAMMAR', 'GRM'),
      'All four fields parsed'
    );
    check(
      'Grammarian task completed',
      devE2eTaskState_(9903, 'GRAMMAR', member, 'GRM') === 'Completed',
      devE2eTaskState_(9903, 'GRAMMAR', member, 'GRM')
    );
    devE2eRestoreConfig_(config, 'GRAMMAR_REMINDER_HOURS',
      configBackup.GRAMMAR_REMINDER_HOURS);

    // 5) Speech slot booking -> speaker-detail request/reminder -> all-in-one
    // reply -> TMOD handoff + evaluator assignment + 2-day coordination checks.
    devE2ePrepareMeeting_(planner, rowNo, 9904, 6);
    syncPhase1Tasks_();
    clearConversation_(phone);

    const invite = sendSpeechSlotInvite_(
      {name: member, phone: phone},
      9904,
      planner.getRange(rowNo, 2).getValue(),
      3,
      'BROAD'
    );
    recordWorkflowOutbound_(
      member, phone, 'SPEECH_INVITE', invite.body || '',
      9904, invite, 'DEV_E2E_SPEECH_INVITE|9904|' + member
    );
    check('Speech invitation workflow send', Boolean(invite && invite.ok), 'BROAD');

    devE2eInjectButton_(phone, 'SPEAK|9904', 'Reserve speech slot');
    check(
      'Speech slot reserved',
      String(planner.getRange(rowNo, 7).getValue() || '') === member,
      'Speaker 1=' + planner.getRange(rowNo, 7).getValue()
    );

    planner.getRange(rowNo, 19).setValue(member); // TMOD for handoff test
    planner.getRange(rowNo, 26).setValue(member); // Evaluator 1
    SpreadsheetApp.flush();
    syncPhase1Tasks_();
    clearConversation_(phone);
    processSpeakerDetailPrompts_();

    const speakerPromptKey = 'SPEAKER_DETAILS_PROMPT|9904|S1|' + member;
    check(
      'Six-day speaker-details request',
      devE2eHasOutboxPrefix_(speakerPromptKey),
      'Active=' + templateNameForKind_('SPEAKER_DETAILS')
    );

    configBackup.SPEAKER_DETAILS_REMINDER_HOURS =
      devE2eSetConfig_(config, 'SPEAKER_DETAILS_REMINDER_HOURS', devE2eCurrentHour_());
    devE2eBackdateOutbox_(speakerPromptKey, 2);
    processSpeakerDetailReminders_();
    check(
      'Speaker-details repeated reminder',
      devE2eHasOutboxPrefix_('SPEAKER_DETAILS_REMINDER|9904|S1|' + member + '|'),
      'Prompt backdated 2h for DEV test'
    );

    devE2eInjectText_(
      phone,
      'Path: Presentation Mastery\n' +
      'Level: 2\n' +
      'Project: Understanding Your Communication Style\n' +
      'Title: The Test Run'
    );

    check(
      'All-in-one speaker-details parser',
      devE2eHasSubmission_(9904, member, 'SPEAKER_DETAILS', 'S1'),
      'Path/Level/Project/Title'
    );

    const handoff = devE2eOutboxTextByPrefix_(
      'SPEECH_TO_TMOD|9904|S1|' + member + '|'
    );
    check(
      'Speaker handoff identifies slot + name',
      handoff.indexOf('Speaker 1 — ' + member) !== -1,
      handoff
    );

    check(
      'Speaker evaluator assignment sent',
      devE2eHasOutboxPrefix_(
        'EVALUATOR_PAIR_SPEAKER|9904|S1|' + member + '|' + member
      ),
      'Speaker-side pairing'
    );
    check(
      'Evaluator speaker assignment sent',
      devE2eHasOutboxPrefix_(
        'EVALUATOR_PAIR_EVALUATOR|9904|S1|' + member + '|' + member
      ),
      'Evaluator-side assignment'
    );

    planner.getRange(rowNo, 2).setValue(devE2eDateAhead_(2));
    SpreadsheetApp.flush();
    processEvaluatorNotifications_();

    check(
      'Speaker Yes/No coordination check sent',
      devE2eHasOutboxPrefix_(
        'EVALUATOR_COORD_CHECK_SPEAKER|9904|S1|D2|' + member
      ),
      'D2'
    );
    check(
      'Evaluator Yes/No coordination check sent',
      devE2eHasOutboxPrefix_(
        'EVALUATOR_COORD_CHECK_EVALUATOR|9904|S1|D2|' + member
      ),
      'D2'
    );

    devE2eInjectButton_(
      phone,
      'EVAL_COORD|SPEAKER|9904|YES',
      'Yes'
    );
    devE2eInjectButton_(
      phone,
      'EVAL_COORD|EVALUATOR|9904|YES',
      'Yes'
    );
    check(
      'Coordination Yes replies parsed',
      devE2eCoordinationConfirmedCount_(9904) >= 2,
      'Confirmed=' + devE2eCoordinationConfirmedCount_(9904)
    );

    devE2eRestoreConfig_(
      config,
      'SPEAKER_DETAILS_REMINDER_HOURS',
      configBackup.SPEAKER_DETAILS_REMINDER_HOURS
    );

  } catch (error) {
    fatal = String(error && error.stack ? error.stack : error);
    check('Fatal execution', false, fatal);
  } finally {
    DEV_E2E_SUPPRESS_OUTBOUND_ = false;

    try {
      Object.keys(configBackup).forEach(function(key) {
        if (configBackup[key] !== undefined) {
          devE2eRestoreConfig_(config, key, configBackup[key]);
        }
      });
    } catch (_) {}

    try {
      devE2eRestorePlannerRow_(planner, rowNo, originalValues, originalFormulas);
    } catch (restoreError) {
      check('Planner row restored', false, String(restoreError));
    }

    try { clearConversation_(phone); } catch (_) {}
    try {
      devE2eCleanupWebhookRows_();
      syncPhase1Tasks_();
    } catch (cleanupError) {
      check('Test cleanup', false, String(cleanupError));
    }
  }

  const result = {
    runToken: runToken,
    finishedAt: new Date().toISOString(),
    templates: templates,
    checks: checks,
    fatal: fatal,
    passed: !fatal &&
      templates.every(function(t) { return t.approved && t.ok; }) &&
      checks.every(function(c) { return c.ok; })
  };

  logAutomation_(
    'DEV E2E',
    '',
    member,
    'End-to-end automation test',
    result.passed ? 'PASS' : 'FAIL',
    JSON.stringify(result)
  );

  return result;
}

function devE2ePrepareMeeting_(planner, rowNo, meetingNo, daysAhead) {
  clearConversation_(configValue_('VPE_MEMBER_NAME', 'Riyas AF'));
  planner.getRange(rowNo, 4, 1, 25).clearContent(); // D:AB only; preserve status formula C.
  planner.getRange(rowNo, 1).setValue(meetingNo);
  planner.getRange(rowNo, 2).setValue(devE2eDateAhead_(daysAhead));
  SpreadsheetApp.flush();
}

function devE2eDateAhead_(days) {
  const d = startOfDay_(new Date());
  d.setDate(d.getDate() + Number(days || 0));
  return d;
}

function devE2eCurrentHour_() {
  return Utilities.formatDate(
    new Date(),
    configValue_('TIMEZONE', 'Asia/Kolkata'),
    'H'
  );
}

function devE2eSetConfig_(sheet, key, value) {
  const last = sheet.getLastRow();
  const keys = sheet.getRange(1, 1, last, 1).getValues().flat();
  const idx = keys.findIndex(function(v) { return String(v || '').trim() === key; });
  if (idx < 0) throw new Error('Config key not found: ' + key);
  const cell = sheet.getRange(idx + 1, 2);
  const original = cell.getValue();
  cell.setValue(value);
  CONFIG_CACHE_ = null;
  SpreadsheetApp.flush();
  return original;
}

function devE2eRestoreConfig_(sheet, key, value) {
  const last = sheet.getLastRow();
  const keys = sheet.getRange(1, 1, last, 1).getValues().flat();
  const idx = keys.findIndex(function(v) { return String(v || '').trim() === key; });
  if (idx < 0) return;
  sheet.getRange(idx + 1, 2).setValue(value);
  CONFIG_CACHE_ = null;
  SpreadsheetApp.flush();
}

function devE2eRestorePlannerRow_(planner, rowNo, values, formulas) {
  for (let col = 1; col <= 28; col++) {
    const cell = planner.getRange(rowNo, col);
    const formula = formulas[col - 1];
    if (formula) cell.setFormula(formula);
    else cell.setValue(values[col - 1]);
  }
  SpreadsheetApp.flush();
}

function devE2eInjectText_(phone, text) {
  return devE2eInjectInbound_(phone, {
    type: 'text',
    text: {body: text}
  });
}

function devE2eInjectButton_(phone, id, title) {
  return devE2eInjectInbound_(phone, {
    type: 'interactive',
    interactive: {
      button_reply: {
        id: id,
        title: title || id
      }
    }
  });
}

function devE2eInjectInbound_(phone, message) {
  const id = 'wamid.DEV_E2E_' + Utilities.getUuid();
  message.id = id;
  message.from = phone;
  const payload = {
    entry: [{
      changes: [{
        value: {
          messages: [message]
        }
      }]
    }]
  };
  logWhatsAppWebhook_(payload);
  processWhatsAppAutomation_(payload);
  return id;
}

function devE2eHasSubmission_(meetingNo, member, type, related) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('WA Submissions');
  if (!sheet || sheet.getLastRow() < 2) return false;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
  return rows.some(function(r) {
    return Number(r[1]) === Number(meetingNo) &&
      String(r[2] || '').trim() === member &&
      String(r[3] || '') === type &&
      String(r[4] || '') === related;
  });
}

function devE2eHasOutboxPrefix_(prefix) {
  return Boolean(devE2eFindOutboxRow_(prefix));
}

function devE2eOutboxTextByPrefix_(prefix) {
  const row = devE2eFindOutboxRow_(prefix);
  return row ? String(row[5] || '') : '';
}

function devE2eFindOutboxRow_(prefix) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('WA Outbox');
  if (!sheet || sheet.getLastRow() < 2) return null;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).getValues();
  for (let i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][15] || '').indexOf(prefix) === 0) return rows[i];
  }
  return null;
}

function devE2eTaskState_(meetingNo, type, member, related) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Automation Tasks');
  if (!sheet || sheet.getLastRow() < 2) return '';
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 18).getValues();
  const row = rows.find(function(r) {
    return Number(r[1]) === Number(meetingNo) &&
      String(r[2] || '') === type &&
      String(r[3] || '').trim() === member &&
      String(r[4] || '') === related;
  });
  return row ? String(row[5] || '') : '';
}

function devE2eBackdateOutbox_(prefix, hours) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('WA Outbox');
  if (!sheet || sheet.getLastRow() < 2) return false;
  const keys = sheet.getRange(2, 16, sheet.getLastRow() - 1, 1).getValues().flat();
  for (let i = keys.length - 1; i >= 0; i--) {
    if (String(keys[i] || '').indexOf(prefix) === 0) {
      sheet.getRange(i + 2, 13).setValue(
        new Date(Date.now() - Number(hours || 0) * 3600000)
      );
      SpreadsheetApp.flush();
      return true;
    }
  }
  return false;
}

function devE2eCoordinationConfirmedCount_(meetingNo) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Automation Log');
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  return rows.filter(function(r) {
    return Number(r[2]) === Number(meetingNo) &&
      String(r[1] || '') === 'Evaluation Coordination' &&
      String(r[5] || '') === 'Confirmed';
  }).length;
}

function devE2eCleanupWebhookRows_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('WA Webhook Log');
  if (!sheet || sheet.getLastRow() < 2) return;
  const ids = sheet.getRange(2, 5, sheet.getLastRow() - 1, 1).getValues().flat();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i] || '').indexOf('wamid.DEV_E2E_') === 0) {
      sheet.deleteRow(i + 2);
    }
  }
}

function devE2eCleanupArtifacts_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const submissions = ss.getSheetByName('WA Submissions');
  if (submissions && submissions.getLastRow() >= 2) {
    const nums = submissions.getRange(2, 2, submissions.getLastRow() - 1, 1).getValues().flat();
    for (let i = nums.length - 1; i >= 0; i--) {
      const n = Number(nums[i]);
      if (n >= 9900 && n <= 9999) submissions.deleteRow(i + 2);
    }
  }

  const outbox = ss.getSheetByName('WA Outbox');
  if (outbox && outbox.getLastRow() >= 2) {
    const nums = outbox.getRange(2, 11, outbox.getLastRow() - 1, 1).getValues().flat();
    for (let i = nums.length - 1; i >= 0; i--) {
      const n = Number(nums[i]);
      if (n >= 9900 && n <= 9999) outbox.deleteRow(i + 2);
    }
  }

  const automation = ss.getSheetByName('Automation Log');
  if (automation && automation.getLastRow() >= 2) {
    const nums = automation.getRange(2, 3, automation.getLastRow() - 1, 1).getValues().flat();
    for (let i = nums.length - 1; i >= 0; i--) {
      const n = Number(nums[i]);
      if (n >= 9900 && n <= 9999) automation.deleteRow(i + 2);
    }
  }

  devE2eCleanupWebhookRows_();
}
