/**
 * One-time Meta WhatsApp template setup for ATK Phase 1.
 * Safe to rerun: existing templates are read and skipped.
 */

function createAtkWhatsAppTemplates_() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('META_ACCESS_TOKEN');
  const wabaId = props.getProperty('WA_WABA_ID');

  if (!token) throw new Error('META_ACCESS_TOKEN is missing.');
  if (!wabaId) throw new Error('WA_WABA_ID is missing.');

  const specs = getAtkWhatsAppTemplateSpecs_();
  const results = specs.map(spec => createOrReadTemplate_(token, wabaId, spec));

  return {
    ok: results.every(r => r.ok),
    results: results
  };
}

function getAtkWhatsAppTemplateSpecs_() {
  return [
    roleReminderV2Spec_(),
    speechInviteV2Spec_(),
    themeRequestV2Spec_(),
    themeReminderV2Spec_(),
    speakerEvaluatorPairingV2Spec_(),
    evaluatorAssignmentV2Spec_(),
    speakerCoordinationCheckV2Spec_(),
    evaluatorCoordinationCheckV2Spec_(),
    speakerDetailsV3Spec_(),
    {
      name: 'atk_role_reminder',
      language: 'en_US',
      category: 'UTILITY',
      body:
        'Hi {{1}}, ATK Meeting {{2}} still has these meeting roles open: {{3}}.\n\n' +
        'Reply ROLE to view the currently available roles and book one.',
      examples: ['Riyas AF', '450', 'Timer, Grammarian, Ah-Counter']
    },
    {
      name: 'atk_speech_invite',
      language: 'en_US',
      category: 'UTILITY',
      body:
        'Hi {{1}}, ATK Meeting {{2}} on {{3}} has {{4}} prepared-speech slot(s) open.\n\n' +
        'If you would like to take a slot, use the button below.',
      examples: ['Riyas AF', '450', 'Mon, 28 Sep', '2'],
      buttons: [{ type: 'QUICK_REPLY', text: 'Reserve speech slot' }]
    },
    {
      name: 'atk_evaluator_pairing',
      language: 'en_US',
      category: 'UTILITY',
      body:
        'Hi {{1}}, you are paired with {{2}} for ATK Meeting {{3}}. ' +
        'Please connect directly and coordinate the project and evaluation requirements before the meeting.',
      examples: ['Anil Mathew', 'Antu Dominic', '450']
    },
    {
      name: 'atk_evaluator_coord_reminder',
      language: 'en_US',
      category: 'UTILITY',
      body:
        'Hi {{1}}, reminder for ATK Meeting {{2}}: please connect with {{3}} ' +
        'and complete your speaker-evaluator coordination before the meeting.',
      examples: ['Anil Mathew', '450', 'Antu Dominic']
    },
    {
      name: 'atk_intro_bundle',
      language: 'en_US',
      category: 'UTILITY',
      body:
        'Hi {{1}}, here are the role-player introductions received for ATK Meeting {{2}}:\n\n' +
        '{{3}}\n\nPlease use these while introducing the role players.',
      examples: [
        'Riyas AF',
        '450',
        'Timer — Member introduction text\nGrammarian — Member introduction text'
      ]
    },
    {
      name: 'atk_vpe_exception',
      language: 'en_US',
      category: 'UTILITY',
      body:
        'Hi {{1}}, there are ATK VPE items that need your attention. ' +
        'Use the button below to view the current action list.',
      examples: ['Riyas AF'],
      buttons: [{ type: 'QUICK_REPLY', text: 'View action list' }]
    },
    {
      name: 'atk_theme_guided',
      language: 'en_US', category: 'UTILITY',
      body: 'Hi {{1}}, for ATK Meeting {{2}}, please reply with your theme in your own words. We will save it for the agenda.',
      examples: ['Riyas AF', '450']
    },
    {
      name: 'atk_intro_guided',
      language: 'en_US', category: 'UTILITY',
      body: 'Hi {{1}}, for your {{3}} role in ATK Meeting {{2}}, please reply to choose your saved introduction or share something new. We will guide you.',
      examples: ['Riyas AF', '450', 'Timer']
    },
    {
      name: 'atk_speaker_guided',
      language: 'en_US', category: 'UTILITY',
      body: 'Hi {{1}}, for ATK Meeting {{2}}, please reply with your speech title. We will ask for Path, Level, Project and your introduction one by one.',
      examples: ['Anil Mathew', '450']
    },
    {
      name: 'atk_member_intro',
      language: 'en_US', category: 'UTILITY',
      body: 'Hi {{1}}, welcome to ATK! Please reply with a short introduction about yourself. We will save it for future meeting introductions.',
      examples: ['Riyas AF']
    },
    {
      name: 'atk_tmod_update',
      language: 'en_US', category: 'UTILITY',
      body: 'Hi {{1}}, here are details received for ATK Meeting {{2}}:\n\n{{3}}\n\nPlease use these details while preparing for the meeting.',
      examples: ['Riyas AF', '450', 'Speaker Anil Mathew — speech title and project details']
    }
  ];
}

function createOrReadTemplate_(token, wabaId, spec) {
  const base =
    'https://graph.facebook.com/v26.0/' +
    encodeURIComponent(wabaId) +
    '/message_templates';

  const lookupUrl =
    base +
    '?name=' + encodeURIComponent(spec.name) +
    '&fields=id,name,status,category,language';

  const lookup = UrlFetchApp.fetch(lookupUrl, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (lookup.getResponseCode() >= 200 && lookup.getResponseCode() < 300) {
    const data = JSON.parse(lookup.getContentText() || '{}').data || [];
    const exact = data.find(t => t.name === spec.name && t.language === spec.language);
    if (exact) {
      return {
        ok: true,
        name: spec.name,
        action: 'EXISTS',
        id: exact.id || '',
        status: exact.status || '',
        category: exact.category || ''
      };
    }
  }

  const bodyComponent = {
    type: 'BODY',
    text: spec.body
  };

  if (spec.examples && spec.examples.length) {
    bodyComponent.example = {
      body_text: [spec.examples]
    };
  }

  const components = [bodyComponent];

  if (spec.buttons && spec.buttons.length) {
    components.push({
      type: 'BUTTONS',
      buttons: spec.buttons
    });
  }

  const response = UrlFetchApp.fetch(base, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({
      name: spec.name,
      language: spec.language,
      category: spec.category,
      components: components
    }),
    muteHttpExceptions: true
  });

  const http = response.getResponseCode();
  const body = response.getContentText();
  let parsed = {};
  try { parsed = JSON.parse(body || '{}'); } catch (_) {}

  if (http >= 200 && http < 300) {
    return {
      ok: true,
      name: spec.name,
      action: 'CREATED',
      id: parsed.id || '',
      status: parsed.status || '',
      category: parsed.category || spec.category
    };
  }

  return {
    ok: false,
    name: spec.name,
    action: 'ERROR',
    http: http,
    error: parsed.error ? JSON.stringify(parsed.error) : body
  };
}


function syncAtkTemplateApprovalStatus_() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('META_ACCESS_TOKEN');
  const wabaId = props.getProperty('WA_WABA_ID');
  if (!token || !wabaId) return;

  const specs = getAtkWhatsAppTemplateSpecs_();
  const wanted = new Set(specs.map(s => s.name));

  const url =
    'https://graph.facebook.com/v26.0/' +
    encodeURIComponent(wabaId) +
    '/message_templates?fields=id,name,status,category,language&limit=100';

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  const http = response.getResponseCode();
  if (http < 200 || http >= 300) {
    logAutomation_(
      'WhatsApp Templates',
      '',
      '',
      'Approval status sync',
      'Error',
      'HTTP ' + http + ': ' + response.getContentText()
    );
    return;
  }

  const data = JSON.parse(response.getContentText() || '{}').data || [];
  const statusMap = {};
  data.forEach(t => {
    if (wanted.has(t.name) && t.language === 'en_US') {
      statusMap[t.name] = {
        status: String(t.status || 'UNKNOWN'),
        category: String(t.category || '')
      };
    }
  });

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Automation Config');
  if (!sheet) return;

  const rows = sheet.getRange(48, 1, Math.max(10, sheet.getLastRow() - 47), 4).getValues();
  const updates = [];
  let changed = false;

  rows.forEach(row => {
    const key = String(row[0] || '');
    const templateName = String(row[1] || '');
    const oldStatus = String(row[3] || '');
    const meta = statusMap[templateName];

    let newStatus = oldStatus;
    if (/^RETIRED$/i.test(oldStatus)) {
      newStatus = oldStatus;
    } else if (meta) {
      newStatus = meta.status + (meta.category ? ' · ' + meta.category : '');
    } else if (templateName) {
      newStatus = 'NOT FOUND';
    }

    updates.push([newStatus]);
    if (newStatus !== oldStatus) changed = true;
  });

  if (changed) {
    sheet.getRange(48, 4, updates.length, 1).setValues(updates);
    logAutomation_(
      'WhatsApp Templates',
      '',
      '',
      'Approval status sync',
      'Success',
      'Template review statuses updated'
    );
  }
}


function testApprovedTemplatesToVpe_() {
  if (configValue_('SYSTEM_MODE', '') !== 'DEV' ||
      configValue_('WA_SEND_MODE', '') !== 'TEST') {
    throw new Error('Template test is restricted to DEV TEST mode.');
  }

  const member = configValue_('VPE_MEMBER_NAME', 'Riyas AF');
  const phone = normalizePhone_(phoneForMember_(member));
  if (!phone) throw new Error('VPE WhatsApp number not found.');
  if (!isWhatsAppRecipientAllowed_(phone)) {
    throw new Error('VPE is not listed in WA_TEST_RECIPIENTS.');
  }

  const meetingNo = 'DEV TEST';
  const tests = [
    {
      key: 'WA_TEMPLATE_SPEAKER_DETAILS',
      fallback: 'atk_speaker_guided',
      params: [member, meetingNo]
    },
    {
      key: 'WA_TEMPLATE_THEME_REQUEST',
      fallback: 'atk_theme_guided',
      params: [member, meetingNo]
    },
    {
      key: 'WA_TEMPLATE_INTRO_REQUEST',
      fallback: 'atk_intro_guided',
      params: [member, meetingNo, 'Timer']
    },
    {
      key: 'WA_TEMPLATE_DEFAULT_INTRO_REQUEST',
      fallback: 'atk_member_intro',
      params: [member]
    },
    {
      key: 'WA_TEMPLATE_TMOD_UPDATE',
      fallback: 'atk_tmod_update',
      params: [member, meetingNo, 'DEV template verification only']
    }
  ];

  const results = [];
  tests.forEach((test, index) => {
    const templateName = configValue_(test.key, test.fallback);
    if (!isApprovedTemplate_(templateName)) {
      results.push({
        ok: false,
        key: test.key,
        template: templateName,
        error: 'Template is not marked APPROVED in Automation Config.'
      });
      return;
    }

    if (index > 0) Utilities.sleep(1200);
    const result = sendWhatsAppTemplate_(phone, templateName, test.params, null);
    results.push({
      ok: Boolean(result && result.ok),
      key: test.key,
      template: templateName,
      messageId: result && result.messageId ? result.messageId : '',
      error: result && result.error ? result.error : ''
    });
  });

  return {
    ok: results.length === 5 && results.every(result => result.ok),
    recipient: phone,
    mode: 'DEV TEST',
    results: results
  };
}


function testRoleReminderTemplateToVpe_() {
  const member = configValue_('VPE_MEMBER_NAME', 'Riyas AF');
  const phone = phoneForMember_(member);
  if (!phone) throw new Error('VPE WhatsApp number not found.');

  const meetingNo = 450;
  const result = sendWhatsAppTemplate_(
    phone,
    configValue_('WA_TEMPLATE_ROLE_REMINDER', 'atk_role_reminder'),
    [member, meetingNo, 'Timer, Grammarian'],
    null
  );

  return {
    ok: Boolean(result && result.ok),
    result: result
  };
}


function themeRequestV2Spec_() {
  return {
    name: 'atk_theme_request_v2',
    language: 'en_US',
    category: 'UTILITY',
    body:
      'Hi {{1}},\n\n' +
      'You’re the Toastmaster of the Day for ATK Meeting {{2}} 🎤\n\n' +
      'Please share the theme today itself so that all the role players can align with it.\n\n' +
      'Just reply with your theme. We’ll save it for the meeting agenda.\n\n' +
      'Thank you.',
    examples: ['Riyas AF', '452']
  };
}

function themeReminderV2Spec_() {
  return {
    name: 'atk_theme_reminder_v3',
    language: 'en_US',
    category: 'UTILITY',
    body:
      'Hi {{1}},\n\n' +
      'Just a reminder about the meeting theme for ATK Meeting {{2}} 🎤\n\n' +
      'We haven’t received the theme yet.\n\n' +
      'Please share it today so that all the role players can align with it.\n\n' +
      'Just reply with your theme.\n\n' +
      'Thank you.',
    examples: ['Riyas AF', '452']
  };
}

function speakerEvaluatorPairingV2Spec_() {
  return {
    name: 'atk_speaker_evaluator_pairing_v2',
    language: 'en_US',
    category: 'UTILITY',
    body:
      'Hi {{1}},\n\n' +
      'you will be evaluated by {{2}} for ATK Meeting {{3}}. Please connect directly and coordinate the project and evaluation requirements before the meeting.\n\n' +
      'Please ensure you send:\n' +
      'Your full Project and evaluation form.\n\n' +
      'Thank you.',
    examples: ['Anil Mathew', 'Antu Dominic', '452']
  };
}

function evaluatorAssignmentV2Spec_() {
  return {
    name: 'atk_evaluator_assignment_v2',
    language: 'en_US',
    category: 'UTILITY',
    body:
      'Hi {{1}},\n\n' +
      'You will be evaluating {{2}} for ATK Meeting {{3}}.\n\n' +
      'Please connect directly with the speaker before the meeting and ensure you receive their complete project and evaluation form.\n\n' +
      'Thank you.',
    examples: ['Antu Dominic', 'Anil Mathew', '452']
  };
}

function speakerCoordinationCheckV2Spec_() {
  return {
    name: 'atk_speaker_coord_check_v2',
    language: 'en_US',
    category: 'UTILITY',
    body:
      'Hi {{1}},\n\n' +
      'Hope you have coordinated with your evaluator, {{2}}, for ATK Meeting {{3}}.\n\n' +
      'Have you coordinated?',
    examples: ['Anil Mathew', 'Antu Dominic', '452'],
    buttons: [
      {type: 'QUICK_REPLY', text: 'Yes'},
      {type: 'QUICK_REPLY', text: 'No'}
    ]
  };
}

function evaluatorCoordinationCheckV2Spec_() {
  return {
    name: 'atk_evaluator_coord_check_v2',
    language: 'en_US',
    category: 'UTILITY',
    body:
      'Hi {{1}},\n\n' +
      'Hope you have coordinated with your speaker, {{2}}, for ATK Meeting {{3}}.\n\n' +
      'Have you coordinated?',
    examples: ['Antu Dominic', 'Anil Mathew', '452'],
    buttons: [
      {type: 'QUICK_REPLY', text: 'Yes'},
      {type: 'QUICK_REPLY', text: 'No'}
    ]
  };
}

function speakerDetailsV3Spec_() {
  return {
    name: 'atk_speaker_details_v3',
    language: 'en_US',
    category: 'UTILITY',
    body:
      'Hi {{1}},\n\n' +
      'You’re confirmed as a speaker for ATK Meeting {{2}} 🎤\n\n' +
      'Please send your speech details in one message using this format:\n\n' +
      'Path:\n' +
      'Level:\n' +
      'Project:\n' +
      'Title:\n\n' +
      'Thank you.',
    examples: ['Anil Mathew', '452']
  };
}

function speechInviteV2Spec_() {
  return {
    name: 'atk_speech_invite_v2',
    language: 'en_US',
    category: 'MARKETING',
    body:
      'Hi {{1}},\n\n' +
      'Want to give a speech?\n\n' +
      'Select a slot before it’s gone! 🎤\n\n' +
      'If you’d like to take a speaking slot, choose a convenient meeting date using the button below.',
    examples: ['Riyas AF'],
    buttons: [{type: 'QUICK_REPLY', text: 'Reserve speech slot'}]
  };
}

