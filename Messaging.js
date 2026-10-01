/**
 * Scheduled WhatsApp messaging layer.
 *
 * Free-form messages are used only when the member has an open service window.
 * Otherwise an approved template name must be configured for that message kind.
 * The actual template definitions/approval remain in Meta.
 */

let WA_INBOUND_CACHE_ = null;

function getLastInboundMap_() {
  if (WA_INBOUND_CACHE_) return WA_INBOUND_CACHE_;

  const map = {};
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('WA Webhook Log');

  if (sheet && sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    rows.forEach(row => {
      const at = row[0];
      const event = String(row[1] || '').trim();
      const phone = normalizePhone_(row[2] || '');
      if (event !== 'Incoming Message' || !phone || !(at instanceof Date)) return;

      if (!map[phone] || at > map[phone]) {
        map[phone] = at;
      }
    });
  }

  WA_INBOUND_CACHE_ = map;
  return WA_INBOUND_CACHE_;
}

function hasOpenWhatsAppServiceWindow_(phone) {
  const destination = normalizePhone_(phone);
  const lastInbound = getLastInboundMap_()[destination];
  if (!(lastInbound instanceof Date)) return false;

  // Small safety margin so a message is not attempted on the exact boundary.
  const hours = Math.max(
    1,
    Number(configValue_('WA_SERVICE_WINDOW_HOURS', '23.5')) || 23.5
  );

  return (new Date().getTime() - lastInbound.getTime()) <= hours * 60 * 60 * 1000;
}

function templateNameForKind_(kind) {
  const normalizedKind = String(kind || '').toUpperCase();

  const preferredConfig = {
    ROLE_REMINDER: ['WA_TEMPLATE_ROLE_REMINDER_V2', 'atk_role_reminder_v2'],
    SPEECH_INVITE: ['WA_TEMPLATE_SPEECH_INVITE_V2', 'atk_speech_invite_v2'],
    THEME_REQUEST: ['WA_TEMPLATE_THEME_REQUEST_V2', 'atk_theme_request_v2'],
    SPEAKER_DETAILS: ['WA_TEMPLATE_SPEAKER_DETAILS_V3', 'atk_speaker_details_v3']
  };

  const preferred = preferredConfig[normalizedKind];
  if (preferred) {
    const preferredName = String(configValue_(preferred[0], preferred[1])).trim();
    if (preferredName && isApprovedTemplate_(preferredName)) return preferredName;
  }

  if (normalizedKind === 'EVALUATOR_PAIRING_SPEAKER' ||
      normalizedKind === 'EVALUATOR_ASSIGNMENT') {
    const current = String(
      configValue_('WA_TEMPLATE_' + normalizedKind, '')
    ).trim();
    if (current && isApprovedTemplate_(current)) return current;

    const fallback = String(
      configValue_('WA_TEMPLATE_EVALUATOR_PAIRING', 'atk_evaluator_pairing')
    ).trim();
    if (fallback && isApprovedTemplate_(fallback)) return fallback;

    return current;
  }

  return String(
    configValue_('WA_TEMPLATE_' + normalizedKind, '')
  ).trim();
}

function sendAutomatedText_(phone, kind, freeformBody, templateParams) {
  const destination = normalizePhone_(phone);
  if (!destination) return { ok: false, error: 'Missing destination phone.' };

  if (hasOpenWhatsAppServiceWindow_(destination)) {
    const result = sendWhatsAppText_(destination, freeformBody);
    result.channelMode = 'FREEFORM';
    return result;
  }

  const templateName = templateNameForKind_(kind);
  if (!templateName) {
    logAutomation_(
      'WhatsApp',
      '',
      destination,
      'Template required for scheduled message: ' + kind,
      'Deferred',
      'No approved template configured'
    );

    return {
      ok: false,
      deferred: true,
      templateRequired: true,
      error: 'No approved WhatsApp template configured for ' + kind + '.'
    };
  }

  if (/^atk_(theme_guided|theme_request_v2|theme_reminder_v3|grammar_request|intro_guided|speaker_guided|speaker_details_v3|speaker_evaluator_pairing_v2|evaluator_assignment_v2|member_intro|tmod_update)$/.test(templateName) &&
      !isApprovedTemplate_(templateName)) {
    return { ok: false, deferred: true, error: 'Template awaiting Meta approval: ' + templateName };
  }

  const params = templateParams || [];
  const parameterCount = {
    atk_role_reminder_v2: 1, atk_speech_invite_v2: 1,
    atk_theme_guided: 2, atk_theme_request_v2: 2, atk_theme_reminder_v3: 2,
    atk_grammar_request: 3,
    atk_intro_guided: 3, atk_speaker_guided: 2,
    atk_speaker_details_v3: 2,
    atk_speaker_evaluator_pairing_v2: 3,
    atk_evaluator_assignment_v2: 3,
    atk_member_intro: 1, atk_tmod_update: 3
  }[templateName];
  const result = sendWhatsAppTemplate_(
    destination,
    templateName,
    parameterCount ? params.slice(0, parameterCount) : params,
    null
  );
  result.channelMode = 'TEMPLATE';
  return result;
}

function sendAutomatedQuickReply_(phone, kind, freeformPayload, templateParams, buttonPayload) {
  const destination = normalizePhone_(phone);
  if (!destination) return { ok: false, error: 'Missing destination phone.' };

  if (hasOpenWhatsAppServiceWindow_(destination)) {
    const result = sendWhatsAppPayload_(freeformPayload);
    result.channelMode = 'FREEFORM';
    return result;
  }

  const templateName = templateNameForKind_(kind);
  if (!templateName) {
    logAutomation_(
      'WhatsApp',
      '',
      destination,
      'Template required for scheduled interactive message: ' + kind,
      'Deferred',
      'No approved template configured'
    );

    return {
      ok: false,
      deferred: true,
      templateRequired: true,
      error: 'No approved WhatsApp template configured for ' + kind + '.'
    };
  }

  if (!isApprovedTemplate_(templateName)) {
    return {ok: false, deferred: true, error: 'Template awaiting Meta approval: ' + templateName};
  }

  const result = sendWhatsAppTemplate_(
    destination,
    templateName,
    templateParams || [],
    buttonPayload || null
  );
  result.channelMode = 'TEMPLATE';
  return result;
}

function sanitizeWhatsAppTemplateParam_(value) {
  return String(value === null || value === undefined ? '' : value)
    .replace(/[\r\n\t]+/g, ' | ')
    .replace(/ {5,}/g, '    ')
    .trim();
}

function sendWhatsAppTemplate_(phone, templateName, bodyParams, quickReplyPayload) {
  const components = [];
  const params = (bodyParams || []).map(value => ({
    type: 'text',
    text: sanitizeWhatsAppTemplateParam_(value)
  }));

  if (params.length) {
    components.push({
      type: 'body',
      parameters: params
    });
  }

  if (quickReplyPayload) {
    const payloads = Array.isArray(quickReplyPayload)
      ? quickReplyPayload
      : [quickReplyPayload];

    payloads.forEach((payload, index) => {
      components.push({
        type: 'button',
        sub_type: 'quick_reply',
        index: String(index),
        parameters: [{
          type: 'payload',
          payload: String(payload)
        }]
      });
    });
  }

  const template = {
    name: templateName,
    language: {
      code: configValue_('WA_TEMPLATE_LANGUAGE', 'en_US')
    }
  };

  if (components.length) template.components = components;

  return sendWhatsAppPayload_({
    messaging_product: 'whatsapp',
    recipient_type: 'individual',
    to: normalizePhone_(phone),
    type: 'template',
    template: template
  });
}

function isApprovedTemplate_(name) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Automation Config');
  if (!sheet || sheet.getLastRow() < 48) return false;
  const rows = sheet.getRange(48, 1, sheet.getLastRow() - 47, 4).getValues();
  return rows.some(r => String(r[1] || '') === name &&
    String(r[3] || '').toUpperCase().indexOf('APPROVED') === 0);
}
