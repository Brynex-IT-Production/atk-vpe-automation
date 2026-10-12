// DEV deployment refresh 2026-10-11.25
// DEV immediate-test deployment marker 2026-10-01
// DEV immediate-test redeploy after POST redirect fix
// Immediate test retry 2
const SPREADSHEET_ID = '1qZnZov1l9DfvhTzruqlLubGhydcOOFKLf2kG8xplNLM';

// Must match the Verify Token already entered in Meta.
const VERIFY_TOKEN = 'ATK_VPE_VERIFY_2026_X7K9M4P2';


/**
 * Meta uses this to verify our webhook.
 */
function doGet(e) {

  if (e && e.parameter && e.parameter.view === 'programme' && e.parameter.token) {
    return renderProgrammeReviewPage_(String(e.parameter.token || ''));
  }

  // CREATE_WA_TEMPLATES_GET — DEV deployment helper
  if (e && e.parameter &&
      e.parameter.system_action === 'CREATE_WA_TEMPLATES' &&
      e.parameter.verify_token === VERIFY_TOKEN) {
    const result = createAtkWhatsAppTemplates_();
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (e && e.parameter &&
      e.parameter.system_action === 'PHASE1_TICK' &&
      e.parameter.verify_token === VERIFY_TOKEN) {
    phase1ScheduledTick();
    return ContentService.createTextOutput('PHASE1_TICK_OK')
      .setMimeType(ContentService.MimeType.TEXT);
  }


  const mode = e.parameter['hub.mode'];
  const token = e.parameter['hub.verify_token'];
  const challenge = e.parameter['hub.challenge'];

  if (
    mode === 'subscribe' &&
    token === VERIFY_TOKEN &&
    challenge
  ) {
    return ContentService
      .createTextOutput(challenge)
      .setMimeType(ContentService.MimeType.TEXT);
  }

  return ContentService
    .createTextOutput('Webhook verification failed')
    .setMimeType(ContentService.MimeType.TEXT);
}


/**
 * Meta sends real WhatsApp messages and delivery events here.
 */
function doPost(e) {

  try {

    const incoming = JSON.parse(e.postData.contents);

    // WEB_API_V1: Flutter/Web requests are explicitly namespaced so they can
    // share the Apps Script deployment without entering Meta webhook logic.
    if (incoming && incoming.api_version === 'v1') {
      const apiResult = apiHandleRequest_(incoming);
      return apiJson_(apiResult.status || 200, apiResult.data, apiResult.error);
    }

    const relaySecret = PropertiesService.getScriptProperties()
      .getProperty('RELAY_SHARED_SECRET');
    let payload = incoming;
    if (incoming && incoming.relay_payload !== undefined) {
      if (!relaySecret || !verifyRelayEnvelope_(incoming, relaySecret)) {
        return ContentService.createTextOutput('EVENT_REJECTED');
      }
      payload = JSON.parse(incoming.relay_payload);
    } else if (relaySecret && (!incoming || !incoming.system_action)) {
      // Once the relay secret is configured, direct unsigned Meta events fail closed.
      return ContentService.createTextOutput('EVENT_REJECTED');
    }

    if (
      payload &&
      payload.system_action === 'PHASE1_TICK' &&
      payload.verify_token === VERIFY_TOKEN
    ) {
      phase1ScheduledTick();

      return ContentService
        .createTextOutput('PHASE1_TICK_OK')
        .setMimeType(ContentService.MimeType.TEXT);
    }

    if (
      payload &&
      payload.system_action === 'CREATE_WA_TEMPLATES' &&
      payload.verify_token === VERIFY_TOKEN
    ) {
      const result = createAtkWhatsAppTemplates_();

      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (
      payload &&
      payload.system_action === 'TEST_APPROVED_WA_TEMPLATES' &&
      payload.verify_token === VERIFY_TOKEN
    ) {
      const result = testApprovedTemplatesToVpe_();

      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (
      payload &&
      payload.system_action === 'TEST_ROLE_TEMPLATE_ONLY' &&
      payload.verify_token === VERIFY_TOKEN
    ) {
      const result = sendDevRoleReminderForNextVacancy_();

      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (payload && payload.system_action === 'RESEND_DEV_VPE_INTRO' &&
        payload.verify_token === VERIFY_TOKEN) {
      return ContentService.createTextOutput(JSON.stringify(resendDevVpeIntro_()))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // DEV-only HTTP bridge for the final API/WhatsApp regression gate.
    // This exists because Apps Script Execution API calls are unreliable from
    // GitHub Actions for this project. The shared verify token plus DEV mode
    // prevents these helpers from being callable in LIVE.
    if (payload && payload.system_action === 'DEV_MASTER_REGRESSION' &&
        payload.verify_token === VERIFY_TOKEN &&
        configValue_('SYSTEM_MODE','') === 'DEV') {
      return ContentService
        .createTextOutput(JSON.stringify(apiDevMasterRegressionSuite()))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (payload && payload.system_action === 'DEV_WHATSAPP_GOLDEN_START' &&
        payload.verify_token === VERIFY_TOKEN &&
        configValue_('SYSTEM_MODE','') === 'DEV') {
      return ContentService
        .createTextOutput(JSON.stringify(apiDevWhatsAppGoldenPathTest()))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (payload && payload.system_action === 'DEV_WHATSAPP_GOLDEN_RESULT' &&
        payload.verify_token === VERIFY_TOKEN &&
        configValue_('SYSTEM_MODE','') === 'DEV') {
      return ContentService
        .createTextOutput(JSON.stringify(apiDevGoldenPathResult()))
        .setMimeType(ContentService.MimeType.JSON);
    }

    logWhatsAppWebhook_(payload);
    if (!processAgendaDecisionButton_(payload) &&
        !processAgendaApprovalButton_(payload) &&
        !processGrammarApprovalButton_(payload)) {
      processWhatsAppAutomation_(payload);
    }

    return ContentService
      .createTextOutput('EVENT_RECEIVED')
      .setMimeType(ContentService.MimeType.TEXT);

  } catch (error) {

    console.error(error);

    return ContentService
      .createTextOutput('EVENT_RECEIVED')
      .setMimeType(ContentService.MimeType.TEXT);
  }
}


/**
 * The Worker verifies Meta's signed header. Apps Script receives this
 * short-lived HMAC envelope because doPost does not expose HTTP headers.
 */
function verifyRelayEnvelope_(envelope, secret) {
  const timestamp = String(envelope.relay_timestamp || '');
  const raw = envelope.relay_payload;
  const mac = String(envelope.relay_mac || '').toLowerCase();
  if (!/^\d{13}$/.test(timestamp) || typeof raw !== 'string' ||
      !/^[0-9a-f]{64}$/.test(mac)) return false;
  if (Math.abs(Date.now() - Number(timestamp)) > 5 * 60 * 1000) return false;
  const bytes = Utilities.computeHmacSha256Signature(
    timestamp + '.' + raw, secret, Utilities.Charset.UTF_8
  );
  const expected = bytes.map(v => ('0' + (v & 255).toString(16)).slice(-2)).join('');
  let difference = 0;
  for (let i = 0; i < expected.length; i++) {
    difference |= expected.charCodeAt(i) ^ mac.charCodeAt(i);
  }
  return difference === 0;
}


/**
 * Finds the ATK member who owns a WhatsApp number.
 *
 * Member Master:
 * A = Member ID
 * B = Member Name
 * C = WhatsApp Number
 * D = Status
 */
function findMemberByWhatsApp_(phone) {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Member Master');

  if (!sheet) {
    return 'Unknown';
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return 'Unknown';
  }

  const rows = sheet
    .getRange(2, 1, lastRow - 1, 4)
    .getValues();

  const cleanPhone = String(phone || '')
    .replace(/\D/g, '');

  for (const row of rows) {

    const memberName = String(row[1] || '').trim();

    const whatsapp = String(row[2] || '')
      .replace(/\D/g, '');

    const status = String(row[3] || '').trim();

    if (
      status === 'Active' &&
      whatsapp &&
      whatsapp === cleanPhone
    ) {
      return memberName;
    }
  }

  return 'Unknown';
}


/**
 * Logs incoming WhatsApp messages into the DEV spreadsheet.
 *
 * It also identifies the member automatically using Member Master.
 */
function logWhatsAppWebhook_(payload) {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let sheet = ss.getSheetByName('WA Webhook Log');


  /*
   * Create log sheet if it does not exist.
   */
  if (!sheet) {

    sheet = ss.insertSheet('WA Webhook Log');

    sheet.appendRow([
      'Timestamp',
      'Event',
      'WhatsApp Number',
      'Member Name',
      'Message ID',
      'Message Type',
      'Message / Selection'
    ]);

    sheet.hideSheet();
  }


  /*
   * Upgrade the existing old log.
   *
   * Your existing log did not have Member Name.
   * This inserts the column without deleting old test data.
   */
  else if (sheet.getRange('D1').getValue() !== 'Member Name') {

    sheet.insertColumnBefore(4);

    sheet.getRange('D1')
      .setValue('Member Name');
  }


  const entry =
    payload.entry &&
    payload.entry[0];

  if (!entry || !entry.changes) {
    return;
  }


  entry.changes.forEach(change => {

    const value = change.value || {};


    /*
     * INCOMING WHATSAPP MESSAGE
     */
    if (
      value.messages &&
      value.messages.length > 0
    ) {

      value.messages.forEach(message => {

        let content = '';


        /*
         * Normal text message
         */
        if (message.type === 'text') {

          content =
            message.text?.body || '';
        }


        /*
         * Button response
         */
        else if (message.type === 'button') {

          content =
            message.button?.text || '';
        }


        /*
         * Interactive response
         *
         * This will later be used for role selection.
         */
        else if (message.type === 'image') {
          content = '[Image] ' + String(message.image?.id || '') +
            (message.image?.caption ? ' · ' + message.image.caption : '');
        }

        else if (message.type === 'interactive') {

          if (
            message.interactive?.button_reply
          ) {

            content =
              message.interactive.button_reply.title ||
              message.interactive.button_reply.id ||
              '';
          }

          else if (
            message.interactive?.list_reply
          ) {

            content =
              message.interactive.list_reply.title ||
              message.interactive.list_reply.id ||
              '';
          }
        }


        const phone =
          message.from || '';

        const memberName =
          findMemberByWhatsApp_(phone);


        sheet.appendRow([
          new Date(),
          'Incoming Message',
          phone,
          memberName,
          message.id || '',
          message.type || '',
          content
        ]);

      });
    }


    /*
     * MESSAGE DELIVERY STATUS
     */
    if (
      value.statuses &&
      value.statuses.length > 0
    ) {

      value.statuses.forEach(status => {

        const errorText = (status.errors || []).map(error => {
          const parts = [];
          if (error.code !== undefined) parts.push('code=' + error.code);
          if (error.title) parts.push('title=' + error.title);
          if (error.message) parts.push('message=' + error.message);
          if (error.error_data && error.error_data.details) {
            parts.push('details=' + error.error_data.details);
          }
          return parts.join(' | ');
        }).join(' || ');

        sheet.appendRow([
          new Date(),
          'Status: ' + (status.status || ''),
          status.recipient_id || '',
          '',
          status.id || '',
          '',
          errorText
        ]);

      });
    }

  });
}


/**
 * One-time utility already used successfully
 * to subscribe the VPE app to the ATK WABA.
 */
function subscribeVpeAppToWaba() {

  const props =
    PropertiesService.getScriptProperties();

  const token =
    props.getProperty('META_ACCESS_TOKEN');

  const wabaId =
    props.getProperty('WA_WABA_ID');

  if (!token) {
    throw new Error(
      'META_ACCESS_TOKEN is missing from Script Properties.'
    );
  }

  if (!wabaId) {
    throw new Error(
      'WA_WABA_ID is missing from Script Properties.'
    );
  }

  const url =
    'https://graph.facebook.com/v26.0/' +
    encodeURIComponent(wabaId) +
    '/subscribed_apps';

  const response =
    UrlFetchApp.fetch(url, {

      method: 'post',

      headers: {
        Authorization: 'Bearer ' + token
      },

      muteHttpExceptions: true

    });

  const status =
    response.getResponseCode();

  const body =
    response.getContentText();

  console.log(
    'HTTP ' + status
  );

  console.log(body);

  if (
    status < 200 ||
    status >= 300
  ) {

    throw new Error(
      'Meta subscription failed. HTTP ' +
      status +
      ': ' +
      body
    );
  }

  return body;
}


/**
 * Safe diagnostic utility.
 * Shows only Script Property names.
 */
function checkMetaProperties() {

  const props =
    PropertiesService
      .getScriptProperties()
      .getProperties();

  console.log(
    Object.keys(props)
  );
}


/**
 * Safe diagnostic utility.
 * Shows only value lengths, never the token.
 */
function checkMetaPropertyValues() {

  const props =
    PropertiesService.getScriptProperties();

  const token =
    props.getProperty('META_ACCESS_TOKEN') || '';

  const phoneId =
    props.getProperty('WA_PHONE_NUMBER_ID') || '';

  const wabaId =
    props.getProperty('WA_WABA_ID') || '';

  console.log(
    'META_ACCESS_TOKEN length: ' +
    token.length
  );

  console.log(
    'WA_PHONE_NUMBER_ID length: ' +
    phoneId.length
  );

  console.log(
    'WA_WABA_ID length: ' +
    wabaId.length
  );
}

/** Explicit DEV-only resend requested by the VPE; never changes the allowlist. */
function resendDevVpeIntro_() {
  if (configValue_('SYSTEM_MODE', '') !== 'DEV' ||
      configValue_('WA_SEND_MODE', '') !== 'TEST') {
    throw new Error('Introduction resend is restricted to DEV TEST mode.');
  }
  const member = configValue_('VPE_MEMBER_NAME', '');
  const phone = phoneForMember_(member);
  if (!member || !phone || !isWhatsAppRecipientAllowed_(phone)) {
    throw new Error('VPE must be an allowlisted DEV recipient.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    const sentAt = Number(props.getProperty('DEV_VPE_INTRO_RESENT_AT') || 0);
    if (Date.now() - sentAt < 5 * 60 * 1000) return {ok: true, skipped: 'Recently resent'};
    const active = activeConversation_(phone);
    if (active && active.kind !== 'DEFAULT_INTRO') {
      return {ok: false, error: 'Another guided request is active', kind: active.kind};
    }
    if (defaultIntro_(member)) return {ok: false, error: 'Default introduction already saved'};
    const body = introductionPrompt_(member, '', 0);
    const result = sendAutomatedText_(phone, 'DEFAULT_INTRO_REQUEST', body, [member]);
    recordWorkflowOutbound_(member, phone, 'DEFAULT_INTRO_REQUEST', body, 0, result,
      'DEV_INTRO_RESEND|' + Date.now());
    if (result && result.ok) {
      setConversation_(phone, {kind: 'DEFAULT_INTRO', meetingNo: 0, related: '', step: 'CHOICE', data: {}});
      props.setProperty('DEV_VPE_INTRO_RESENT_AT', String(Date.now()));
    }
    return result;
  } finally {
    lock.releaseLock();
  }
}


/* ATK_AGENDA_DECISION_BUTTON_V1 */
function processAgendaDecisionButton_(payload) {
  const entry = payload && payload.entry && payload.entry[0];
  if (!entry || !entry.changes) return false;
  let handled = false;
  entry.changes.forEach(change => {
    const messages = (change.value && change.value.messages) || [];
    messages.forEach(message => {
      const reply = message.type === 'interactive' &&
        message.interactive && message.interactive.button_reply;
      if (!reply || !reply.id) return;
      const match = String(reply.id).match(/^AGENDA_TT_(KEEP|REMOVE)\|(\d+)\|(\d+(?:\.\d+)?)$/);
      if (!match) return;
      recordAgendaTableTopicsDecision_(Number(match[2]), match[1], Number(match[3]));
      handled = true;
    });
  });
  return handled;
}


/* ATK_GRAMMAR_APPROVAL_BUTTON_V1 */
function processGrammarApprovalButton_(payload) {
  const entry = payload && payload.entry && payload.entry[0];
  if (!entry || !entry.changes) return false;
  let handled = false;
  entry.changes.forEach(change => {
    const messages = (change.value && change.value.messages) || [];
    messages.forEach(message => {
      const reply = message.type === 'interactive' &&
        message.interactive && message.interactive.button_reply;
      if (!reply || !reply.id) return;
      const match = String(reply.id).match(
        /^GRAMMAR_APPROVAL\|(\d+)\|(WORD|IDIOM)\|(APPROVE|EDIT)$/
      );
      if (!match) return;
      const phone = String(message.from || '');
      const member = findMemberByWhatsApp_(phone);
      handleGrammarApprovalReply_(
        member, phone, Number(match[1]), match[2], match[3], String(message.id || '')
      );
      handled = true;
    });
  });
  return handled;
}


/* ATK_AGENDA_APPROVAL_BUTTON_V1 */
function processAgendaApprovalButton_(payload) {
  const entry = payload && payload.entry && payload.entry[0];
  if (!entry || !entry.changes) return false;
  let handled = false;
  entry.changes.forEach(change => {
    const messages = (change.value && change.value.messages) || [];
    messages.forEach(message => {
      const reply = message.type === 'interactive' &&
        message.interactive && message.interactive.button_reply;
      if (!reply || !reply.id) return;
      const match = String(reply.id).match(/^AGENDA_APPROVE\|(\d+)\|(\d+)\|([0-9a-f]{64})$/i);
      if (!match) return;
      const phone = String(message.from || '');
      const member = findMemberByWhatsApp_(phone);
      recordAgendaApproval_(
        Number(match[1]), Number(match[2]), String(match[3]).toLowerCase(),
        member, String(message.id || '')
      );
      handled = true;
    });
  });
  return handled;
}
