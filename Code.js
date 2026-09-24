const SPREADSHEET_ID = '1qZnZov1l9DfvhTzruqlLubGhydcOOFKLf2kG8xplNLM';

// Must match the Verify Token already entered in Meta.
const VERIFY_TOKEN = 'ATK_VPE_VERIFY_2026_X7K9M4P2';


/**
 * Meta uses this to verify our webhook.
 */
function doGet(e) {

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

    const payload = JSON.parse(e.postData.contents);

    logWhatsAppWebhook_(payload);

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

        sheet.appendRow([
          new Date(),
          'Status: ' + (status.status || ''),
          status.recipient_id || '',
          '',
          status.id || '',
          '',
          ''
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