/**
 * TMOD theme, Grammarian and role-player introduction workflows.
 * These flows use explicit reply prefixes so submissions are unambiguous.
 */

function parseSpeakerDetailsBlock_(text) {
  const fields = {};
  let recognized = 0;

  String(text || '').split(/\r?\n/).forEach(line => {
    const match = String(line || '').match(/^\s*(Path|Level|Project|Title|Time)\s*:\s*(.*?)\s*$/i);
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
    title: fields.title || '',
    time: fields.time || ''
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
  let currentKey = '';
  let section = '';

  raw.split(/\r?\n/).forEach(line => {
    const match = String(line || '').match(
      /^\s*(Word(?: of the Day)?|Word meaning|Meaning of (?:the )?word|Idiom(?: of the Day)?|Idiom meaning|Meaning of (?:the )?idiom|Meaning)\s*:\s*(.*?)\s*$/i
    );

    if (match) {
      const label = String(match[1] || '').toLowerCase();
      const value = String(match[2] || '').trim();
      let key = '';

      if (/^word(?: of the day)?$/.test(label)) {
        key = 'word';
        section = 'word';
      } else if (/^(word meaning|meaning of (?:the )?word)$/.test(label)) {
        key = 'wordMeaning';
        section = 'word';
      } else if (/^idiom(?: of the day)?$/.test(label)) {
        key = 'idiom';
        section = 'idiom';
      } else if (/^(idiom meaning|meaning of (?:the )?idiom)$/.test(label)) {
        key = 'idiomMeaning';
        section = 'idiom';
      } else if (label === 'meaning') {
        // Natural replies commonly use "Meaning:" twice. Resolve each one
        // from the section that immediately precedes it, rather than requiring
        // the stricter "Word meaning:" / "Idiom meaning:" labels.
        if (section === 'word' && !fields.wordMeaning) {
          key = 'wordMeaning';
        } else if (section === 'idiom' && !fields.idiomMeaning) {
          key = 'idiomMeaning';
        } else if (fields.word && !fields.wordMeaning && !fields.idiom) {
          key = 'wordMeaning';
          section = 'word';
        } else if (fields.idiom && !fields.idiomMeaning) {
          key = 'idiomMeaning';
          section = 'idiom';
        }
      }

      currentKey = key;
      if (key && value) fields[key] = value;
      return;
    }

    // Preserve continuation lines inside a labelled field. This lets members
    // send a longer definition over multiple lines without breaking parsing.
    const continuation = String(line || '').trim();
    if (continuation && currentKey) {
      fields[currentKey] = fields[currentKey]
        ? fields[currentKey] + '\n' + continuation
        : continuation;
    }
  });

  let recognized = ['word', 'wordMeaning', 'idiom', 'idiomMeaning']
    .filter(key => String(fields[key] || '').trim()).length;

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


function parseGrammarCombinedBlock_(text) {
  const raw = String(text || '').trim();
  const fields = {};
  let section = '';
  let currentKey = '';

  raw.split(/\r?\n/).forEach(line => {
    const textLine = String(line || '');
    const match = textLine.match(
      /^\s*(Word(?: of the Day)?|Idiom(?: of the Day)?|Meaning|Usage|Example|Example usage)\s*(?::|[-–—])\s*(.*?)\s*$/i
    );

    if (match) {
      const label = String(match[1] || '').trim().toLowerCase();
      const value = String(match[2] || '').trim();
      let key = '';

      if (/^word(?: of the day)?$/.test(label)) {
        section = 'word';
        key = 'word';
      } else if (/^idiom(?: of the day)?$/.test(label)) {
        section = 'idiom';
        key = 'idiom';
      } else if (label === 'meaning' && section) {
        key = section + 'Meaning';
      } else if (/^(usage|example|example usage)$/.test(label) && section) {
        key = section + 'Usage';
      }

      currentKey = key;
      if (key && value) fields[key] = value;
      return;
    }

    const continuation = textLine.trim();
    if (continuation && currentKey) {
      fields[currentKey] = fields[currentKey]
        ? fields[currentKey] + '\n' + continuation
        : continuation;
    }
  });

  const keys = [
    'word','wordMeaning','wordUsage',
    'idiom','idiomMeaning','idiomUsage'
  ];
  return {
    recognized: keys.filter(k => String(fields[k] || '').trim()).length,
    word: String(fields.word || '').trim(),
    wordMeaning: String(fields.wordMeaning || '').trim(),
    wordUsage: String(fields.wordUsage || '').trim(),
    idiom: String(fields.idiom || '').trim(),
    idiomMeaning: String(fields.idiomMeaning || '').trim(),
    idiomUsage: String(fields.idiomUsage || '').trim()
  };
}

function handleCombinedGrammarSubmission_(member, phone, requestedMeetingNo, body, messageId) {
  if (!configBool_('GRAMMARIAN_WORKFLOW_ENABLED')) {
    sendWhatsAppText_(phone, 'Grammarian collection is not active yet.');
    return false;
  }

  const task = findPendingTaskForMember_('GRAMMAR', member, requestedMeetingNo);
  if (!task) return false;

  const details = body && typeof body === 'object'
    ? body
    : parseGrammarCombinedBlock_(body);
  if (Number(details.recognized || 0) < 6) return false;

  const assignment = verifyRoleAssignment_(task.meetingNo, 'GRM', member);
  if (!assignment.ok) {
    clearConversation_(phone);
    sendWhatsAppText_(phone,
      'Your Grammarian assignment for Meeting ' + task.meetingNo +
      ' has changed, so I did not save this.');
    syncPhase1Tasks_();
    return true;
  }

  const source = String(messageId || '');
  appendSubmission_({
    meetingNo:task.meetingNo,
    member:member,
    type:'GRAMMAR_WORD',
    related:'GRM',
    values:[details.word, details.wordMeaning, details.wordUsage, ''],
    sourceMessageId:source ? source + '|WORD' : ''
  });
  appendSubmission_({
    meetingNo:task.meetingNo,
    member:member,
    type:'GRAMMAR_IDIOM',
    related:'GRM',
    values:[details.idiom, details.idiomMeaning, details.idiomUsage, ''],
    sourceMessageId:source ? source + '|IDIOM' : ''
  });

  clearConversation_(phone);
  logAutomation_(
    'Grammarian', task.meetingNo, member,
    'Combined Word + Idiom received',
    'Success', details.word + ' / ' + details.idiom
  );
  syncPhase1Tasks_();
  sendGrammarReviewButtons_(member, phone, task.meetingNo);
  return true;
}


function parseGrammarStageBlock_(text, stage) {
  const normalizedStage = String(stage || '').toUpperCase();
  if (normalizedStage !== 'WORD' && normalizedStage !== 'IDIOM') {
    return { recognized: 0, item: '', meaning: '', usage: '', oppositeStage: '' };
  }

  if (text && typeof text === 'object') {
    const item = String(text.item || '').trim();
    const meaning = String(text.meaning || '').trim();
    const usage = String(text.usage || '').trim();
    return {
      recognized: [item, meaning, usage].filter(Boolean).length,
      item:item, meaning:meaning, usage:usage,
      label: normalizedStage === 'WORD' ? 'Word' : 'Idiom',
      oppositeStage:''
    };
  }

  const raw = String(text || '').trim();
  const fields = {};
  let currentKey = '';
  let section = '';
  let sawExpected = false;
  let sawOpposite = false;
  const itemLabel = normalizedStage === 'WORD' ? 'Word' : 'Idiom';
  const expectedPattern = normalizedStage === 'WORD'
    ? /^(Word(?: of the Day)?)$/i
    : /^(Idiom(?: of the Day)?)$/i;
  const oppositePattern = normalizedStage === 'WORD'
    ? /^(Idiom(?: of the Day)?)$/i
    : /^(Word(?: of the Day)?)$/i;

  raw.split(/\r?\n/).forEach(line => {
    const textLine = String(line || '');
    const match = textLine.match(
      /(?:^|\s)(Word(?: of the Day)?|Idiom(?: of the Day)?|Meaning|Usage|Example|Example usage)\s*(?::|[-–—])\s*(.*?)\s*$/i
    );

    if (match) {
      const label = String(match[1] || '').trim();
      const value = String(match[2] || '').trim();
      let key = '';

      if (expectedPattern.test(label)) {
        sawExpected = true;
        section = 'EXPECTED';
        key = 'item';
      } else if (oppositePattern.test(label)) {
        sawOpposite = true;
        section = 'OPPOSITE';
        currentKey = '';
        return;
      } else if (/^meaning$/i.test(label)) {
        if (section !== 'OPPOSITE') key = 'meaning';
      } else if (/^(usage|example|example usage)$/i.test(label)) {
        if (section !== 'OPPOSITE') key = 'usage';
      }

      currentKey = key;
      if (key && value) fields[key] = value;
      return;
    }

    const continuation = textLine.trim();
    if (continuation && currentKey) {
      fields[currentKey] = fields[currentKey]
        ? fields[currentKey] + '\n' + continuation
        : continuation;
    }
  });

  if (sawOpposite && !sawExpected) {
    return {
      recognized:0, item:'', meaning:'', usage:'',
      label:itemLabel,
      oppositeStage: normalizedStage === 'WORD' ? 'IDIOM' : 'WORD'
    };
  }

  if (!fields.item && !fields.meaning && !fields.usage &&
      raw && raw.indexOf('\n') === -1 &&
      raw.indexOf(':') === -1 &&
      !/\s[-–—]\s/.test(raw) &&
      raw.length <= 120) {
    fields.item = raw;
  }

  return {
    recognized:['item','meaning','usage']
      .filter(key => String(fields[key] || '').trim()).length,
    item:String(fields.item || '').trim(),
    meaning:String(fields.meaning || '').trim(),
    usage:String(fields.usage || '').trim(),
    label:itemLabel,
    oppositeStage:''
  };
}

function grammarStageForMember_(meetingNo, member) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('WA Submissions');
  if (!sheet || sheet.getLastRow() < 2) return 'WORD';

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
  let legacyDone = false;
  let wordDone = false;
  let idiomDone = false;
  let confirmed = false;

  rows.forEach(row => {
    if (Number(row[1]) !== Number(meetingNo) ||
        String(row[2] || '').trim() !== String(member || '').trim() ||
        String(row[4] || '').trim() !== 'GRM') return;

    const type = String(row[3] || '').trim().toUpperCase();
    if (type === 'GRAMMAR') legacyDone = true;
    if (type === 'GRAMMAR_WORD') wordDone = true;
    if (type === 'GRAMMAR_IDIOM') idiomDone = true;
    if (type === 'GRAMMAR_CONFIRMED') confirmed = true;
  });

  if (legacyDone || confirmed) return 'DONE';
  if (wordDone && idiomDone) return 'REVIEW';
  if (wordDone) return 'IDIOM';
  return 'WORD';
}

function grammarStageIntroMessage_(member, meetingNo, stage) {
  if (String(stage || '').toUpperCase() === 'IDIOM') {
    return (
      'Thank you! Your Word of the Day is received. ✅\n\n' +
      '*IDIOM OF THE DAY*\n\n' +
      'Please copy the next message, fill in the details, and send it back.'
    );
  }

  return (
    'Thank you for being the Grammarian for ATK Meeting ' + meetingNo + '. 😊\n\n' +
    '*WORD OF THE DAY*\n\n' +
    'Please copy the next message, fill in the details, and send it back.'
  );
}

function grammarStageCopyMessage_(stage, editMode) {
  const label = String(stage || '').toUpperCase() === 'IDIOM' ? 'Idiom' : 'Word';
  return label + ':\nMeaning:\nUsage:';
}

function grammarMissingFieldsMessage_(stage, details) {
  const label = String(stage || '').toUpperCase() === 'IDIOM' ? 'Idiom' : 'Word';
  const missing = [];
  if (!String(details.item || '').trim()) missing.push(label);
  if (!String(details.meaning || '').trim()) missing.push('Meaning');
  if (!String(details.usage || '').trim()) missing.push('Usage');

  const lines = missing.map(field => field + ':').join('\n');
  const human = missing.length === 1 ? missing[0] :
    missing.length === 2 ? missing[0] + ' and ' + missing[1] :
    missing.slice(0, -1).join(', ') + ' and ' + missing[missing.length - 1];

  return 'Got it. I still need the ' + human + '.\n\n' + lines;
}

/**
 * DEV contract for the AI content step. The provider/model adapter should
 * return {meaning, usage}; workflow code remains provider-independent.
 */
function grammarAiGenerationRequest_(stage, item, meetingNo) {
  return {
    task: 'GRAMMAR_MEANING_USAGE',
    stage: String(stage || '').toUpperCase(),
    item: String(item || '').trim(),
    meetingNo: Number(meetingNo || 0),
    output: ['meaning', 'usage']
  };
}

function grammarApprovalMessage_(stage, item, meaning, usage) {
  const label = String(stage || '').toUpperCase() === 'IDIOM' ? 'Idiom' : 'Word';
  return label + ': ' + item + '\n' +
    'Meaning: ' + meaning + '\n' +
    'Usage: ' + usage + '\n\n' +
    'Is this okay to use for the meeting?';
}

function grammarApprovalEditFormat_(stage) {
  return grammarStageCopyMessage_(stage, true);
}

function grammarStagePromptSentAt_(meetingNo, member, stage) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('WA Outbox');
  if (!sheet || sheet.getLastRow() < 2) return null;

  const expectedKind = String(stage || '').toUpperCase() === 'IDIOM'
    ? 'GRAMMAR_IDIOM_PROMPT'
    : 'GRAMMAR_WORD_PROMPT';
  const successfulStates = new Set(['SENT', 'DELIVERED', 'READ', 'ACCEPTED']);
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).getValues();

  const sent = rows
    .filter(row =>
      String(row[1] || '').trim() === String(member || '').trim() &&
      String(row[3] || '') === expectedKind &&
      Number(row[10]) === Number(meetingNo) &&
      successfulStates.has(String(row[6] || '').toUpperCase()) &&
      row[12] instanceof Date
    )
    .map(row => row[12])
    .sort((a, b) => b.getTime() - a.getTime());

  return sent.length ? sent[0] : null;
}


function grammarReviewPromptSentAt_(meetingNo, member) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('WA Outbox');
  if (!sheet || sheet.getLastRow() < 2) return null;

  const successfulStates = new Set(['SENT','DELIVERED','READ','ACCEPTED']);
  const rows = sheet.getRange(2,1,sheet.getLastRow()-1,16).getValues();
  const sent = rows
    .filter(row =>
      String(row[1] || '').trim() === String(member || '').trim() &&
      String(row[3] || '') === 'GRAMMAR_REVIEW' &&
      Number(row[10]) === Number(meetingNo) &&
      successfulStates.has(String(row[6] || '').toUpperCase()) &&
      row[12] instanceof Date
    )
    .map(row => row[12])
    .sort((a,b) => b.getTime() - a.getTime());

  return sent.length ? sent[0] : null;
}


function grammarStageReminderMessage_(member, meetingNo, stage) {
  const label = String(stage || '').toUpperCase() === 'IDIOM'
    ? 'Idiom of the Day'
    : 'Word of the Day';
  return (
    'Hi ' + member + ',\n\n' +
    'Just a reminder about the ' + label + ' for ATK Meeting ' + meetingNo + '.\n\n' +
    grammarStageCopyMessage_(stage)
  );
}

function sendGrammarStagePrompt_(member, phone, meetingNo, stage, sourceKind) {
  const normalizedStage = String(stage || '').toUpperCase();
  if (normalizedStage !== 'WORD' && normalizedStage !== 'IDIOM') return false;
  if (!phone || !isWhatsAppRecipientAllowed_(phone)) return false;

  const active = activeConversation_(phone);
  if (active && !(
      active.kind === 'GRAMMAR' &&
      Number(active.meetingNo) === Number(meetingNo) &&
      String(active.step || '').toUpperCase() === normalizedStage
    )) return false;

  // This collection flow is intentionally free-form. It should only start
  // while the member has an open WhatsApp customer-service window.
  if (!hasOpenWhatsAppServiceWindow_(phone)) return false;

  const intro = grammarStageIntroMessage_(member, meetingNo, normalizedStage);
  const format = grammarStageCopyMessage_(normalizedStage);
  const baseKey = [
    sourceKind || 'GRAMMAR_STAGE_PROMPT',
    meetingNo,
    'GRM',
    member,
    normalizedStage
  ].join('|');

  const first = sendWhatsAppText_(phone, intro);
  recordWorkflowOutbound_(
    member,
    phone,
    normalizedStage === 'WORD' ? 'GRAMMAR_WORD_PROMPT' : 'GRAMMAR_IDIOM_PROMPT',
    intro,
    meetingNo,
    first,
    baseKey + '|INTRO'
  );
  if (!first.ok) return false;

  const second = sendWhatsAppText_(phone, format);
  recordWorkflowOutbound_(
    member,
    phone,
    normalizedStage === 'WORD' ? 'GRAMMAR_WORD_FORMAT' : 'GRAMMAR_IDIOM_FORMAT',
    format,
    meetingNo,
    second,
    baseKey + '|FORMAT'
  );
  if (!second.ok) return false;

  setConversation_(phone, {
    kind: 'GRAMMAR',
    meetingNo: Number(meetingNo),
    related: 'GRM',
    step: normalizedStage,
    data: {}
  });
  return true;
}

/**
 * Preserve an optional Grammarian image in the configured DEV Drive folder.
 * Each image is a separate submission linked to its meeting, member and stage;
 * it does not complete the Word or Idiom text task.
 */
function handleGrammarImage_(member, phone, message) {
  const active = activeConversation_(phone);
  if (!active || active.kind !== 'GRAMMAR' || !active.meetingNo) {
    sendWhatsAppText_(phone,
      'I received your image, but there is no active Grammarian request for this chat.');
    return;
  }

  const caption = String((message.image && message.image.caption) || '').trim();
  const explicit = /^\\s*(word|idiom)(?:\\b|\\s*:)/i.exec(caption);
  const stage = explicit ? explicit[1].toUpperCase() :
    String(active.step || '').toUpperCase();
  const meetingNo = Number(active.meetingNo);
  const mediaId = String((message.image && message.image.id) || '').trim();
  if ((stage !== 'WORD' && stage !== 'IDIOM') || !mediaId ||
      !verifyRoleAssignment_(meetingNo, 'GRM', member).ok) {
    sendWhatsAppText_(phone,
      'I could not link that image to your Grammarian assignment. Please try again.');
    return;
  }

  // After the Word has been submitted, an uncaptioned image could belong
  // either to that Word or to the upcoming Idiom. Ask instead of guessing.
  if (!explicit && stage === 'IDIOM') {
    active.data = active.data || {};
    active.data.pendingImage = {
      mediaId: mediaId,
      mimeType: String(message.image.mime_type || ''),
      caption: caption,
      messageId: String(message.id || '')
    };
    setConversation_(phone, active);
    sendWhatsAppText_(phone, 'Is this image for the Word or the Idiom? Reply WORD or IDIOM.');
    return;
  }

  saveGrammarImageSubmission_(member, phone, meetingNo, stage, {
    mediaId: mediaId,
    mimeType: String(message.image.mime_type || ''),
    caption: caption,
    messageId: String(message.id || '')
  });
}

function saveGrammarImageSubmission_(member, phone, meetingNo, stage, image) {
  try {
    const link = saveGrammarImage_(image.mediaId, image.mimeType,
      meetingNo, member, stage, image.messageId);
    const appended = appendSubmission_({
      meetingNo: meetingNo,
      member: member,
      type: 'GRAMMAR_' + stage + '_IMAGE',
      related: 'GRM',
      values: [image.caption, '', '', ''],
      fileOrLink: link,
      sourceMessageId: image.messageId
    });
    if (appended) {
      sendWhatsAppText_(phone,
        'Image saved for the ' + (stage === 'WORD' ? 'Word' : 'Idiom') +
        ' of the Day, ATK Meeting ' + meetingNo + ' ✅');
      logAutomation_('Grammarian', meetingNo, member,
        stage + ' image saved', 'Success', link);
    }
  } catch (error) {
    logAutomation_('Grammarian', meetingNo, member,
      stage + ' image save failed', 'Error', String(error));
    sendWhatsAppText_(phone,
      'I received the image but could not save it. Please resend it later.');
  }
}

function saveGrammarImage_(mediaId, mimeHint, meetingNo, member, stage, messageId) {
  const folderId = String(configValue_('GRAMMAR_IMAGE_FOLDER_ID', '')).trim();
  if (!folderId) throw new Error('GRAMMAR_IMAGE_FOLDER_ID is missing.');
  const token = PropertiesService.getScriptProperties()
    .getProperty('META_ACCESS_TOKEN');
  if (!token) throw new Error('META_ACCESS_TOKEN is missing.');

  const metadataResponse = UrlFetchApp.fetch(
    'https://graph.facebook.com/v26.0/' + encodeURIComponent(mediaId), {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
  if (metadataResponse.getResponseCode() !== 200) {
    throw new Error('Meta media metadata HTTP ' + metadataResponse.getResponseCode());
  }
  const metadata = JSON.parse(metadataResponse.getContentText());
  if (!metadata.url || !/^image\\//i.test(String(metadata.mime_type || mimeHint || ''))) {
    throw new Error('WhatsApp media is not an image.');
  }
  if (Number(metadata.file_size || 0) > 10 * 1024 * 1024) {
    throw new Error('WhatsApp image exceeds 10 MB.');
  }

  const download = UrlFetchApp.fetch(metadata.url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  if (download.getResponseCode() !== 200) {
    throw new Error('Meta media download HTTP ' + download.getResponseCode());
  }
  const mime = String(metadata.mime_type || mimeHint || 'image/jpeg');
  const extension = mime === 'image/png' ? 'png' :
    mime === 'image/webp' ? 'webp' : 'jpg';
  const safeMember = String(member || '').replace(/[^A-Za-z0-9_-]/g, '_');
  const safeMessage = String(messageId || mediaId).replace(/[^A-Za-z0-9_-]/g, '_');
  const filename = 'ATK_' + meetingNo + '_' + safeMember + '_' +
    stage + '_' + safeMessage + '.' + extension;
  const blob = download.getBlob().setName(filename);
  const file = DriveApp.getFolderById(folderId).createFile(blob);
  return file.getUrl();
}

function generateGrammarMeaningUsage_(stage, item, meetingNo) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = String(props.getProperty('OPENAI_API_KEY') || '').trim();
  const model = String(props.getProperty('GRAMMAR_AI_MODEL') || 'gpt-6-luna').trim();
  if (!apiKey) return {ok:false, error:'OPENAI_API_KEY missing'};

  const label = String(stage || '').toUpperCase() === 'IDIOM' ? 'idiom' : 'word';
  const theme = meetingTheme_(meetingNo);
  const instructions =
    'You create concise Toastmasters Grammarian content. Return JSON only with keys meaning and usage. ' +
    'Meaning must be plain, accurate English. Usage must be one natural sentence that clearly demonstrates the ' +
    label + '. Do not add markdown or commentary.';
  const input =
    'Meeting theme: ' + (theme || 'not provided') + '\n' +
    (label === 'idiom' ? 'Idiom: ' : 'Word: ') + String(item || '').trim();

  try {
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
      method: 'post',
      contentType: 'application/json',
      headers: {Authorization: 'Bearer ' + apiKey},
      payload: JSON.stringify({
        model: model,
        instructions: instructions,
        input: input,
        text: {format: {type:'json_schema', name:'grammar_content', strict:true,
          schema:{type:'object', properties:{
            meaning:{type:'string'}, usage:{type:'string'}
          }, required:['meaning','usage'], additionalProperties:false}}}
      }),
      muteHttpExceptions: true
    });
    const status = response.getResponseCode();
    if (status < 200 || status >= 300) {
      let detail = '';
      try {
        const errorBody = JSON.parse(response.getContentText() || '{}');
        detail = String(
          errorBody && errorBody.error &&
          (errorBody.error.code || errorBody.error.type || errorBody.error.message) || ''
        ).trim();
      } catch (ignore) {}
      return {ok:false, error:'AI HTTP ' + status + (detail ? ' | ' + detail : '')};
    }
    const data = JSON.parse(response.getContentText());
    let output = String(data.output_text || '').trim();
    if (!output && data.output) {
      data.output.forEach(part => (part.content || []).forEach(content => {
        if (!output && content.text) output = String(content.text).trim();
      }));
    }
    const parsed = JSON.parse(output);
    const meaning = String(parsed.meaning || '').trim();
    const usage = String(parsed.usage || '').trim();
    if (!meaning || !usage) return {ok:false, error:'AI returned incomplete content'};
    return {ok:true, meaning:meaning, usage:usage, model:model};
  } catch (error) {
    return {ok:false, error:String(error)};
  }
}

function sendGrammarApprovalButtons_(member, phone, meetingNo, stage, item, meaning, usage) {
  const normalizedStage = String(stage || '').toUpperCase();
  const body = grammarApprovalMessage_(normalizedStage, item, meaning, usage);
  const payload = {
    messaging_product:'whatsapp', recipient_type:'individual', to:phone,
    type:'interactive',
    interactive:{type:'button', body:{text:body}, action:{buttons:[
      {type:'reply', reply:{id:['GRAMMAR_APPROVAL',meetingNo,normalizedStage,'APPROVE'].join('|'), title:'Approve'}},
      {type:'reply', reply:{id:['GRAMMAR_APPROVAL',meetingNo,normalizedStage,'EDIT'].join('|'), title:'Edit'}}
    ]}}
  };
  const result = sendWhatsAppPayload_(payload);
  recordWorkflowOutbound_(member, phone, 'GRAMMAR_' + normalizedStage + '_APPROVAL',
    body, meetingNo, result,
    ['GRAMMAR_APPROVAL',meetingNo,normalizedStage,member,item].join('|'));
  if (result && result.ok) {
    setConversation_(phone, {kind:'GRAMMAR', meetingNo:Number(meetingNo), related:'GRM',
      step:normalizedStage + '_APPROVAL',
      data:{item:item, meaning:meaning, usage:usage}});
  }
  return result;
}

function approvedGrammarBundle_(meetingNo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('WA Submissions');
  if (!sheet || sheet.getLastRow() < 2) return null;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
  const out = {};
  rows.forEach(row => {
    if (Number(row[1]) !== Number(meetingNo) || String(row[4] || '').trim() !== 'GRM') return;
    const type = String(row[3] || '').trim().toUpperCase();
    if (type === 'GRAMMAR_WORD') out.word = {item:String(row[5]||'').trim(), meaning:String(row[6]||'').trim(), usage:String(row[7]||'').trim()};
    if (type === 'GRAMMAR_IDIOM') out.idiom = {item:String(row[5]||'').trim(), meaning:String(row[6]||'').trim(), usage:String(row[7]||'').trim()};
  });
  return out.word && out.idiom ? out : null;
}

function grammarReviewMessage_(bundle) {
  return (
    'Thank you! Your Grammarian details are ready. ✅\n\n' +
    'Please review:\n\n' +
    '*WORD*\n' +
    'Word: ' + bundle.word.item + '\n' +
    'Meaning: ' + bundle.word.meaning + '\n' +
    'Usage: ' + bundle.word.usage + '\n\n' +
    '*IDIOM*\n' +
    'Idiom: ' + bundle.idiom.item + '\n' +
    'Meaning: ' + bundle.idiom.meaning + '\n' +
    'Usage: ' + bundle.idiom.usage
  );
}

function sendGrammarReviewButtons_(member, phone, meetingNo, keyKind) {
  const bundle = approvedGrammarBundle_(meetingNo);
  if (!bundle || !phone || !isWhatsAppRecipientAllowed_(phone)) return false;

  const key = [
    String(keyKind || 'GRAMMAR_REVIEW'),
    meetingNo,
    member,
    bundle.word.item,
    bundle.word.meaning,
    bundle.word.usage,
    bundle.idiom.item,
    bundle.idiom.meaning,
    bundle.idiom.usage
  ].join('|');
  if (hasReminderBeenSent_(key)) return false;

  const body = grammarReviewMessage_(bundle);
  const payload = {
    messaging_product:'whatsapp',
    recipient_type:'individual',
    to:phone,
    type:'interactive',
    interactive:{
      type:'button',
      body:{text:body},
      action:{buttons:[
        {type:'reply', reply:{id:['GRAMMAR_REVIEW',meetingNo,'CONFIRM'].join('|'), title:'Confirm'}},
        {type:'reply', reply:{id:['GRAMMAR_REVIEW',meetingNo,'EDIT_WORD'].join('|'), title:'Edit Word'}},
        {type:'reply', reply:{id:['GRAMMAR_REVIEW',meetingNo,'EDIT_IDIOM'].join('|'), title:'Edit Idiom'}}
      ]}
    }
  };

  const result = sendWhatsAppPayload_(payload);
  recordWorkflowOutbound_(
    member, phone, 'GRAMMAR_REVIEW', body, meetingNo, result, key
  );

  if (result && result.ok) {
    setConversation_(phone, {
      kind:'GRAMMAR',
      meetingNo:Number(meetingNo),
      related:'GRM',
      step:'REVIEW',
      data:{}
    });
  }
  return Boolean(result && result.ok);
}

function handleGrammarReviewReply_(member, phone, meetingNo, choice, messageId) {
  const assignment = verifyRoleAssignment_(meetingNo, 'GRM', member);
  if (!assignment.ok) {
    clearConversation_(phone);
    sendWhatsAppText_(phone,
      'Your Grammarian assignment for Meeting ' + meetingNo +
      ' has changed, so I did not save this.');
    syncPhase1Tasks_();
    return;
  }

  const normalizedChoice = String(choice || '').toUpperCase();
  if (normalizedChoice === 'CONFIRM') {
    appendSubmission_({
      meetingNo:meetingNo,
      member:member,
      type:'GRAMMAR_CONFIRMED',
      related:'GRM',
      values:['Confirmed','','',''],
      sourceMessageId:messageId || ''
    });
    clearConversation_(phone);
    sendWhatsAppText_(phone,
      'Thank you. Your Grammarian details are confirmed ✅');
    // Do not echo the compiled handoff back to the Grammarian who just confirmed it.
    // Other handoff recipients (for example TMOD/VPE) still receive it.
    sendApprovedGrammarBundle_(meetingNo, member);
    syncPhase1Tasks_();
    logAutomation_('Grammarian', meetingNo, member,
      'Final Grammarian review confirmed', 'Success', '');
    return;
  }

  const stage = normalizedChoice === 'EDIT_IDIOM' ? 'IDIOM' :
    normalizedChoice === 'EDIT_WORD' ? 'WORD' : '';
  if (!stage) return;

  setConversation_(phone, {
    kind:'GRAMMAR',
    meetingNo:Number(meetingNo),
    related:'GRM',
    step:'EDIT_' + stage,
    data:{}
  });
  sendWhatsAppText_(phone,
    'Please copy the next message, make the changes, and send it back.');
  sendWhatsAppText_(phone, grammarStageCopyMessage_(stage, true));
}

function grammarHandoffRecipients_(meetingNo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  const recipients = [];
  if (planner && planner.getLastRow() >= 3) {
    const ids = planner.getRange(3, 1, Math.max(0, planner.getLastRow() - 2), 1).getValues().flat();
    const idx = ids.findIndex(v => Number(v) === Number(meetingNo));
    if (idx >= 0) {
      const row = planner.getRange(idx + 3, 1, 1, 28).getValues()[0];
      const tmod = String(row[18] || '').trim();
      if (tmod) recipients.push({member:tmod, role:'Toastmaster of the Day'});
    }
  }
  const vpe = String(configValue_('VPE_MEMBER_NAME', 'Riyas AF') || '').trim();
  if (vpe && !recipients.some(r => r.member === vpe)) recipients.push({member:vpe, role:'VP Education'});
  return recipients;
}

function grammarBundleIsFinalConfirmed_(meetingNo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('WA Submissions');
  if (!sheet || sheet.getLastRow() < 2) return false;

  return sheet.getRange(2, 2, sheet.getLastRow() - 1, 4).getValues()
    .some(row =>
      Number(row[0]) === Number(meetingNo) &&
      String(row[2] || '').trim() === 'GRAMMAR_CONFIRMED' &&
      String(row[3] || '').trim() === 'GRM'
    );
}

function sendApprovedGrammarBundle_(meetingNo, excludeMember) {
  const bundle = approvedGrammarBundle_(meetingNo);
  if (!bundle || !grammarBundleIsFinalConfirmed_(meetingNo)) return false;
  const body =
    'The Grammarian details for ATK Meeting ' + meetingNo + ' are ready.\n\n' +
    'Word of the Day: ' + bundle.word.item + '\n' +
    'Meaning: ' + bundle.word.meaning + '\n' +
    'Usage: ' + bundle.word.usage + '\n\n' +
    'Idiom of the Day: ' + bundle.idiom.item + '\n' +
    'Meaning: ' + bundle.idiom.meaning + '\n' +
    'Usage: ' + bundle.idiom.usage + '\n\n' +
    'These are the final details approved by the Grammarian.';
  let anySent = false;
  grammarHandoffRecipients_(meetingNo).forEach(recipient => {
    if (excludeMember &&
        String(recipient.member || '').trim() === String(excludeMember || '').trim()) return;
    const phone = normalizePhone_(phoneForMember_(recipient.member));
    if (!phone || !isWhatsAppRecipientAllowed_(phone)) return;
    const key = ['GRAMMAR_HANDOFF', meetingNo, recipient.member,
      bundle.word.item, bundle.idiom.item].join('|');
    if (hasReminderBeenSent_(key)) return;
    const personalized = 'Hi ' + recipient.member + ',\n\n' + body;
    const result = sendAutomatedText_(phone, 'GRAMMAR_HANDOFF', personalized,
      [recipient.member, meetingNo, bundle.word.item, bundle.idiom.item]);
    recordWorkflowOutbound_(recipient.member, phone, 'GRAMMAR_HANDOFF',
      personalized, meetingNo, result, key);
    logAutomation_('Grammarian', meetingNo, recipient.member,
      'Approved Word + Idiom shared with ' + recipient.role,
      result && result.ok ? 'Sent' : (result && result.blocked ? 'Blocked' : 'Error'),
      bundle.word.item + ' / ' + bundle.idiom.item);
    if (result && result.ok) anySent = true;
  });
  return anySent;
}

function saveApprovedGrammarStage_(member, phone, meetingNo, stage, details, sourceMessageId, returnToReview) {
  const normalizedStage = String(stage || '').toUpperCase();
  const saved = appendSubmission_({
    meetingNo:meetingNo, member:member,
    type:normalizedStage === 'WORD' ? 'GRAMMAR_WORD' : 'GRAMMAR_IDIOM',
    related:'GRM',
    values:[details.item, details.meaning, details.usage, ''],
    sourceMessageId:sourceMessageId || ''
  });

  if (saved !== false) {
    logAutomation_('Grammarian', meetingNo, member,
      (normalizedStage === 'WORD' ? 'Word' : 'Idiom') + ' received',
      'Success', details.item);
  }

  if (returnToReview) {
    sendGrammarReviewButtons_(member, phone, meetingNo);
    syncPhase1Tasks_();
    return;
  }

  if (normalizedStage === 'WORD') {
    setConversation_(phone, {kind:'GRAMMAR', meetingNo:Number(meetingNo), related:'GRM',
      step:'IDIOM', data:{}});
    sendGrammarStagePrompt_(member, phone, meetingNo, 'IDIOM', 'GRAMMAR_AFTER_WORD');
  } else {
    sendGrammarReviewButtons_(member, phone, meetingNo);
  }
  syncPhase1Tasks_();
}

function handleGrammarApprovalReply_(member, phone, meetingNo, stage, choice, messageId) {
  const active = activeConversation_(phone);
  const normalizedStage = String(stage || '').toUpperCase();
  if (!active || active.kind !== 'GRAMMAR' ||
      Number(active.meetingNo) !== Number(meetingNo) ||
      String(active.step || '').toUpperCase() !== normalizedStage + '_APPROVAL') {
    sendWhatsAppText_(phone, 'That Grammarian approval is no longer active.');
    return;
  }
  if (String(choice || '').toUpperCase() === 'APPROVE') {
    saveApprovedGrammarStage_(member, phone, meetingNo, normalizedStage, active.data || {}, messageId);
    return;
  }
  if (String(choice || '').toUpperCase() === 'EDIT') {
    active.step = normalizedStage + '_EDIT';
    setConversation_(phone, active);
    sendWhatsAppText_(phone, grammarApprovalEditFormat_(normalizedStage));
  }
}

function handleGrammarStageSubmission_(member, phone, requestedMeetingNo, stage, body, messageId) {
  if (!configBool_('GRAMMARIAN_WORKFLOW_ENABLED')) {
    sendWhatsAppText_(phone, 'Grammarian collection is not active yet.'); return;
  }

  const task = findPendingTaskForMember_('GRAMMAR', member, requestedMeetingNo);
  if (!task) {
    sendWhatsAppText_(phone, 'I could not find a pending Grammarian submission for you.');
    return;
  }

  const normalizedStage = String(stage || '').toUpperCase();
  if (normalizedStage !== 'WORD' && normalizedStage !== 'IDIOM') return;

  const assignment = verifyRoleAssignment_(task.meetingNo, 'GRM', member);
  if (!assignment.ok) {
    clearConversation_(phone);
    sendWhatsAppText_(phone, 'Your Grammarian assignment for Meeting ' + task.meetingNo +
      ' has changed, so I did not save this.');
    syncPhase1Tasks_();
    return;
  }

  const active = activeConversation_(phone);
  const parsed = parseGrammarStageBlock_(body, normalizedStage);
  const rawStep = active && active.kind === 'GRAMMAR' &&
    Number(active.meetingNo) === Number(task.meetingNo)
      ? String(active.step || '').toUpperCase() : '';
  const editing = rawStep === 'EDIT_' + normalizedStage;
  const previous = rawStep.indexOf(normalizedStage) !== -1
    ? (active.data || {}) : {};

  if (parsed.oppositeStage) {
    const wanted = normalizedStage === 'WORD' ? 'Word' : 'Idiom';
    const received = parsed.oppositeStage === 'WORD' ? 'Word' : 'Idiom';
    sendWhatsAppText_(
      phone,
      received + ' details were received, but I am currently waiting for the ' +
      wanted + ' of the Day.\n\n' + grammarStageCopyMessage_(normalizedStage)
    );
    return;
  }

  const details = {
    item: String(parsed.item || previous.item || '').trim(),
    meaning: String(parsed.meaning || previous.meaning || '').trim(),
    usage: String(parsed.usage || previous.usage || '').trim()
  };
  details.recognized = [details.item, details.meaning, details.usage].filter(Boolean).length;

  // Manual-first canonical flow: never depend on AI. Preserve partial input and
  // ask only for the fields still missing.
  if (details.recognized < 3) {
    setConversation_(phone, {
      kind:'GRAMMAR',
      meetingNo:Number(task.meetingNo),
      related:'GRM',
      step:normalizedStage,
      data:details
    });
    sendWhatsAppText_(phone, grammarMissingFieldsMessage_(normalizedStage, details));
    return;
  }

  saveApprovedGrammarStage_(
    member, phone, task.meetingNo, normalizedStage, details, messageId, editing
  );
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

  const combinedGrammarBlock = parseGrammarCombinedBlock_(text);
  if (combinedGrammarBlock.recognized >= 6) {
    if (handleCombinedGrammarSubmission_(
      member, phone, 0, combinedGrammarBlock, messageId
    )) return true;
  }

  const grammarBlock = parseGrammarDetailsBlock_(text);
  if (grammarBlock.recognized >= 4) {
    handleGrammarSubmission_(member, phone, 0, grammarBlock, messageId);
    return true;
  }

  const pendingGrammarTask = findPendingTaskForMember_('GRAMMAR', member, 0);
  if (pendingGrammarTask) {
    const grammarStage = grammarStageForMember_(pendingGrammarTask.meetingNo, member);
    if (grammarStage === 'WORD' || grammarStage === 'IDIOM') {
      const stageBlock = parseGrammarStageBlock_(text, grammarStage);
      if (stageBlock.item) {
        handleGrammarStageSubmission_(
          member, phone, pendingGrammarTask.meetingNo, grammarStage, stageBlock, messageId);
        return true;
      }
    }
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

function introChoiceMessage_(member, meetingNo, isReminder) {
  if (isReminder) {
    return (
      'Hey ' + member + ',\n\n' +
      'Just a reminder about your introduction for ATK Meeting ' + meetingNo + '.\n\n' +
      'Anything new to add — a recognition, achievement, update, or something we may not know about you?\n\n' +
      'If not, we’ll use your saved introduction.\n\n' +
      'Thank you!'
    );
  }

  return (
    'Hey ' + member + ',\n\n' +
    'For the upcoming ATK Meeting ' + meetingNo + ', we already have your saved introduction.\n\n' +
    'But to avoid repeating the same introduction every time, is there anything new with you that we can add?\n\n' +
    'A new recognition, achievement, update, or simply something about you that we may not know yet.\n\n' +
    'Anything that comes to your mind will do. 😊\n\n' +
    'Thank you!'
  );
}

function sendIntroChoiceButtons_(member, phone, meetingNo, related, isReminder) {
  const saved = defaultIntro_(member);
  if (!saved || !phone) return null;

  const text = introChoiceMessage_(member, meetingNo, Boolean(isReminder));
  const payload = {
    messaging_product: 'whatsapp',
    recipient_type: 'individual',
    to: phone,
    type: 'interactive',
    interactive: {
      type: 'button',
      body: { text: text },
      action: {
        buttons: [
          {
            type: 'reply',
            reply: {
              id: ['INTRO_CHOICE', meetingNo, related, 'NEW'].join('|'),
              title: 'I have something new'
            }
          },
          {
            type: 'reply',
            reply: {
              id: ['INTRO_CHOICE', meetingNo, related, 'DEFAULT'].join('|'),
              title: 'Use my default'
            }
          }
        ]
      }
    }
  };

  const result = sendWhatsAppPayload_(payload);
  if (result.ok) {
    setConversation_(phone, {
      kind: 'INTRO',
      meetingNo: Number(meetingNo),
      related: String(related || ''),
      step: 'CHOICE',
      data: {}
    });
  }

  return { result: result, text: text };
}

function handleIntroChoiceReply_(member, phone, interaction, messageId) {
  if (!configBool_('INTRO_WORKFLOW_ENABLED')) return;

  const meetingNo = Number(interaction.meetingNo || 0);
  const related = String(interaction.related || '');
  const choice = String(interaction.choice || '').toUpperCase();
  const task = findPendingTaskForMember_('INTRO', member, meetingNo, related);

  if (!task) {
    sendWhatsAppText_(phone, 'Your introduction for this meeting is already settled. Thank you.');
    return;
  }

  const saved = defaultIntro_(member);
  if (!saved) {
    clearConversation_(phone);
    const request = startGuidedRequest_(member, phone, 'INTRO', meetingNo, related, 'role player');
    if (request) {
      const result = sendWhatsAppText_(phone, request.text);
      if (result.ok) setConversation_(phone, request.state);
    }
    return;
  }

  if (choice === 'DEFAULT') {
    const active = activeConversation_(phone);
    if (active && active.kind === 'INTRO' &&
        Number(active.meetingNo) === meetingNo &&
        String(active.related || '') === related) {
      clearConversation_(phone);
    }
    handleIntroSubmission_(member, phone, meetingNo, saved, messageId, related);
    return;
  }

  if (choice === 'NEW') {
    setConversation_(phone, {
      kind: 'INTRO',
      meetingNo: meetingNo,
      related: related,
      step: 'MEETING_INTRO',
      data: {}
    });
    sendWhatsAppText_(phone, 'Great! Just send it here in your own words. 😊');
  }
}

function finalizeIntroWithDefault_(member, phone, meetingNo, related) {
  const saved = defaultIntro_(member);
  if (!saved) return false;

  const task = findPendingTaskForMember_('INTRO', member, meetingNo, related);
  if (!task) return false;

  const assignment = /^S[123]$/.test(related)
    ? verifySpeakerAssignment_(meetingNo, related, member)
    : verifyRoleAssignment_(meetingNo, related, member);

  if (!assignment.ok) {
    syncPhase1Tasks_();
    return false;
  }

  const appended = appendSubmission_({
    meetingNo: meetingNo,
    member: member,
    type: 'INTRO',
    related: related,
    values: [saved, '', '', ''],
    sourceMessageId: ''
  });

  if (!appended) return false;

  const active = activeConversation_(phone);
  if (active && active.kind === 'INTRO' &&
      Number(active.meetingNo) === Number(meetingNo) &&
      String(active.related || '') === String(related || '')) {
    clearConversation_(phone);
  }

  syncPhase1Tasks_();
  logAutomation_(
    'Introduction',
    meetingNo,
    member,
    'Saved default introduction used automatically after no response',
    'Success',
    related
  );
  return true;
}

function processIntroChoiceTimeouts_() {
  if (!configBool_('INTRO_WORKFLOW_ENABLED')) return;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tasks = ss.getSheetByName('Automation Tasks');
  const outbox = ss.getSheetByName('WA Outbox');
  if (!tasks || !outbox || tasks.getLastRow() < 2 || outbox.getLastRow() < 2) return;

  const reminderAfterHours = Math.max(
    1,
    Number(configValue_('INTRO_REMINDER_AFTER_HOURS', '23')) || 23
  );
  const defaultAfterHours = Math.max(
    reminderAfterHours,
    Number(configValue_('INTRO_DEFAULT_AFTER_HOURS', '24')) || 24
  );

  const now = new Date();
  const successfulStates = new Set(['SENT', 'DELIVERED', 'READ', 'ACCEPTED']);
  const taskRows = tasks.getRange(2, 1, tasks.getLastRow() - 1, 18).getValues();
  const outboxRows = outbox.getRange(2, 1, outbox.getLastRow() - 1, 16).getValues();

  taskRows.forEach(row => {
    if (String(row[2] || '') !== 'INTRO' || String(row[5] || '') !== 'Pending') return;

    const meetingNo = Number(row[1]);
    const member = String(row[3] || '').trim();
    const related = String(row[4] || '').trim();
    if (!meetingNo || !member || !related || !defaultIntro_(member)) return;

    const promptKeys = new Set([
      ['ASSIGNMENT_PROMPT', meetingNo, related, member].join('|'),
      ['INTRO_PROMPT', meetingNo, related, member].join('|'),
      ['INTRO_REQUEST', meetingNo, related, member].join('|')
    ]);

    const promptAt = outboxRows
      .filter(r =>
        promptKeys.has(String(r[15] || '')) &&
        successfulStates.has(String(r[6] || '').toUpperCase()) &&
        r[12] instanceof Date
      )
      .map(r => r[12])
      .sort((a, b) => a.getTime() - b.getTime())[0];

    if (!(promptAt instanceof Date)) return;

    const elapsedHours = (now.getTime() - promptAt.getTime()) / 3600000;
    const phone = phoneForMember_(member);

    if (elapsedHours >= defaultAfterHours) {
      finalizeIntroWithDefault_(member, phone, meetingNo, related);
      return;
    }

    if (elapsedHours < reminderAfterHours) return;

    const reminderKey = ['INTRO_CHOICE_REMINDER', meetingNo, related, member].join('|');
    if (hasReminderBeenSent_(reminderKey)) return;
    if (!phone || !isWhatsAppRecipientAllowed_(phone) || !hasOpenWhatsAppServiceWindow_(phone)) return;

    const sent = sendIntroChoiceButtons_(member, phone, meetingNo, related, true);
    if (!sent) return;

    recordWorkflowOutbound_(
      member,
      phone,
      'INTRO_CHOICE_REMINDER',
      sent.text,
      meetingNo,
      sent.result,
      reminderKey
    );
  });
}


function sendImmediateAssignmentPrompts_(meetingNo, roleDef, member, phone) {
  if (roleDef.code === 'TMOD' && !configBool_('TMOD_WORKFLOW_ENABLED')) return;
  if (roleDef.code === 'GRM' && !configBool_('GRAMMARIAN_WORKFLOW_ENABLED')) return;
  if (roleDef.code !== 'TMOD' && roleDef.code !== 'GRM' &&
      !configBool_('INTRO_WORKFLOW_ENABLED')) return;
  const kind = roleDef.code === 'TMOD' ? 'THEME' :
    roleDef.code === 'GRM' ? 'GRAMMAR' : 'INTRO';

  if (kind === 'GRAMMAR') {
    sendGrammarStagePrompt_(member, phone, meetingNo, 'WORD', 'ASSIGNMENT_PROMPT');
    return;
  }

  if (kind === 'INTRO' && defaultIntro_(member)) {
    const sent = sendIntroChoiceButtons_(
      member, phone, meetingNo, roleDef.code, false);
    if (!sent) return;
    recordWorkflowOutbound_(
      member,
      phone,
      'ASSIGNMENT_PROMPT',
      sent.text,
      meetingNo,
      sent.result,
      ['ASSIGNMENT_PROMPT', meetingNo, roleDef.code, member].join('|')
    );
    return;
  }

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
  const lastRow = sheet.getLastRow();
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

  const lastRow = sheet.getLastRow();
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


function sendThemeAndTmodToVppr_(meetingNo) {
  if (!configBool_('VPPR_THEME_HANDOFF_ENABLED')) return false;

  const vppr = String(configValue_('VPPR_MEMBER_NAME', '') || '').trim();
  if (!vppr) return false;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  if (!planner || planner.getLastRow() < 3) return false;

  const ids = planner.getRange(3, 1, Math.max(0, planner.getLastRow() - 2), 1)
    .getValues().flat();
  const idx = ids.findIndex(v => Number(v) === Number(meetingNo));
  if (idx < 0) return false;

  const row = planner.getRange(idx + 3, 1, 1, 28).getValues()[0];
  const theme = String(row[3] || '').trim();
  const tmod = String(row[18] || '').trim();
  if (!theme || !tmod) return false;

  const phone = normalizePhone_(phoneForMember_(vppr));
  if (!/^\d{10,15}$/.test(phone || '')) return false;
  if (!isWhatsAppRecipientAllowed_(phone)) return false;

  const key = ['VPPR_THEME_HANDOFF', meetingNo, vppr, theme, tmod].join('|');
  if (hasReminderBeenSent_(key)) return false;

  const body =
    'Hi ' + vppr + ',\n\n' +
    'ATK Meeting ' + meetingNo + ' update for promotion:\n\n' +
    'Theme: ' + theme + '\n' +
    'Toastmaster of the Day: ' + tmod + '\n\n' +
    'Please use these details for the meeting promotion/poster.\n\n' +
    'Thank you.';

  const result = sendAutomatedText_(
    phone,
    'TMOD_UPDATE',
    body,
    [
      vppr,
      meetingNo,
      'For VPPR promotion — Theme: ' + theme +
        ' | Toastmaster of the Day: ' + tmod
    ]
  );

  recordWorkflowOutbound_(
    vppr,
    phone,
    'VPPR_THEME_HANDOFF',
    body,
    meetingNo,
    result,
    key
  );

  logAutomation_(
    'VPPR Handoff',
    meetingNo,
    vppr,
    'Theme + TMOD shared with VPPR',
    result.ok ? 'Sent' : (result.blocked ? 'Blocked' : 'Error'),
    'Theme=' + theme + ' | TMOD=' + tmod +
      ' | channel=' + (result.channelMode || 'UNKNOWN')
  );

  return Boolean(result && result.ok);
}


function processVpprThemeHandoffs_() {
  if (!configBool_('VPPR_THEME_HANDOFF_ENABLED')) return;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  if (!planner || planner.getLastRow() < 3) return;

  const today = startOfDay_(new Date());
  const rows = planner.getRange(
    3, 1, Math.max(0, planner.getLastRow() - 2), 28
  ).getValues();

  rows.forEach(row => {
    const meetingNo = Number(row[0]);
    const meetingDate = row[1];
    const status = String(row[2] || '').trim();
    const theme = String(row[3] || '').trim();
    const tmod = String(row[18] || '').trim();

    if (!meetingNo || !(meetingDate instanceof Date)) return;
    if (startOfDay_(meetingDate) < today) return;
    if (['Completed', 'Cancelled'].indexOf(status) !== -1) return;
    if (!theme || !tmod) return;

    sendThemeAndTmodToVppr_(meetingNo);
  });
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
  sendThemeAndTmodToVppr_(task.meetingNo);
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
    values: parts.slice(0, 5),
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

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return { ok: false };
  const ids = sheet.getRange(3, 1, lastRow - 2, 1).getValues().flat();
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
    '',
    values[4] || ''
  ]);

  // The fifth submission value is used for speaker Time. Sheets can coerce
  // strings such as "07:30" into a Date/time serial, which breaks a later
  // TMOD catch-up replay. Force that cell back to plain text after append.
  if (values.length >= 5 && values[4] !== '' && values[4] !== null &&
      values[4] !== undefined) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 15)
      .setNumberFormat('@')
      .setValue(String(values[4]));
  }

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
    3, 1, Math.max(0, planner.getLastRow() - 2), 1
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
    const missing = ['path', 'level', 'project', 'title', 'time']
      .filter(key => !String(body[key] || '').trim());

    if (missing.length) {
      sendWhatsAppText_(phone,
        'Please add the missing detail' + (missing.length > 1 ? 's' : '') +
        ': ' + missing.map(key => key.charAt(0).toUpperCase() + key.slice(1)).join(', ') +
        '\n\nPath:\nLevel:\nProject:\nTitle:\nTime:');
      return;
    }

    parts = [
      String(body.title || '').trim(),
      String(body.path || '').trim(),
      String(body.level || '').trim(),
      String(body.project || '').trim(),
      String(body.time || '').trim()
    ];
  } else {
    parts = String(body || '').split('|').map(v => v.trim());
    if (parts.length < 5 || parts.slice(0, 5).some(v => !v)) {
      sendWhatsAppText_(phone,
        'Please use this format:\n\nPath:\nLevel:\nProject:\nTitle:\nTime:');
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
    values: parts.slice(0, 5),
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
    'Title: ' + parts[0] + '\n' +
    'Time: ' + parts[4]);

  logAutomation_(
    'Speaker Details',
    task.meetingNo,
    member,
    'Speech details received',
    'Success',
    task.related
  );

  sendSpeakerDetailsToTmod_(task.meetingNo, member, task.related, parts);
  if (configBool_('INTRO_WORKFLOW_ENABLED')) {
    promptSpeakerIntro_(member, phone, task.meetingNo, task.related);
  }
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

  const lastRow = sheet.getLastRow();
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
  const blockFlow = activeTemplate === 'atk_speaker_details_v3';
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
        const body = activeTemplate === 'atk_speaker_details_v3'
          ? (
              'Hi ' + member + ',\n\n' +
              'You’re confirmed as a speaker for ATK Meeting ' + meetingNo + ' 🎤\n\n' +
              'Please share your speech details in the format below. Here’s an example:\n\n' +
              'Path: Presentation Mastery\n' +
              'Level: 1\n' +
              'Project: 3\n' +
              'Title: My Ambition\n' +
              'Time: 7 minutes'
            )
          : (
              'Hi ' + member + ',\n\n' +
              'You’re confirmed as a speaker for ATK Meeting ' + meetingNo + ' 🎤\n\n' +
              'Let’s get your speech details ready for the agenda.\n\n' +
              'Please fill the below details.\n\n' +
              'Thank you.'
            );

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

        const copyBody =
          'Please copy the message below, replace the details with yours, and send it back:\n\n' +
          'Path:\n' +
          'Level:\n' +
          'Project:\n' +
          'Title:\n' +
          'Time:';
        const copyResult = sendWhatsAppText_(phone, copyBody);
        if (!copyResult || !copyResult.ok) return;

        setConversation_(phone, {
          kind: 'SPEAKER_DETAILS',
          meetingNo: meetingNo,
          related: related,
          member: member
        });
        introSent = true;
      }

      if (!introSent) return;

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


function processSpeakerDetailReminders_() {
  if (!configBool_('SPEAKER_DETAILS_WORKFLOW_ENABLED')) return;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tasks = ss.getSheetByName('Automation Tasks');
  const outbox = ss.getSheetByName('WA Outbox');
  if (!tasks || !outbox || tasks.getLastRow() < 2 || outbox.getLastRow() < 2) return;

  const timezone = configValue_('TIMEZONE', 'Asia/Kolkata');
  const now = new Date();
  const localHour = Number(Utilities.formatDate(now, timezone, 'H'));
  const reminderHours = String(
    configValue_('SPEAKER_DETAILS_REMINDER_HOURS', '7,20')
  ).split(',')
    .map(v => Number(String(v).trim()))
    .filter(v => Number.isFinite(v));

  if (reminderHours.indexOf(localHour) === -1) return;

  const successfulStates = new Set(['SENT', 'DELIVERED', 'READ', 'ACCEPTED']);
  const taskRows = tasks.getRange(2, 1, tasks.getLastRow() - 1, 18).getValues();
  const outboxRows = outbox.getRange(2, 1, outbox.getLastRow() - 1, 16).getValues();

  taskRows.forEach(row => {
    if (String(row[2] || '') !== 'SPEAKER_DETAILS' ||
        String(row[5] || '') !== 'Pending') return;

    const meetingNo = Number(row[1]);
    const member = String(row[3] || '').trim();
    const related = String(row[4] || '').trim();
    if (!meetingNo || !member || !related) return;

    const firstKey = ['SPEAKER_DETAILS_PROMPT', meetingNo, related, member].join('|');
    const firstPromptAt = outboxRows
      .filter(r =>
        String(r[15] || '') === firstKey &&
        successfulStates.has(String(r[6] || '').toUpperCase()) &&
        r[12] instanceof Date
      )
      .map(r => r[12])
      .sort((a, b) => a.getTime() - b.getTime())[0];

    if (!(firstPromptAt instanceof Date)) return;

    // Avoid sending the first reminder immediately after the initial request
    // if both happen to land in the same configured reminder hour.
    if ((now.getTime() - firstPromptAt.getTime()) < 60 * 60 * 1000) return;

    const phone = phoneForMember_(member);
    if (!phone || !isWhatsAppRecipientAllowed_(phone)) return;

    const slotDate = Utilities.formatDate(now, timezone, 'yyyyMMdd');
    const reminderKey = [
      'SPEAKER_DETAILS_REMINDER',
      meetingNo,
      related,
      member,
      slotDate,
      'H' + localHour
    ].join('|');

    if (hasReminderBeenSent_(reminderKey)) return;

    const body =
      'Hi ' + member + ',\n\n' +
      'Just a reminder about your speech details for ATK Meeting ' + meetingNo + '.\n\n' +
      'We’re still waiting for:\n\n' +
      'Path:\n' +
      'Level:\n' +
      'Project:\n' +
      'Title:\n' +
      'Time:\n\n' +
      'Please send the pending details when you can.\n\n' +
      'Thank you.';

    const result = sendAutomatedText_(
      phone,
      'SPEAKER_DETAILS',
      body,
      [member, meetingNo]
    );

    recordWorkflowOutbound_(
      member,
      phone,
      'SPEAKER_DETAILS_REMINDER',
      body,
      meetingNo,
      result,
      reminderKey
    );

    logAutomation_(
      'Speaker Details',
      meetingNo,
      member,
      'Speaker details reminder at ' + (localHour === 20 ? '8 PM' : localHour + ' AM'),
      result.ok ? 'Sent' : (result.deferred ? 'Deferred' : 'Error'),
      related
    );
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
    Math.max(0, planner.getLastRow() - 2),
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

  const lastRow = planner.getLastRow();
  if (lastRow < 3) return false;
  const ids = planner.getRange(3, 1, lastRow - 2, 1).getValues().flat();
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
    const ids = planner.getRange(3, 1, Math.max(0, planner.getLastRow() - 2), 1)
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
      if (pair) {
        counterpart = role === 'SPEAKER' ? pair.evaluator : pair.speaker;

        if (answer === 'YES') {
          const slotIndex = pairs.indexOf(pair) + 1;
          appendSubmission_({
            meetingNo: meetingNo,
            member: pair.speaker,
            type: 'EVALUATOR_COORDINATION_CONFIRMED',
            related: 'S' + slotIndex,
            values: [member, role, pair.evaluator, 'YES'],
            sourceMessageId: interaction.messageId || ''
          });
          syncPhase1Tasks_();
        }
      }
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
    const kind = type === 'THEME' ? 'THEME_REQUEST' : 'INTRO_REQUEST';
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

function grammarDetailsRequestBody_(member, meetingNo, isReminder, stage) {
  const currentStage = String(stage || 'WORD').toUpperCase();
  return isReminder
    ? grammarStageReminderMessage_(member, meetingNo, currentStage)
    : grammarStageIntroMessage_(member, meetingNo, currentStage) +
      '\n\n' + grammarStageCopyMessage_(currentStage);
}

/**
 * Catch-up for Grammarians who were entered directly in the planner instead
 * of booking the role through WhatsApp. This intentionally uses the same
 * ASSIGNMENT_PROMPT idempotency key expected by the reminder flow.
 *
 * No Meta template is used. If the WhatsApp customer-service window is not
 * open, this waits rather than sending anything out of policy.
 */
function processMissingGrammarAssignmentPrompts_() {
  if (!configBool_('GRAMMARIAN_WORKFLOW_ENABLED')) return;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tasks = ss.getSheetByName('Automation Tasks');
  if (!tasks || tasks.getLastRow() < 2) return;

  const rows = tasks.getRange(2, 1, tasks.getLastRow() - 1, 18).getValues();

  rows.forEach(row => {
    if (String(row[2] || '') !== 'GRAMMAR' ||
        String(row[5] || '') !== 'Pending') return;

    const meetingNo = Number(row[1]);
    const member = String(row[3] || '').trim();
    const related = String(row[4] || '').trim();
    if (!meetingNo || !member || related !== 'GRM') return;

    const assignment = verifyRoleAssignment_(meetingNo, 'GRM', member);
    if (!assignment.ok) return;

    const stage = grammarStageForMember_(meetingNo, member);
    if (stage === 'DONE') return;

    const phone = phoneForMember_(member);
    if (!phone || !isWhatsAppRecipientAllowed_(phone)) return;
    if (!hasOpenWhatsAppServiceWindow_(phone)) return;

    const active = activeConversation_(phone);
    if (active && !(
      active.kind === 'GRAMMAR' &&
      Number(active.meetingNo) === Number(meetingNo)
    )) return;

    if (stage === 'REVIEW') {
      const sent = sendGrammarReviewButtons_(member, phone, meetingNo);
      if (sent) {
        logAutomation_(
          'Grammarian', meetingNo, member,
          'Missing Grammarian final review catch-up',
          'Sent', 'REVIEW'
        );
      }
      return;
    }

    const assignmentKey = [
      'ASSIGNMENT_PROMPT', meetingNo, related, member, stage, 'INTRO'
    ].join('|');
    if (hasReminderBeenSent_(assignmentKey)) return;

    const sent = sendGrammarStagePrompt_(
      member, phone, meetingNo, stage, 'ASSIGNMENT_PROMPT'
    );

    logAutomation_(
      'Grammarian',
      meetingNo,
      member,
      'Missing Grammarian stage prompt catch-up',
      sent ? 'Sent' : 'Error',
      'Stage=' + stage + ' | free-form only'
    );
  });
}

function processGrammarAssignmentReminder_() {
  if (!configBool_('GRAMMARIAN_WORKFLOW_ENABLED')) return;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tasks = ss.getSheetByName('Automation Tasks');
  if (!tasks || tasks.getLastRow() < 2) return;

  const now = new Date();
  const timezone = configValue_('TIMEZONE', 'Asia/Kolkata');
  const localHour = Number(Utilities.formatDate(now, timezone, 'H'));
  const reminderHours = String(
    configValue_('GRAMMAR_REMINDER_HOURS', '7,20')
  ).split(',')
    .map(v => Number(String(v).trim()))
    .filter(v => Number.isFinite(v));

  if (reminderHours.indexOf(localHour) === -1) return;

  const rows = tasks.getRange(2, 1, tasks.getLastRow() - 1, 18).getValues();
  rows.forEach(row => {
    if (String(row[2] || '') !== 'GRAMMAR' ||
        String(row[5] || '') !== 'Pending') return;

    const meetingNo = Number(row[1]);
    const member = String(row[3] || '').trim();
    const related = String(row[4] || '').trim();
    if (!meetingNo || !member || related !== 'GRM') return;

    const stage = grammarStageForMember_(meetingNo, member);
    if (stage === 'DONE') return;

    const phone = phoneForMember_(member);
    if (!phone || !isWhatsAppRecipientAllowed_(phone)) return;

    const minHours = Math.max(
      1,
      Number(configValue_('GRAMMAR_REMINDER_MIN_HOURS', '2')) || 2
    );

    if (stage === 'REVIEW') {
      const reviewAt = grammarReviewPromptSentAt_(meetingNo, member);
      if (!(reviewAt instanceof Date)) return;
      if ((now.getTime() - reviewAt.getTime()) < minHours * 60 * 60 * 1000) return;
      if (!hasOpenWhatsAppServiceWindow_(phone)) return;

      const slotDate = Utilities.formatDate(now, timezone, 'yyyyMMdd');
      const reminderKind = [
        'GRAMMAR_REVIEW_REMINDER',
        slotDate,
        'H' + localHour
      ].join('|');

      const sent = sendGrammarReviewButtons_(
        member, phone, meetingNo, reminderKind
      );
      if (sent) {
        logAutomation_(
          'Grammarian', meetingNo, member,
          'Grammarian final review reminder at ' +
            (localHour === 20 ? '8 PM' : localHour + ' AM'),
          'Sent', 'REVIEW'
        );
      }
      return;
    }

    const promptAt = grammarStagePromptSentAt_(meetingNo, member, stage);
    if (!(promptAt instanceof Date)) return;
    if ((now.getTime() - promptAt.getTime()) < minHours * 60 * 60 * 1000) return;
    if (!hasOpenWhatsAppServiceWindow_(phone)) return;

    const slotDate = Utilities.formatDate(now, timezone, 'yyyyMMdd');
    const reminderKey = [
      'GRAMMAR_STAGE_REMINDER',
      meetingNo,
      related,
      member,
      stage,
      slotDate,
      'H' + localHour
    ].join('|');
    if (hasReminderBeenSent_(reminderKey)) return;

    const body = grammarStageReminderMessage_(member, meetingNo, stage);
    const result = sendWhatsAppText_(phone, body);

    recordWorkflowOutbound_(
      member,
      phone,
      'GRAMMAR_STAGE_REMINDER',
      body,
      meetingNo,
      result,
      reminderKey
    );

    logAutomation_(
      'Grammarian',
      meetingNo,
      member,
      'Grammarian ' + stage + ' reminder at ' +
        (localHour === 20 ? '8 PM' : localHour + ' AM'),
      result.ok ? 'Sent' : (result.blocked ? 'Blocked' : 'Error'),
      'Free-form split flow'
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
  if (!planner || planner.getLastRow() < 3) return false;

  const ids = planner.getRange(3, 1, planner.getLastRow() - 2, 1)
    .getValues().flat();
  const idx = ids.findIndex(v => Number(v) === Number(meetingNo));
  if (idx < 0) return false;

  const tmod = String(planner.getRange(idx + 3, 19).getValue() || '').trim();
  const phone = normalizePhone_(phoneForMember_(tmod));
  if (!tmod || !phone || !isWhatsAppRecipientAllowed_(phone)) return false;

  const normalizedParts = (parts || []).map(v => String(v || '').trim());
  const speechTime = normalizedParts[4] || 'TBC';
  const slotNumber = Number(String(slotRef || '').replace(/^S/i, ''));
  const speakerLabel = 'Speaker ' +
    (slotNumber || String(slotRef || '').replace(/^S/i, '')) +
    ' — ' + member;

  const body =
    'ATK Meeting ' + meetingNo + '\n\n' +
    speakerLabel + '\n' +
    'Title: ' + normalizedParts[0] + '\n' +
    'Path: ' + normalizedParts[1] + '\n' +
    'Level: ' + normalizedParts[2] + '\n' +
    'Project: ' + normalizedParts[3] + '\n' +
    'Time: ' + speechTime;

  // TMOD is part of the key so a replacement TMOD receives the same stored
  // speaker details even if the previous TMOD already received them.
  const key = [
    'SPEECH_TO_TMOD',
    meetingNo,
    tmod,
    slotRef,
    member,
    normalizedParts.join('|')
  ].join('|');
  if (hasReminderBeenSent_(key)) return false;

  const result = sendAutomatedText_(
    phone,
    'TMOD_UPDATE',
    body,
    [
      tmod,
      meetingNo,
      speakerLabel + '\n' +
        'Title: ' + normalizedParts[0] + '\n' +
        'Path: ' + normalizedParts[1] + '\n' +
        'Level: ' + normalizedParts[2] + '\n' +
        'Project: ' + normalizedParts[3] + '\n' +
        'Time: ' + speechTime
    ]
  );

  recordWorkflowOutbound_(
    tmod, phone, 'SPEECH_TO_TMOD', body, meetingNo, result, key
  );
  return Boolean(result && result.ok);
}

function storedSubmissionText_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' &&
      !isNaN(value.getTime())) {
    return Utilities.formatDate(
      value,
      configValue_('TIMEZONE', 'Asia/Kolkata'),
      'HH:mm'
    );
  }
  return String(value === null || value === undefined ? '' : value).trim();
}

function sendStoredSpeakerDetailsToTmod_(meetingNo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const subs = ss.getSheetByName('WA Submissions');
  if (!subs || subs.getLastRow() < 2) return false;

  let anySent = false;
  subs.getRange(2, 1, subs.getLastRow() - 1, 15).getValues().forEach(row => {
    if (Number(row[1]) !== Number(meetingNo) ||
        String(row[3] || '') !== 'SPEAKER_DETAILS') return;

    const sent = sendSpeakerDetailsToTmod_(
      Number(row[1]),
      String(row[2] || ''),
      String(row[4] || ''),
      [
        storedSubmissionText_(row[5]),
        storedSubmissionText_(row[6]),
        storedSubmissionText_(row[7]),
        storedSubmissionText_(row[8]),
        storedSubmissionText_(row[14])
      ]
    );
    if (sent) anySent = true;
  });
  return anySent;
}

function processReadySpeakerDetailsToTmod_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const subs = ss.getSheetByName('WA Submissions');
  if (!subs || subs.getLastRow() < 2) return;

  subs.getRange(2, 1, subs.getLastRow() - 1, 15).getValues().forEach(row => {
    if (String(row[3] || '') !== 'SPEAKER_DETAILS') return;
    sendSpeakerDetailsToTmod_(
      Number(row[1]),
      String(row[2] || ''),
      String(row[4] || ''),
      [
        storedSubmissionText_(row[5]),
        storedSubmissionText_(row[6]),
        storedSubmissionText_(row[7]),
        storedSubmissionText_(row[8]),
        storedSubmissionText_(row[14])
      ]
    );
  });
}

function sendLateTmodCatchUp_(meetingNo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  if (!planner || planner.getLastRow() < 3) return false;

  const ids = planner.getRange(3, 1, planner.getLastRow() - 2, 1)
    .getValues().flat();
  const idx = ids.findIndex(v => Number(v) === Number(meetingNo));
  if (idx < 0) return false;

  const row = planner.getRange(idx + 3, 1, 1, 28).getValues()[0];
  const meetingDate = row[1];
  const status = String(row[2] || '').trim();
  const tmod = String(row[18] || '').trim();

  if (!tmod || ['Completed', 'Cancelled'].indexOf(status) !== -1) return false;
  if (meetingDate instanceof Date &&
      startOfDay_(meetingDate) < startOfDay_(new Date())) return false;

  let anySent = false;

  // Grammarian handoff is replayed only after the Grammarian's final review
  // confirmation. Recipient-specific keys allow a replacement TMOD to receive it.
  if (configBool_('GRAMMARIAN_WORKFLOW_ENABLED') &&
      grammarBundleIsFinalConfirmed_(meetingNo)) {
    if (sendApprovedGrammarBundle_(meetingNo)) anySent = true;
  }

  // Speaker details are stored independently of TMOD assignment, so replay all
  // available submissions to the newly assigned/replacement TMOD.
  if (sendStoredSpeakerDetailsToTmod_(meetingNo)) anySent = true;

  // Introductions are handed over once the meeting is inside the configured
  // bundle window. A TMOD assigned earlier will receive them on the normal tick.
  if (configBool_('INTRO_WORKFLOW_ENABLED') && meetingDate instanceof Date) {
    const dueDays = Math.max(
      0,
      Number(configValue_('INTRO_BUNDLE_DUE_DAYS', '2')) || 2
    );
    const daysLeft = Math.round(
      (startOfDay_(meetingDate) - startOfDay_(new Date())) / 86400000
    );
    if (daysLeft >= 0 && daysLeft <= dueDays &&
        sendConsolidatedIntroductionsIfReady_(meetingNo)) {
      anySent = true;
    }
  }

  // If a theme already exists (for example after a TMOD replacement), refresh
  // the promotion handoff so VPPR receives the current TMOD + theme pairing.
  if (String(row[3] || '').trim() && sendThemeAndTmodToVppr_(meetingNo)) {
    anySent = true;
  }

  return anySent;
}

function processLateTmodCatchUps_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  if (!planner || planner.getLastRow() < 3) return;

  const today = startOfDay_(new Date());
  const rows = planner.getRange(
    3, 1, planner.getLastRow() - 2, 28
  ).getValues();

  rows.forEach(row => {
    const meetingNo = Number(row[0]);
    const meetingDate = row[1];
    const status = String(row[2] || '').trim();
    const tmod = String(row[18] || '').trim();

    if (!meetingNo || !tmod || !(meetingDate instanceof Date)) return;
    if (startOfDay_(meetingDate) < today) return;
    if (['Completed', 'Cancelled'].indexOf(status) !== -1) return;

    sendLateTmodCatchUp_(meetingNo);
  });
}


/* ATK_AGENDA_TIMING_V1 */
/**
 * Agenda timing guardrail.
 * Fixed rule: reserve a 9-minute transition/buffer allowance for every meeting,
 * regardless of the number of speakers or presenter changes.
 * Table Topics is the first/only flexible block. If the resulting TT time is
 * below 10 minutes, the VPE must explicitly KEEP or REMOVE it.
 */
function agendaTimingDecision_(baseTableTopicsMinutes, presenters) {
  const transitionMinutes = 9;
  const availableTableTopicsMinutes = Math.max(0, Number(baseTableTopicsMinutes || 0) - transitionMinutes);
  return {
    totalTransitions: null,
    coveredTransitions: null,
    transitionMinutes: transitionMinutes,
    availableTableTopicsMinutes: availableTableTopicsMinutes,
    needsVpeDecision: availableTableTopicsMinutes < 10
  };
}

/**
 * Rebalances the agenda when a reviewer increases/decreases programme time.
 * Priority when extra time is needed:
 *   1) Table Topics down to 12 min
 *   2) Remove VPE Meeting Planning block
 *   3) Transition buffer down from 9 to 7 min
 *   4) Table Topics again; below 10 min requires VPE KEEP/REMOVE decision
 * When time is recovered, restore transition buffer to 9 min first, then TT.
 */
function rebalanceAgendaTiming_(tableTopicsMinutes, transitionMinutes, meetingPlanningMinutes, deltaMinutes) {
  let tt = Math.max(0, Number(tableTopicsMinutes || 0));
  let buffer = Math.min(9, Math.max(7, Number(transitionMinutes == null ? 9 : transitionMinutes)));
  let planning = Math.max(0, Number(meetingPlanningMinutes || 0));
  let need = Number(deltaMinutes || 0);
  let meetingPlanningRemoved = false;

  if (need > 0) {
    // 1. Protect a 12-minute TT block initially.
    const fromTtAbove12 = Math.min(need, Math.max(0, tt - 12));
    tt -= fromTtAbove12;
    need -= fromTtAbove12;

    // 2. Once TT is at/below 12, remove Meeting Planning before touching buffer.
    if (need > 0 && tt <= 12 && planning > 0) {
      need = Math.max(0, need - planning);
      planning = 0;
      meetingPlanningRemoved = true;
    }

    // 3. Reduce transition buffer, but never below 7 minutes.
    if (need > 0) {
      const fromBuffer = Math.min(need, Math.max(0, buffer - 7));
      buffer -= fromBuffer;
      need -= fromBuffer;
    }

    // 4. Last resort: take the balance from Table Topics.
    if (need > 0) {
      const fromTtFinal = Math.min(need, tt);
      tt -= fromTtFinal;
      need -= fromTtFinal;
    }
  } else if (need < 0) {
    let recovered = Math.abs(need);
    // Restore standard buffer first, then return remaining time to Table Topics.
    const restoreBuffer = Math.min(recovered, 9 - buffer);
    buffer += restoreBuffer;
    recovered -= restoreBuffer;
    tt += recovered;
    need = 0;
  }

  return {
    tableTopicsMinutes: tt,
    transitionMinutes: buffer,
    meetingPlanningMinutes: planning,
    meetingPlanningRemoved: meetingPlanningRemoved,
    unresolvedMinutes: Math.max(0, need),
    needsVpeDecision: tt < 10
  };
}

function sendAgendaTableTopicsDecision_(meetingNo, availableMinutes) {
  const vpe = configValue_('VPE_MEMBER_NAME', 'Riyas AF');
  const phone = phoneForMember_(vpe);
  if (!phone || !isWhatsAppRecipientAllowed_(phone)) return {ok:false, blocked:true, error:'VPE not allowlisted'};

  const minutes = Math.max(0, Number(availableMinutes || 0));
  const body =
    'ATK Meeting ' + meetingNo + ' agenda needs your decision.\n\n' +
    'After adding the required transition time, only ' + minutes +
    ' minutes are available for Table Topics.\n\n' +
    'Our minimum is 10 minutes. Would you like to keep Table Topics with the available ' +
    minutes + ' minutes, or remove it?';

  // Interactive reply buttons can only be sent inside an open customer-service window.
  // Outside the window we deliberately do not fall back to a template until the VPE
  // has reviewed and approved that Meta template.
  if (!hasOpenWhatsAppServiceWindow_(phone)) {
    logAutomation_('Agenda', meetingNo, vpe, 'Table Topics decision required',
      'Blocked', 'Available=' + minutes + ' min | Meta button template approval required');
    return {ok:false, blocked:true, error:'No open WhatsApp service window'};
  }

  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('META_ACCESS_TOKEN');
  const phoneId = props.getProperty('WA_PHONE_NUMBER_ID');
  if (!token || !phoneId) return {ok:false, error:'Meta credentials missing'};

  const payload = {
    messaging_product: 'whatsapp',
    to: String(phone).replace(/\D/g, ''),
    type: 'interactive',
    interactive: {
      type: 'button',
      body: {text: body},
      action: {buttons: [
        {type:'reply', reply:{id:'AGENDA_TT_KEEP|' + meetingNo + '|' + minutes, title:'Keep it'}},
        {type:'reply', reply:{id:'AGENDA_TT_REMOVE|' + meetingNo + '|' + minutes, title:'Remove it'}}
      ]}
    }
  };
  const response = UrlFetchApp.fetch(
    'https://graph.facebook.com/v26.0/' + encodeURIComponent(phoneId) + '/messages',
    {method:'post', contentType:'application/json',
     headers:{Authorization:'Bearer ' + token},
     payload:JSON.stringify(payload), muteHttpExceptions:true}
  );
  const ok = response.getResponseCode() >= 200 && response.getResponseCode() < 300;
  logAutomation_('Agenda', meetingNo, vpe, 'Table Topics decision required',
    ok ? 'Sent' : 'Error', 'Available=' + minutes + ' min');
  return {ok:ok, status:response.getResponseCode()};
}

function recordAgendaTableTopicsDecision_(meetingNo, action, availableMinutes) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Agenda Preview');
  if (!sheet) return false;
  const keep = String(action || '').toUpperCase() === 'KEEP';
  sheet.getRange('I4').setValue(
    keep
      ? 'VPE DECISION: KEEP TABLE TOPICS (' + availableMinutes + ' min)'
      : 'VPE DECISION: REMOVE TABLE TOPICS'
  );
  sheet.getRange('I4').setNote(
    'Transition rule: fixed 9-minute meeting buffer. ' +
    'Decision recorded ' + new Date()
  );
  logAutomation_('Agenda', meetingNo, configValue_('VPE_MEMBER_NAME','Riyas AF'),
    'Table Topics decision', keep ? 'Keep' : 'Remove',
    'Available=' + availableMinutes + ' min');
  return true;
}


/* ATK_AGENDA_DEADLINE_V1 */
/**
 * Hard agenda publication rule.
 * The agenda must not be blocked by missing role-player inputs.
 * Default deadline: Saturday 8:00 PM, Asia/Kolkata.
 *
 * Readiness can publish earlier; once the deadline is reached this returns
 * publishNow=true even when required details are still pending. Missing values
 * should be rendered as "Pending" / "TBC" by the agenda builder and can be
 * revised later when submissions arrive.
 */
function agendaPublicationDecision_(meetingDate, readiness) {
  const timezone = configValue_('TIMEZONE', 'Asia/Kolkata');
  const deadlineDay = String(configValue_('AGENDA_PUBLISH_DEADLINE_DAY', 'SATURDAY')).toUpperCase();
  const deadlineHour = Math.max(0, Math.min(23,
    Number(configValue_('AGENDA_PUBLISH_DEADLINE_HOUR', '20')) || 20
  ));

  const meeting = meetingDate instanceof Date ? new Date(meetingDate) : new Date(meetingDate);
  if (isNaN(meeting.getTime())) {
    return {publishNow:false, deadlineReached:false, ready:false, error:'Invalid meeting date'};
  }

  const dayMap = {SUNDAY:0, MONDAY:1, TUESDAY:2, WEDNESDAY:3, THURSDAY:4, FRIDAY:5, SATURDAY:6};
  const targetDay = Object.prototype.hasOwnProperty.call(dayMap, deadlineDay) ? dayMap[deadlineDay] : 6;

  // ATK meetings are normally Monday; walk backwards to the configured
  // publication weekday so the rule also behaves correctly if a meeting date changes.
  const deadline = new Date(meeting);
  deadline.setHours(deadlineHour, 0, 0, 0);
  while (deadline.getDay() !== targetDay || deadline.getTime() >= meeting.getTime()) {
    deadline.setDate(deadline.getDate() - 1);
  }

  const now = new Date();
  const ready = Boolean(readiness && readiness.ready);
  const deadlineReached = now.getTime() >= deadline.getTime();

  return {
    publishNow: ready || deadlineReached,
    deadlineReached: deadlineReached,
    ready: ready,
    deadline: deadline,
    timezone: timezone,
    allowPendingFields: deadlineReached && !ready,
    publicationReason: ready ? 'READY' : (deadlineReached ? 'HARD_DEADLINE' : 'WAITING_FOR_INPUTS')
  };
}

/**
 * Normalizes unresolved agenda values for a deadline publication.
 * This deliberately keeps publication moving instead of treating a missing
 * Theme / Grammarian / speaker field as a blocker.
 */
function agendaValueForPublication_(value, deadlineReached) {
  const clean = String(value == null ? '' : value).trim();
  if (clean) return clean;
  return deadlineReached ? 'TBC' : '[Pending]';
}


/* ATK_AGENDA_APPROVAL_V1
 * Agenda approval is satisfied by EITHER the VPE OR the President.
 * The exact PDF version sent for review is fingerprinted so a later edit
 * cannot inherit an earlier approval.
 */
function agendaApprovalSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Agenda Approvals');
  if (!sheet) {
    sheet = ss.insertSheet('Agenda Approvals');
    sheet.appendRow([
      'Meeting No','Version','PDF File ID','PDF URL','Fingerprint','Status',
      'VPE Status','President Status','Approved By','Approved At','Created At','Notes'
    ]);
    sheet.hideSheet();
  }
  return sheet;
}

function agendaPreviewFingerprint_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Agenda Preview');
  if (!sheet) throw new Error('Agenda Preview not found.');
  const values = sheet.getDataRange().getDisplayValues();
  const raw = JSON.stringify(values);
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return bytes.map(v => ('0' + (v & 255).toString(16)).slice(-2)).join('');
}

function latestAgendaApprovalRecord_(meetingNo) {
  const sheet = agendaApprovalSheet_();
  if (sheet.getLastRow() < 2) return null;
  const rows = sheet.getRange(2,1,sheet.getLastRow()-1,12).getValues();
  for (let i=rows.length-1;i>=0;i--) {
    if (Number(rows[i][0]) === Number(meetingNo)) return {row:i+2, values:rows[i]};
  }
  return null;
}

function createAgendaPdfForReview_(meetingNo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const preview = ss.getSheetByName('Agenda Preview');
  if (!preview) throw new Error('Agenda Preview not found.');

  const fingerprint = agendaPreviewFingerprint_();
  const existing = latestAgendaApprovalRecord_(meetingNo);
  if (existing && String(existing.values[4] || '') === fingerprint &&
      String(existing.values[2] || '')) {
    const file = DriveApp.getFileById(String(existing.values[2]));
    return {file:file, version:Number(existing.values[1] || 1), fingerprint:fingerprint,
      row:existing.row, reused:true};
  }

  const version = existing ? Number(existing.values[1] || 0) + 1 : 1;
  const base = ss.getUrl().replace(/edit.*$/, '');
  const exportUrl = base + 'export?format=pdf&gid=' + preview.getSheetId() +
    '&portrait=true&size=A4&fitw=true&sheetnames=false&printtitle=false' +
    '&pagenumbers=false&gridlines=false&fzr=false&top_margin=0.25&bottom_margin=0.25' +
    '&left_margin=0.25&right_margin=0.25';

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(exportUrl, {
    headers:{Authorization:'Bearer ' + token}, muteHttpExceptions:true
  });
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    throw new Error('Agenda PDF export failed. HTTP ' + response.getResponseCode());
  }

  const name = 'ATK_Meeting_' + meetingNo + '_Agenda_v' + version + '_REVIEW.pdf';
  const blob = response.getBlob().setName(name);
  const folderId = String(configValue_('AGENDA_PDF_FOLDER_ID','')).trim();
  const file = folderId ? DriveApp.getFolderById(folderId).createFile(blob) : DriveApp.createFile(blob);

  const approvals = agendaApprovalSheet_();
  approvals.appendRow([
    Number(meetingNo), version, file.getId(), file.getUrl(), fingerprint, 'AWAITING_APPROVAL',
    'PENDING','PENDING','','',new Date(),'Either VPE or President approval is sufficient.'
  ]);
  return {file:file, version:version, fingerprint:fingerprint, row:approvals.getLastRow(), reused:false};
}

function uploadAgendaPdfToMeta_(file) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('META_ACCESS_TOKEN');
  const phoneId = props.getProperty('WA_PHONE_NUMBER_ID');
  if (!token || !phoneId) throw new Error('Meta credentials missing.');
  const response = UrlFetchApp.fetch(
    'https://graph.facebook.com/v26.0/' + encodeURIComponent(phoneId) + '/media',
    {method:'post', headers:{Authorization:'Bearer ' + token},
     payload:{messaging_product:'whatsapp', type:'application/pdf', file:file.getBlob()},
     muteHttpExceptions:true}
  );
  const http=response.getResponseCode();
  let data={}; try{data=JSON.parse(response.getContentText()||'{}');}catch(_){}
  if (http<200 || http>=300 || !data.id) throw new Error('Meta PDF upload failed. HTTP ' + http);
  return String(data.id);
}

function sendAgendaApprovalReview_(member, phone, meetingNo, pdf, roleLabel) {
  if (!phone || !isWhatsAppRecipientAllowed_(phone)) {
    return {ok:false, blocked:true, error:'Recipient not allowlisted in current send mode'};
  }
  if (!hasOpenWhatsAppServiceWindow_(phone)) {
    return {ok:false, deferred:true, templateRequired:true,
      error:'Agenda PDF approval template required outside service window'};
  }

  const mediaId = uploadAgendaPdfToMeta_(pdf.file);
  const caption = 'ATK Meeting ' + meetingNo + ' — Agenda v' + pdf.version +
    '\n\nPlease review the attached agenda. If everything is okay, tap Approve Agenda. ' +
    'Approval from either the VPE or President will finalize this version.';

  const documentResult = sendWhatsAppPayload_({
    messaging_product:'whatsapp', recipient_type:'individual', to:normalizePhone_(phone),
    type:'document', document:{id:mediaId, filename:pdf.file.getName(), caption:caption}
  });
  recordWorkflowOutbound_(member, phone, 'AGENDA_PDF_REVIEW', caption, meetingNo, documentResult,
    ['AGENDA_PDF_REVIEW',meetingNo,pdf.version,member].join('|'));
  if (!documentResult.ok) return documentResult;

  const body = 'ATK Meeting ' + meetingNo + ' Agenda v' + pdf.version +
    '\nApprove this exact PDF version?';
  const buttonResult = sendWhatsAppPayload_({
    messaging_product:'whatsapp', recipient_type:'individual', to:normalizePhone_(phone),
    type:'interactive', interactive:{type:'button', body:{text:body}, action:{buttons:[
      {type:'reply',reply:{id:['AGENDA_APPROVE',meetingNo,pdf.version,pdf.fingerprint].join('|'),title:'Approve Agenda'}}
    ]}}
  });
  recordWorkflowOutbound_(member, phone, 'AGENDA_APPROVAL_REQUEST', body, meetingNo, buttonResult,
    ['AGENDA_APPROVAL_REQUEST',meetingNo,pdf.version,member].join('|'));
  return buttonResult;
}

function sendAgendaForApproval_(meetingNo) {
  const pdf = createAgendaPdfForReview_(meetingNo);
  const vpe = configValue_('VPE_MEMBER_NAME','Riyas AF');
  const president = configValue_('PRESIDENT_MEMBER_NAME','Abhilasha Kishore');
  const recipients = [
    {member:vpe, role:'VPE'},
    {member:president, role:'President'}
  ];
  const results = recipients.map(r => {
    const phone = phoneForMember_(r.member);
    const result = sendAgendaApprovalReview_(r.member, phone, meetingNo, pdf, r.role);
    return {member:r.member, role:r.role, result:result};
  });
  logAutomation_('Agenda',meetingNo,'VPE + President','Agenda PDF approval dispatch',
    results.some(x=>x.result&&x.result.ok)?'Sent':'Deferred',
    'Version=' + pdf.version + ' | either approval finalizes');
  return {meetingNo:Number(meetingNo), version:pdf.version, fingerprint:pdf.fingerprint, results:results};
}

function recordAgendaApproval_(meetingNo, version, fingerprint, member, messageId) {
  const record = latestAgendaApprovalRecord_(meetingNo);
  if (!record || Number(record.values[1]) !== Number(version) ||
      String(record.values[4] || '') !== String(fingerprint || '')) {
    sendWhatsAppText_(phoneForMember_(member),
      'That agenda version is no longer current. Please review the latest agenda.');
    return {ok:false, stale:true};
  }
  if (String(record.values[5] || '') === 'APPROVED') {
    sendWhatsAppText_(phoneForMember_(member),
      'This agenda version has already been approved and finalized.');
    return {ok:true, alreadyApproved:true};
  }

  const sheet = agendaApprovalSheet_();
  const vpe = configValue_('VPE_MEMBER_NAME','Riyas AF');
  const president = configValue_('PRESIDENT_MEMBER_NAME','Abhilasha Kishore');
  const role = member === vpe ? 'VPE' : (member === president ? 'President' : '');
  if (!role) return {ok:false, error:'Unauthorized agenda approver'};

  if (role === 'VPE') sheet.getRange(record.row,7).setValue('APPROVED');
  if (role === 'President') sheet.getRange(record.row,8).setValue('APPROVED');
  sheet.getRange(record.row,6).setValue('APPROVED');
  sheet.getRange(record.row,9).setValue(member + ' (' + role + ')');
  sheet.getRange(record.row,10).setValue(new Date());
  sheet.getRange(record.row,12).setValue('Approved by ' + role + '. Either approver is sufficient.');

  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const preview=ss.getSheetByName('Agenda Preview');
  if (preview) {
    preview.getRange('I1').setValue('APPROVED — ' + role);
    preview.getRange('I2').setValue('Approved by: ' + member);
    preview.getRange('I3').setValue('Agenda v' + version + ' finalized');
  }

  sendWhatsAppText_(phoneForMember_(member),
    'Agenda approved ✅\nATK Meeting ' + meetingNo + ' · Version ' + version +
    '\nThis version is now finalized.');
  logAutomation_('Agenda',meetingNo,member,'Agenda approval','Approved',
    'Version=' + version + ' | Message=' + String(messageId||''));
  return {ok:true, approvedBy:member, role:role};
}
