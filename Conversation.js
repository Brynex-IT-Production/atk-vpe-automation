/**
 * One active guided WhatsApp request per member. State survives Apps Script runs.
 * Explicit ROLE/ADMIN commands still take precedence in Engine.js.
 */
function conversationKey_(phone) {
  return 'ATK_CONV_' + normalizePhone_(phone);
}

function activeConversation_(phone) {
  const props = PropertiesService.getScriptProperties();
  const key = conversationKey_(phone);
  const raw = props.getProperty(key);
  if (!raw) return null;
  let state;
  try { state = JSON.parse(raw); } catch (ignore) {}
  if (!state || Date.now() - Number(state.at || 0) > 72 * 3600000) {
    props.deleteProperty(key);
    return null;
  }
  return state;
}

function setConversation_(phone, state) {
  state.at = Date.now();
  PropertiesService.getScriptProperties()
    .setProperty(conversationKey_(phone), JSON.stringify(state));
}

function clearConversation_(phone) {
  PropertiesService.getScriptProperties().deleteProperty(conversationKey_(phone));
}

function defaultIntro_(member) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Member Master');
  if (!sheet || sheet.getLastRow() < 2) return '';
  const rows = sheet.getRange(2, 2, sheet.getLastRow() - 1, 5).getValues();
  const row = rows.find(r => String(r[0] || '').trim() === member);
  return row ? String(row[4] || '').trim() : '';
}

function saveDefaultIntro_(member, intro) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Member Master');
  const names = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().flat();
  const idx = names.findIndex(v => String(v || '').trim() === member);
  if (idx < 0) throw new Error('Member not found for default introduction.');
  const cell = sheet.getRange(idx + 2, 6);
  cell.setValue(intro);
  SpreadsheetApp.flush();
  if (String(cell.getValue() || '') !== String(intro)) {
    throw new Error('Introduction save could not be verified; please try again.');
  }
  logAutomation_('Introduction', '', member, 'Standard introduction saved and verified',
    'Success', 'Member Master!' + cell.getA1Notation());
}

function introductionPrompt_(member, role, meetingNo) {
  const saved = defaultIntro_(member);
  const context = meetingNo
    ? 'For ATK Meeting ' + meetingNo + ' as ' + role + ','
    : 'For your ATK member profile,';
  return saved
    ? 'Hey ' + member + ',\n\n' +
      'For the upcoming ATK Meeting ' + meetingNo + ', we already have your saved introduction.\n\n' +
      'But to avoid repeating the same introduction every time, is there anything new with you that we can add?\n\n' +
      'A new recognition, achievement, update, or simply something about you that we may not know yet.\n\n' +
      'Anything that comes to your mind will do. 😊\n\n' +
      'Reply “I have something new” or “Use my default”.\n\n' +
      'Thank you!'
    : context + ' please reply with a short introduction about yourself. We will save it for future meetings, so you need not write it each time.';
}

function startGuidedRequest_(member, phone, kind, meetingNo, related, role) {
  if (activeConversation_(phone)) return null;
  const state = {
    kind: kind, meetingNo: meetingNo || 0, related: related || '',
    step: kind === 'GRAMMAR' ? 'WORD' :
      kind === 'SPEAKER_DETAILS' ? 'TITLE' :
      kind === 'INTRO' || kind === 'DEFAULT_INTRO' ? 'CHOICE' : 'THEME',
    data: {}
  };
  let text = '';
  if (kind === 'THEME') {
    text =
      'Hi ' + member + ',\n\n' +
      'Please copy the next message, add the theme for ATK Meeting ' + meetingNo +
      ', and send it back.\n\nThank you.';
  } else if (kind === 'GRAMMAR') {
    text = grammarStageIntroMessage_(member, meetingNo, 'WORD') +
      '\n\n' + grammarStageCopyMessage_('WORD');
  } else if (kind === 'SPEAKER_DETAILS') {
    text = 'Hi ' + member + ', we are collecting the agenda details for ATK Meeting ' + meetingNo +
      '. What is your speech title? I will ask for your Path, Level and Project next, then your introduction.';
  } else {
    text = introductionPrompt_(member, role || 'role player', meetingNo);
  }
  return { state: state, text: text };
}

function handleGuidedReply_(member, phone, text, messageId) {
  const state = activeConversation_(phone);
  if (!state) return false;

  // Conversation records can survive workflow/template changes. Always
  // normalize mutable state before any handler writes into it so a legacy or
  // partially-created conversation can never throw on state.data[...] access.
  if (!state.data || typeof state.data !== 'object' || Array.isArray(state.data)) {
    state.data = {};
  }

  const answer = String(text || '').trim();
  if (!answer || /^(ROLE|ROLES|ADMIN STATUS|ADMIN SETUP|ADMIN EXCEPTIONS)$/i.test(answer)) return false;

  const speakerBlock = parseSpeakerDetailsBlock_(answer);
  if (speakerBlock && speakerBlock.recognized >= 2) return false;

  // Existing approved templates may still show the earlier structured syntax
  // until Meta approves the new ones. Let the legacy parser handle that reply.
  // Recognize structured replies with single-backslash regex escapes. If the
  // member copies "Theme: ...", let handleStructuredSubmission_ save it.
  if (/^(THEME|GRAMMAR|INTRO|SPEECH)(?:\s+\d+)?\s*:/i.test(answer)) {
    clearConversation_(phone);
    return false;
  }
  const kind = state.kind;
  const n = state.meetingNo;
  if (kind === 'THEME') {
    // Theme is a structured input too. Do not interpret ordinary chat as a
    // meeting theme. The accepted shape is explicitly "Theme: ...".
    sendStructuredInputFallbackOnce_(
      member,
      phone,
      Number(n || 0),
      String(state.related || 'TMOD'),
      'THEME',
      '',
      themeCopyMessage_()
    );
    return true;
  }

  if (kind === 'GRAMMAR') {
    const combinedGrammar = parseGrammarCombinedBlock_(answer);
    if (combinedGrammar.recognized >= 6) {
      handleCombinedGrammarSubmission_(
        member, phone, n, combinedGrammar, messageId
      );
      return true;
    }

    if (state.data && state.data.pendingImage) {
      if (!/^(WORD|IDIOM)$/i.test(answer)) {
        sendWhatsAppText_(phone, 'Please reply WORD or IDIOM to label the image.');
        return true;
      }
      const pendingImage = state.data.pendingImage;
      delete state.data.pendingImage;
      setConversation_(phone, state);
      saveGrammarImageSubmission_(
        member, phone, n, answer.toUpperCase(), pendingImage);
      return true;
    }

    const rawStage = String(state.step || grammarStageForMember_(n, member)).toUpperCase();
    const stage = rawStage.indexOf('IDIOM') !== -1 ? 'IDIOM' :
      rawStage.indexOf('WORD') !== -1 ? 'WORD' : '';

    if (!stage) {
      clearConversation_(phone);
      return false;
    }

    const details = parseGrammarStageBlock_(answer, stage);

    if (!details.oppositeStage && Number(details.recognized || 0) === 0) {
      const label = stage === 'IDIOM' ? 'Idiom of the Day' : 'Word of the Day';
      sendStructuredInputFallbackOnce_(
        member,
        phone,
        Number(n || 0),
        'GRM|' + stage,
        'GRAMMAR_' + stage,
        'Hi ' + member + ',\n\n' +
          'This is an automated message, so I may not be able to understand regular chat replies.\n\n' +
          'Please use the format shared in the next message so I can record your details correctly.\n\n' +
          'Thank you.',
        grammarStageCopyMessage_(stage)
      );
      return true;
    }

    // Canonical manual-first design: accept complete or partial structured
    // replies here. The workflow layer merges partial fields and asks only for
    // whatever is still missing.
    handleGrammarStageSubmission_(member, phone, n, stage, details, messageId);
    return true;
  }

  if (kind === 'SPEAKER_DETAILS') {
    // The current speaker-details template uses a single copy/fill block.
    // Older guided conversations stored step/data, while the block flow only
    // stored meeting/slot/member. Never assume guided state exists.
    const block = parseSpeakerDetailsBlock_(answer);
    if (block && block.recognized >= 2) {
      handleSpeakerDetailsSubmission_(member, phone, n, block, messageId);
      return true;
    }

    const steps = ['TITLE', 'PATH', 'LEVEL', 'PROJECT'];
    const prompts = {
      TITLE: 'Which Path is this speech in?',
      PATH: 'What is the Level?',
      LEVEL: 'What is the Project name?'
    };
    const guidedStep = String(state.step || '').toUpperCase();

    if (steps.indexOf(guidedStep) === -1) {
      // Block-flow state: ordinary conversational replies are not speech-detail
      // submissions. Set the right expectation once and keep the copy/paste
      // block as its own message.
      sendStructuredInputFallbackOnce_(
        member,
        phone,
        Number(n || 0),
        String(state.related || ''),
        'SPEAKER_DETAILS',
        'Hi ' + member + ',\n\n' +
          'This is an automated message, so I may not be able to understand regular chat replies.\n\n' +
          'Please use the format shared in the next message so I can record your details correctly.\n\n' +
          'Thank you.',
        speakerDetailsCopyMessage_()
      );
      return true;
    }

    if (!state.data || typeof state.data !== 'object' || Array.isArray(state.data)) {
      state.data = {};
    }

    state.data[guidedStep] = answer;
    const index = steps.indexOf(guidedStep);
    if (index < 3) {
      state.step = steps[index + 1];
      setConversation_(phone, state);
      sendWhatsAppText_(phone, prompts[steps[index]]);
    } else {
      clearConversation_(phone);
      handleSpeakerDetailsSubmission_(
        member, phone, n, steps.map(s => state.data[s]).join(' | '), messageId
      );
    }
    return true;
  }

  if (kind === 'INTRO' || kind === 'DEFAULT_INTRO') {
    const saved = defaultIntro_(member);
    if (state.step === 'CHOICE' && saved) {
      if (answer === '1' || /^use (my )?(saved|default)/i.test(answer)) {
        clearConversation_(phone);
        if (kind === 'INTRO') handleIntroSubmission_(member, phone, n, saved, messageId, state.related);
        else sendWhatsAppText_(phone, 'Your saved introduction is ready ✅');
        return true;
      }
      if (/^(i\s+have\s+something\s+new|something\s+new)$/i.test(answer) || answer === '3') {
        state.step = 'MEETING_INTRO';
        setConversation_(phone, state);
        sendWhatsAppText_(phone, 'Great! Just send it here in your own words. 😊');
        return true;
      }
      if (answer === '2') {
        state.step = 'UPDATE_DEFAULT';
        setConversation_(phone, state);
        sendWhatsAppText_(phone, 'Please reply with your new standard introduction. We will save it for future meetings.');
        return true;
      }
      if (answer.length >= 10) {
        clearConversation_(phone);
        if (kind === 'INTRO') handleIntroSubmission_(member, phone, n, answer, messageId, state.related);
        return true;
      }
      sendWhatsAppText_(phone, 'Please choose “I have something new” or “Use my default”.');
      return true;
    }
    if (answer.length < 10) {
      sendWhatsAppText_(phone, 'Please add a little more about yourself so the introduction is useful.');
      return true;
    }
    if (state.step !== 'MEETING_INTRO') saveDefaultIntro_(member, answer);
    clearConversation_(phone);
    if (kind === 'INTRO') handleIntroSubmission_(member, phone, n, answer, messageId, state.related);
    else sendWhatsAppText_(phone, 'Your standard ATK introduction is saved ✅');
    return true;
  }
  return false;
}

function promptSpeakerIntro_(member, phone, meetingNo, slotRef) {
  const task = findPendingTaskForMember_('INTRO', member, meetingNo, slotRef);
  if (!task || task.related !== slotRef || activeConversation_(phone)) return;

  const key = ['INTRO_PROMPT', meetingNo, slotRef, member].join('|');
  if (hasReminderBeenSent_(key)) return;

  if (defaultIntro_(member)) {
    const sent = sendIntroChoiceButtons_(member, phone, meetingNo, slotRef, false);
    if (!sent) return;
    recordWorkflowOutbound_(member, phone, 'INTRO_PROMPT',
      sent.text, meetingNo, sent.result, key);
    return;
  }

  const prompt = startGuidedRequest_(member, phone, 'INTRO', meetingNo, slotRef, 'speaker');
  if (!prompt) return;
  const result = sendWhatsAppText_(phone, prompt.text);
  if (result.ok) setConversation_(phone, prompt.state);
  recordWorkflowOutbound_(member, phone, 'INTRO_PROMPT',
    prompt.text, meetingNo, result, key);
}

function processMemberIntroOnboarding_() {
  if (!configBool_('INTRO_WORKFLOW_ENABLED')) return;
  if (String(configValue_('MEMBER_INTRO_ONBOARDING_ENABLED', 'TRUE')).toUpperCase() === 'FALSE') return;
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Member Master');
  if (!sheet || sheet.getLastRow() < 2) return;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  rows.forEach(row => {
    const member = String(row[1] || '').trim();
    const phone = normalizePhone_(row[2]);
    if (!member || !phone || String(row[3] || '') !== 'Active' ||
        String(row[5] || '').trim() || !isWhatsAppRecipientAllowed_(phone) ||
        activeConversation_(phone)) return;
    const key = ['DEFAULT_INTRO_REQUEST', String(row[0] || member)].join('|');
    if (hasReminderBeenSent_(key)) return;
    const request = startGuidedRequest_(member, phone, 'DEFAULT_INTRO', 0, '', '');
    if (!request) return;
    const result = sendAutomatedText_(phone, 'DEFAULT_INTRO_REQUEST',
      request.text, [member]);
    recordWorkflowOutbound_(member, phone, 'DEFAULT_INTRO_REQUEST',
      request.text, 0, result, key);
    if (result.ok) setConversation_(phone, request.state);
  });
}
