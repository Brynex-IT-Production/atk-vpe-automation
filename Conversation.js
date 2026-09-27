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
  sheet.getRange(idx + 2, 6).setValue(intro);
}

function introductionPrompt_(member, role, meetingNo) {
  const saved = defaultIntro_(member);
  const context = meetingNo
    ? 'For ATK Meeting ' + meetingNo + ' as ' + role + ','
    : 'For your ATK member profile,';
  return saved
    ? context + ' would you like us to use your saved introduction?\n\n' +
      'Reply 1 to use it, 2 to update your standard introduction, or 3 to share a recent achievement or something ATK may not know about you for this meeting.\n\nSaved introduction: ' + saved
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
    text = 'Hi ' + member + ', please reply with the theme for ATK Meeting ' + meetingNo +
      '. A plain reply is fine; we will save it for the agenda.';
  } else if (kind === 'GRAMMAR') {
    text = 'Hi ' + member + ', as Grammarian for ATK Meeting ' + meetingNo +
      ', what is your Word of the Day? Just reply with the word. I will ask for the rest one by one.';
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
  const answer = String(text || '').trim();
  if (!answer || /^(ROLE|ROLES|ADMIN STATUS|ADMIN SETUP|ADMIN EXCEPTIONS)$/i.test(answer)) return false;

  // Existing approved templates may still show the earlier structured syntax
  // until Meta approves the new ones. Let the legacy parser handle that reply.
  if (/^(THEME|GRAMMAR|INTRO|SPEECH)(?:\\s+\\d+)?\\s*:/i.test(answer)) {
    clearConversation_(phone);
    return false;
  }
  const kind = state.kind;
  const n = state.meetingNo;
  if (kind === 'THEME') {
    if (answer.length < 3) {
      sendWhatsAppText_(phone, 'Please reply with the full theme.');
      return true;
    }
    clearConversation_(phone);
    handleThemeSubmission_(member, phone, n, answer, messageId);
    return true;
  }

  if (kind === 'GRAMMAR') {
    const steps = ['WORD', 'WORD_MEANING', 'IDIOM', 'IDIOM_MEANING'];
    const prompts = {
      WORD: 'What does the word mean?',
      WORD_MEANING: 'Thank you. What is your Idiom of the Day?',
      IDIOM: 'What does the idiom mean?'
    };
    if (answer.length < 2) {
      sendWhatsAppText_(phone, 'Please send a little more detail.');
      return true;
    }
    state.data[state.step] = answer;
    const index = steps.indexOf(state.step);
    if (index < 3) {
      state.step = steps[index + 1];
      setConversation_(phone, state);
      sendWhatsAppText_(phone, prompts[steps[index]]);
    } else {
      clearConversation_(phone);
      handleGrammarSubmission_(member, phone, n, steps.map(s => state.data[s]).join(' | '), messageId);
    }
    return true;
  }

  if (kind === 'SPEAKER_DETAILS') {
    const steps = ['TITLE', 'PATH', 'LEVEL', 'PROJECT'];
    const prompts = {
      TITLE: 'Which Path is this speech in?',
      PATH: 'What is the Level?',
      LEVEL: 'What is the Project name?'
    };
    if (answer.length < 1) return true;
    state.data[state.step] = answer;
    const index = steps.indexOf(state.step);
    if (index < 3) {
      state.step = steps[index + 1];
      setConversation_(phone, state);
      sendWhatsAppText_(phone, prompts[steps[index]]);
    } else {
      clearConversation_(phone);
      handleSpeakerDetailsSubmission_(member, phone, n, steps.map(s => state.data[s]).join(' | '), messageId);
      promptSpeakerIntro_(member, phone, n, state.related);
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
      if (answer === '2') {
        state.step = 'UPDATE_DEFAULT';
        setConversation_(phone, state);
        sendWhatsAppText_(phone, 'Please reply with your new standard introduction. We will save it for future meetings.');
        return true;
      }
      if (answer === '3') {
        state.step = 'MEETING_INTRO';
        setConversation_(phone, state);
        sendWhatsAppText_(phone, 'Please share your recent achievement or something you would like ATK to know. We will use it for this meeting.');
        return true;
      }
      sendWhatsAppText_(phone, 'Please reply 1, 2 or 3 so I know which introduction to use.');
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
  const prompt = startGuidedRequest_(member, phone, 'INTRO', meetingNo, slotRef, 'speaker');
  if (!prompt) return;
  const result = sendWhatsAppText_(phone, prompt.text);
  if (result.ok) setConversation_(phone, prompt.state);
}

function processAgendaSecondReminders_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  const tasks = ss.getSheetByName('Automation Tasks');
  if (!planner || !tasks || tasks.getLastRow() < 2) return;
  const meetings = {};
  const rows = planner.getRange(3, 1, Math.max(0, Math.min(60, planner.getLastRow() - 2)), 28).getValues();
  rows.forEach(row => {
    const date = row[1];
    if (!(date instanceof Date) || String(row[2]) === 'Cancelled' || String(row[2]) === 'Completed') return;
    if (Utilities.formatDate(date, configValue_('TIMEZONE', 'Asia/Kolkata'), 'EEE') !== 'Mon') return;
    meetings[Number(row[0])] = Math.round((startOfDay_(date) - startOfDay_(new Date())) / 86400000);
  });
  const pending = tasks.getRange(2, 1, tasks.getLastRow() - 1, 18).getValues();
  pending.forEach(row => {
    const meetingNo = Number(row[1]);
    const days = meetings[meetingNo];
    if (days !== 3 && days !== 2 || String(row[5]) !== 'Pending') return;
    const kind = String(row[2] || '');
    if (['INTRO', 'GRAMMAR', 'THEME', 'SPEAKER_DETAILS'].indexOf(kind) < 0) return;
    const member = String(row[3] || '').trim();
    const phone = phoneForMember_(member);
    if (!member || !phone || !isWhatsAppRecipientAllowed_(phone)) return;
    const key = ['AGENDA_SECOND', meetingNo, kind, row[4], member].join('|');
    if (hasReminderBeenSent_(key)) return;
    const firstKind = kind === 'INTRO' ? 'INTRO_REQUEST' :
      kind === 'GRAMMAR' ? 'GRAMMAR_REQUEST' :
      kind === 'THEME' ? 'THEME_REQUEST' : 'SPEAKER_DETAILS_PROMPT';
    const firstKey = [firstKind, meetingNo, row[4], member].join('|');
    if (!hasReminderBeenSent_(firstKey)) return;
    const item = kind === 'SPEAKER_DETAILS' ? 'speech title, Path, Level and Project' :
      kind === 'INTRO' ? 'introduction' :
      kind === 'GRAMMAR' ? 'Word and Idiom details' : 'meeting theme';
    const body = 'Hi ' + member + ', this is a second reminder for ATK Meeting ' +
      meetingNo + '. Please send your ' + item +
      ' today. We can publish the agenda by Saturday evening only after receiving the pending details. Thank you.';
    const result = sendAutomatedText_(phone, 'AGENDA_REMINDER', body, [member, meetingNo, item]);
    recordWorkflowOutbound_(member, phone, 'AGENDA_SECOND', body, meetingNo, result, key);
  });
}

function processMemberIntroOnboarding_() {
  if (!configBool_('INTRO_WORKFLOW_ENABLED')) return;
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
