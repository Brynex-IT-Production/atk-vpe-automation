/**
 * Web API V1 for the VPE Assistant.
 *
 * This is deliberately an adapter layer. It gives the Flutter/Web client an
 * ordered meeting model without changing the proven ATK automation tables.
 * Existing automation continues to read VPE Weekly Planner.
 */

const API_V1_VERSION_ = '2026-10-08.1';
const API_MEETING_ITEMS_SHEET_ = 'API Meeting Items';
const API_MEETINGS_SHEET_ = 'API Meetings';
const API_DEFAULT_CLUB_ID_ = 'ATK';
const API_ALLOW_LEGACY_SYNC_PROPERTY_ = 'WEB_API_ALLOW_LEGACY_SYNC';
const API_CLUB_SETTINGS_SHEET_ = 'API Club Settings';
const API_EXCOM_SHEET_ = 'API ExCom';

function apiJson_(status, data, error) {
  return ContentService.createTextOutput(JSON.stringify({
    ok: status >= 200 && status < 300,
    status: status,
    version: API_V1_VERSION_,
    data: data === undefined ? null : data,
    error: error || null
  })).setMimeType(ContentService.MimeType.JSON);
}

function apiAuthenticate_(payload) {
  const expected = String(PropertiesService.getScriptProperties().getProperty('WEB_API_KEY') || '');
  if (!expected) return false;
  const supplied = String(payload && payload.api_key || '');
  return supplied && supplied === expected;
}

function apiHandleRequest_(payload) {
  payload = payload || {};
  const action = String(payload.action || '').trim();
  if (action !== 'health' && !apiAuthenticate_(payload)) {
    return {status:401,error:{code:'UNAUTHORIZED',message:'Valid API key required'}};
  }
  const clubId = String(payload.club_id || API_DEFAULT_CLUB_ID_).trim();

  if (!action) return {status:400,error:{code:'MISSING_ACTION',message:'action is required'}};
  if (payload.api_version && String(payload.api_version) !== 'v1') {
    return {status:400,error:{code:'UNSUPPORTED_API_VERSION',message:'api_version must be v1'}};
  }

  try {
    switch (action) {
      case 'health':
        return {status:200,data:{service:'vpe-api',status:'ok'}};
      case 'members.list':
        return {status:200,data:{members:apiListMembers_(clubId).map(apiPublicMember_)}};
      case 'member.save':
        return {status:200,data:apiSaveMember_(clubId,payload)};
      case 'club.settings.get':
        return {status:200,data:apiGetClubSettings_(clubId)};
      case 'club.settings.save':
        return {status:200,data:apiSaveClubSettings_(clubId,payload)};
      case 'excom.list':
        return {status:200,data:{excom:apiListExCom_(clubId)}};
      case 'excom.save':
        return {status:200,data:apiSaveExCom_(clubId,payload)};
      case 'meeting_template.get':
        return {status:200,data:apiDefaultMeetingTemplate_(clubId)};
      case 'club.bootstrap':
        return {status:200,data:apiClubBootstrap_(clubId)};
      case 'meeting.create':
        return {status:201,data:apiCreateMeeting_(clubId,payload)};
      case 'meeting.get':
        return {status:200,data:apiMeetingDetail_(clubId,String(payload.meeting_id || ''))};
      case 'meetings.list':
        return {status:200,data:{meetings:apiListMeetings_(clubId)}};
      case 'meeting.save':
        return {status:200,data:apiSaveMeeting_(clubId,payload)};
      case 'meeting.sync':
        return {status:200,data:apiSyncMeetingToLegacyPlanner_(clubId,String(payload.meeting_id || ''))};
      case 'dashboard.get':
        return {status:200,data:apiDashboard_(clubId)};
      case 'meeting.activate':
        return {status:200,data:apiActivateMeeting_(clubId,String(payload.meeting_id || ''))};
      default:
        return {status:404,error:{code:'UNKNOWN_ACTION',message:'Unknown API action: '+action}};
    }
  } catch (err) {
    const code = String(err && err.apiCode || 'INTERNAL_ERROR');
    const clientCodes = ['BAD_REQUEST','NOT_FOUND','MEETING_NOT_READY','LEGACY_SYNC_DISABLED'];
    const status = code === 'NOT_FOUND' ? 404 : (clientCodes.indexOf(code) >= 0 ? 400 : 500);
    return {status:status,error:{
      code:code,
      message:String(err && err.message || err),
      details:(err && err.details) || null
    }};
  }
}

function apiPublicMember_(m) {
  return {member_id:m.member_id,name:m.name,status:m.status};
}

function apiListMembers_(clubId) {
  // V1 reads the existing Member Master. Multi-club storage comes next.
  if (clubId !== API_DEFAULT_CLUB_ID_) throw new Error('Club not provisioned: '+clubId);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('Member Master');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,4).getValues()
    .filter(r => String(r[3] || '').trim() === 'Active')
    .map(r => ({
      member_id:String(r[0] || '').trim(),
      name:String(r[1] || '').trim(),
      whatsapp:String(r[2] || '').trim(),
      status:String(r[3] || '').trim()
    }))
    .filter(m => m.member_id && m.name);
}


function apiSaveMember_(clubId,payload) {
  if (clubId !== API_DEFAULT_CLUB_ID_) throw new Error('Club not provisioned: '+clubId);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('Member Master');
  if (!sh) throw new Error('Member Master missing');
  const memberId = String(payload.member_id || '').trim() || apiNewId_('MEM');
  const name = String(payload.name || '').trim();
  const whatsapp = String(payload.whatsapp || '').trim();
  const status = String(payload.status || 'Active').trim();
  if (!name) throw new Error('name is required');
  if (!whatsapp) throw new Error('whatsapp is required');
  const rows = sh.getLastRow() >= 2 ? sh.getRange(2,1,sh.getLastRow()-1,4).getValues() : [];
  const idx = rows.findIndex(r => String(r[0] || '').trim() === memberId);
  if (idx >= 0) sh.getRange(idx+2,1,1,4).setValues([[memberId,name,whatsapp,status]]);
  else sh.appendRow([memberId,name,whatsapp,status]);
  SpreadsheetApp.flush();
  return {member_id:memberId,name:name,status:status};
}

function apiEnsureOnboardingSheets_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let settings = ss.getSheetByName(API_CLUB_SETTINGS_SHEET_);
  if (!settings) {
    settings = ss.insertSheet(API_CLUB_SETTINGS_SHEET_);
    settings.appendRow(['Club ID','Club Name','Timezone','Meeting Time','Updated At']);
    settings.hideSheet();
  }
  let excom = ss.getSheetByName(API_EXCOM_SHEET_);
  if (!excom) {
    excom = ss.insertSheet(API_EXCOM_SHEET_);
    excom.appendRow(['Club ID','Officer Role','Member ID','Member Name','Updated At']);
    excom.hideSheet();
  }
  return {settings:settings,excom:excom};
}

function apiGetClubSettings_(clubId) {
  const sheets = apiEnsureOnboardingSheets_();
  const rows = sheets.settings.getLastRow() >= 2
    ? sheets.settings.getRange(2,1,sheets.settings.getLastRow()-1,5).getValues() : [];
  const row = rows.find(r => String(r[0]) === clubId);
  return row ? {club_id:clubId,name:String(row[1]),timezone:String(row[2]),meeting_time:String(row[3])}
    : {club_id:clubId,name:clubId === API_DEFAULT_CLUB_ID_ ? 'ATK' : clubId,
       timezone:configValue_('TIMEZONE','Asia/Kolkata'),meeting_time:''};
}

function apiSaveClubSettings_(clubId,payload) {
  const sheets = apiEnsureOnboardingSheets_();
  const current = apiGetClubSettings_(clubId);
  const name = String(payload.name || current.name || '').trim();
  const timezone = String(payload.timezone || current.timezone || 'Asia/Kolkata').trim();
  const meetingTime = String(payload.meeting_time || current.meeting_time || '').trim();
  if (!name) throw new Error('Club name is required');
  if (meetingTime && !/^([01]\\d|2[0-3]):[0-5]\\d$/.test(meetingTime)) throw new Error('meeting_time must be HH:MM');
  const rows = sheets.settings.getLastRow() >= 2
    ? sheets.settings.getRange(2,1,sheets.settings.getLastRow()-1,5).getValues() : [];
  const idx = rows.findIndex(r => String(r[0]) === clubId);
  const values = [[clubId,name,timezone,meetingTime,new Date()]];
  if (idx >= 0) sheets.settings.getRange(idx+2,1,1,5).setValues(values);
  else sheets.settings.appendRow(values[0]);
  return apiGetClubSettings_(clubId);
}

function apiListExCom_(clubId) {
  const sheets = apiEnsureOnboardingSheets_();
  if (sheets.excom.getLastRow() < 2) return [];
  return sheets.excom.getRange(2,1,sheets.excom.getLastRow()-1,5).getValues()
    .filter(r => String(r[0]) === clubId)
    .map(r => ({officer_role:String(r[1]),member_id:String(r[2]),member_name:String(r[3])}));
}

function apiSaveExCom_(clubId,payload) {
  const role = String(payload.officer_role || '').trim();
  const memberId = String(payload.member_id || '').trim();
  if (!role || !memberId) throw new Error('officer_role and member_id are required');
  const member = apiMemberIndex_(clubId)[memberId];
  if (!member) throw new Error('Unknown or inactive member_id: '+memberId);
  const sheets = apiEnsureOnboardingSheets_();
  const rows = sheets.excom.getLastRow() >= 2
    ? sheets.excom.getRange(2,1,sheets.excom.getLastRow()-1,5).getValues() : [];
  const idx = rows.findIndex(r => String(r[0]) === clubId && String(r[1]) === role);
  const values = [[clubId,role,memberId,member.name,new Date()]];
  if (idx >= 0) sheets.excom.getRange(idx+2,1,1,5).setValues(values);
  else sheets.excom.appendRow(values[0]);
  return {officer_role:role,member_id:memberId,member_name:member.name};
}

function apiDefaultMeetingTemplate_(clubId) {
  return {
    club_id:clubId,
    items:[
      ['TMOD','TMOD','role'],
      ['GE','General Evaluator','role'],
      ['TTM','Table Topics Master','role'],
      ['S1','Speaker 1','speech'],
      ['S2','Speaker 2','speech'],
      ['S3','Speaker 3','speech'],
      ['EV1','Evaluator 1','role'],
      ['EV2','Evaluator 2','role'],
      ['EV3','Evaluator 3','role'],
      ['Timer','Timer','role'],
      ['Ah','Ah-Counter','role'],
      ['GRM','Grammarian','role']
    ].map((x,i)=>({
      item_id:'DEFAULT_'+x[0],
      position:i+1,
      code:x[0],
      label:x[1],
      type:x[2],
      member_id:'',
      member_name:'',
      source:'club_default'
    }))
  };
}

function apiEnsureSheets_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let meetings = ss.getSheetByName(API_MEETINGS_SHEET_);
  if (!meetings) {
    meetings = ss.insertSheet(API_MEETINGS_SHEET_);
    meetings.appendRow(['Meeting ID','Club ID','Meeting Date','Status','Created At','Updated At','Planner Meeting No']);
    meetings.hideSheet();
  }
  let items = ss.getSheetByName(API_MEETING_ITEMS_SHEET_);
  if (!items) {
    items = ss.insertSheet(API_MEETING_ITEMS_SHEET_);
    items.appendRow(['Item ID','Meeting ID','Club ID','Position','Type','Code','Label','Member ID','Member Name','Created At','Updated At']);
    items.hideSheet();
  }
  return {meetings:meetings,items:items};
}

function apiNewId_(prefix) {
  return prefix+'_'+Utilities.getUuid().replace(/-/g,'').slice(0,16).toUpperCase();
}

function apiCreateMeeting_(clubId,payload) {
  const dateText = String(payload.meeting_date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateText)) throw new Error('meeting_date must be YYYY-MM-DD');
  const sheets = apiEnsureSheets_();
  const id = apiNewId_('MTG');
  const now = new Date();
  sheets.meetings.appendRow([id,clubId,dateText,'DRAFT',now,now,'']);
  const defaults = apiDefaultMeetingTemplate_(clubId).items;
  apiReplaceMeetingItems_(sheets.items,clubId,id,defaults);
  return apiGetMeeting_(clubId,id);
}

function apiReplaceMeetingItems_(sheet,clubId,meetingId,items) {
  const now = new Date();
  const normalized = (items || []).map((item,index)=>({
    item_id:String((item.item_id && String(item.item_id).indexOf('DEFAULT_') !== 0) ? item.item_id : apiNewId_('ITM')),
    position:index+1,
    type:String(item.type || 'role'),
    code:String(item.code || ''),
    label:String(item.label || '').trim(),
    member_id:String(item.member_id || ''),
    member_name:String(item.member_name || '')
  }));
  if (normalized.some(x => !x.label)) throw new Error('Every meeting item requires a label');

  // Remove prior rows for this meeting, bottom-up.
  for (let r=sheet.getLastRow(); r>=2; r--) {
    if (String(sheet.getRange(r,2).getValue()) === meetingId) sheet.deleteRow(r);
  }
  normalized.forEach(x => sheet.appendRow([
    x.item_id,meetingId,clubId,x.position,x.type,x.code,x.label,
    x.member_id,x.member_name,now,now
  ]));
  return normalized;
}

function apiGetMeeting_(clubId,meetingId) {
  if (!meetingId) {
    const e = new Error('meeting_id is required'); e.apiCode='BAD_REQUEST'; throw e;
  }
  const sheets = apiEnsureSheets_();
  const mr = sheets.meetings.getLastRow() >= 2
    ? sheets.meetings.getRange(2,1,sheets.meetings.getLastRow()-1,7).getValues()
      .find(r => String(r[0])===meetingId && String(r[1])===clubId)
    : null;
  if (!mr) {
    const e = new Error('Meeting not found'); e.apiCode='NOT_FOUND'; throw e;
  }

  const rows = sheets.items.getLastRow() >= 2
    ? sheets.items.getRange(2,1,sheets.items.getLastRow()-1,11).getValues()
      .filter(r => String(r[1])===meetingId && String(r[2])===clubId)
    : [];
  const items = rows.map(r=>({
    item_id:String(r[0]), position:Number(r[3]), type:String(r[4]), code:String(r[5]),
    label:String(r[6]), member_id:String(r[7]), member_name:String(r[8])
  })).sort((a,b)=>a.position-b.position);

  const meetingDate = mr[2] instanceof Date
    ? Utilities.formatDate(mr[2], configValue_('TIMEZONE','Asia/Kolkata'), 'yyyy-MM-dd')
    : String(mr[2] || '').slice(0,10);
  return {
    meeting_id:String(mr[0]), club_id:String(mr[1]), meeting_date:meetingDate,
    status:String(mr[3]), planner_meeting_no:String(mr[6] || ''), items:items
  };
}


function apiListMeetings_(clubId) {
  const sheets = apiEnsureSheets_();
  if (sheets.meetings.getLastRow() < 2) return [];
  return sheets.meetings.getRange(2,1,sheets.meetings.getLastRow()-1,7).getValues()
    .filter(r => String(r[1]) === clubId)
    .map(r => ({
      meeting_id:String(r[0]),meeting_date:String(r[2]),status:String(r[3]),
      updated_at:r[5] instanceof Date ? r[5].toISOString() : String(r[5] || '')
    }))
    .sort((a,b)=>a.meeting_date.localeCompare(b.meeting_date));
}

function apiMeetingDetail_(clubId,meetingId) {
  const meeting = apiGetMeeting_(clubId,meetingId);
  const check = apiActivationCheck_(meeting);
  const actions = meeting.planner_meeting_no ? apiPendingActionsForMeeting_(Number(meeting.planner_meeting_no)) : [];
  return {
    meeting_id:meeting.meeting_id,club_id:meeting.club_id,meeting_date:meeting.meeting_date,
    status:meeting.status,items:meeting.items,
    activation:{ready:check.ready,issues:check.errors},
    pending_actions:actions
  };
}

function apiMemberIndex_(clubId) {
  const map = {};
  apiListMembers_(clubId).forEach(m => { map[m.member_id] = m; });
  return map;
}

function apiValidateMeetingItems_(items, clubId) {
  const knownCodes = {TMOD:1,GE:1,TTM:1,S1:1,S2:1,S3:1,EV1:1,EV2:1,EV3:1,Timer:1,Ah:1,GRM:1};
  const seenCodes = {};
  const members = apiMemberIndex_(clubId);
  (items || []).forEach(item => {
    const code = String(item.code || '');
    const type = String(item.type || 'role');
    if (code && !knownCodes[code]) throw new Error('Unsupported standard code: '+code);
    if (code && seenCodes[code]) throw new Error('Duplicate standard code: '+code);
    if (code) seenCodes[code] = true;
    if (type === 'session' && code) throw new Error('Custom sessions must not use a standard role code');
    const memberId = String(item.member_id || '');
    if (memberId && !members[memberId]) throw new Error('Unknown or inactive member_id: '+memberId);
    if (memberId && !String(item.member_name || '').trim()) item.member_name = members[memberId].name;
  });
}

function apiSaveMeeting_(clubId,payload) {
  const meetingId = String(payload.meeting_id || '');
  const sheets = apiEnsureSheets_();
  const meeting = apiGetMeeting_(clubId,meetingId);
  if (meeting.status === 'ACTIVE') throw new Error('Active meeting cannot be structurally edited in API V1');
  apiValidateMeetingItems_(payload.items || [],clubId);
  const saved = apiReplaceMeetingItems_(sheets.items,clubId,meetingId,payload.items || []);
  // touch updated timestamp
  for (let r=2;r<=sheets.meetings.getLastRow();r++) {
    if (String(sheets.meetings.getRange(r,1).getValue())===meetingId &&
        String(sheets.meetings.getRange(r,2).getValue())===clubId) {
      sheets.meetings.getRange(r,6).setValue(new Date());
      break;
    }
  }
  return apiGetMeeting_(clubId,meetingId);
}

function apiActivationCheck_(meeting) {
  const byCode = {};
  meeting.items.forEach(x => { if (x.code) byCode[x.code] = x; });
  const errors = [];
  if (!byCode.TMOD || !byCode.TMOD.member_id) errors.push({code:'TMOD_REQUIRED',message:'Assign a TMOD before starting automation.'});
  const speakerCodes = ['S1','S2','S3'].filter(code => byCode[code] && byCode[code].member_id);
  speakerCodes.forEach(code => {
    const ev = 'EV'+code.slice(1);
    if (!byCode[ev] || !byCode[ev].member_id) {
      errors.push({code:'EVALUATOR_REQUIRED',related:code,message:'Assign an evaluator for '+code+' before starting automation.'});
    }
  });
  return {ready:errors.length===0,errors:errors};
}

function apiActivateMeeting_(clubId,meetingId) {
  const meeting = apiGetMeeting_(clubId,meetingId);
  const check = apiActivationCheck_(meeting);
  if (!check.ready) {
    const err = new Error('Meeting is not ready for activation');
    err.apiCode = 'MEETING_NOT_READY';
    err.details = check.errors;
    throw err;
  }
  if (meeting.status === 'ACTIVE') return meeting;

  // Safety: API V1 activation does NOT send WhatsApp. It only marks the web plan
  // ready for the legacy-planner adapter. Messaging is enabled after adapter tests.
  const sheets = apiEnsureSheets_();
  for (let r=2;r<=sheets.meetings.getLastRow();r++) {
    if (String(sheets.meetings.getRange(r,1).getValue())===meetingId &&
        String(sheets.meetings.getRange(r,2).getValue())===clubId) {
      sheets.meetings.getRange(r,4,1,3).setValues([['READY_FOR_SYNC',sheets.meetings.getRange(r,5).getValue(),new Date()]]);
      break;
    }
  }
  return apiGetMeeting_(clubId,meetingId);
}


/**
 * Synchronizes an ordered API meeting into the legacy planner.
 * This writes only the fields the proven automation already understands.
 * Custom sessions remain in API Meeting Items and are handled by Programme Planning.
 */
function apiSyncMeetingToLegacyPlanner_(clubId, meetingId) {
  const allowed = String(PropertiesService.getScriptProperties().getProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_) || '').toUpperCase() === 'TRUE';
  if (!allowed) {
    const err = new Error('Legacy sync is disabled in this environment');
    err.apiCode = 'LEGACY_SYNC_DISABLED';
    throw err;
  }
  if (clubId !== API_DEFAULT_CLUB_ID_) throw new Error('Legacy adapter currently supports ATK only');
  const meeting = apiGetMeeting_(clubId, meetingId);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  if (!planner) throw new Error('VPE Weekly Planner missing');

  const rows = planner.getRange(3,1,Math.max(0,planner.getLastRow()-2),28).getValues();
  const wanted = String(meeting.meeting_date);
  let idx = rows.findIndex(r => {
    if (!(r[1] instanceof Date)) return false;
    return Utilities.formatDate(r[1], configValue_('TIMEZONE','Asia/Kolkata'), 'yyyy-MM-dd') === wanted;
  });
  if (idx < 0) throw new Error('No legacy planner row exists for meeting date '+wanted);

  const rowNo = idx + 3;
  const byCode = {};
  meeting.items.forEach(x => { if (x.code) byCode[x.code] = x; });
  const memberName = code => String((byCode[code] && byCode[code].member_name) || '').trim();

  // Speakers and their status columns.
  [['S1',7,9],['S2',11,13],['S3',15,17]].forEach(x => {
    planner.getRange(rowNo,x[1]).setValue(memberName(x[0]));
    planner.getRange(rowNo,x[2]).setValue(memberName(x[0]) ? 'Planned' : '');
  });

  // Standard role columns.
  ROLE_DEFS_.forEach(def => planner.getRange(rowNo,def.column).setValue(memberName(def.code)));
  SpreadsheetApp.flush();

  const meetingNo = Number(planner.getRange(rowNo,1).getValue());
  const sheets = apiEnsureSheets_();
  for (let r=2;r<=sheets.meetings.getLastRow();r++) {
    if (String(sheets.meetings.getRange(r,1).getValue())===meetingId) {
      sheets.meetings.getRange(r,7).setValue(meetingNo);
      sheets.meetings.getRange(r,6).setValue(new Date());
      break;
    }
  }
  return {meeting_id:meetingId,planner_meeting_no:meetingNo,synced:true};
}


/**
 * VPE-facing dashboard projection. Sheets remain an implementation detail:
 * the client receives meetings/readiness/actions as domain objects only.
 */

function apiPendingActionLabel_(type,member,related,state) {
  const names = {
    THEME:'Theme pending', GRAMMAR:'Grammarian input pending',
    SPEAKER_DETAILS:'Speech details pending', INTRO:'Introduction pending',
    EVALUATOR_COORDINATION:'Evaluator coordination pending'
  };
  const title = names[type] || (String(type || '').replace(/_/g,' ').toLowerCase() + ' pending');
  return {
    title:title.charAt(0).toUpperCase()+title.slice(1),
    description:[member,related].filter(Boolean).join(' · '),
    severity:state === 'Escalated' ? 'attention' : 'pending'
  };
}

function apiPendingActionsForMeeting_(meetingNo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tasks = ss.getSheetByName(TASK_SHEET_);
  if (!tasks || tasks.getLastRow() < 2) return [];
  return tasks.getRange(2,1,tasks.getLastRow()-1,18).getValues()
    .filter(r => Number(r[1]) === Number(meetingNo) &&
      ['Pending','Sent','Escalated'].indexOf(String(r[5] || '')) >= 0)
    .map(r => {
      const ui = apiPendingActionLabel_(String(r[2] || ''),String(r[3] || ''),String(r[4] || ''),String(r[5] || ''));
      return {
        action_id:String(r[0] || ''),kind:String(r[2] || ''),member_name:String(r[3] || ''),
        related:String(r[4] || ''),state:String(r[5] || ''),title:ui.title,
        description:ui.description,severity:ui.severity,
        due_at:r[6] instanceof Date ? r[6].toISOString() : String(r[6] || '')
      };
    });
}

function apiDashboard_(clubId) {
  if (clubId !== API_DEFAULT_CLUB_ID_) throw new Error('Club not provisioned: '+clubId);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  const tasks = ss.getSheetByName(TASK_SHEET_);
  if (!planner) return {club_id:clubId,meetings:[],pending_actions:[]};

  const today = startOfDay_(new Date());
  const rows = planner.getLastRow() >= 3
    ? planner.getRange(3,1,Math.min(60,planner.getLastRow()-2),28).getValues() : [];
  const taskRows = tasks && tasks.getLastRow() >= 2
    ? tasks.getRange(2,1,tasks.getLastRow()-1,18).getValues() : [];

  const meetings = rows.filter(row => {
    return Number(row[0]) && row[1] instanceof Date &&
      startOfDay_(row[1]) >= today &&
      ['Completed','Cancelled'].indexOf(String(row[2] || '')) < 0;
  }).slice(0,4).map(row => {
    const readiness = buildReadinessRow_(row,taskRows);
    return {
      meeting_no:Number(row[0]),
      meeting_date:Utilities.formatDate(row[1],configValue_('TIMEZONE','Asia/Kolkata'),'yyyy-MM-dd'),
      status:String(row[2] || ''),
      readiness:{
        roles:readiness[1],
        programme:readiness[2],
        theme:readiness[3],
        grammar:readiness[4],
        introductions:readiness[5],
        speaker_details:readiness[6],
        evaluators:readiness[7],
        coordination:readiness[8],
        overall:readiness[9]
      }
    };
  });

  const pending = taskRows.filter(r =>
    ['Pending','Sent','Escalated'].indexOf(String(r[5] || '')) >= 0
  ).map(r => {
    const ui = apiPendingActionLabel_(String(r[2] || ''),String(r[3] || ''),String(r[4] || ''),String(r[5] || ''));
    return {
      action_id:String(r[0] || ''),meeting_no:Number(r[1] || 0),
      kind:String(r[2] || ''),member_name:String(r[3] || ''),related:String(r[4] || ''),
      state:String(r[5] || ''),title:ui.title,description:ui.description,severity:ui.severity,
      due_at:r[6] instanceof Date ? r[6].toISOString() : String(r[6] || '')
    };
  });

  return {club_id:clubId,meetings:meetings,pending_actions:pending};
}


function apiClubBootstrap_(clubId) {
  return {
    club:apiGetClubSettings_(clubId),
    excom:apiListExCom_(clubId),
    meeting_template:apiDefaultMeetingTemplate_(clubId),
    capabilities:{
      meeting_planning:true,
      custom_sessions:true,
      ordered_items:true,
      dashboard:true,
      whatsapp_activation:false,
      legacy_sync:false
    }
  };
}


/**
 * Non-destructive DEV smoke test. Does not sync the legacy planner or send WhatsApp.
 * Run manually after deploying the DEV Apps Script version.
 */
function apiDevSmokeTest_() {
  const report = [];
  const check = (name,fn) => {
    try { const value=fn(); report.push({name:name,ok:true,value:value}); }
    catch (e) { report.push({name:name,ok:false,error:String(e && e.message || e)}); }
  };
  check('members',()=>({count:apiListMembers_(API_DEFAULT_CLUB_ID_).length}));
  check('club_settings',()=>apiGetClubSettings_(API_DEFAULT_CLUB_ID_));
  check('excom',()=>({count:apiListExCom_(API_DEFAULT_CLUB_ID_).length}));
  check('template',()=>({items:apiDefaultMeetingTemplate_(API_DEFAULT_CLUB_ID_).items.length}));
  check('dashboard',()=>{
    const d=apiDashboard_(API_DEFAULT_CLUB_ID_);
    return {meetings:d.meetings.length,pending_actions:d.pending_actions.length};
  });
  check('legacy_sync_guard',()=>{
    const enabled=String(PropertiesService.getScriptProperties().getProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_) || '').toUpperCase()==='TRUE';
    if (enabled) throw new Error('WEB_API_ALLOW_LEGACY_SYNC must be FALSE/absent during frontend integration');
    return {blocked:true};
  });
  check('api_key_configured',()=>{
    if (!String(PropertiesService.getScriptProperties().getProperty('WEB_API_KEY') || '')) throw new Error('WEB_API_KEY missing');
    return {configured:true};
  });
  return {ok:report.every(x=>x.ok),checks:report};
}


/**
 * Public editor entrypoint for the non-destructive DEV API smoke test.
 */
function apiDevSmokeTest() {
  const result = apiDevSmokeTest_();
  console.log(JSON.stringify(result, null, 2));
  return result;
}
