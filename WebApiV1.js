/**
 * Web API V1 for the VPE Assistant.
 *
 * This is deliberately an adapter layer. It gives the Flutter/Web client an
 * ordered meeting model without changing the proven ATK automation tables.
 * Existing automation continues to read VPE Weekly Planner.
 */

const API_V1_VERSION_ = '2026-10-11.25';
let DEV_SUPPRESS_TEST_DETAIL_LOGS_ = false;
const API_MEETING_ITEMS_SHEET_ = 'API Meeting Items';
const API_MEETINGS_SHEET_ = 'API Meetings';
const API_DEFAULT_CLUB_ID_ = 'ATK';
const API_ALLOW_LEGACY_SYNC_PROPERTY_ = 'WEB_API_ALLOW_LEGACY_SYNC';
const API_AUTO_SYNC_PLANNED_PROPERTY_ = 'WEB_API_AUTO_SYNC_PLANNED';
const API_CLUB_SETTINGS_SHEET_ = 'API Club Settings';
const API_EXCOM_SHEET_ = 'API ExCom';
const API_CACHE_SECONDS_ = 900;
let API_SS_CACHE_ = null;

function apiSpreadsheet_() {
  if (!API_SS_CACHE_) API_SS_CACHE_ = SpreadsheetApp.openById(SPREADSHEET_ID);
  return API_SS_CACHE_;
}

function apiCache_() {
  return CacheService.getScriptCache();
}

function apiCacheKey_(kind,clubId,suffix) {
  return ['VPE_API_V1',kind,String(clubId || ''),String(suffix || '')].join('|');
}

function apiCacheGetJson_(key) {
  try {
    const raw = apiCache_().get(key);
    return raw ? JSON.parse(raw) : null;
  } catch (ignore) {
    return null;
  }
}

function apiCachePutJson_(key,value,seconds) {
  try {
    apiCache_().put(key,JSON.stringify(value),seconds || API_CACHE_SECONDS_);
  } catch (ignore) {}
}

function apiCacheRemove_(keys) {
  try { apiCache_().removeAll(keys); } catch (ignore) {}
}

function apiInvalidateClubReadCaches_(clubId) {
  apiCacheRemove_([
    apiCacheKey_('members',clubId),
    apiCacheKey_('settings',clubId),
    apiCacheKey_('excom',clubId),
    apiCacheKey_('meetings',clubId)
  ]);
}

function apiInvalidateMeetingCaches_(clubId,meetingId) {
  apiCacheRemove_([
    apiCacheKey_('meetings',clubId),
    apiCacheKey_('meeting',clubId,meetingId)
  ]);
}

function apiCacheMeeting_(meeting) {
  if (!meeting || !meeting.club_id || !meeting.meeting_id) return meeting;
  apiCachePutJson_(
    apiCacheKey_('meeting',meeting.club_id,meeting.meeting_id),
    meeting,
    300
  );
  return meeting;
}

function apiFindMeetingRecord_(sheets,clubId,meetingId) {
  if (!meetingId) {
    const e = new Error('meeting_id is required'); e.apiCode='BAD_REQUEST'; throw e;
  }
  const rows = sheets.meetings.getLastRow() >= 2
    ? sheets.meetings.getRange(2,1,sheets.meetings.getLastRow()-1,7).getValues()
    : [];
  const idx = rows.findIndex(r => String(r[0])===meetingId && String(r[1])===clubId);
  if (idx < 0) {
    const e = new Error('Meeting not found'); e.apiCode='NOT_FOUND'; throw e;
  }
  return {row_number:idx+2,values:rows[idx]};
}


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
      case 'meeting.ensure': {
        const ensured = apiEnsureMeetingForPlanner_(
          clubId,Number(payload.planner_meeting_no || 0)
        );
        return {
          status:200,
          data:apiMeetingDetail_(clubId,ensured.meeting_id)
        };
      }
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
    const clientCodes = [
      'BAD_REQUEST','NOT_FOUND','MEETING_NOT_READY','LEGACY_SYNC_DISABLED',
      'SYNC_CONFLICT','SYNC_MEMBER_NOT_FOUND','ANOTHER_MEETING_ACTIVE'
    ];
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
  // Member roster is small and operationally sensitive. Always read the
  // current Member Master so UI dropdowns reflect roster/status changes
  // immediately instead of serving a stale cached subset.
  if (clubId !== API_DEFAULT_CLUB_ID_) throw new Error('Club not provisioned: '+clubId);

  const sh = apiSpreadsheet_().getSheetByName('Member Master');
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
  const ss = apiSpreadsheet_();
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
  apiInvalidateClubReadCaches_(clubId);
  return {member_id:memberId,name:name,status:status};
}

function apiEnsureOnboardingSheets_() {
  const ss = apiSpreadsheet_();
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
  const cacheKey = apiCacheKey_('settings',clubId);
  const cached = apiCacheGetJson_(cacheKey);
  if (cached) return cached;

  const sheets = apiEnsureOnboardingSheets_();
  const rows = sheets.settings.getLastRow() >= 2
    ? sheets.settings.getRange(2,1,sheets.settings.getLastRow()-1,5).getValues() : [];
  const row = rows.find(r => String(r[0]) === clubId);
  const settings = row
    ? {club_id:clubId,name:String(row[1]),timezone:String(row[2]),meeting_time:String(row[3])}
    : {club_id:clubId,name:clubId === API_DEFAULT_CLUB_ID_ ? 'ATK' : clubId,
       timezone:configValue_('TIMEZONE','Asia/Kolkata'),meeting_time:''};
  apiCachePutJson_(cacheKey,settings);
  return settings;
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
  apiInvalidateClubReadCaches_(clubId);
  return apiGetClubSettings_(clubId);
}

function apiListExCom_(clubId) {
  const cacheKey = apiCacheKey_('excom',clubId);
  const cached = apiCacheGetJson_(cacheKey);
  if (cached) return cached;

  const sheets = apiEnsureOnboardingSheets_();
  if (sheets.excom.getLastRow() < 2) {
    apiCachePutJson_(cacheKey,[]);
    return [];
  }
  const excom = sheets.excom.getRange(2,1,sheets.excom.getLastRow()-1,5).getValues()
    .filter(r => String(r[0]) === clubId)
    .map(r => ({officer_role:String(r[1]),member_id:String(r[2]),member_name:String(r[3])}));
  apiCachePutJson_(cacheKey,excom);
  return excom;
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
  apiInvalidateClubReadCaches_(clubId);
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
  const ss = apiSpreadsheet_();
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

function apiMemberByName_(clubId, memberName) {
  const wanted = String(memberName || '').trim();
  if (!wanted) return null;
  return apiListMembers_(clubId).find(m =>
    String(m.name || '').trim() === wanted
  ) || null;
}

function apiFindPlannerMeetingByDate_(dateText) {
  const planner = apiSpreadsheet_().getSheetByName('VPE Weekly Planner');
  if (!planner || planner.getLastRow() < 3) return null;

  const timezone = configValue_('TIMEZONE','Asia/Kolkata');
  const rows = planner.getRange(3,1,planner.getLastRow()-2,28).getValues();
  const idx = rows.findIndex(row =>
    row[1] instanceof Date &&
    Utilities.formatDate(row[1], timezone, 'yyyy-MM-dd') === String(dateText || '')
  );
  if (idx < 0) return null;

  return {
    row_number: idx + 3,
    meeting_no: Number(rows[idx][0] || 0),
    values: rows[idx]
  };
}

function apiPlannerAssignments_(plannerRow) {
  const row = plannerRow || [];
  const assignments = {
    S1:String(row[6] || '').trim(),
    S2:String(row[10] || '').trim(),
    S3:String(row[14] || '').trim()
  };
  ROLE_DEFS_.forEach(def => {
    assignments[def.code] = String(row[def.column - 1] || '').trim();
  });
  return assignments;
}

function apiReconcileItemsWithPlanner_(clubId, meetingDate, items) {
  const plannerMeeting = apiFindPlannerMeetingByDate_(meetingDate);
  const cloned = (items || []).map(item => Object.assign({}, item));
  if (!plannerMeeting) {
    return {items:cloned, planner_meeting_no:'', imported:[], conflicts:[]};
  }

  const plannerAssignments = apiPlannerAssignments_(plannerMeeting.values);
  const imported = [];
  const conflicts = [];

  cloned.forEach(item => {
    const code = String(item.code || '');
    if (!code || !Object.prototype.hasOwnProperty.call(plannerAssignments, code)) return;

    const plannerName = String(plannerAssignments[code] || '').trim();
    const apiName = String(item.member_name || '').trim();
    if (!plannerName) return;

    if (!apiName) {
      const member = apiMemberByName_(clubId, plannerName);
      if (!member) {
        const e = new Error(
          'Planner member is not active in Member Master: ' + plannerName
        );
        e.apiCode = 'SYNC_MEMBER_NOT_FOUND';
        throw e;
      }
      item.member_id = member.member_id;
      item.member_name = member.name;
      imported.push({code:code,member_name:member.name});
      return;
    }

    if (apiName !== plannerName) {
      conflicts.push({
        code:code,
        api_member:apiName,
        planner_member:plannerName
      });
    }
  });

  if (conflicts.length) {
    const e = new Error(
      'API/planner assignment conflict: ' +
      conflicts.map(x =>
        x.code + ' (' + x.api_member + ' vs ' + x.planner_member + ')'
      ).join(', ')
    );
    e.apiCode = 'SYNC_CONFLICT';
    e.details = conflicts;
    throw e;
  }

  return {
    items:cloned,
    planner_meeting_no:plannerMeeting.meeting_no || '',
    imported:imported,
    conflicts:[]
  };
}

function apiFindMeetingForPlannerMeeting_(clubId, meetingNo) {
  const sheets = apiEnsureSheets_();
  const rows = sheets.meetings.getLastRow() >= 2
    ? sheets.meetings.getRange(2,1,sheets.meetings.getLastRow()-1,7).getValues()
    : [];

  let idx = rows.findIndex(r =>
    String(r[1] || '') === String(clubId || '') &&
    Number(r[6]) === Number(meetingNo)
  );
  if (idx >= 0) {
    return {row_number:idx+2, values:rows[idx]};
  }

  const planner = apiSpreadsheet_().getSheetByName('VPE Weekly Planner');
  if (!planner || planner.getLastRow() < 3) return null;
  const ids = planner.getRange(3,1,planner.getLastRow()-2,1).getValues().flat();
  const plannerIdx = ids.findIndex(v => Number(v) === Number(meetingNo));
  if (plannerIdx < 0) return null;

  const date = planner.getRange(plannerIdx+3,2).getValue();
  if (!(date instanceof Date)) return null;
  const dateText = Utilities.formatDate(
    date, configValue_('TIMEZONE','Asia/Kolkata'), 'yyyy-MM-dd'
  );

  idx = rows.findIndex(r => {
    if (String(r[1] || '') !== String(clubId || '')) return false;
    const raw = r[2];
    const apiDate = raw instanceof Date
      ? Utilities.formatDate(raw, configValue_('TIMEZONE','Asia/Kolkata'), 'yyyy-MM-dd')
      : String(raw || '').slice(0,10);
    return apiDate === dateText;
  });

  if (idx < 0) return null;

  if (!Number(rows[idx][6])) {
    sheets.meetings.getRange(idx+2,7).setValue(Number(meetingNo));
    sheets.meetings.getRange(idx+2,6).setValue(new Date());
    SpreadsheetApp.flush();
    apiInvalidateMeetingCaches_(clubId, String(rows[idx][0] || ''));
    rows[idx][6] = Number(meetingNo);
  }

  return {row_number:idx+2, values:rows[idx]};
}

function apiMirrorPlannerAssignmentToMeeting_(clubId, meetingNo, code, memberName) {
  const found = apiFindMeetingForPlannerMeeting_(clubId, meetingNo);
  if (!found) return {synced:false, reason:'NO_API_MEETING'};

  const meetingId = String(found.values[0] || '');
  const member = apiMemberByName_(clubId, memberName);
  if (!member) {
    throw new Error(
      'Active member not found while syncing planner booking: ' + memberName
    );
  }

  const sheets = apiEnsureSheets_();
  const rows = sheets.items.getLastRow() >= 2
    ? sheets.items.getRange(2,1,sheets.items.getLastRow()-1,11).getValues()
    : [];
  const idx = rows.findIndex(r =>
    String(r[1] || '') === meetingId &&
    String(r[2] || '') === String(clubId || '') &&
    String(r[5] || '') === String(code || '')
  );
  if (idx < 0) {
    throw new Error(
      'API meeting item missing for Meeting ' + meetingNo + ': ' + code
    );
  }

  const currentName = String(rows[idx][8] || '').trim();
  if (currentName && currentName !== member.name) {
    const e = new Error(
      'API/planner assignment conflict for ' + code +
      ': ' + currentName + ' vs ' + member.name
    );
    e.apiCode = 'SYNC_CONFLICT';
    e.details = [{
      code:String(code || ''),
      api_member:currentName,
      planner_member:member.name
    }];
    throw e;
  }

  const now = new Date();
  sheets.items.getRange(idx+2,8,1,2)
    .setValues([[member.member_id, member.name]]);
  sheets.items.getRange(idx+2,11).setValue(now);
  sheets.meetings.getRange(found.row_number,6).setValue(now);
  SpreadsheetApp.flush();
  apiInvalidateMeetingCaches_(clubId,meetingId);

  return {
    synced:true,
    meeting_id:meetingId,
    code:String(code || ''),
    member_id:member.member_id
  };
}


function apiEnsureMeetingForPlanner_(clubId, meetingNo) {
  const wantedNo = Number(meetingNo || 0);
  if (!wantedNo) {
    const e = new Error('planner_meeting_no is required');
    e.apiCode = 'BAD_REQUEST';
    throw e;
  }

  const existing = apiFindMeetingForPlannerMeeting_(clubId,wantedNo);
  if (existing) {
    return apiGetMeeting_(clubId,String(existing.values[0] || ''));
  }

  const planner = apiSpreadsheet_().getSheetByName('VPE Weekly Planner');
  if (!planner || planner.getLastRow() < 3) {
    const e = new Error('Planner meeting not found: ' + wantedNo);
    e.apiCode = 'NOT_FOUND';
    throw e;
  }

  const ids = planner.getRange(
    3,1,planner.getLastRow()-2,1
  ).getValues().flat();
  const idx = ids.findIndex(v => Number(v) === wantedNo);
  if (idx < 0) {
    const e = new Error('Planner meeting not found: ' + wantedNo);
    e.apiCode = 'NOT_FOUND';
    throw e;
  }

  const date = planner.getRange(idx+3,2).getValue();
  if (!(date instanceof Date)) {
    const e = new Error('Planner meeting date is missing for Meeting ' + wantedNo);
    e.apiCode = 'BAD_REQUEST';
    throw e;
  }

  const dateText = Utilities.formatDate(
    date, configValue_('TIMEZONE','Asia/Kolkata'), 'yyyy-MM-dd'
  );
  return apiCreateMeeting_(clubId,{meeting_date:dateText});
}


function apiCreateMeeting_(clubId,payload) {
  const dateText = String(payload.meeting_date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateText)) {
    throw new Error('meeting_date must be YYYY-MM-DD');
  }

  const sheets = apiEnsureSheets_();

  // Do not create duplicate API meetings for the same club/date.
  const existing = sheets.meetings.getLastRow() >= 2
    ? sheets.meetings.getRange(2,1,sheets.meetings.getLastRow()-1,7).getValues()
    : [];
  const duplicate = existing.find(r => {
    if (String(r[1] || '') !== clubId) return false;
    const raw = r[2];
    const existingDate = raw instanceof Date
      ? Utilities.formatDate(raw, configValue_('TIMEZONE','Asia/Kolkata'), 'yyyy-MM-dd')
      : String(raw || '').slice(0,10);
    return existingDate === dateText;
  });
  if (duplicate) {
    const e = new Error('An API meeting already exists for ' + dateText);
    e.apiCode = 'BAD_REQUEST';
    e.details = {meeting_id:String(duplicate[0] || '')};
    throw e;
  }

  const id = apiNewId_('MTG');
  const now = new Date();
  const defaults = apiDefaultMeetingTemplate_(clubId).items;
  const reconciled = apiReconcileItemsWithPlanner_(clubId,dateText,defaults);

  sheets.meetings.getRange(sheets.meetings.getLastRow()+1,1,1,7)
    .setValues([[
      id,clubId,dateText,'DRAFT',now,now,
      reconciled.planner_meeting_no || ''
    ]]);

  const normalized = apiReplaceMeetingItems_(
    sheets.items,clubId,id,reconciled.items
  );
  apiInvalidateMeetingCaches_(clubId,id);

  return apiCacheMeeting_({
    meeting_id:id,
    club_id:clubId,
    meeting_date:dateText,
    status:'DRAFT',
    planner_meeting_no:String(reconciled.planner_meeting_no || ''),
    items:normalized
  });
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
  if (normalized.some(x => !x.label)) {
    const e = new Error('Every meeting item requires a label'); e.apiCode='BAD_REQUEST'; throw e;
  }

  const added = normalized.map(x => [
    x.item_id,meetingId,clubId,x.position,x.type,x.code,x.label,
    x.member_id,x.member_name,now,now
  ]);

  // Targeted update: touch only this meeting's rows instead of rewriting
  // the entire API Meeting Items sheet on every save.
  const lastRow = sheet.getLastRow();
  const existing = lastRow >= 2
    ? sheet.getRange(2,1,lastRow-1,11).getValues()
    : [];
  const rowNumbers = [];
  existing.forEach((r,i) => {
    if (String(r[1]) === meetingId && String(r[2]) === clubId) rowNumbers.push(i + 2);
  });

  if (!rowNumbers.length) {
    if (added.length) {
      sheet.getRange(sheet.getLastRow()+1,1,added.length,11).setValues(added);
    }
  } else {
    const contiguous = rowNumbers.every((r,i) => i === 0 || r === rowNumbers[i-1] + 1);

    if (contiguous && added.length <= rowNumbers.length) {
      // Fast path for normal edits/deletions: overwrite existing block once,
      // then clear only any surplus old rows.
      if (added.length) {
        sheet.getRange(rowNumbers[0],1,added.length,11).setValues(added);
      }
      const surplus = rowNumbers.length - added.length;
      if (surplus > 0) {
        sheet.getRange(rowNumbers[0] + added.length,1,surplus,11).clearContent();
      }
    } else {
      // If more rows are needed (e.g. custom session added), clear the old
      // meeting rows only and append the new block once at the bottom.
      if (contiguous) {
        sheet.getRange(rowNumbers[0],1,rowNumbers.length,11).clearContent();
      } else {
        // Legacy/test data may be fragmented; clear only matched rows.
        rowNumbers.forEach(r => sheet.getRange(r,1,1,11).clearContent());
      }
      if (added.length) {
        sheet.getRange(sheet.getLastRow()+1,1,added.length,11).setValues(added);
      }
    }
  }

  apiInvalidateMeetingCaches_(clubId,meetingId);
  return normalized;
}

function apiGetMeeting_(clubId,meetingId) {
  if (!meetingId) {
    const e = new Error('meeting_id is required'); e.apiCode='BAD_REQUEST'; throw e;
  }
  const cacheKey = apiCacheKey_('meeting',clubId,meetingId);
  const cached = apiCacheGetJson_(cacheKey);
  if (cached) return cached;

  const sheets = apiEnsureSheets_();
  const rec = apiFindMeetingRecord_(sheets,clubId,meetingId);
  const mr = rec.values;
  const beforeItems = apiGetMeeting_(clubId,meetingId).items || [];

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
  return apiCacheMeeting_({
    meeting_id:String(mr[0]), club_id:String(mr[1]), meeting_date:meetingDate,
    status:String(mr[3]), planner_meeting_no:String(mr[6] || ''), items:items
  });
}

function apiListMeetings_(clubId) {
  const cacheKey = apiCacheKey_('meetings',clubId);
  const cached = apiCacheGetJson_(cacheKey);
  if (cached) return cached;

  const sheets = apiEnsureSheets_();
  if (sheets.meetings.getLastRow() < 2) {
    apiCachePutJson_(cacheKey,[]);
    return [];
  }
  const meetings = sheets.meetings.getRange(2,1,sheets.meetings.getLastRow()-1,7).getValues()
    .filter(r => String(r[1]) === clubId)
    .map(r => ({
      meeting_id:String(r[0]),
      meeting_date:r[2] instanceof Date
        ? Utilities.formatDate(r[2],configValue_('TIMEZONE','Asia/Kolkata'),'yyyy-MM-dd')
        : String(r[2] || '').slice(0,10),
      status:String(r[3]),
      planner_meeting_no:String(r[6] || ''),
      updated_at:r[5] instanceof Date ? r[5].toISOString() : String(r[5] || '')
    }))
    .sort((a,b)=>a.meeting_date.localeCompare(b.meeting_date));
  apiCachePutJson_(cacheKey,meetings,300);
  return meetings;
}

function apiMeetingDetail_(clubId,meetingId) {
  const meeting = apiGetMeeting_(clubId,meetingId);
  const check = apiActivationCheck_(meeting);
  const actions = meeting.planner_meeting_no ? apiPendingActionsForMeeting_(Number(meeting.planner_meeting_no)) : [];
  return {
    meeting_id:meeting.meeting_id,
    club_id:meeting.club_id,
    meeting_date:meeting.meeting_date,
    status:meeting.status,
    planner_meeting_no:meeting.planner_meeting_no || '',
    items:meeting.items,
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
    if (code && !knownCodes[code]) { const e=new Error('Unsupported standard code: '+code); e.apiCode='BAD_REQUEST'; throw e; }
    if (code && seenCodes[code]) { const e=new Error('Duplicate standard code: '+code); e.apiCode='BAD_REQUEST'; throw e; }
    if (code) seenCodes[code] = true;
    if (type === 'session' && code) { const e=new Error('Custom sessions must not use a standard role code'); e.apiCode='BAD_REQUEST'; throw e; }
    const memberId = String(item.member_id || '');
    if (memberId && !members[memberId]) { const e=new Error('Unknown or inactive member_id: '+memberId); e.apiCode='BAD_REQUEST'; throw e; }
    if (memberId && !String(item.member_name || '').trim()) item.member_name = members[memberId].name;
  });
}

function apiRunLateTmodCatchUpSafely_(meetingNo) {
  if (!meetingNo || typeof sendLateTmodCatchUp_ !== 'function') return false;
  try {
    return Boolean(sendLateTmodCatchUp_(meetingNo));
  } catch (error) {
    if (typeof logAutomation_ === 'function') {
      logAutomation_(
        'TMOD Catch-up',
        meetingNo,
        '',
        'Late TMOD handoff',
        'Error',
        String(error && error.message || error)
      );
    }
    return false;
  }
}

function apiRunMeetingActivationKickoffSafely_(meetingNo) {
  const report = {
    meeting_no:Number(meetingNo || 0),
    speaker_details:false,
    assigned_role_prompts:false,
    grammar_prompt:false,
    evaluator_pairing:false,
    role_invites:null
  };
  if (!report.meeting_no) return report;

  const run = (label, fn) => {
    try {
      return fn();
    } catch (error) {
      if (typeof logAutomation_ === 'function') {
        logAutomation_(
          'Meeting Activation',
          report.meeting_no,
          '',
          label,
          'Error',
          String(error && error.message || error)
        );
      }
      return false;
    }
  };

  if (typeof processSpeakerDetailPrompts_ === 'function') {
    run('Start speaker detail collection', () => {
      processSpeakerDetailPrompts_(report.meeting_no);
      report.speaker_details = true;
      return true;
    });
  }

  if (typeof processAssignedRolePrompts_ === 'function') {
    run('Start assigned-role collection', () => {
      processAssignedRolePrompts_(report.meeting_no);
      report.assigned_role_prompts = true;
      return true;
    });
  }

  if (typeof processMissingGrammarAssignmentPrompts_ === 'function') {
    run('Start Grammarian collection', () => {
      processMissingGrammarAssignmentPrompts_(report.meeting_no);
      report.grammar_prompt = true;
      return true;
    });
  }

  if (typeof processEvaluatorNotificationsForMeeting_ === 'function') {
    run('Start evaluator coordination', () => {
      processEvaluatorNotificationsForMeeting_(report.meeting_no);
      report.evaluator_pairing = true;
      return true;
    });
  }

  if (typeof sendVacantRoleInvitesNow_ === 'function') {
    report.role_invites = run(
      'Send vacant-role invitations',
      () => sendVacantRoleInvitesNow_(report.meeting_no)
    );
  }

  return report;
}


function apiBridgePlannedMeetingIfEnabled_(clubId,meetingId) {
  const props = PropertiesService.getScriptProperties();
  const autoSync = String(props.getProperty(API_AUTO_SYNC_PLANNED_PROPERTY_) || '').toUpperCase() === 'TRUE';
  if (!autoSync) return {enabled:false,synced:false};

  const allowLegacy = String(props.getProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_) || '').toUpperCase() === 'TRUE';
  if (!allowLegacy) {
    const e = new Error('Auto-sync is enabled but legacy sync permission is disabled');
    e.apiCode = 'LEGACY_SYNC_DISABLED';
    throw e;
  }

  const result = apiSyncMeetingToLegacyPlanner_(clubId,meetingId);

  // Refresh legacy-derived programme/task projections after the planner write.
  if (typeof syncProgrammeSlots_ === 'function') syncProgrammeSlots_();
  if (typeof syncPhase1Tasks_ === 'function' && configBool_('TASK_ENGINE_ENABLED')) syncPhase1Tasks_();

  apiRunLateTmodCatchUpSafely_(result.planner_meeting_no);

  return {enabled:true,synced:true,planner_meeting_no:result.planner_meeting_no};
}

function apiSaveMeeting_(clubId,payload) {
  const meetingId = String(payload.meeting_id || '');
  const saveMode = String(payload.save_mode || 'DRAFT').toUpperCase();
  if (['DRAFT','CONFIRMED'].indexOf(saveMode) < 0) {
    const e = new Error('save_mode must be DRAFT or CONFIRMED');
    e.apiCode='BAD_REQUEST';
    throw e;
  }

  const sheets = apiEnsureSheets_();
  const rec = apiFindMeetingRecord_(sheets,clubId,meetingId);
  const mr = rec.values;
  // Snapshot assignments before replacing items. Confirmation notifications compare
  // old vs new speaker bookings and must not fail after persisting the plan.
  const beforeItems = apiGetMeeting_(clubId,meetingId).items || [];
  if (String(mr[3]) === 'ACTIVE') {
    throw new Error('Active meeting cannot be structurally edited in API V1');
  }

  const meetingDate = mr[2] instanceof Date
    ? Utilities.formatDate(mr[2], configValue_('TIMEZONE','Asia/Kolkata'), 'yyyy-MM-dd')
    : String(mr[2] || '').slice(0,10);

  // Import planner/WhatsApp assignments into any still-blank API slots.
  // Conflicting non-blank assignments are rejected rather than overwritten.
  const reconciled = apiReconcileItemsWithPlanner_(
    clubId, meetingDate, payload.items || []
  );

  apiValidateMeetingItems_(reconciled.items,clubId);
  const normalized = apiReplaceMeetingItems_(
    sheets.items,clubId,meetingId,reconciled.items
  );

  const nextStatus = saveMode === 'CONFIRMED' ? 'PLANNED' : 'DRAFT';
  const now = new Date();
  const plannerNo = reconciled.planner_meeting_no || mr[6] || '';
  sheets.meetings.getRange(rec.row_number,4,1,4)
    .setValues([[nextStatus,mr[4],now,plannerNo]]);

  apiInvalidateMeetingCaches_(clubId,meetingId);

  let plannerMeetingNo = String(plannerNo || '');
  if (saveMode === 'CONFIRMED') {
    const bridge = apiBridgePlannedMeetingIfEnabled_(clubId,meetingId);
    if (bridge && bridge.synced) {
      plannerMeetingNo = String(bridge.planner_meeting_no || '');
    }

    apiSendNewSpeechBookingConfirmations_(
      clubId,
      Number(plannerMeetingNo || 0),
      beforeItems,
      normalized
    );
  }

  return apiCacheMeeting_({
    meeting_id:meetingId,
    club_id:clubId,
    meeting_date:meetingDate,
    status:nextStatus,
    planner_meeting_no:plannerMeetingNo,
    items:normalized
  });
}

function apiSendNewSpeechBookingConfirmations_(clubId, meetingNo, beforeItems, afterItems) {
  if (String(clubId || '') !== 'ATK' ||
      !Number(meetingNo) ||
      typeof sendSpeechBookingConfirmation_ !== 'function') return;

  const before = {};
  (beforeItems || []).forEach(item => {
    const code = String(item.code || '');
    if (/^S[123]$/.test(code)) {
      before[code] = String(item.member_name || '').trim();
    }
  });

  const after = {};
  (afterItems || []).forEach(item => {
    const code = String(item.code || '');
    if (/^S[123]$/.test(code)) {
      after[code] = String(item.member_name || '').trim();
    }
  });

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const planner = ss.getSheetByName('VPE Weekly Planner');
  let dateText = '';
  if (planner && planner.getLastRow() >= 3) {
    const ids = planner.getRange(3, 1, planner.getLastRow() - 2, 1).getValues().flat();
    const idx = ids.findIndex(v => Number(v) === Number(meetingNo));
    if (idx >= 0) {
      const d = planner.getRange(idx + 3, 2).getValue();
      if (d instanceof Date) {
        dateText = Utilities.formatDate(
          d,
          configValue_('TIMEZONE', 'Asia/Kolkata'),
          'EEE, d MMM'
        );
      }
    }
  }

  ['S1','S2','S3'].forEach(code => {
    const member = after[code] || '';
    if (!member || member === (before[code] || '')) return;

    const phone = normalizePhone_(phoneForMember_(member));
    if (!phone) return;

    sendSpeechBookingConfirmation_(
      member,
      phone,
      Number(meetingNo),
      dateText,
      Number(code.slice(1))
    );
  });
}


function apiActivationCheck_(meeting) {
  const errors = [];
  if (!meeting || !meeting.meeting_date) {
    errors.push({code:'MEETING_DATE_REQUIRED',message:'Meeting date is required.'});
  }
  if (!meeting || !meeting.items || !meeting.items.length) {
    errors.push({code:'MEETING_STRUCTURE_REQUIRED',message:'Add at least one meeting item.'});
  }
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

  const sheets = apiEnsureSheets_();

  if (sheets.meetings.getLastRow() >= 2) {
    const rows = sheets.meetings.getRange(
      2, 1, sheets.meetings.getLastRow() - 1, 7
    ).getValues();

    const otherActive = rows.find(row =>
      String(row[1] || '') === String(clubId || '') &&
      String(row[0] || '') !== String(meetingId || '') &&
      String(row[3] || '').toUpperCase() === 'ACTIVE'
    );

    if (otherActive) {
      const err = new Error(
        'Another meeting is already active. Complete or close it before activating this meeting.'
      );
      err.apiCode = 'ANOTHER_MEETING_ACTIVE';
      err.details = [{
        meeting_id: String(otherActive[0] || ''),
        planner_meeting_no: Number(otherActive[6] || 0) || '',
        message: 'Only one meeting can be ACTIVE at a time.'
      }];
      throw err;
    }
  }

  let meetingRow = 0;
  for (let r=2;r<=sheets.meetings.getLastRow();r++) {
    if (String(sheets.meetings.getRange(r,1).getValue())===meetingId &&
        String(sheets.meetings.getRange(r,2).getValue())===clubId) {
      meetingRow = r;
      sheets.meetings.getRange(r,4,1,3).setValues([[
        'READY_FOR_SYNC',
        sheets.meetings.getRange(r,5).getValue(),
        new Date()
      ]]);
      apiInvalidateMeetingCaches_(clubId,meetingId);
      break;
    }
  }
  if (!meetingRow) {
    const err = new Error('Meeting not found during activation');
    err.apiCode = 'NOT_FOUND';
    throw err;
  }

  if (configValue_('SYSTEM_MODE','') === 'DEV') {
    const props = PropertiesService.getScriptProperties();
    const originalAllow = props.getProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_);
    try {
      props.setProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_,'TRUE');
      const synced = apiSyncMeetingToLegacyPlanner_(clubId,meetingId);

      if (typeof syncProgrammeSlots_ === 'function') syncProgrammeSlots_();
      if (typeof syncPhase1Tasks_ === 'function' &&
          configBool_('TASK_ENGINE_ENABLED')) {
        syncPhase1Tasks_();
      }

      sheets.meetings.getRange(meetingRow,4).setValue('ACTIVE');
      sheets.meetings.getRange(meetingRow,6).setValue(new Date());
      if (synced && synced.planner_meeting_no) {
        sheets.meetings.getRange(meetingRow,7)
          .setValue(synced.planner_meeting_no);
      }
      SpreadsheetApp.flush();
      apiInvalidateMeetingCaches_(clubId,meetingId);
      const activatedMeetingNo = synced && synced.planner_meeting_no;
      apiRunLateTmodCatchUpSafely_(activatedMeetingNo);
      apiRunMeetingActivationKickoffSafely_(activatedMeetingNo);
    } catch (error) {
      // Activation is atomic: a failed reconciliation/sync must not leave
      // the meeting stranded in READY_FOR_SYNC.
      sheets.meetings.getRange(meetingRow,4).setValue(
        meeting.status || 'PLANNED'
      );
      sheets.meetings.getRange(meetingRow,6).setValue(new Date());
      SpreadsheetApp.flush();
      apiInvalidateMeetingCaches_(clubId,meetingId);
      throw error;
    } finally {
      if (originalAllow===null) {
        props.deleteProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_);
      } else {
        props.setProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_,originalAllow);
      }
    }
  }

  return apiGetMeeting_(clubId,meetingId);
}


/**
 * Builds the exact legacy-planner values implied by an ordered API meeting.
 * Pure mapping: no sheet writes, safe for previews/tests.
 */
function apiBuildLegacyPlannerProjection_(meeting) {
  const byCode = {};
  (meeting.items || []).forEach(x => { if (x.code) byCode[x.code] = x; });
  const name = code => String((byCode[code] && byCode[code].member_name) || '').trim();

  const speakers = {
    S1:name('S1'), S2:name('S2'), S3:name('S3')
  };
  const roles = {};
  ROLE_DEFS_.forEach(def => { roles[def.code] = name(def.code); });

  return {speakers:speakers,roles:roles};
}

function apiApplyProjectionToPlannerRow_(row, projection) {
  const out = (row || []).slice();
  while (out.length < 28) out.push('');

  [['S1',7,9],['S2',11,13],['S3',15,17]].forEach(x => {
    out[x[1]-1] = projection.speakers[x[0]] || '';
    out[x[2]-1] = projection.speakers[x[0]] ? 'Planned' : '';
  });
  ROLE_DEFS_.forEach(def => {
    out[def.column-1] = projection.roles[def.code] || '';
  });
  return out;
}

/**
 * Synchronizes an ordered API meeting into the legacy planner.
 * This writes only the fields the proven automation already understands.
 * Custom sessions remain in API Meeting Items and are handled by Programme Planning.
 */
function apiSyncMeetingToLegacyPlanner_(clubId, meetingId) {
  const allowed = String(
    PropertiesService.getScriptProperties()
      .getProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_) || ''
  ).toUpperCase() === 'TRUE';
  if (!allowed) {
    const err = new Error('Legacy sync is disabled in this environment');
    err.apiCode = 'LEGACY_SYNC_DISABLED';
    throw err;
  }
  if (clubId !== API_DEFAULT_CLUB_ID_) {
    throw new Error('Legacy adapter currently supports ATK only');
  }

  let meeting = apiGetMeeting_(clubId, meetingId);
  const ss = apiSpreadsheet_();
  const planner = ss.getSheetByName('VPE Weekly Planner');
  if (!planner) throw new Error('VPE Weekly Planner missing');

  const rows = planner.getRange(
    3,1,Math.max(0,planner.getLastRow()-2),28
  ).getValues();
  const wanted = String(meeting.meeting_date);
  const idx = rows.findIndex(r => {
    if (!(r[1] instanceof Date)) return false;
    return Utilities.formatDate(
      r[1], configValue_('TIMEZONE','Asia/Kolkata'), 'yyyy-MM-dd'
    ) === wanted;
  });
  if (idx < 0) {
    throw new Error('No legacy planner row exists for meeting date ' + wanted);
  }

  const rowNo = idx + 3;

  // Reconcile just before the write. This catches WhatsApp/planner bookings
  // made after the API draft was opened and prevents silent overwrites.
  const reconciled = apiReconcileItemsWithPlanner_(
    clubId, meeting.meeting_date, meeting.items
  );
  if (reconciled.imported.length) {
    apiReplaceMeetingItems_(
      apiEnsureSheets_().items, clubId, meetingId, reconciled.items
    );
    apiInvalidateMeetingCaches_(clubId,meetingId);
    meeting = apiGetMeeting_(clubId,meetingId);
  }

  const projection = apiBuildLegacyPlannerProjection_(meeting);

  [['S1',7,9],['S2',11,13],['S3',15,17]].forEach(x => {
    planner.getRange(rowNo,x[1]).setValue(projection.speakers[x[0]] || '');
    planner.getRange(rowNo,x[2]).setValue(
      projection.speakers[x[0]] ? 'Planned' : ''
    );
  });

  ROLE_DEFS_.forEach(def =>
    planner.getRange(rowNo,def.column)
      .setValue(projection.roles[def.code] || '')
  );
  SpreadsheetApp.flush();

  const meetingNo = Number(planner.getRange(rowNo,1).getValue());
  const sheets = apiEnsureSheets_();
  for (let r=2;r<=sheets.meetings.getLastRow();r++) {
    if (String(sheets.meetings.getRange(r,1).getValue())===meetingId) {
      sheets.meetings.getRange(r,7).setValue(meetingNo);
      sheets.meetings.getRange(r,6).setValue(new Date());
      apiInvalidateMeetingCaches_(clubId,meetingId);
      break;
    }
  }

  return {
    meeting_id:meetingId,
    planner_meeting_no:meetingNo,
    synced:true,
    imported_from_planner:reconciled.imported
  };
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
  const ss = apiSpreadsheet_();
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

function apiVacantPositionsFromPlannerRow_(row) {
  const defs = [
    {code:'TMOD', label:'Toastmaster of the Day', column:19, type:'role'},
    {code:'GE', label:'General Evaluator', column:20, type:'role'},
    {code:'TTM', label:'Table Topics Master', column:21, type:'role'},
    {code:'S1', label:'Speaker 1', column:7, type:'speech'},
    {code:'S2', label:'Speaker 2', column:11, type:'speech'},
    {code:'S3', label:'Speaker 3', column:15, type:'speech'},
    {code:'EV1', label:'Evaluator 1', column:26, type:'evaluator'},
    {code:'EV2', label:'Evaluator 2', column:27, type:'evaluator'},
    {code:'EV3', label:'Evaluator 3', column:28, type:'evaluator'},
    {code:'Timer', label:'Timer', column:22, type:'role'},
    {code:'Ah', label:'Ah-Counter', column:23, type:'role'},
    {code:'GRM', label:'Grammarian', column:24, type:'role'}
  ];

  return defs
    .filter(def => !String(row[def.column - 1] || '').trim())
    .map(def => ({
      code:def.code,
      label:def.label,
      type:def.type
    }));
}


function apiDashboard_(clubId) {
  if (clubId !== API_DEFAULT_CLUB_ID_) throw new Error('Club not provisioned: '+clubId);
  const ss = apiSpreadsheet_();
  const planner = ss.getSheetByName('VPE Weekly Planner');
  const tasks = ss.getSheetByName(TASK_SHEET_);
  if (!planner) return {club_id:clubId,meetings:[],pending_actions:[]};

  const today = startOfDay_(new Date());
  const rows = planner.getLastRow() >= 3
    ? planner.getRange(3,1,Math.max(0, planner.getLastRow() - 2),28).getValues() : [];
  const taskRows = tasks && tasks.getLastRow() >= 2
    ? tasks.getRange(2,1,tasks.getLastRow()-1,18).getValues() : [];

  const upcomingRows = rows.filter(row => {
    return Number(row[0]) && row[1] instanceof Date &&
      startOfDay_(row[1]) >= today &&
      ['Completed','Cancelled'].indexOf(String(row[2] || '')) < 0;
  }).slice(0,4);

  const meetings = upcomingRows.map(row => {
    const readiness = buildReadinessRow_(row,taskRows);
    const vacancies = apiVacantPositionsFromPlannerRow_(row);
    return {
      meeting_no:Number(row[0]),
      meeting_date:Utilities.formatDate(row[1],configValue_('TIMEZONE','Asia/Kolkata'),'yyyy-MM-dd'),
      status:String(row[2] || ''),
      vacant_positions:vacancies,
      vacant_count:vacancies.length,
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

  // Backward-compatible Home-page vacancy projection.
  // Every unfilled meeting position is surfaced as VACANT_ROLE regardless of
  // how the automation engine internally models speakers/evaluators.
  const vacancyActions = [];
  upcomingRows.forEach(row => {
    const meetingNo = Number(row[0]);
    const meetingDate = row[1];
    apiVacantPositionsFromPlannerRow_(row).forEach(v => {
      const key = meetingNo + '|' + v.code;
      const alreadyPresent = pending.some(p =>
        Number(p.meeting_no) === meetingNo &&
        String(p.kind || '') === 'VACANT_ROLE' &&
        String(p.related || '') === v.code
      );
      if (alreadyPresent) return;

      vacancyActions.push({
        action_id:'VACANCY|' + key,
        meeting_no:meetingNo,
        kind:'VACANT_ROLE',
        member_name:'',
        related:v.code,
        state:'Pending',
        title:v.label + ' vacant',
        description:v.label,
        severity:'pending',
        due_at:meetingDate instanceof Date ? meetingDate.toISOString() : ''
      });
    });
  });

  const allPending = pending.concat(vacancyActions);
  return {
    club_id:clubId,
    meetings:meetings,
    vacant_positions:meetings.length ? meetings[0].vacant_positions : [],
    pending_actions:allPending
  };
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
      planner_lifecycle_sync:true,
      meeting_ensure:true,
      whatsapp_booking_sync:true,
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


/**
 * Public DEV integration test runner.
 * Safe by design: uses API storage only, does not sync legacy planner and does not send WhatsApp.
 */
function apiDevIntegrationTest() {
  const clubId = API_DEFAULT_CLUB_ID_;
  const results = [];
  const record = (name, ok, detail) => results.push({name:name,ok:!!ok,detail:detail || ''});
  let meetingId = '';

  try {
    const members = apiListMembers_(clubId);
    if (members.length < 4) throw new Error('Need at least 4 active members for integration test');

    const created = apiCreateMeeting_(clubId,{meeting_date:'2026-12-31'});
    meetingId = created.meeting_id;
    record('create_meeting', created.status === 'DRAFT' && created.items.length === 12,
      created.meeting_id);

    // Draft save with a manual TMOD assignment.
    const draftItems = created.items.map(x => Object.assign({},x));
    draftItems[0].member_id = members[0].member_id;
    draftItems[0].member_name = members[0].name;
    const draft = apiSaveMeeting_(clubId,{
      meeting_id:meetingId,save_mode:'DRAFT',items:draftItems
    });
    record('draft_save', draft.status === 'DRAFT' && draft.items[0].member_id === members[0].member_id,
      draft.status);

    // Confirmed save with vacancies intentionally left open.
    const confirmed = apiSaveMeeting_(clubId,{
      meeting_id:meetingId,save_mode:'CONFIRMED',items:draftItems
    });
    record('confirmed_save_allows_vacancies', confirmed.status === 'PLANNED', confirmed.status);

    // Reload integrity.
    const reloaded = apiGetMeeting_(clubId,meetingId);
    record('reload_integrity',
      reloaded.status === 'PLANNED' && reloaded.items[0].member_id === members[0].member_id,
      reloaded.status);

    // Invalid member must be rejected as BAD_REQUEST.
    let invalidMemberOk = false;
    try {
      const bad = reloaded.items.map(x => Object.assign({},x));
      bad[1].member_id = 'DOES_NOT_EXIST';
      bad[1].member_name = 'Invalid';
      apiSaveMeeting_(clubId,{meeting_id:meetingId,save_mode:'DRAFT',items:bad});
    } catch (e) {
      invalidMemberOk = e && e.apiCode === 'BAD_REQUEST';
    }
    record('invalid_member_rejected', invalidMemberOk);

    // Duplicate standard code must be rejected as BAD_REQUEST.
    let duplicateRoleOk = false;
    try {
      const dup = reloaded.items.map(x => Object.assign({},x));
      dup[1].code = dup[0].code;
      apiSaveMeeting_(clubId,{meeting_id:meetingId,save_mode:'DRAFT',items:dup});
    } catch (e) {
      duplicateRoleOk = e && e.apiCode === 'BAD_REQUEST';
    }
    record('duplicate_role_rejected', duplicateRoleOk);

    // Legacy sync must remain disabled during frontend integration.
    const syncEnabled = String(PropertiesService.getScriptProperties()
      .getProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_) || '').toUpperCase() === 'TRUE';
    record('legacy_sync_guard', !syncEnabled, syncEnabled ? 'ENABLED - unsafe for DEV test' : 'blocked');
    const autoSyncEnabled = String(PropertiesService.getScriptProperties()
      .getProperty(API_AUTO_SYNC_PLANNED_PROPERTY_) || '').toUpperCase() === 'TRUE';
    record('auto_sync_guard', !autoSyncEnabled, autoSyncEnabled ? 'ENABLED - unsafe for DEV test' : 'blocked');

  } catch (e) {
    record('runner_exception', false, String(e && e.message || e));
  }

  const report = {
    ok: results.length > 0 && results.every(x => x.ok),
    meeting_id: meetingId,
    checks: results
  };
  console.log(JSON.stringify(report,null,2));
  return report;
}


/**
 * Verifies the web-plan -> legacy-role-call bridge without sheet writes or messaging.
 */
function apiDevRoleCallBridgeTest() {
  const clubId = API_DEFAULT_CLUB_ID_;
  const members = apiListMembers_(clubId);
  if (members.length < 6) throw new Error('Need at least 6 active members');

  const meeting = {
    meeting_id:'TEST_BRIDGE',
    meeting_date:'2026-12-30',
    items:apiDefaultMeetingTemplate_(clubId).items.map(x => Object.assign({},x))
  };
  const byCode = {};
  meeting.items.forEach(x => byCode[x.code] = x);

  byCode.TMOD.member_id=members[0].member_id; byCode.TMOD.member_name=members[0].name;
  byCode.GE.member_id=members[1].member_id; byCode.GE.member_name=members[1].name;
  byCode.S1.member_id=members[2].member_id; byCode.S1.member_name=members[2].name;
  byCode.EV1.member_id=members[3].member_id; byCode.EV1.member_name=members[3].name;
  byCode.S2.member_id=members[4].member_id; byCode.S2.member_name=members[4].name;
  // S3 deliberately vacant, therefore EV3 must not be offered.

  const projection = apiBuildLegacyPlannerProjection_(meeting);
  const syntheticRow = apiApplyProjectionToPlannerRow_(new Array(28).fill(''),projection);
  const available = availableRoleDefs_(syntheticRow).map(x => x.code);

  const checks = [
    {name:'manual_tmod_hidden',ok:available.indexOf('TMOD') < 0},
    {name:'manual_ge_hidden',ok:available.indexOf('GE') < 0},
    {name:'manual_ev1_hidden',ok:available.indexOf('EV1') < 0},
    {name:'vacant_ttm_offered',ok:available.indexOf('TTM') >= 0},
    {name:'ev2_offered_for_assigned_s2',ok:available.indexOf('EV2') >= 0},
    {name:'ev3_hidden_without_s3',ok:available.indexOf('EV3') < 0}
  ];

  const report = {ok:checks.every(x=>x.ok),available_roles:available,checks:checks};
  console.log(JSON.stringify(report,null,2));
  return report;
}


/**
 * One-shot DEV backend readiness test.
 * Safe: no legacy planner writes and no WhatsApp sends.
 */
function apiDevBackendReadinessTest() {
  const sections = [];
  const run = (name, fn) => {
    try {
      const result = fn();
      sections.push({name:name,ok:!!(result && result.ok),result:result});
    } catch (e) {
      sections.push({name:name,ok:false,error:String(e && e.message || e)});
    }
  };

  run('smoke', apiDevSmokeTest_);
  run('meeting_planner', apiDevIntegrationTest);
  run('role_call_bridge', apiDevRoleCallBridgeTest);
  run('dashboard_contract', () => {
    const d = apiDashboard_(API_DEFAULT_CLUB_ID_);
    const valid = d && Array.isArray(d.meetings) && Array.isArray(d.pending_actions);
    return {ok:valid,meetings:d.meetings.length,pending_actions:d.pending_actions.length};
  });

  const report = {
    ok: sections.every(x => x.ok),
    version: API_V1_VERSION_,
    sections: sections
  };
  console.log(JSON.stringify(report,null,2));
  return report;
}


/**
 * DEV-only internal timing check. Does not send WhatsApp or write the legacy planner.
 * It times backend operations only; browser->Apps Script redirect latency is not included.
 */
function apiDevPerformanceTest() {
  const clubId = API_DEFAULT_CLUB_ID_;
  const time = (name,fn) => {
    const start = Date.now();
    const value = fn();
    return {name:name,ms:Date.now()-start,value:value};
  };

  apiInvalidateClubReadCaches_(clubId);
  const coldMembers = time('members_cold',()=>apiListMembers_(clubId).length);
  const warmMembers = time('members_cached',()=>apiListMembers_(clubId).length);
  const coldMeetings = time('meetings_cold',()=>apiListMeetings_(clubId).length);
  const warmMeetings = time('meetings_cached',()=>apiListMeetings_(clubId).length);
  const settings = time('club_settings',()=>apiGetClubSettings_(clubId).club_id);

  const meetingList = apiListMeetings_(clubId);
  let meetingTimings = [];
  if (meetingList.length) {
    const id = meetingList[0].meeting_id;
    apiCacheRemove_([apiCacheKey_('meeting',clubId,id)]);
    meetingTimings = [
      time('meeting_get_cold',()=>apiGetMeeting_(clubId,id).items.length),
      time('meeting_get_cached',()=>apiGetMeeting_(clubId,id).items.length)
    ];
  }

  const report = {
    ok:true,
    version:API_V1_VERSION_,
    timings:[coldMembers,warmMembers,coldMeetings,warmMeetings,settings].concat(meetingTimings)
  };
  console.log(JSON.stringify(report,null,2));
  return report;
}


/**
 * Measures internal save-path time without HTTP/redirect overhead.
 * Uses an existing DEV API meeting and does not touch legacy planner/WhatsApp.
 */
function apiDevSavePerformanceTest() {
  const clubId = API_DEFAULT_CLUB_ID_;
  const meetings = apiListMeetings_(clubId);
  if (!meetings.length) throw new Error('No API meetings available for save performance test');

  const meeting = apiGetMeeting_(clubId,meetings[0].meeting_id);
  const before = Date.now();
  const saved = apiSaveMeeting_(clubId,{
    meeting_id:meeting.meeting_id,
    save_mode:meeting.status === 'PLANNED' ? 'CONFIRMED' : 'DRAFT',
    items:meeting.items
  });
  const ms = Date.now()-before;

  const report = {
    ok:true,
    version:API_V1_VERSION_,
    meeting_id:saved.meeting_id,
    internal_save_ms:ms,
    item_count:saved.items.length
  };
  console.log(JSON.stringify(report,null,2));
  return report;
}


/**
 * Reversible end-to-end bridge test:
 * API meeting -> actual legacy planner row -> existing role availability logic.
 *
 * Safety:
 * - Uses only a historical Completed/Cancelled planner row.
 * - Backs up and restores the entire 28-column planner row.
 * - Temporarily enables legacy sync only inside this invocation and restores
 *   the original property value in finally.
 * - Requires WEB_API_AUTO_SYNC_PLANNED to remain FALSE/absent.
 * - Does not send WhatsApp.
 * - Deletes the temporary API meeting/items after verification.
 */
function apiDevControlledLegacySyncTest() {
  const clubId = API_DEFAULT_CLUB_ID_;
  const props = PropertiesService.getScriptProperties();
  const autoSync = String(props.getProperty(API_AUTO_SYNC_PLANNED_PROPERTY_) || '').toUpperCase() === 'TRUE';
  if (autoSync) throw new Error('Set WEB_API_AUTO_SYNC_PLANNED to FALSE before running this test.');

  const ss = apiSpreadsheet_();
  const planner = ss.getSheetByName('VPE Weekly Planner');
  if (!planner) throw new Error('VPE Weekly Planner missing');

  const members = apiListMembers_(clubId);
  if (members.length < 6) throw new Error('Need at least 6 active members for bridge test');

  // Create our own temporary, cancelled planner row so the test never depends
  // on or overwrites an existing live/historical meeting.
  const existingCount = Math.max(0, planner.getLastRow() - 2);
  const existingRows = existingCount
    ? planner.getRange(3,1,existingCount,28).getValues()
    : [];
  const maxMeetingNo = existingRows.reduce((m,r)=>Math.max(m,Number(r[0])||0),0);
  const meetingNo = maxMeetingNo + 1000;
  const meetingDate = new Date(2099,11,31);
  const dateText = Utilities.formatDate(
    meetingDate,
    configValue_('TIMEZONE','Asia/Kolkata'),
    'yyyy-MM-dd'
  );

  const tempPlannerRow = new Array(28).fill('');
  tempPlannerRow[0] = meetingNo;
  tempPlannerRow[1] = meetingDate;
  tempPlannerRow[2] = 'Cancelled';

  const tempPlannerRowNumber = planner.getLastRow() + 1;
  planner.getRange(tempPlannerRowNumber,1,1,28).setValues([tempPlannerRow]);

  let tempMeetingId = '';
  const originalAllow = props.getProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_);
  let report = null;

  try {
    const created = apiCreateMeeting_(clubId,{meeting_date:dateText});
    tempMeetingId = created.meeting_id;
    const items = created.items.map(x => Object.assign({},x));
    const byCode = {};
    items.forEach(x => { byCode[x.code] = x; });

    byCode.TMOD.member_id=members[0].member_id; byCode.TMOD.member_name=members[0].name;
    byCode.GE.member_id=members[1].member_id; byCode.GE.member_name=members[1].name;
    byCode.S1.member_id=members[2].member_id; byCode.S1.member_name=members[2].name;
    byCode.EV1.member_id=members[3].member_id; byCode.EV1.member_name=members[3].name;
    byCode.S2.member_id=members[4].member_id; byCode.S2.member_name=members[4].name;
    // Leave TTM/Timer/Ah/GRM/EV2 vacant; S3 vacant means EV3 must stay hidden.

    apiSaveMeeting_(clubId,{
      meeting_id:tempMeetingId,
      save_mode:'DRAFT',
      items:items
    });

    props.setProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_,'TRUE');
    const syncResult = apiSyncMeetingToLegacyPlanner_(clubId,tempMeetingId);

    const projectedRow = planner.getRange(tempPlannerRowNumber,1,1,28).getValues()[0];
    const available = availableRoleDefs_(projectedRow).map(x => x.code);

    const checks = [
      {name:'synced_to_temporary_meeting',ok:Number(syncResult.planner_meeting_no)===meetingNo},
      {name:'speaker_1_written',ok:String(projectedRow[6]||'')===members[2].name},
      {name:'speaker_2_written',ok:String(projectedRow[10]||'')===members[4].name},
      {name:'manual_tmod_written',ok:String(projectedRow[18]||'')===members[0].name},
      {name:'manual_ge_written',ok:String(projectedRow[19]||'')===members[1].name},
      {name:'manual_ev1_written',ok:String(projectedRow[25]||'')===members[3].name},
      {name:'manual_tmod_hidden_from_role_call',ok:available.indexOf('TMOD')<0},
      {name:'manual_ge_hidden_from_role_call',ok:available.indexOf('GE')<0},
      {name:'manual_ev1_hidden_from_role_call',ok:available.indexOf('EV1')<0},
      {name:'vacant_ttm_available',ok:available.indexOf('TTM')>=0},
      {name:'vacant_timer_available',ok:available.indexOf('Timer')>=0},
      {name:'vacant_ah_available',ok:available.indexOf('Ah')>=0},
      {name:'vacant_grammarian_available',ok:available.indexOf('GRM')>=0},
      {name:'ev2_available_for_s2',ok:available.indexOf('EV2')>=0},
      {name:'ev3_hidden_without_s3',ok:available.indexOf('EV3')<0}
    ];

    report = {
      ok:checks.every(x=>x.ok),
      version:API_V1_VERSION_,
      temporary_meeting_no:meetingNo,
      temporary_meeting_date:dateText,
      temporary_api_meeting_id:tempMeetingId,
      available_roles:available,
      checks:checks
    };
  } finally {
    // Restore sync permission exactly as it was.
    if (originalAllow === null) props.deleteProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_);
    else props.setProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_,originalAllow);

    // Remove temporary API meeting/items.
    if (tempMeetingId) {
      const sheets = apiEnsureSheets_();
      for (let r=sheets.items.getLastRow(); r>=2; r--) {
        if (String(sheets.items.getRange(r,2).getValue())===tempMeetingId) sheets.items.deleteRow(r);
      }
      for (let r=sheets.meetings.getLastRow(); r>=2; r--) {
        if (String(sheets.meetings.getRange(r,1).getValue())===tempMeetingId) sheets.meetings.deleteRow(r);
      }
      apiInvalidateMeetingCaches_(clubId,tempMeetingId);
    }

    // Delete the temporary planner row we created.
    if (planner.getLastRow() >= tempPlannerRowNumber) {
      planner.deleteRow(tempPlannerRowNumber);
    }
  }

  console.log(JSON.stringify(report,null,2));
  return report;
}


const API_DEV_ROLE_PREVIEW_PROPERTY_ = 'DEV_WEB_ROLE_PREVIEW_V1';
const API_DEV_ROLE_BOOKING_PROPERTY_ = 'DEV_WEB_ROLE_BOOKING_V1';
const API_DEV_GOLDEN_PROPERTY_ = 'DEV_WEB_GOLDEN_ROLE_V1';
const API_DEV_GOLDEN_RESULT_PROPERTY_ = 'DEV_WEB_GOLDEN_ROLE_RESULT_V1';

/**
 * Sends one real approved "View roles" WhatsApp template to the sole DEV
 * test recipient. The role list behind it is generated from a temporary API
 * meeting that is actually synced through the legacy bridge first.
 *
 * The temporary API/planner records are removed before this function returns.
 * A short-lived projection is kept only so the test recipient can tap
 * "View roles" and see the exact list. Role booking is intercepted in DEV.
 */
function apiDevWhatsAppBridgePreviewTest() {
  const clubId = API_DEFAULT_CLUB_ID_;
  if (configValue_('SYSTEM_MODE','') !== 'DEV' ||
      String(configValue_('WA_SEND_MODE','TEST')).toUpperCase() !== 'TEST') {
    throw new Error('WhatsApp bridge preview requires SYSTEM_MODE=DEV and WA_SEND_MODE=TEST.');
  }
  if (!configBool_('ROLE_SELECTION_ENABLED')) {
    throw new Error('ROLE_SELECTION_ENABLED must be TRUE for this preview.');
  }
  if (!isApprovedTemplate_('atk_role_reminder_v3')) {
    throw new Error('atk_role_reminder_v3 is not marked APPROVED.');
  }

  const member = configValue_('VPE_MEMBER_NAME','Riyas AF');
  const phone = normalizePhone_(phoneForMember_(member));
  const allowlist = String(configValue_('WA_TEST_RECIPIENTS',''))
    .split(',').map(normalizePhone_).filter(Boolean);

  if (member !== 'Riyas AF' || !phone ||
      allowlist.indexOf(phone) === -1 || !isWhatsAppRecipientAllowed_(phone)) {
    throw new Error('DEV WhatsApp bridge preview requires Riyas to be an allowlisted test recipient.');
  }

  const props = PropertiesService.getScriptProperties();
  const autoSync = String(props.getProperty(API_AUTO_SYNC_PLANNED_PROPERTY_) || '').toUpperCase() === 'TRUE';
  if (autoSync) throw new Error('Keep WEB_API_AUTO_SYNC_PLANNED=FALSE for this controlled test.');

  const ss = apiSpreadsheet_();
  const planner = ss.getSheetByName('VPE Weekly Planner');
  if (!planner) throw new Error('VPE Weekly Planner missing');

  const members = apiListMembers_(clubId).filter(m => m.name !== member);
  if (members.length < 5) throw new Error('Need at least five other active members for preview.');

  // Remove any leftover DEV preview rows from an interrupted earlier test.
  // These are identifiable by status TEST and a far-future 2099 date.
  const preCount = Math.max(0, planner.getLastRow() - 2);
  if (preCount) {
    const preRows = planner.getRange(3,1,preCount,28).getValues();
    preRows.forEach((row,i) => {
      const d=row[1];
      if (String(row[2]||'')==='TEST' && d instanceof Date && d.getFullYear()===2099) {
        planner.getRange(i+3,1,1,28).clearContent();
      }
    });
  }

  const existingCount = Math.max(0, planner.getLastRow() - 2);
  const existingRows = existingCount ? planner.getRange(3,1,existingCount,28).getValues() : [];
  const maxMeetingNo = existingRows.reduce((m,r)=>Math.max(m,Number(r[0])||0),0);
  const meetingNo = maxMeetingNo + 2000;

  // Far-future DEV-only date. Leftover 2099 TEST rows are cleared above first,
  // so the date match used by the real bridge is unambiguous.
  const meetingDate = new Date(2099,11,30);
  const dateText = Utilities.formatDate(meetingDate,configValue_('TIMEZONE','Asia/Kolkata'),'yyyy-MM-dd');
  const tempPlannerRow = new Array(28).fill('');
  tempPlannerRow[0]=meetingNo;
  tempPlannerRow[1]=meetingDate;
  tempPlannerRow[2]='TEST';
  const plannerRowNo=planner.getLastRow()+1;
  planner.getRange(plannerRowNo,1,1,28).setValues([tempPlannerRow]);

  const originalAllow=props.getProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_);
  let tempMeetingId='';
  let preview=null;

  try {
    const created=apiCreateMeeting_(clubId,{meeting_date:dateText});
    tempMeetingId=created.meeting_id;
    const items=created.items.map(x=>Object.assign({},x));
    const byCode={}; items.forEach(x=>{byCode[x.code]=x;});

    // Fixed assignments; expected vacancies:
    // TTM, Timer, Ah, GRM and EV2. S3 is vacant, so EV3 must not appear.
    byCode.TMOD.member_id=members[0].member_id; byCode.TMOD.member_name=members[0].name;
    byCode.GE.member_id=members[1].member_id; byCode.GE.member_name=members[1].name;
    byCode.S1.member_id=members[2].member_id; byCode.S1.member_name=members[2].name;
    byCode.EV1.member_id=members[3].member_id; byCode.EV1.member_name=members[3].name;
    byCode.S2.member_id=members[4].member_id; byCode.S2.member_name=members[4].name;

    apiSaveMeeting_(clubId,{meeting_id:tempMeetingId,save_mode:'DRAFT',items:items});
    props.setProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_,'TRUE');
    const sync=apiSyncMeetingToLegacyPlanner_(clubId,tempMeetingId);

    const syncedRow=planner.getRange(plannerRowNo,1,1,28).getValues()[0];

    // Fail closed if the bridge did not write to the exact temporary row.
    if (String(syncedRow[18]||'') !== members[0].name ||
        String(syncedRow[19]||'') !== members[1].name ||
        String(syncedRow[6]||'') !== members[2].name ||
        String(syncedRow[25]||'') !== members[3].name ||
        String(syncedRow[10]||'') !== members[4].name) {
      throw new Error('Controlled bridge preview did not update the expected temporary planner row.');
    }

    const available=availableRoleDefs_(syncedRow).map(def=>({
      code:def.code,
      name:def.name,
      speakerName:def.speakerName || ''
    }));

    preview={
      member:member,
      phone:phone,
      meetingNo:Number(sync.planner_meeting_no),
      meetingDate:dateText,
      available:available,
      expiresAt:Date.now()+15*60*1000
    };
    props.setProperty(API_DEV_ROLE_PREVIEW_PROPERTY_,JSON.stringify(preview));

    const result=sendApprovedRoleCall_(
      phone,preview.meetingNo,preview.meetingDate
    );
    if (!result || !result.ok) {
      props.deleteProperty(API_DEV_ROLE_PREVIEW_PROPERTY_);
      throw new Error('WhatsApp preview template failed: '+String(result && result.error || 'Unknown Meta error'));
    }

    const report={
      ok:true,
      version:API_V1_VERSION_,
      message_id:result.messageId || '',
      test_recipient:member,
      temporary_meeting_no:preview.meetingNo,
      expected_available_roles:available.map(x=>x.code),
      instruction:'Tap View roles in WhatsApp. Do not choose a role; selection is DEV-blocked anyway.'
    };
    if (!DEV_SUPPRESS_TEST_DETAIL_LOGS_) {
      console.log(JSON.stringify(report,null,2));
    }
    return report;
  } finally {
    if (originalAllow === null) props.deleteProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_);
    else props.setProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_,originalAllow);

    if (tempMeetingId) {
      const sheets=apiEnsureSheets_();

      // Bulk cleanup only. Avoid per-row Spreadsheet calls because Meta delivery
      // webhooks can arrive concurrently with this DEV test.
      const itemLast=sheets.items.getLastRow();
      if (itemLast>=2) {
        const itemIds=sheets.items.getRange(2,2,itemLast-1,1).getValues().flat();
        const itemRows=[];
        itemIds.forEach((v,i)=>{ if (String(v)===tempMeetingId) itemRows.push(i+2); });
        if (itemRows.length) {
          const first=itemRows[0];
          const contiguous=itemRows.every((v,i)=>i===0 || v===itemRows[i-1]+1);
          if (contiguous) {
            sheets.items.getRange(first,1,itemRows.length,11).clearContent();
          } else {
            // Extremely unlikely for this temporary record; one bulk rewrite is
            // still safer than hundreds of row-by-row calls.
            const all=sheets.items.getRange(2,1,itemLast-1,11).getValues();
            const kept=all.filter(row=>String(row[1])!==tempMeetingId);
            const area=sheets.items.getRange(2,1,itemLast-1,11);
            area.clearContent();
            if (kept.length) sheets.items.getRange(2,1,kept.length,11).setValues(kept);
          }
        }
      }

      const meetingLast=sheets.meetings.getLastRow();
      if (meetingLast>=2) {
        const ids=sheets.meetings.getRange(2,1,meetingLast-1,1).getValues().flat();
        const idx=ids.findIndex(v=>String(v)===tempMeetingId);
        if (idx>=0) sheets.meetings.getRange(idx+2,1,1,7).clearContent();
      }
      apiInvalidateMeetingCaches_(clubId,tempMeetingId);
    }

    // Clear, rather than structurally delete, the temporary planner row.
    if (planner.getMaxRows() >= plannerRowNo) {
      planner.getRange(plannerRowNo,1,1,28).clearContent();
    }
  }
}

/**
 * DEV-only hook used by Engine.js after the approved View roles button is
 * pressed. Returns true only for the short-lived controlled preview.
 */
function apiDevTrySendPendingRolePreview_(member,phone) {
  if (configValue_('SYSTEM_MODE','') !== 'DEV' ||
      String(configValue_('WA_SEND_MODE','TEST')).toUpperCase() !== 'TEST') return false;

  const props=PropertiesService.getScriptProperties();
  const raw=props.getProperty(API_DEV_ROLE_PREVIEW_PROPERTY_);
  if (!raw) return false;

  let p;
  try { p=JSON.parse(raw); } catch(ignore) {
    props.deleteProperty(API_DEV_ROLE_PREVIEW_PROPERTY_);
    return false;
  }
  if (Date.now()>Number(p.expiresAt||0)) {
    props.deleteProperty(API_DEV_ROLE_PREVIEW_PROPERTY_);
    return false;
  }
  if (String(p.member)!==String(member) ||
      normalizePhone_(p.phone)!==normalizePhone_(phone)) return false;

  const parts=String(p.meetingDate||'').split('-').map(Number);
  const meetingDate=new Date(parts[0],parts[1]-1,parts[2]);
  const sendResult=sendRoleList_(phone,member,{
    meetingNo:Number(p.meetingNo),
    meetingDate:meetingDate,
    available:p.available || [],
    reason:'AVAILABLE'
  });
  return {handled:true,sendResult:sendResult};
}

/**
 * DEV-only booking guard for role-list preview rows.
 */
function apiDevHandleRolePreviewSelection_(member,phone,interaction) {
  if (configValue_('SYSTEM_MODE','') !== 'DEV' ||
      String(configValue_('WA_SEND_MODE','TEST')).toUpperCase() !== 'TEST') return false;
  const raw=PropertiesService.getScriptProperties().getProperty(API_DEV_ROLE_PREVIEW_PROPERTY_);
  if (!raw) return false;
  let p; try { p=JSON.parse(raw); } catch(ignore) { return false; }
  if (Date.now()>Number(p.expiresAt||0) ||
      Number(interaction.meetingNo)!==Number(p.meetingNo) ||
      String(member)!==String(p.member) ||
      normalizePhone_(phone)!==normalizePhone_(p.phone)) return false;

  sendWhatsAppText_(phone,
    'DEV preview only ✅ No role was booked and no assignment workflow was triggered.');
  return true;
}


/**
 * DEV diagnostic for the controlled WhatsApp role preview.
 * Read-only: does not send anything or mutate assignments.
 */
function apiDevRolePreviewDiagnostic() {
  const props=PropertiesService.getScriptProperties();
  const raw=props.getProperty(API_DEV_ROLE_PREVIEW_PROPERTY_);
  let preview=null;
  try { preview=raw ? JSON.parse(raw) : null; } catch(ignore) {}

  const ss=apiSpreadsheet_();
  const webhook=ss.getSheetByName('WA Webhook Log');
  const automation=ss.getSheetByName('Automation Log');

  const member=configValue_('VPE_MEMBER_NAME','Riyas AF');
  const phone=normalizePhone_(phoneForMember_(member));
  const masked=phone ? ('******'+phone.slice(-4)) : '';

  const recentWebhook=[];
  if (webhook && webhook.getLastRow()>=2) {
    const start=Math.max(2,webhook.getLastRow()-19);
    const rows=webhook.getRange(start,1,webhook.getLastRow()-start+1,7).getValues();
    rows.forEach(r=>{
      const rowPhone=normalizePhone_(r[2]||'');
      if (!phone || rowPhone===phone) {
        recentWebhook.push({
          at:r[0] instanceof Date ? r[0].toISOString() : String(r[0]||''),
          event:String(r[1]||''),
          message_id:String(r[4]||''),
          message_type:String(r[5]||''),
          content:String(r[6]||'')
        });
      }
    });
  }

  const recentAutomation=[];
  if (automation && automation.getLastRow()>=2) {
    const start=Math.max(2,automation.getLastRow()-29);
    const rows=automation.getRange(start,1,automation.getLastRow()-start+1,7).getValues();
    rows.forEach(r=>{
      const module=String(r[1]||'');
      const action=String(r[4]||'');
      if (/WhatsApp|Role Selection/i.test(module) ||
          /role menu|inbound processing/i.test(action)) {
        recentAutomation.push({
          at:r[0] instanceof Date ? r[0].toISOString() : String(r[0]||''),
          module:module,
          meeting_no:String(r[2]||''),
          action:action,
          result:String(r[5]||''),
          notes:String(r[6]||'')
        });
      }
    });
  }

  const now=Date.now();
  const result={
    ok:true,
    version:API_V1_VERSION_,
    member:member,
    recipient_masked:masked,
    preview_present:Boolean(preview),
    preview_expired:preview ? now>Number(preview.expiresAt||0) : null,
    preview_meeting_no:preview ? Number(preview.meetingNo||0) : null,
    expected_roles:preview && preview.available ? preview.available.map(x=>x.code) : [],
    recent_webhook:recentWebhook,
    recent_automation:recentAutomation
  };
  console.log(JSON.stringify(result,null,2));
  return result;
}

/**
 * DEV-only direct-send test for the stored role preview.
 * Bypasses inbound webhook parsing but uses the exact same sendRoleList_ path.
 */
function apiDevSendStoredRolePreviewNow() {
  if (configValue_('SYSTEM_MODE','') !== 'DEV' ||
      String(configValue_('WA_SEND_MODE','TEST')).toUpperCase() !== 'TEST') {
    throw new Error('Direct role preview requires DEV / TEST mode.');
  }

  const member=configValue_('VPE_MEMBER_NAME','Riyas AF');
  const phone=normalizePhone_(phoneForMember_(member));
  if (!phone || !isWhatsAppRecipientAllowed_(phone)) {
    throw new Error('VPE test recipient is not allowlisted.');
  }

  const previewResult=apiDevTrySendPendingRolePreview_(member,phone);
  if (!previewResult || previewResult === false) {
    throw new Error('No active stored role preview is available. Run apiDevWhatsAppBridgePreviewTest again first.');
  }

  const sendResult=previewResult.sendResult || {};
  const result={
    ok:Boolean(sendResult.ok),
    version:API_V1_VERSION_,
    member:member,
    meta_status:sendResult.status || null,
    message_id:sendResult.messageId || '',
    blocked:Boolean(sendResult.blocked),
    error:sendResult.error || '',
    message:sendResult.ok
      ? 'Meta accepted the role-list message.'
      : 'Role-list send was not accepted.'
  };
  console.log(JSON.stringify(result,null,2));
  return result;
}


/**
 * Controlled DEV booking test:
 * API meeting -> legacy planner -> WhatsApp role list -> real reserveRole_ write.
 * Downstream role workflows are intercepted by apiDevHandleRoleBookingSelection_.
 */
function apiDevWhatsAppRoleBookingTest() {
  const clubId=API_DEFAULT_CLUB_ID_;
  if (configValue_('SYSTEM_MODE','') !== 'DEV' ||
      String(configValue_('WA_SEND_MODE','TEST')).toUpperCase() !== 'TEST') {
    throw new Error('Role booking test requires DEV / TEST mode.');
  }
  if (!configBool_('ROLE_SELECTION_ENABLED')) {
    throw new Error('ROLE_SELECTION_ENABLED must be TRUE.');
  }

  const member=configValue_('VPE_MEMBER_NAME','Riyas AF');
  const phone=normalizePhone_(phoneForMember_(member));
  if (!phone || !isWhatsAppRecipientAllowed_(phone)) {
    throw new Error('VPE test recipient is not allowlisted.');
  }

  const props=PropertiesService.getScriptProperties();
  if (String(props.getProperty(API_AUTO_SYNC_PLANNED_PROPERTY_)||'').toUpperCase()==='TRUE') {
    throw new Error('Keep WEB_API_AUTO_SYNC_PLANNED=FALSE for this controlled test.');
  }

  const ss=apiSpreadsheet_();
  const planner=ss.getSheetByName('VPE Weekly Planner');
  if (!planner) throw new Error('VPE Weekly Planner missing');

  // Clean stale DEV booking rows from previous interrupted runs.
  const preCount=Math.max(0,planner.getLastRow()-2);
  if (preCount) {
    const preRows=planner.getRange(3,1,preCount,28).getValues();
    preRows.forEach((row,i)=>{
      const d=row[1];
      if (String(row[2]||'')==='TEST_BOOKING' && d instanceof Date && d.getFullYear()===2099) {
        planner.getRange(i+3,1,1,28).clearContent();
      }
    });
  }

  const rows=Math.max(0,planner.getLastRow()-2)
    ? planner.getRange(3,1,Math.max(0,planner.getLastRow()-2),28).getValues() : [];
  const maxMeetingNo=rows.reduce((m,r)=>Math.max(m,Number(r[0])||0),0);
  const meetingNo=maxMeetingNo+3000;
  const meetingDate=new Date(2099,11,29);
  const dateText=Utilities.formatDate(meetingDate,configValue_('TIMEZONE','Asia/Kolkata'),'yyyy-MM-dd');

  const tempRow=new Array(28).fill('');
  tempRow[0]=meetingNo; tempRow[1]=meetingDate; tempRow[2]='TEST_BOOKING';
  const plannerRowNo=planner.getLastRow()+1;
  planner.getRange(plannerRowNo,1,1,28).setValues([tempRow]);

  const members=apiListMembers_(clubId).filter(m=>m.name!==member);
  if (members.length<5) throw new Error('Need five other active members for booking test.');

  const originalAllow=props.getProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_);
  let tempMeetingId='';

  try {
    const created=apiCreateMeeting_(clubId,{meeting_date:dateText});
    tempMeetingId=created.meeting_id;
    const items=created.items.map(x=>Object.assign({},x));
    const byCode={}; items.forEach(x=>{byCode[x.code]=x;});

    byCode.TMOD.member_id=members[0].member_id; byCode.TMOD.member_name=members[0].name;
    byCode.GE.member_id=members[1].member_id; byCode.GE.member_name=members[1].name;
    byCode.S1.member_id=members[2].member_id; byCode.S1.member_name=members[2].name;
    byCode.EV1.member_id=members[3].member_id; byCode.EV1.member_name=members[3].name;
    byCode.S2.member_id=members[4].member_id; byCode.S2.member_name=members[4].name;

    apiSaveMeeting_(clubId,{meeting_id:tempMeetingId,save_mode:'DRAFT',items:items});
    props.setProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_,'TRUE');
    const sync=apiSyncMeetingToLegacyPlanner_(clubId,tempMeetingId);

    const syncedRow=planner.getRange(plannerRowNo,1,1,28).getValues()[0];
    const available=availableRoleDefs_(syncedRow).map(def=>({
      code:def.code,name:def.name,speakerName:def.speakerName||''
    }));
    const codes=available.map(x=>x.code);
    if (codes.indexOf('Timer')<0) throw new Error('Timer is not available in controlled booking test.');

    const state={
      member:member,phone:phone,meetingNo:Number(sync.planner_meeting_no),
      plannerRowNo:plannerRowNo,tempMeetingId:tempMeetingId,
      available:available,expiresAt:Date.now()+15*60*1000
    };
    props.setProperty(API_DEV_ROLE_BOOKING_PROPERTY_,JSON.stringify(state));

    // Reuse the DEV preview hook for the View roles button.
    props.setProperty(API_DEV_ROLE_PREVIEW_PROPERTY_,JSON.stringify({
      member:member,phone:phone,meetingNo:state.meetingNo,
      meetingDate:dateText,available:available,expiresAt:state.expiresAt
    }));

    const result=sendApprovedRoleCall_(phone,state.meetingNo,dateText);
    if (!result || !result.ok) throw new Error('Role booking template failed: '+String(result&&result.error||'Unknown Meta error'));

    const report={
      ok:true,version:API_V1_VERSION_,message_id:result.messageId||'',
      temporary_meeting_no:state.meetingNo,
      available_roles:codes,
      instruction:'Tap View roles, then choose Timer.'
    };
    console.log(JSON.stringify(report,null,2));
    return report;
  } catch(err) {
    props.deleteProperty(API_DEV_ROLE_BOOKING_PROPERTY_);
    props.deleteProperty(API_DEV_ROLE_PREVIEW_PROPERTY_);
    if (planner.getMaxRows()>=plannerRowNo) planner.getRange(plannerRowNo,1,1,28).clearContent();
    throw err;
  } finally {
    if (originalAllow===null) props.deleteProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_);
    else props.setProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_,originalAllow);
  }
}

function apiDevCleanupRoleBookingState_(state) {
  if (!state) return;
  const props=PropertiesService.getScriptProperties();
  const ss=apiSpreadsheet_();
  const planner=ss.getSheetByName('VPE Weekly Planner');

  if (planner && Number(state.plannerRowNo)>0 &&
      planner.getMaxRows()>=Number(state.plannerRowNo)) {
    planner.getRange(Number(state.plannerRowNo),1,1,28).clearContent();
  }

  if (state.tempMeetingId) {
    const sheets=apiEnsureSheets_();
    const itemLast=sheets.items.getLastRow();
    if (itemLast>=2) {
      const all=sheets.items.getRange(2,1,itemLast-1,11).getValues();
      const kept=all.filter(row=>String(row[1])!==String(state.tempMeetingId));
      const area=sheets.items.getRange(2,1,itemLast-1,11);
      area.clearContent();
      if (kept.length) sheets.items.getRange(2,1,kept.length,11).setValues(kept);
    }
    const meetingLast=sheets.meetings.getLastRow();
    if (meetingLast>=2) {
      const ids=sheets.meetings.getRange(2,1,meetingLast-1,1).getValues().flat();
      const idx=ids.findIndex(v=>String(v)===String(state.tempMeetingId));
      if (idx>=0) sheets.meetings.getRange(idx+2,1,1,7).clearContent();
    }
    apiInvalidateMeetingCaches_(API_DEFAULT_CLUB_ID_,String(state.tempMeetingId));
  }

  props.deleteProperty(API_DEV_ROLE_BOOKING_PROPERTY_);
  props.deleteProperty(API_DEV_ROLE_PREVIEW_PROPERTY_);
}

/**
 * Intercepts only the controlled DEV booking meeting.
 * Uses the real reserveRole_ write, verifies it, then cleans up and stops all
 * downstream assignment workflows.
 */
function apiDevHandleRoleBookingSelection_(member,phone,interaction) {
  if (configValue_('SYSTEM_MODE','')!=='DEV' ||
      String(configValue_('WA_SEND_MODE','TEST')).toUpperCase()!=='TEST') return false;

  const props=PropertiesService.getScriptProperties();
  const raw=props.getProperty(API_DEV_ROLE_BOOKING_PROPERTY_);
  if (!raw) return false;

  let state;
  try { state=JSON.parse(raw); } catch(ignore) { return false; }

  if (Date.now()>Number(state.expiresAt||0)) {
    apiDevCleanupRoleBookingState_(state);
    return false;
  }
  if (String(member)!==String(state.member) ||
      normalizePhone_(phone)!==normalizePhone_(state.phone) ||
      Number(interaction.meetingNo)!==Number(state.meetingNo)) return false;

  const roleDef=ROLE_DEFS_.find(r=>r.code===interaction.roleCode);
  if (!roleDef) return false;

  const result=reserveRole_(interaction.meetingNo,roleDef,member);
  let verified=false;
  if (result && result.status==='BOOKED') {
    const planner=apiSpreadsheet_().getSheetByName('VPE Weekly Planner');
    const row=planner.getRange(Number(state.plannerRowNo),1,1,28).getValues()[0];
    verified=String(row[roleDef.column-1]||'').trim()===member;
  }

  const message=result && result.status==='BOOKED' && verified
    ? 'DEV booking test passed ✅ ' + roleDef.name + ' was written correctly to the synced planner row. No downstream workflow was triggered.'
    : 'DEV booking test failed. Result: ' + String(result && result.status || 'UNKNOWN');

  apiDevCleanupRoleBookingState_(state);
  sendWhatsAppText_(phone,message);
  return true;
}


/**
 * Full internal end-to-end integration self-test.
 *
 * Covers:
 * API meeting creation -> save -> legacy bridge -> role availability ->
 * WhatsApp interaction parsing -> real reserveRole_ booking -> planner write ->
 * updated availability -> cleanup.
 *
 * It does NOT send WhatsApp and does NOT use a live meeting.
 */
function apiDevFullEndToEndSelfTest() {
  const clubId=API_DEFAULT_CLUB_ID_;
  const props=PropertiesService.getScriptProperties();
  const ss=apiSpreadsheet_();
  const planner=ss.getSheetByName('VPE Weekly Planner');
  if (!planner) throw new Error('VPE Weekly Planner missing');

  if (String(props.getProperty(API_AUTO_SYNC_PLANNED_PROPERTY_)||'').toUpperCase()==='TRUE') {
    throw new Error('Set WEB_API_AUTO_SYNC_PLANNED=FALSE before running the self-test.');
  }

  const vpe=configValue_('VPE_MEMBER_NAME','Riyas AF');
  const members=apiListMembers_(clubId).filter(m=>m.name!==vpe);
  if (members.length<5) throw new Error('Need at least five active non-VPE members.');

  // Clear stale self-test rows from interrupted prior runs.
  const preCount=Math.max(0,planner.getLastRow()-2);
  if (preCount) {
    const pre=planner.getRange(3,1,preCount,28).getValues();
    pre.forEach((row,i)=>{
      const d=row[1];
      if (String(row[2]||'')==='SELF_TEST' && d instanceof Date && d.getFullYear()===2099) {
        planner.getRange(i+3,1,1,28).clearContent();
      }
    });
  }

  const currentCount=Math.max(0,planner.getLastRow()-2);
  const currentRows=currentCount ? planner.getRange(3,1,currentCount,28).getValues() : [];
  const maxMeetingNo=currentRows.reduce((m,r)=>Math.max(m,Number(r[0])||0),0);
  const meetingNo=maxMeetingNo+4000;

  const meetingDate=new Date(2099,11,28);
  const dateText=Utilities.formatDate(
    meetingDate,configValue_('TIMEZONE','Asia/Kolkata'),'yyyy-MM-dd'
  );

  const row=new Array(28).fill('');
  row[0]=meetingNo;
  row[1]=meetingDate;
  row[2]='SELF_TEST';
  const plannerRowNo=planner.getLastRow()+1;
  planner.getRange(plannerRowNo,1,1,28).setValues([row]);

  const originalAllow=props.getProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_);
  let tempMeetingId='';
  const checks=[];

  function check(name,ok,detail) {
    checks.push({name:name,ok:Boolean(ok),detail:detail||''});
    if (!ok) {
      const e=new Error('Self-test failed: '+name+(detail ? ' — '+detail : ''));
      e.selfTestChecks=checks.slice();
      throw e;
    }
  }

  try {
    // 1. API meeting creation.
    const created=apiCreateMeeting_(clubId,{meeting_date:dateText});
    tempMeetingId=created.meeting_id;
    check('meeting_created',Boolean(tempMeetingId),tempMeetingId);

    // 2. Plan assignments. Leave TTM/Timer/Ah/GRM/EV2 vacant; S3 vacant.
    const items=created.items.map(x=>Object.assign({},x));
    const byCode={}; items.forEach(x=>{byCode[x.code]=x;});
    byCode.TMOD.member_id=members[0].member_id; byCode.TMOD.member_name=members[0].name;
    byCode.GE.member_id=members[1].member_id; byCode.GE.member_name=members[1].name;
    byCode.S1.member_id=members[2].member_id; byCode.S1.member_name=members[2].name;
    byCode.EV1.member_id=members[3].member_id; byCode.EV1.member_name=members[3].name;
    byCode.S2.member_id=members[4].member_id; byCode.S2.member_name=members[4].name;

    const saved=apiSaveMeeting_(clubId,{
      meeting_id:tempMeetingId,
      save_mode:'DRAFT',
      items:items
    });
    check('meeting_saved',saved.status==='DRAFT',saved.status);

    // 3. Real legacy bridge.
    props.setProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_,'TRUE');
    const sync=apiSyncMeetingToLegacyPlanner_(clubId,tempMeetingId);
    check('bridge_synced',Boolean(sync&&sync.synced),JSON.stringify(sync));
    check('bridge_meeting_number',Number(sync.planner_meeting_no)===meetingNo,
      String(sync.planner_meeting_no));

    const synced=planner.getRange(plannerRowNo,1,1,28).getValues()[0];
    check('tmod_written',String(synced[18]||'')===members[0].name);
    check('ge_written',String(synced[19]||'')===members[1].name);
    check('speaker1_written',String(synced[6]||'')===members[2].name);
    check('evaluator1_written',String(synced[25]||'')===members[3].name);
    check('speaker2_written',String(synced[10]||'')===members[4].name);

    // 4. Availability from the same production function used by WhatsApp.
    const availableBefore=availableRoleDefs_(synced);
    const beforeCodes=availableBefore.map(x=>x.code);
    const expected=['TTM','Timer','Ah','GRM','EV2'];
    check('vacancies_exact',
      JSON.stringify(beforeCodes)===JSON.stringify(expected),
      'got '+JSON.stringify(beforeCodes));

    // 5. Parse the real template quick-reply shape.
    const menuInteraction=parseInteraction_({
      type:'button',
      button:{payload:'ROLE_MENU',text:'View roles'}
    });
    check('view_roles_parsed',
      menuInteraction&&menuInteraction.kind==='ROLE_MENU',
      JSON.stringify(menuInteraction));

    // Also verify visible-text fallback.
    const menuFallback=parseInteraction_({
      type:'button',
      button:{payload:'',text:'View roles'}
    });
    check('view_roles_text_fallback',
      menuFallback&&menuFallback.kind==='ROLE_MENU',
      JSON.stringify(menuFallback));

    // 6. Parse the same list-reply shape used when Timer is tapped.
    const timerInteraction=parseInteraction_({
      type:'interactive',
      interactive:{
        list_reply:{
          id:'ROLE|'+meetingNo+'|Timer',
          title:'Timer'
        }
      }
    });
    check('timer_selection_parsed',
      timerInteraction&&timerInteraction.kind==='ROLE_SELECTION' &&
      timerInteraction.roleCode==='Timer' &&
      Number(timerInteraction.meetingNo)===meetingNo,
      JSON.stringify(timerInteraction));

    // 7. Real booking function.
    const timerDef=ROLE_DEFS_.find(x=>x.code==='Timer');
    const booking=reserveRole_(meetingNo,timerDef,vpe);
    check('timer_booking_result',
      booking&&booking.status==='BOOKED',
      JSON.stringify(booking));

    const bookedRow=planner.getRange(plannerRowNo,1,1,28).getValues()[0];
    check('timer_written_to_planner',
      String(bookedRow[timerDef.column-1]||'').trim()===vpe,
      String(bookedRow[timerDef.column-1]||''));

    // 8. Availability immediately after booking.
    const afterCodes=availableRoleDefs_(bookedRow).map(x=>x.code);
    check('timer_removed_after_booking',
      afterCodes.indexOf('Timer')<0,
      JSON.stringify(afterCodes));
    check('remaining_vacancies_correct',
      JSON.stringify(afterCodes)===JSON.stringify(['TTM','Ah','GRM','EV2']),
      JSON.stringify(afterCodes));

    // 9. Same member cannot take a second role in the same meeting.
    const ahDef=ROLE_DEFS_.find(x=>x.code==='Ah');
    const secondBooking=reserveRole_(meetingNo,ahDef,vpe);
    check('duplicate_member_role_blocked',
      secondBooking&&secondBooking.status==='ALREADY_ASSIGNED',
      JSON.stringify(secondBooking));

    const report={
      ok:true,
      version:API_V1_VERSION_,
      temporary_meeting_id:tempMeetingId,
      temporary_meeting_no:meetingNo,
      checks:checks,
      available_before:beforeCodes,
      available_after_timer:afterCodes,
      booking_status:booking.status
    };
    if (!DEV_SUPPRESS_TEST_DETAIL_LOGS_) {
      console.log(JSON.stringify(report,null,2));
    }
    return report;

  } catch(err) {
    console.error(JSON.stringify({
      ok:false,
      version:API_V1_VERSION_,
      error:String(err&&err.message||err),
      checks:(err&&err.selfTestChecks)||checks
    },null,2));
    throw err;
  } finally {
    if (originalAllow===null) props.deleteProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_);
    else props.setProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_,originalAllow);

    // Clean planner test row.
    if (planner.getMaxRows()>=plannerRowNo) {
      planner.getRange(plannerRowNo,1,1,28).clearContent();
    }

    // Clean API test records in bulk.
    if (tempMeetingId) {
      const sheets=apiEnsureSheets_();

      const itemLast=sheets.items.getLastRow();
      if (itemLast>=2) {
        const all=sheets.items.getRange(2,1,itemLast-1,11).getValues();
        const kept=all.filter(r=>String(r[1])!==String(tempMeetingId));
        const area=sheets.items.getRange(2,1,itemLast-1,11);
        area.clearContent();
        if (kept.length) sheets.items.getRange(2,1,kept.length,11).setValues(kept);
      }

      const meetingLast=sheets.meetings.getLastRow();
      if (meetingLast>=2) {
        const rows=sheets.meetings.getRange(2,1,meetingLast-1,7).getValues();
        const kept=rows.filter(r=>String(r[0])!==String(tempMeetingId));
        const area=sheets.meetings.getRange(2,1,meetingLast-1,7);
        area.clearContent();
        if (kept.length) sheets.meetings.getRange(2,1,kept.length,7).setValues(kept);
      }

      apiInvalidateMeetingCaches_(clubId,tempMeetingId);
    }
  }
}


/**
 * Full dry-run post-booking workflow test.
 *
 * Uses the real handleRoleSelection_ path, including reserveRole_, task sync,
 * immediate assignment prompt branching and theme handoff. All WhatsApp payloads
 * are captured in-memory by DEV_OUTBOUND_CAPTURE_, so nothing is sent to Meta.
 *
 * A temporary API meeting and planner row are created on an unused date inside
 * the current task-engine scan window. Test records are cleaned afterwards.
 */
function apiDevPostBookingWorkflowSelfTest() {
  const clubId=API_DEFAULT_CLUB_ID_;
  const props=PropertiesService.getScriptProperties();
  const ss=apiSpreadsheet_();
  const planner=ss.getSheetByName('VPE Weekly Planner');
  const taskSheet=ss.getSheetByName(TASK_SHEET_);
  const outbox=ss.getSheetByName('WA Outbox');
  if (!planner || !taskSheet || !outbox) {
    throw new Error('Planner, Automation Tasks and WA Outbox are required.');
  }

  if (!configBool_('ROLE_SELECTION_ENABLED')) {
    throw new Error('ROLE_SELECTION_ENABLED must be TRUE.');
  }
  if (!configBool_('TASK_ENGINE_ENABLED')) {
    throw new Error('TASK_ENGINE_ENABLED must be TRUE.');
  }
  if (String(props.getProperty(API_AUTO_SYNC_PLANNED_PROPERTY_)||'').toUpperCase()==='TRUE') {
    throw new Error('Keep WEB_API_AUTO_SYNC_PLANNED=FALSE for this controlled self-test.');
  }

  const vpe=configValue_('VPE_MEMBER_NAME','Riyas AF');
  const phone=normalizePhone_(phoneForMember_(vpe));
  if (!phone) throw new Error('VPE WhatsApp number not found.');
  if (!isWhatsAppRecipientAllowed_(phone)) {
    throw new Error('VPE must be an allowlisted DEV recipient for post-booking workflow validation.');
  }

  const members=apiListMembers_(clubId).filter(m=>m.name!==vpe);
  if (members.length<5) throw new Error('Need five active non-VPE members.');

  const originalAllow=props.getProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_);
  const originalPreview=props.getProperty(API_DEV_ROLE_PREVIEW_PROPERTY_);
  const originalBooking=props.getProperty(API_DEV_ROLE_BOOKING_PROPERTY_);
  const convKey=conversationKey_(phone);
  const originalConversation=props.getProperty(convKey);

  let tempMeetingId='';
  let plannerRowNo=0;
  let meetingNo=0;
  let testDateText='';
  const checks=[];
  const captures=[];
  const introEnabled=configBool_('INTRO_WORKFLOW_ENABLED');

  function check(name,ok,detail) {
    checks.push({name:name,ok:Boolean(ok),detail:detail||''});
    if (!ok) {
      const err=new Error('Post-booking self-test failed: '+name+
        (detail ? ' — '+detail : ''));
      err.selfTestChecks=checks.slice();
      throw err;
    }
  }

  function removeMeetingRows_(sheet,meetingNoValue,width,meetingColumnIndex) {
    if (!sheet || sheet.getLastRow()<2) return;
    const rowCount=sheet.getLastRow()-1;
    const values=sheet.getRange(2,1,rowCount,width).getValues();
    const idx=Number(meetingColumnIndex)||0;
    const kept=values.filter(row=>Number(row[idx])!==Number(meetingNoValue));
    const area=sheet.getRange(2,1,rowCount,width);
    area.clearContent();
    if (kept.length) sheet.getRange(2,1,kept.length,width).setValues(kept);
  }

  try {
    // Ensure old controlled preview state cannot intercept this handler call.
    props.deleteProperty(API_DEV_ROLE_PREVIEW_PROPERTY_);
    props.deleteProperty(API_DEV_ROLE_BOOKING_PROPERTY_);
    props.deleteProperty(convKey);

    const today=startOfDay_(new Date());
    const scanDays=Math.max(
      7,
      Number(configValue_('REMINDER_WINDOW_DAYS','30')) || 30
    );

    const existingCount=Math.max(0,planner.getLastRow()-2);
    const existingRows=existingCount
      ? planner.getRange(3,1,existingCount,28).getValues()
      : [];
    const usedDates=new Set(
      existingRows
        .filter(r=>r[1] instanceof Date)
        .map(r=>Utilities.formatDate(
          r[1],configValue_('TIMEZONE','Asia/Kolkata'),'yyyy-MM-dd'
        ))
    );

    let testDate=null;
    for (let offset=Math.min(scanDays-1,20);offset>=1;offset--) {
      const candidate=new Date(today.getTime()+offset*86400000);
      const text=Utilities.formatDate(
        candidate,configValue_('TIMEZONE','Asia/Kolkata'),'yyyy-MM-dd'
      );
      if (!usedDates.has(text)) {
        testDate=candidate;
        testDateText=text;
        break;
      }
    }
    check('unused_test_date_found',Boolean(testDate),testDateText);

    const maxMeetingNo=existingRows.reduce(
      (m,row)=>Math.max(m,Number(row[0])||0),0
    );
    meetingNo=maxMeetingNo+5000;

    const tempPlannerRow=new Array(28).fill('');
    tempPlannerRow[0]=meetingNo;
    tempPlannerRow[1]=testDate;
    tempPlannerRow[2]='SELF_TEST_POST';
    tempPlannerRow[3]='E2E SELF TEST THEME';
    plannerRowNo=planner.getLastRow()+1;
    planner.getRange(plannerRowNo,1,1,28).setValues([tempPlannerRow]);

    // Build a genuine API meeting and push it through the real bridge first.
    const created=apiCreateMeeting_(clubId,{meeting_date:testDateText});
    tempMeetingId=created.meeting_id;
    const items=created.items.map(x=>Object.assign({},x));
    const byCode={}; items.forEach(x=>{byCode[x.code]=x;});

    byCode.TMOD.member_id=members[0].member_id; byCode.TMOD.member_name=members[0].name;
    byCode.GE.member_id=members[1].member_id; byCode.GE.member_name=members[1].name;
    byCode.S1.member_id=members[2].member_id; byCode.S1.member_name=members[2].name;
    byCode.EV1.member_id=members[3].member_id; byCode.EV1.member_name=members[3].name;
    byCode.S2.member_id=members[4].member_id; byCode.S2.member_name=members[4].name;

    apiSaveMeeting_(clubId,{
      meeting_id:tempMeetingId,
      save_mode:'DRAFT',
      items:items
    });

    props.setProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_,'TRUE');
    const sync=apiSyncMeetingToLegacyPlanner_(clubId,tempMeetingId);
    check('bridge_synced',Boolean(sync&&sync.synced),JSON.stringify(sync));
    check('bridge_target_correct',
      Number(sync.planner_meeting_no)===meetingNo,
      String(sync.planner_meeting_no));

    // The bridge intentionally doesn't own the theme column. Keep our test theme.
    planner.getRange(plannerRowNo,4).setValue('E2E SELF TEST THEME');

    const before=planner.getRange(plannerRowNo,1,1,28).getValues()[0];
    check('timer_initially_vacant',
      !String(before[21]||'').trim(),
      String(before[21]||''));

    // Capture every outbound payload from the real handler in-memory.
    DEV_OUTBOUND_CAPTURE_=captures;

    handleRoleSelection_(vpe,phone,{
      kind:'ROLE_SELECTION',
      meetingNo:meetingNo,
      roleCode:'Timer',
      text:'Timer'
    });

    const after=planner.getRange(plannerRowNo,1,1,28).getValues()[0];
    check('timer_written',
      String(after[21]||'').trim()===vpe,
      String(after[21]||''));

    const payloadText=captures.map(x=>JSON.stringify(x));
    check('confirmation_generated',
      payloadText.some(x=>x.indexOf('Confirmed')>=0 && x.indexOf('Timer')>=0),
      'captured='+captures.length);

    check('theme_message_generated',
      payloadText.some(x=>x.indexOf('E2E SELF TEST THEME')>=0),
      'captured='+captures.length);

    // Task engine should mark the Timer vacancy completed for Riyas.
    const taskRows=taskSheet.getLastRow()>=2
      ? taskSheet.getRange(2,1,taskSheet.getLastRow()-1,18).getValues()
      : [];
    const timerVacancy=taskRows.find(r=>
      Number(r[1])===meetingNo &&
      String(r[2])==='VACANT_ROLE' &&
      String(r[4])==='Timer'
    );
    check('timer_task_exists',Boolean(timerVacancy));
    check('timer_task_completed',
      Boolean(timerVacancy) &&
      String(timerVacancy[3]||'')===vpe &&
      String(timerVacancy[5]||'')==='Completed',
      timerVacancy ? JSON.stringify({
        member:timerVacancy[3],state:timerVacancy[5]
      }) : 'missing');

    const introTask=taskRows.find(r=>
      Number(r[1])===meetingNo &&
      String(r[2])==='INTRO' &&
      String(r[3])===vpe &&
      String(r[4])==='Timer'
    );

    if (introEnabled) {
      check('intro_task_created',Boolean(introTask));
      check('intro_task_pending',
        Boolean(introTask) && String(introTask[5])==='Pending',
        introTask ? String(introTask[5]) : 'missing');
      check('intro_prompt_generated',
        payloadText.some(x=>x.toLowerCase().indexOf('introduction')>=0),
        'captured='+captures.length);
    } else {
      check('intro_task_not_created_when_disabled',!introTask);
      check('intro_prompt_not_generated_when_disabled',
        !payloadText.some(x=>x.toLowerCase().indexOf('introduction')>=0),
        'captured='+captures.length);
    }

    // No evaluator workflow should run for Timer.
    check('no_evaluator_specific_message',
      !payloadText.some(x=>
        x.indexOf('Evaluator 2')>=0 ||
        x.indexOf('EVAL_COORD')>=0 ||
        x.indexOf('evaluator coordination')>=0
      ),
      'captured='+captures.length);

    const expectedMinimum=introEnabled ? 3 : 2;
    check('expected_outbound_count',
      captures.length>=expectedMinimum,
      'captured='+captures.length+', expected_at_least='+expectedMinimum);

    const report={
      ok:true,
      version:API_V1_VERSION_,
      temporary_meeting_id:tempMeetingId,
      temporary_meeting_no:meetingNo,
      temporary_meeting_date:testDateText,
      intro_workflow_enabled:introEnabled,
      captured_outbound_count:captures.length,
      captured_types:captures.map(p=>String(p.type||'')),
      checks:checks
    };
    console.log(JSON.stringify(report,null,2));
    return report;

  } catch(err) {
    console.error(JSON.stringify({
      ok:false,
      version:API_V1_VERSION_,
      error:String(err&&err.message||err),
      intro_workflow_enabled:introEnabled,
      captured_outbound_count:captures.length,
      checks:(err&&err.selfTestChecks)||checks
    },null,2));
    throw err;

  } finally {
    DEV_OUTBOUND_CAPTURE_=null;

    // Restore the VPE's prior guided-conversation state exactly.
    if (originalConversation===null) props.deleteProperty(convKey);
    else props.setProperty(convKey,originalConversation);

    if (originalPreview===null) props.deleteProperty(API_DEV_ROLE_PREVIEW_PROPERTY_);
    else props.setProperty(API_DEV_ROLE_PREVIEW_PROPERTY_,originalPreview);

    if (originalBooking===null) props.deleteProperty(API_DEV_ROLE_BOOKING_PROPERTY_);
    else props.setProperty(API_DEV_ROLE_BOOKING_PROPERTY_,originalBooking);

    if (originalAllow===null) props.deleteProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_);
    else props.setProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_,originalAllow);

    if (plannerRowNo && planner.getMaxRows()>=plannerRowNo) {
      planner.getRange(plannerRowNo,1,1,28).clearContent();
    }

    // Remove only this test meeting's generated task/outbox rows.
    if (meetingNo) {
      removeMeetingRows_(taskSheet,meetingNo,18,1);
      removeMeetingRows_(outbox,meetingNo,Math.max(16,outbox.getLastColumn()),10);
    }

    // Remove temporary API records.
    if (tempMeetingId) {
      const sheets=apiEnsureSheets_();

      const itemLast=sheets.items.getLastRow();
      if (itemLast>=2) {
        const all=sheets.items.getRange(2,1,itemLast-1,11).getValues();
        const kept=all.filter(r=>String(r[1])!==String(tempMeetingId));
        const area=sheets.items.getRange(2,1,itemLast-1,11);
        area.clearContent();
        if (kept.length) sheets.items.getRange(2,1,kept.length,11).setValues(kept);
      }

      const meetingLast=sheets.meetings.getLastRow();
      if (meetingLast>=2) {
        const rows=sheets.meetings.getRange(2,1,meetingLast-1,7).getValues();
        const kept=rows.filter(r=>String(r[0])!==String(tempMeetingId));
        const area=sheets.meetings.getRange(2,1,meetingLast-1,7);
        area.clearContent();
        if (kept.length) sheets.meetings.getRange(2,1,kept.length,7).setValues(kept);
      }

      apiInvalidateMeetingCaches_(clubId,tempMeetingId);
    }
  }
}


/**
 * Regression matrix for the four materially different role-booking branches.
 * Uses the real handleRoleSelection_ handler and captures outbound WhatsApp
 * payloads in memory. No Meta message is sent.
 *
 * Branches:
 * - Timer: normal role-player flow
 * - TMOD: theme-request flow
 * - GRM: split Word/Idiom grammarian flow
 * - EV2: evaluator pairing/coordination flow
 */
function apiDevRoleWorkflowMatrixSelfTest(forceAllWorkflows) {
  const ss=apiSpreadsheet_();
  const planner=ss.getSheetByName('VPE Weekly Planner');
  const taskSheet=ss.getSheetByName(TASK_SHEET_);
  const outbox=ss.getSheetByName('WA Outbox');
  const submissions=ss.getSheetByName('WA Submissions');
  const automationLog=ss.getSheetByName('Automation Log');
  if (!planner || !taskSheet || !outbox || !submissions) {
    throw new Error(
      'Planner, Automation Tasks, WA Outbox and WA Submissions are required.'
    );
  }
  if (!configBool_('ROLE_SELECTION_ENABLED')) {
    throw new Error('ROLE_SELECTION_ENABLED must be TRUE.');
  }
  if (!configBool_('TASK_ENGINE_ENABLED')) {
    throw new Error('TASK_ENGINE_ENABLED must be TRUE.');
  }

  const vpe=configValue_('VPE_MEMBER_NAME','Riyas AF');
  const phone=normalizePhone_(phoneForMember_(vpe));
  if (!phone || !isWhatsAppRecipientAllowed_(phone)) {
    throw new Error('VPE must be an allowlisted DEV test recipient.');
  }

  const members=apiListMembers_(API_DEFAULT_CLUB_ID_).filter(m=>m.name!==vpe);
  if (members.length<5) throw new Error('Need at least five active non-VPE members.');

  const props=PropertiesService.getScriptProperties();
  const convKey=conversationKey_(phone);
  const originalConversation=props.getProperty(convKey);
  const originalPreview=props.getProperty(API_DEV_ROLE_PREVIEW_PROPERTY_);
  const originalBooking=props.getProperty(API_DEV_ROLE_BOOKING_PROPERTY_);
  const originalForceWindow=(typeof DEV_FORCE_SERVICE_WINDOW_ === 'boolean')
    ? DEV_FORCE_SERVICE_WINDOW_ : null;

  const workflowKeys=[
    'INTRO_WORKFLOW_ENABLED',
    'TMOD_WORKFLOW_ENABLED',
    'GRAMMARIAN_WORKFLOW_ENABLED',
    'EVALUATION_WORKFLOW_ENABLED'
  ];

  // Automation switches come from the Automation Config sheet via CONFIG_CACHE_.
  // For the forced-enabled branch, override only this invocation's in-memory
  // cache. The sheet and Script Properties are never changed.
  const originalConfigCache = CONFIG_CACHE_
    ? Object.assign({}, CONFIG_CACHE_)
    : null;

  if (forceAllWorkflows===true) {
    const forcedConfig=Object.assign({}, getConfigMap_());
    workflowKeys.forEach(key=>{ forcedConfig[key]='TRUE'; });
    CONFIG_CACHE_=forcedConfig;
  }

  let plannerRowNo=0;
  const allMeetingNos=[];
  const checks=[];
  const branches={};

  const flags={
    intro_workflow_enabled:configBool_('INTRO_WORKFLOW_ENABLED'),
    tmod_workflow_enabled:configBool_('TMOD_WORKFLOW_ENABLED'),
    grammarian_workflow_enabled:configBool_('GRAMMARIAN_WORKFLOW_ENABLED'),
    evaluation_workflow_enabled:configBool_('EVALUATION_WORKFLOW_ENABLED')
  };

  if (forceAllWorkflows===true &&
      !(flags.intro_workflow_enabled &&
        flags.tmod_workflow_enabled &&
        flags.grammarian_workflow_enabled &&
        flags.evaluation_workflow_enabled)) {
    throw new Error('Forced workflow override did not take effect.');
  }

  function check(name,ok,detail) {
    checks.push({name:name,ok:Boolean(ok),detail:detail||''});
    if (!ok) {
      const err=new Error('Role workflow matrix failed: '+name+
        (detail ? ' — '+detail : ''));
      err.selfTestChecks=checks.slice();
      throw err;
    }
  }

  function payloadText_(captures) {
    return captures.map(p=>JSON.stringify(p)).join('\n');
  }

  function taskRowsFor_(meetingNo) {
    if (taskSheet.getLastRow()<2) return [];
    return taskSheet.getRange(2,1,taskSheet.getLastRow()-1,18).getValues()
      .filter(r=>Number(r[1])===Number(meetingNo));
  }

  function task_(rows,type,related,member) {
    return rows.find(r=>
      String(r[2])===String(type) &&
      (related===undefined || String(r[4])===String(related)) &&
      (member===undefined || String(r[3])===String(member))
    );
  }

  function removeByMeetingNos_(sheet,width,meetingColumnIndex) {
    if (!sheet || sheet.getLastRow()<2 || !allMeetingNos.length) return;
    const set=new Set(allMeetingNos.map(Number));
    const count=sheet.getLastRow()-1;
    const values=sheet.getRange(2,1,count,width).getValues();
    const kept=values.filter(r=>!set.has(Number(r[meetingColumnIndex])));
    const area=sheet.getRange(2,1,count,width);
    area.clearContent();
    if (kept.length) sheet.getRange(2,1,kept.length,width).setValues(kept);
  }

  function resetConversation_() {
    props.deleteProperty(convKey);
  }

  function basePlannerRow_(meetingNo,date,theme) {
    const row=new Array(28).fill('');
    row[0]=meetingNo;
    row[1]=date;
    row[2]='SELF_TEST_MATRIX';
    row[3]=theme||'';

    // Speakers.
    row[6]=members[2].name;
    row[8]='Planned';
    row[10]=members[4].name;
    row[12]='Planned';

    // Standard roles. Tested role is blanked per branch below.
    row[18]=members[0].name; // TMOD
    row[19]=members[1].name; // GE
    row[20]=members[3].name; // TTM
    row[21]=members[0].name; // Timer
    row[22]=members[1].name; // Ah
    row[23]=members[3].name; // GRM
    row[25]=members[1].name; // EV1
    row[26]=members[0].name; // EV2
    return row;
  }

  function runBranch_(code,date,meetingNo) {
    resetConversation_();
    props.deleteProperty(API_DEV_ROLE_PREVIEW_PROPERTY_);
    props.deleteProperty(API_DEV_ROLE_BOOKING_PROPERTY_);

    const theme=(code==='TMOD' || code==='GRM') ? '' : 'E2E REGRESSION THEME';
    const row=basePlannerRow_(meetingNo,date,theme);
    const def=ROLE_DEFS_.find(x=>x.code===code);
    if (!def) throw new Error('Unknown role branch '+code);
    row[def.column-1]='';

    // EV2 requires Speaker 2, which is populated in the base row.
    planner.getRange(plannerRowNo,1,1,28).setValues([row]);

    // Prove that content submitted before a TMOD exists is replayed as soon as
    // the TMOD books the role. These rows are cleaned with the test meeting.
    if (code === 'TMOD') {
      appendSubmission_({
        meetingNo:meetingNo,
        member:members[3].name,
        type:'GRAMMAR_WORD',
        related:'GRM',
        values:['Resilient','Able to recover quickly','She remained resilient.',''],
        sourceMessageId:''
      });
      appendSubmission_({
        meetingNo:meetingNo,
        member:members[3].name,
        type:'GRAMMAR_IDIOM',
        related:'GRM',
        values:['Once in a blue moon','Very rarely','We meet once in a blue moon.',''],
        sourceMessageId:''
      });
      appendSubmission_({
        meetingNo:meetingNo,
        member:members[3].name,
        type:'GRAMMAR_CONFIRMED',
        related:'GRM',
        values:['Confirmed','','',''],
        sourceMessageId:''
      });
      appendSubmission_({
        meetingNo:meetingNo,
        member:members[2].name,
        type:'SPEAKER_DETAILS',
        related:'S1',
        values:[
          'Late TMOD Speech',
          'Presentation Mastery',
          '2',
          'Understanding Your Communication Style',
          '07:30'
        ],
        sourceMessageId:''
      });
    }

    const captures=[];
    DEV_OUTBOUND_CAPTURE_=captures;
    DEV_FORCE_SERVICE_WINDOW_=true;

    handleRoleSelection_(vpe,phone,{
      kind:'ROLE_SELECTION',
      meetingNo:meetingNo,
      roleCode:code,
      text:def.name
    });

    DEV_OUTBOUND_CAPTURE_=null;

    const after=planner.getRange(plannerRowNo,1,1,28).getValues()[0];
    const tasks=taskRowsFor_(meetingNo);
    const text=payloadText_(captures);
    const vacancy=task_(tasks,'VACANT_ROLE',code);

    check(code+'_planner_write',
      String(after[def.column-1]||'').trim()===vpe,
      String(after[def.column-1]||''));

    if (!def.evaluatorSlot) {
      check(code+'_vacancy_task_exists',Boolean(vacancy));
      check(code+'_vacancy_task_completed',
        Boolean(vacancy) &&
        String(vacancy[3]||'')===vpe &&
        String(vacancy[5]||'')==='Completed',
        vacancy ? JSON.stringify({member:vacancy[3],state:vacancy[5]}) : 'missing');
    }

    if (code!=='GRM') {
      check(code+'_confirmation_generated',
        text.indexOf('Confirmed')>=0 && text.indexOf(def.name)>=0,
        'captured='+captures.length);
    } else {
      check('GRM_no_redundant_generic_confirmation',
        text.indexOf('Confirmed')<0,
        'captured='+captures.length);
    }

    if (code==='Timer') {
      const intro=task_(tasks,'INTRO','Timer',vpe);
      if (flags.intro_workflow_enabled) {
        check('Timer_intro_task_pending',
          Boolean(intro) && String(intro[5])==='Pending',
          intro ? String(intro[5]) : 'missing');
        check('Timer_intro_prompt_generated',
          text.toLowerCase().indexOf('introduction')>=0,
          'captured='+captures.length);
      } else {
        check('Timer_intro_disabled_no_task',!intro);
      }
      check('Timer_theme_message_generated',
        text.indexOf('E2E REGRESSION THEME')>=0,
        'captured='+captures.length);
    }

    if (code==='TMOD') {
      const themeTask=task_(tasks,'THEME','TMOD',vpe);
      check('TMOD_theme_task_exists',Boolean(themeTask));
      check('TMOD_theme_task_pending',
        Boolean(themeTask) && String(themeTask[5])==='Pending',
        themeTask ? String(themeTask[5]) : 'missing');
      if (flags.tmod_workflow_enabled) {
        check('TMOD_theme_prompt_generated',
          text.toLowerCase().indexOf('theme for atk meeting')>=0,
          'captured='+captures.length);
      } else {
        check('TMOD_theme_prompt_suppressed_when_disabled',
          text.toLowerCase().indexOf('theme for atk meeting')<0,
          'captured='+captures.length);
      }

      if (flags.grammarian_workflow_enabled) {
        check('TMOD_late_grammar_handoff_generated',
          text.indexOf('Word of the Day: Resilient')>=0 &&
          text.indexOf('Idiom of the Day: Once in a blue moon')>=0,
          'captured='+captures.length);
      }

      check('TMOD_late_speaker_handoff_generated',
        text.indexOf('Title: Late TMOD Speech')>=0 &&
        text.indexOf('Time: 07:30')>=0,
        'captured='+captures.length);

      // Replacement-TMOD regression:
      // 1) TMOD A (VPE) has already received the stored handoffs above.
      // 2) Swap the planner to a different active member.
      // 3) Allow that member's phone only in this invocation's in-memory config.
      // 4) Confirm TMOD B receives the same stored handoffs.
      // 5) Replay once more and prove idempotency suppresses duplicates.
      const replacementTmod=members.find(m=>
        m && m.name!==vpe && normalizePhone_(m.whatsapp)
      );
      check('TMOD_replacement_candidate_available',Boolean(replacementTmod));

      const replacementPhone=normalizePhone_(replacementTmod.whatsapp);
      const currentAllowlist=String(configValue_('WA_TEST_RECIPIENTS','') || '');
      const replacementConfig=Object.assign({},getConfigMap_());
      replacementConfig.WA_TEST_RECIPIENTS=[
        currentAllowlist,
        replacementPhone
      ].filter(Boolean).join(',');
      CONFIG_CACHE_=replacementConfig;

      planner.getRange(plannerRowNo,19).setValue(replacementTmod.name);

      const replacementCaptures=[];
      DEV_OUTBOUND_CAPTURE_=replacementCaptures;
      sendLateTmodCatchUp_(meetingNo);
      DEV_OUTBOUND_CAPTURE_=null;

      const replacementText=payloadText_(replacementCaptures);

      if (flags.grammarian_workflow_enabled) {
        check('TMOD_replacement_grammar_handoff_generated',
          replacementText.indexOf('Word of the Day: Resilient')>=0 &&
          replacementText.indexOf('Idiom of the Day: Once in a blue moon')>=0,
          'captured='+replacementCaptures.length);
      }

      check('TMOD_replacement_speaker_handoff_generated',
        replacementText.indexOf('Title: Late TMOD Speech')>=0 &&
        replacementText.indexOf('Time: 07:30')>=0,
        'captured='+replacementCaptures.length);

      const duplicateCaptures=[];
      DEV_OUTBOUND_CAPTURE_=duplicateCaptures;
      sendLateTmodCatchUp_(meetingNo);
      DEV_OUTBOUND_CAPTURE_=null;

      check('TMOD_replacement_duplicate_handoffs_suppressed',
        duplicateCaptures.length===0,
        'captured='+duplicateCaptures.length);
    }

    if (code==='GRM') {
      const grammar=task_(tasks,'GRAMMAR','GRM',vpe);
      check('GRM_task_exists',Boolean(grammar));
      check('GRM_task_pending',
        Boolean(grammar) && String(grammar[5])==='Pending',
        grammar ? String(grammar[5]) : 'missing');
      if (flags.grammarian_workflow_enabled) {
        check('GRM_word_prompt_generated',
          captures.length>=2 && text.toLowerCase().indexOf('word')>=0,
          'captured='+captures.length);
        const state=activeConversation_(phone);
        check('GRM_conversation_started',
          Boolean(state) &&
          state.kind==='GRAMMAR' &&
          Number(state.meetingNo)===meetingNo &&
          String(state.step||'').toUpperCase()==='WORD',
          state ? JSON.stringify(state) : 'missing');
      } else {
        check('GRM_prompt_suppressed_when_disabled',
          captures.length===0,
          'captured='+captures.length);
      }
    }

    if (code==='EV2') {
      const assignment=task_(tasks,'EVALUATOR_ASSIGNMENT','S2');
      const coordination=task_(tasks,'EVALUATOR_COORDINATION','S2');
      check('EV2_assignment_task_completed',
        Boolean(assignment) && String(assignment[5])==='Completed',
        assignment ? String(assignment[5]) : 'missing');
      check('EV2_coordination_task_exists',Boolean(coordination));

      if (flags.evaluation_workflow_enabled) {
        check('EV2_pairing_message_generated',
          text.toLowerCase().indexOf('you will be evaluating')>=0 &&
          text.indexOf(members[4].name)>=0,
          'captured='+captures.length);
      } else {
        check('EV2_pairing_suppressed_when_disabled',
          text.toLowerCase().indexOf('you will be evaluating')<0,
          'captured='+captures.length);
      }

      check('EV2_theme_message_generated',
        text.indexOf('E2E REGRESSION THEME')>=0,
        'captured='+captures.length);
    }

    branches[code]={
      meeting_no:meetingNo,
      captured_outbound_count:captures.length,
      captured_types:captures.map(p=>String(p.type||'')),
      task_count:tasks.length
    };

    resetConversation_();
  }

  try {
    const today=startOfDay_(new Date());
    const scanDays=Math.max(7,Number(configValue_('REMINDER_WINDOW_DAYS','30'))||30);
    const existingCount=Math.max(0,planner.getLastRow()-2);
    const existing=existingCount
      ? planner.getRange(3,1,existingCount,28).getValues()
      : [];
    const usedDates=new Set(existing
      .filter(r=>r[1] instanceof Date)
      .map(r=>Utilities.formatDate(
        r[1],configValue_('TIMEZONE','Asia/Kolkata'),'yyyy-MM-dd'
      )));

    let testDate=null;
    for (let offset=Math.min(scanDays-1,20);offset>=1;offset--) {
      const candidate=new Date(today.getTime()+offset*86400000);
      const text=Utilities.formatDate(
        candidate,configValue_('TIMEZONE','Asia/Kolkata'),'yyyy-MM-dd'
      );
      if (!usedDates.has(text)) { testDate=candidate; break; }
    }
    check('matrix_unused_date_found',Boolean(testDate));

    const maxMeetingNo=existing.reduce(
      (m,row)=>Math.max(m,Number(row[0])||0),0
    );

    plannerRowNo=planner.getLastRow()+1;
    planner.getRange(plannerRowNo,1,1,28).clearContent();

    const branchCodes=['Timer','TMOD','GRM','EV2'];
    branchCodes.forEach((code,index)=>{
      const n=maxMeetingNo+6001+index;
      allMeetingNos.push(n);
      runBranch_(code,testDate,n);
    });

    const report={
      ok:true,
      version:API_V1_VERSION_,
      mode:forceAllWorkflows===true ? 'FORCED_ENABLED' : 'CONFIGURED',
      flags:flags,
      branches:branches,
      checks:checks
    };
    if (!DEV_SUPPRESS_TEST_DETAIL_LOGS_) {
      console.log(JSON.stringify(report,null,2));
    }
    return report;

  } catch(err) {
    console.error(JSON.stringify({
      ok:false,
      version:API_V1_VERSION_,
      error:String(err&&err.message||err),
      flags:flags,
      branches:branches,
      checks:(err&&err.selfTestChecks)||checks
    },null,2));
    throw err;

  } finally {
    DEV_OUTBOUND_CAPTURE_=null;
    DEV_FORCE_SERVICE_WINDOW_=originalForceWindow;

    if (originalConversation===null) props.deleteProperty(convKey);
    else props.setProperty(convKey,originalConversation);

    if (originalPreview===null) props.deleteProperty(API_DEV_ROLE_PREVIEW_PROPERTY_);
    else props.setProperty(API_DEV_ROLE_PREVIEW_PROPERTY_,originalPreview);

    if (originalBooking===null) props.deleteProperty(API_DEV_ROLE_BOOKING_PROPERTY_);
    else props.setProperty(API_DEV_ROLE_BOOKING_PROPERTY_,originalBooking);

    CONFIG_CACHE_=originalConfigCache;

    if (plannerRowNo && planner.getMaxRows()>=plannerRowNo) {
      planner.getRange(plannerRowNo,1,1,28).clearContent();
    }

    removeByMeetingNos_(taskSheet,18,1);
    removeByMeetingNos_(outbox,Math.max(16,outbox.getLastColumn()),10);
    removeByMeetingNos_(
      submissions,
      Math.max(15,submissions.getLastColumn()),
      1
    );
    if (automationLog) {
      removeByMeetingNos_(automationLog,Math.max(7,automationLog.getLastColumn()),2);
    }
  }
}


/**
 * One-button internal regression suite.
 *
 * This is the function to run before asking for any manual WhatsApp testing.
 * It combines:
 * 1) API meeting -> bridge -> vacancy -> booking integrity
 * 2) the four distinct post-booking workflow branches
 */
function apiDevPlannerLifecycleSelfTest_() {
  const clubId = API_DEFAULT_CLUB_ID_;
  const ss = apiSpreadsheet_();
  const planner = ss.getSheetByName('VPE Weekly Planner');
  if (!planner) throw new Error('VPE Weekly Planner missing');

  const members = apiListMembers_(clubId);
  if (members.length < 5) {
    throw new Error('Need at least five active members for lifecycle self-test');
  }

  const checks = [];
  const check = (name, ok, detail) => {
    checks.push({name:name,ok:Boolean(ok),detail:detail || ''});
    if (!ok) {
      const e = new Error('Lifecycle self-test failed: ' + name);
      e.selfTestChecks = checks.slice();
      throw e;
    }
  };

  const existingPlanner = planner.getLastRow() >= 3
    ? planner.getRange(3,1,planner.getLastRow()-2,28).getValues()
    : [];
  const usedDates = new Set(existingPlanner
    .filter(r => r[1] instanceof Date)
    .map(r => Utilities.formatDate(
      r[1], configValue_('TIMEZONE','Asia/Kolkata'), 'yyyy-MM-dd'
    )));

  const apiSheets = apiEnsureSheets_();
  const existingApiDates = apiSheets.meetings.getLastRow() >= 2
    ? apiSheets.meetings.getRange(
        2,1,apiSheets.meetings.getLastRow()-1,7
      ).getValues()
      .filter(r => String(r[1] || '') === clubId)
      .map(r => r[2] instanceof Date
        ? Utilities.formatDate(
            r[2],configValue_('TIMEZONE','Asia/Kolkata'),'yyyy-MM-dd'
          )
        : String(r[2] || '').slice(0,10))
    : [];
  existingApiDates.forEach(d => usedDates.add(d));

  let testDate = null;
  for (let day = 20; day >= 1; day--) {
    const candidate = new Date(2098, 11, day);
    const key = Utilities.formatDate(
      candidate, configValue_('TIMEZONE','Asia/Kolkata'), 'yyyy-MM-dd'
    );
    if (!usedDates.has(key)) {
      testDate = candidate;
      break;
    }
  }
  check('unused_test_date_found', Boolean(testDate));

  const dateText = Utilities.formatDate(
    testDate, configValue_('TIMEZONE','Asia/Kolkata'), 'yyyy-MM-dd'
  );
  const maxMeetingNo = existingPlanner.reduce(
    (m,row) => Math.max(m,Number(row[0]) || 0), 0
  );
  const meetingNo = maxMeetingNo + 8100;
  const rowNo = planner.getLastRow() + 1;
  const row = new Array(28).fill('');
  row[0] = meetingNo;
  row[1] = testDate;
  row[2] = 'SELF_TEST_LIFECYCLE';

  // Planner-first assignments that must hydrate into a later API meeting.
  row[6] = members[0].name;   // S1
  row[8] = 'Planned';
  const timerDef = ROLE_DEFS_.find(x => x.code === 'Timer');
  const ahDef = ROLE_DEFS_.find(x => x.code === 'Ah');
  const grmDef = ROLE_DEFS_.find(x => x.code === 'GRM');
  if (!timerDef || !ahDef || !grmDef) {
    throw new Error('Required role definitions missing');
  }
  row[timerDef.column - 1] = members[1].name;
  planner.getRange(rowNo,1,1,28).setValues([row]);
  SpreadsheetApp.flush();

  let meetingId = '';
  try {
    const created = apiEnsureMeetingForPlanner_(clubId,meetingNo);
    meetingId = created.meeting_id;

    const byCode = {};
    created.items.forEach(item => { byCode[item.code] = item; });

    check(
      'planner_number_linked_on_create',
      Number(created.planner_meeting_no) === meetingNo,
      String(created.planner_meeting_no || '')
    );
    check(
      'planner_speaker_hydrated_on_create',
      byCode.S1 && byCode.S1.member_name === members[0].name,
      byCode.S1 ? byCode.S1.member_name : 'missing'
    );
    check(
      'planner_role_hydrated_on_create',
      byCode.Timer && byCode.Timer.member_name === members[1].name,
      byCode.Timer ? byCode.Timer.member_name : 'missing'
    );

    // API exists first; a later WhatsApp-style speaker booking must mirror in.
    const booked = reserveSpeechSlot_(meetingNo,members[2].name);
    check(
      'speech_booking_succeeds_with_api_present',
      booked && booked.status === 'BOOKED' && booked.slot === 2,
      booked ? JSON.stringify(booked) : 'missing'
    );
    const afterBooking = apiGetMeeting_(clubId,meetingId);
    const s2 = afterBooking.items.find(x => x.code === 'S2');
    check(
      'speech_booking_mirrors_immediately_to_api',
      s2 && s2.member_name === members[2].name,
      s2 ? s2.member_name : 'missing'
    );

    // Simulate a stale UI draft: planner received Ah while UI still has it blank.
    planner.getRange(rowNo,ahDef.column).setValue(members[3].name);
    SpreadsheetApp.flush();
    const stale = apiGetMeeting_(clubId,meetingId);
    const staleItems = stale.items.map(x => Object.assign({},x));
    const staleAh = staleItems.find(x => x.code === 'Ah');
    staleAh.member_id = '';
    staleAh.member_name = '';

    const saved = apiSaveMeeting_(clubId,{
      meeting_id:meetingId,
      save_mode:'DRAFT',
      items:staleItems
    });
    const savedAh = saved.items.find(x => x.code === 'Ah');
    check(
      'stale_ui_blank_imports_planner_assignment',
      savedAh && savedAh.member_name === members[3].name,
      savedAh ? savedAh.member_name : 'missing'
    );

    // A genuine non-blank disagreement must be blocked, never overwritten.
    planner.getRange(rowNo,grmDef.column).setValue(members[4].name);
    SpreadsheetApp.flush();

    const conflictBase = apiGetMeeting_(clubId,meetingId);
    const conflictItems = conflictBase.items.map(x => Object.assign({},x));
    const conflictGrm = conflictItems.find(x => x.code === 'GRM');
    conflictGrm.member_id = members[3].member_id;
    conflictGrm.member_name = members[3].name;

    let conflictBlocked = false;
    try {
      apiSaveMeeting_(clubId,{
        meeting_id:meetingId,
        save_mode:'DRAFT',
        items:conflictItems
      });
    } catch (error) {
      conflictBlocked = Boolean(
        error &&
        error.apiCode === 'SYNC_CONFLICT' &&
        error.details &&
        error.details.some(x => x.code === 'GRM')
      );
    }
    check('true_assignment_conflict_is_blocked', conflictBlocked);

    // Force the conflicting API value internally, then prove activation
    // rejects it and rolls status back instead of stranding READY_FOR_SYNC.
    const forced = apiGetMeeting_(clubId,meetingId);
    const forcedItems = forced.items.map(x => Object.assign({},x));
    const forcedGrm = forcedItems.find(x => x.code === 'GRM');
    forcedGrm.member_id = members[3].member_id;
    forcedGrm.member_name = members[3].name;
    apiReplaceMeetingItems_(
      apiEnsureSheets_().items,clubId,meetingId,forcedItems
    );
    apiInvalidateMeetingCaches_(clubId,meetingId);

    let activationConflictBlocked = false;
    try {
      apiActivateMeeting_(clubId,meetingId);
    } catch (error) {
      activationConflictBlocked = Boolean(
        error && error.apiCode === 'SYNC_CONFLICT'
      );
    }
    check(
      'activation_conflict_is_blocked',
      activationConflictBlocked
    );

    const afterFailedActivation = apiGetMeeting_(clubId,meetingId);
    check(
      'failed_activation_rolls_status_back',
      afterFailedActivation.status === 'DRAFT',
      afterFailedActivation.status
    );

    return {
      ok:checks.every(x => x.ok),
      checks:checks
    };
  } finally {
    // Remove temporary API items/meeting and planner row.
    if (meetingId) {
      const sheets = apiEnsureSheets_();
      for (let r=sheets.items.getLastRow(); r>=2; r--) {
        if (String(sheets.items.getRange(r,2).getValue()) === meetingId) {
          sheets.items.deleteRow(r);
        }
      }
      for (let r=sheets.meetings.getLastRow(); r>=2; r--) {
        if (String(sheets.meetings.getRange(r,1).getValue()) === meetingId) {
          sheets.meetings.deleteRow(r);
        }
      }
      apiInvalidateMeetingCaches_(clubId,meetingId);
    }

    if (planner.getMaxRows() >= rowNo) {
      const currentNo = Number(planner.getRange(rowNo,1).getValue() || 0);
      if (currentNo === meetingNo) planner.deleteRow(rowNo);
    }
    SpreadsheetApp.flush();
    apiInvalidateClubReadCaches_(clubId);
  }
}


function apiDevGrammarParserSelfTest_() {
  const checks = [];
  const check = (name, ok, detail) => {
    checks.push({name:name, ok:Boolean(ok), detail:detail || ''});
  };

  const colon = parseGrammarStageBlock_(
    'Word: Resilient\nMeaning: Staying in the process\nUsage: She remained resilient.',
    'WORD'
  );
  check('word_colon_format',
    colon.recognized === 3 &&
    colon.item === 'Resilient' &&
    !colon.oppositeStage,
    JSON.stringify(colon));

  const dash = parseGrammarStageBlock_(
    'Word - Resilient\nMeaning - Staying in the process\nUsage - She remained resilient.',
    'WORD'
  );
  check('word_dash_format',
    dash.recognized === 3 &&
    dash.item === 'Resilient',
    JSON.stringify(dash));

  const opposite = parseGrammarStageBlock_(
    'Word: Resilient\nMeaning: Staying in the process\nUsage: She remained resilient.',
    'IDIOM'
  );
  check('opposite_stage_protected',
    opposite.recognized === 0 && opposite.oppositeStage === 'WORD',
    JSON.stringify(opposite));

  const combined = parseGrammarCombinedBlock_(
    'Word: Resilient\n' +
    'Meaning: Staying in the process\n' +
    'Usage: She remained resilient.\n' +
    'Idiom: Once in a blue moon\n' +
    'Meaning: Very rarely\n' +
    'Usage: He visits once in a blue moon.'
  );
  check('combined_six_field_format',
    combined.recognized === 6 &&
    combined.word === 'Resilient' &&
    combined.idiom === 'Once in a blue moon' &&
    combined.wordUsage === 'She remained resilient.' &&
    combined.idiomUsage === 'He visits once in a blue moon.',
    JSON.stringify(combined));

  return {
    ok:checks.every(x => x.ok),
    checks:checks
  };
}


/**
 * One-button internal regression suite.
 */
function apiDevMasterRegressionSuite() {
  const started=Date.now();
  const previousSuppress=DEV_SUPPRESS_TEST_DETAIL_LOGS_;
  DEV_SUPPRESS_TEST_DETAIL_LOGS_=true;

  try {
    const grammarParser=apiDevGrammarParserSelfTest_();
    const plannerLifecycle=apiDevPlannerLifecycleSelfTest_();
    const core=apiDevFullEndToEndSelfTest();
    const configured=apiDevRoleWorkflowMatrixSelfTest(false);
    const enabled=apiDevRoleWorkflowMatrixSelfTest(true);

    function summarizeMatrix_(matrix) {
      const branchSummary={};
      Object.keys(matrix.branches||{}).forEach(code=>{
        const b=matrix.branches[code]||{};
        branchSummary[code]={
          ok:true,
          captured_outbound_count:Number(b.captured_outbound_count||0),
          task_count:Number(b.task_count||0)
        };
      });
      return {
        ok:Boolean(matrix&&matrix.ok),
        checks_passed:(matrix&&matrix.checks||[]).filter(x=>x.ok).length,
        checks_total:(matrix&&matrix.checks||[]).length,
        flags:matrix&&matrix.flags||{},
        branches:branchSummary
      };
    }

    const report={
      ok:Boolean(
        grammarParser&&grammarParser.ok&&
        plannerLifecycle&&plannerLifecycle.ok&&
        core&&core.ok&&
        configured&&configured.ok&&
        enabled&&enabled.ok
      ),
      version:API_V1_VERSION_,
      duration_ms:Date.now()-started,
      grammarian_parser:{
        ok:Boolean(grammarParser&&grammarParser.ok),
        checks_passed:(grammarParser&&grammarParser.checks||[]).filter(x=>x.ok).length,
        checks_total:(grammarParser&&grammarParser.checks||[]).length
      },
      planner_api_lifecycle:{
        ok:Boolean(plannerLifecycle&&plannerLifecycle.ok),
        checks_passed:(plannerLifecycle&&plannerLifecycle.checks||[]).filter(x=>x.ok).length,
        checks_total:(plannerLifecycle&&plannerLifecycle.checks||[]).length
      },
      core:{
        ok:Boolean(core&&core.ok),
        checks_passed:(core&&core.checks||[]).filter(x=>x.ok).length,
        checks_total:(core&&core.checks||[]).length
      },
      configured_workflows:summarizeMatrix_(configured),
      forced_enabled_workflows:summarizeMatrix_(enabled)
    };

    console.log(JSON.stringify(report,null,2));
    return report;
  } finally {
    DEV_SUPPRESS_TEST_DETAIL_LOGS_=previousSuppress;
  }
}


/**
 * Final real-device WhatsApp golden-path test.
 *
 * Sends the approved role-reminder template to the configured DEV recipient,
 * serves a role list from a temporary bridged API meeting, and permits exactly
 * one Timer selection to pass through the real production role-booking handler.
 *
 * Current configured downstream workflows are used. Nothing is force-enabled.
 */
function apiDevWhatsAppGoldenPathTest() {
  const clubId=API_DEFAULT_CLUB_ID_;
  if (configValue_('SYSTEM_MODE','')!=='DEV' ||
      String(configValue_('WA_SEND_MODE','TEST')).toUpperCase()!=='TEST') {
    throw new Error('Golden-path test requires SYSTEM_MODE=DEV and WA_SEND_MODE=TEST.');
  }
  if (!configBool_('ROLE_SELECTION_ENABLED')) {
    throw new Error('ROLE_SELECTION_ENABLED must be TRUE.');
  }
  if (!configBool_('TASK_ENGINE_ENABLED')) {
    throw new Error('TASK_ENGINE_ENABLED must be TRUE.');
  }
  if (!isApprovedTemplate_('atk_role_reminder_v3')) {
    throw new Error('atk_role_reminder_v3 must be APPROVED.');
  }

  const member=configValue_('VPE_MEMBER_NAME','Riyas AF');
  const phone=normalizePhone_(phoneForMember_(member));
  if (!phone || !isWhatsAppRecipientAllowed_(phone)) {
    throw new Error('VPE DEV test recipient is not allowlisted.');
  }

  const props=PropertiesService.getScriptProperties();
  if (String(props.getProperty(API_AUTO_SYNC_PLANNED_PROPERTY_)||'').toUpperCase()==='TRUE') {
    throw new Error('Keep WEB_API_AUTO_SYNC_PLANNED=FALSE for the golden-path test.');
  }

  // Do not let older controlled tests intercept the final golden path.
  props.deleteProperty(API_DEV_ROLE_BOOKING_PROPERTY_);
  props.deleteProperty(API_DEV_ROLE_PREVIEW_PROPERTY_);
  props.deleteProperty(API_DEV_GOLDEN_PROPERTY_);
  props.deleteProperty(API_DEV_GOLDEN_RESULT_PROPERTY_);

  const ss=apiSpreadsheet_();
  const planner=ss.getSheetByName('VPE Weekly Planner');
  if (!planner) throw new Error('VPE Weekly Planner missing.');

  const members=apiListMembers_(clubId).filter(m=>m.name!==member);
  if (members.length<5) throw new Error('Need five active non-VPE members.');

  const existingCount=Math.max(0,planner.getLastRow()-2);
  const existing=existingCount
    ? planner.getRange(3,1,existingCount,28).getValues()
    : [];

  // Use a unique future date inside both the role and task-engine windows.
  const today=startOfDay_(new Date());
  const roleWindow=Math.max(
    1,Number(configValue_('ROLE_SELECTION_WINDOW_DAYS','14'))||14
  );
  const taskWindow=Math.max(
    7,Number(configValue_('REMINDER_WINDOW_DAYS','30'))||30
  );
  const maxOffset=Math.max(1,Math.min(roleWindow,taskWindow)-1);
  const usedDates=new Set(existing
    .filter(r=>r[1] instanceof Date)
    .map(r=>Utilities.formatDate(
      r[1],configValue_('TIMEZONE','Asia/Kolkata'),'yyyy-MM-dd'
    )));

  let meetingDate=null;
  let dateText='';
  for (let offset=maxOffset;offset>=1;offset--) {
    const candidate=new Date(today.getTime()+offset*86400000);
    const key=Utilities.formatDate(
      candidate,configValue_('TIMEZONE','Asia/Kolkata'),'yyyy-MM-dd'
    );
    if (!usedDates.has(key)) {
      meetingDate=candidate;
      dateText=key;
      break;
    }
  }
  if (!meetingDate) throw new Error('No unused date is available inside the controlled test window.');

  const maxMeetingNo=existing.reduce(
    (m,row)=>Math.max(m,Number(row[0])||0),0
  );
  const meetingNo=maxMeetingNo+7000;

  const tempRow=new Array(28).fill('');
  tempRow[0]=meetingNo;
  tempRow[1]=meetingDate;
  tempRow[2]='SELF_TEST_GOLDEN';
  const plannerRowNo=planner.getLastRow()+1;
  planner.getRange(plannerRowNo,1,1,28).setValues([tempRow]);

  const originalAllow=props.getProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_);
  let tempMeetingId='';

  try {
    const created=apiCreateMeeting_(clubId,{meeting_date:dateText});
    tempMeetingId=created.meeting_id;
    const items=created.items.map(x=>Object.assign({},x));
    const byCode={}; items.forEach(x=>{byCode[x.code]=x;});

    // Leave TTM/Timer/Ah/GRM/EV2 open. S3 stays vacant, so EV3 is hidden.
    byCode.TMOD.member_id=members[0].member_id;
    byCode.TMOD.member_name=members[0].name;
    byCode.GE.member_id=members[1].member_id;
    byCode.GE.member_name=members[1].name;
    byCode.S1.member_id=members[2].member_id;
    byCode.S1.member_name=members[2].name;
    byCode.EV1.member_id=members[3].member_id;
    byCode.EV1.member_name=members[3].name;
    byCode.S2.member_id=members[4].member_id;
    byCode.S2.member_name=members[4].name;

    apiSaveMeeting_(clubId,{
      meeting_id:tempMeetingId,
      save_mode:'DRAFT',
      items:items
    });

    props.setProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_,'TRUE');
    const sync=apiSyncMeetingToLegacyPlanner_(clubId,tempMeetingId);

    const synced=planner.getRange(plannerRowNo,1,1,28).getValues()[0];
    if (Number(sync.planner_meeting_no)!==meetingNo ||
        String(synced[18]||'')!==members[0].name ||
        String(synced[19]||'')!==members[1].name ||
        String(synced[6]||'')!==members[2].name ||
        String(synced[25]||'')!==members[3].name ||
        String(synced[10]||'')!==members[4].name) {
      throw new Error('Golden-path bridge verification failed.');
    }

    // The bridge does not own the theme column. Add one so the live Timer path
    // proves both confirmation and configured theme handoff on WhatsApp.
    planner.getRange(plannerRowNo,4).setValue('E2E GOLDEN TEST THEME');

    const row=planner.getRange(plannerRowNo,1,1,28).getValues()[0];
    const available=availableRoleDefs_(row).map(def=>({
      code:def.code,
      name:def.name,
      speakerName:def.speakerName||''
    }));
    const codes=available.map(x=>x.code);
    if (JSON.stringify(codes)!==JSON.stringify(['TTM','Timer','Ah','GRM','EV2'])) {
      throw new Error('Unexpected golden-path vacancies: '+JSON.stringify(codes));
    }

    const expiresAt=Date.now()+15*60*1000;
    const state={
      member:member,
      phone:phone,
      meetingNo:meetingNo,
      meetingDate:dateText,
      plannerRowNo:plannerRowNo,
      tempMeetingId:tempMeetingId,
      expectedRole:'Timer',
      expiresAt:expiresAt
    };
    props.setProperty(API_DEV_GOLDEN_PROPERTY_,JSON.stringify(state));
    props.setProperty(API_DEV_ROLE_PREVIEW_PROPERTY_,JSON.stringify({
      member:member,
      phone:phone,
      meetingNo:meetingNo,
      meetingDate:dateText,
      available:available,
      expiresAt:expiresAt
    }));

    const result=sendApprovedRoleCall_(
      phone,meetingNo,dateText
    );
    if (!result || !result.ok) {
      throw new Error(
        'Golden-path template failed: '+
        String(result&&result.error||'Unknown Meta error')
      );
    }

    const report={
      ok:true,
      version:API_V1_VERSION_,
      temporary_meeting_no:meetingNo,
      available_roles:codes,
      expected_selection:'Timer',
      expected_live_messages:[
        'Timer booking confirmation',
        'E2E GOLDEN TEST THEME handoff'
      ],
      expires_in_minutes:15
    };
    console.log(JSON.stringify(report,null,2));
    return report;

  } catch(err) {
    const state={
      plannerRowNo:plannerRowNo,
      tempMeetingId:tempMeetingId,
      meetingNo:meetingNo
    };
    apiDevCleanupGoldenRoleState_(state);
    throw err;
  } finally {
    if (originalAllow===null) props.deleteProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_);
    else props.setProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_,originalAllow);
  }
}

function apiDevIsGoldenRoleSelection_(member,phone,interaction) {
  if (configValue_('SYSTEM_MODE','')!=='DEV' ||
      String(configValue_('WA_SEND_MODE','TEST')).toUpperCase()!=='TEST') {
    return false;
  }

  const raw=PropertiesService.getScriptProperties()
    .getProperty(API_DEV_GOLDEN_PROPERTY_);
  if (!raw) return false;

  let state;
  try { state=JSON.parse(raw); } catch(ignore) { return false; }

  return Date.now()<=Number(state.expiresAt||0) &&
    String(member)===String(state.member) &&
    normalizePhone_(phone)===normalizePhone_(state.phone) &&
    Number(interaction.meetingNo)===Number(state.meetingNo) &&
    String(interaction.roleCode)===String(state.expectedRole);
}

function apiDevFinalizeGoldenRoleSelection_(member,phone,interaction,roleDef,result) {
  const props=PropertiesService.getScriptProperties();
  const raw=props.getProperty(API_DEV_GOLDEN_PROPERTY_);
  if (!raw) return false;

  let state;
  try { state=JSON.parse(raw); } catch(ignore) { return false; }

  const ss=apiSpreadsheet_();
  const planner=ss.getSheetByName('VPE Weekly Planner');
  let plannerVerified=false;
  let taskVerified=false;

  if (planner &&
      Number(state.plannerRowNo)>0 &&
      planner.getMaxRows()>=Number(state.plannerRowNo)) {
    const row=planner.getRange(
      Number(state.plannerRowNo),1,1,28
    ).getValues()[0];
    plannerVerified=
      String(row[roleDef.column-1]||'').trim()===String(member);
  }

  const taskSheet=ss.getSheetByName(TASK_SHEET_);
  if (taskSheet && taskSheet.getLastRow()>=2) {
    const rows=taskSheet.getRange(
      2,1,taskSheet.getLastRow()-1,18
    ).getValues();
    taskVerified=rows.some(r=>
      Number(r[1])===Number(state.meetingNo) &&
      String(r[2])==='VACANT_ROLE' &&
      String(r[3])===String(member) &&
      String(r[4])===String(roleDef.code) &&
      String(r[5])==='Completed'
    );
  }

  const report={
    ok:Boolean(
      result &&
      result.status==='BOOKED' &&
      plannerVerified &&
      taskVerified
    ),
    version:API_V1_VERSION_,
    meeting_no:Number(state.meetingNo),
    role:String(roleDef.code),
    booking_status:String(result&&result.status||''),
    planner_verified:plannerVerified,
    task_verified:taskVerified,
    completed_at:new Date().toISOString()
  };

  props.setProperty(
    API_DEV_GOLDEN_RESULT_PROPERTY_,
    JSON.stringify(report)
  );

  logAutomation_(
    'DEV Golden Path',
    state.meetingNo,
    member,
    'Real WhatsApp Timer golden-path verification',
    report.ok ? 'Success' : 'Error',
    JSON.stringify(report)
  );

  apiDevCleanupGoldenRoleState_(state);
  return report.ok;
}

function apiDevCleanupGoldenRoleState_(state) {
  const props=PropertiesService.getScriptProperties();
  state=state||{};

  const ss=apiSpreadsheet_();
  const planner=ss.getSheetByName('VPE Weekly Planner');
  if (planner &&
      Number(state.plannerRowNo)>0 &&
      planner.getMaxRows()>=Number(state.plannerRowNo)) {
    planner.getRange(Number(state.plannerRowNo),1,1,28).clearContent();
  }

  const meetingNo=Number(state.meetingNo||0);
  const taskSheet=ss.getSheetByName(TASK_SHEET_);
  if (meetingNo && taskSheet && taskSheet.getLastRow()>=2) {
    const count=taskSheet.getLastRow()-1;
    const rows=taskSheet.getRange(2,1,count,18).getValues();
    const kept=rows.filter(r=>Number(r[1])!==meetingNo);
    const area=taskSheet.getRange(2,1,count,18);
    area.clearContent();
    if (kept.length) {
      taskSheet.getRange(2,1,kept.length,18).setValues(kept);
    }
  }

  const outbox=ss.getSheetByName('WA Outbox');
  if (meetingNo && outbox && outbox.getLastRow()>=2) {
    const width=Math.max(16,outbox.getLastColumn());
    const count=outbox.getLastRow()-1;
    const rows=outbox.getRange(2,1,count,width).getValues();
    const kept=rows.filter(r=>Number(r[10])!==meetingNo);
    const area=outbox.getRange(2,1,count,width);
    area.clearContent();
    if (kept.length) {
      outbox.getRange(2,1,kept.length,width).setValues(kept);
    }
  }

  if (state.tempMeetingId) {
    const sheets=apiEnsureSheets_();

    const itemLast=sheets.items.getLastRow();
    if (itemLast>=2) {
      const rows=sheets.items.getRange(2,1,itemLast-1,11).getValues();
      const kept=rows.filter(r=>String(r[1])!==String(state.tempMeetingId));
      const area=sheets.items.getRange(2,1,itemLast-1,11);
      area.clearContent();
      if (kept.length) {
        sheets.items.getRange(2,1,kept.length,11).setValues(kept);
      }
    }

    const meetingLast=sheets.meetings.getLastRow();
    if (meetingLast>=2) {
      const rows=sheets.meetings.getRange(2,1,meetingLast-1,7).getValues();
      const kept=rows.filter(r=>String(r[0])!==String(state.tempMeetingId));
      const area=sheets.meetings.getRange(2,1,meetingLast-1,7);
      area.clearContent();
      if (kept.length) {
        sheets.meetings.getRange(2,1,kept.length,7).setValues(kept);
      }
    }

    apiInvalidateMeetingCaches_(
      API_DEFAULT_CLUB_ID_,
      String(state.tempMeetingId)
    );
  }

  props.deleteProperty(API_DEV_GOLDEN_PROPERTY_);
  props.deleteProperty(API_DEV_ROLE_PREVIEW_PROPERTY_);
  props.deleteProperty(API_DEV_ROLE_BOOKING_PROPERTY_);
}

function apiDevGoldenPathResult() {
  const raw=PropertiesService.getScriptProperties()
    .getProperty(API_DEV_GOLDEN_RESULT_PROPERTY_);
  if (!raw) {
    return {
      ok:false,
      version:API_V1_VERSION_,
      status:'NO_RESULT',
      message:'No completed golden-path result is stored.'
    };
  }
  let result;
  try { result=JSON.parse(raw); }
  catch(ignore) {
    return {
      ok:false,
      version:API_V1_VERSION_,
      status:'INVALID_RESULT'
    };
  }
  console.log(JSON.stringify(result,null,2));
  return result;
}
