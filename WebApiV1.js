/**
 * Web API V1 for the VPE Assistant.
 *
 * This is deliberately an adapter layer. It gives the Flutter/Web client an
 * ordered meeting model without changing the proven ATK automation tables.
 * Existing automation continues to read VPE Weekly Planner.
 */

const API_V1_VERSION_ = '2026-10-11.5';
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
  const cacheKey = apiCacheKey_('members',clubId);
  const cached = apiCacheGetJson_(cacheKey);
  if (cached) return cached;

  const sh = apiSpreadsheet_().getSheetByName('Member Master');
  if (!sh || sh.getLastRow() < 2) return [];
  const members = sh.getRange(2,1,sh.getLastRow()-1,4).getValues()
    .filter(r => String(r[3] || '').trim() === 'Active')
    .map(r => ({
      member_id:String(r[0] || '').trim(),
      name:String(r[1] || '').trim(),
      whatsapp:String(r[2] || '').trim(),
      status:String(r[3] || '').trim()
    }))
    .filter(m => m.member_id && m.name);

  apiCachePutJson_(cacheKey,members);
  return members;
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

function apiCreateMeeting_(clubId,payload) {
  const dateText = String(payload.meeting_date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateText)) throw new Error('meeting_date must be YYYY-MM-DD');
  const sheets = apiEnsureSheets_();
  const id = apiNewId_('MTG');
  const now = new Date();
  sheets.meetings.getRange(sheets.meetings.getLastRow()+1,1,1,7)
    .setValues([[id,clubId,dateText,'DRAFT',now,now,'']]);

  const defaults = apiDefaultMeetingTemplate_(clubId).items;
  const normalized = apiReplaceMeetingItems_(sheets.items,clubId,id,defaults);
  apiInvalidateMeetingCaches_(clubId,id);

  return apiCacheMeeting_({
    meeting_id:id,
    club_id:clubId,
    meeting_date:dateText,
    status:'DRAFT',
    planner_meeting_no:'',
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
    if (code && !knownCodes[code]) { const e=new Error('Unsupported standard code: '+code); e.apiCode='BAD_REQUEST'; throw e; }
    if (code && seenCodes[code]) { const e=new Error('Duplicate standard code: '+code); e.apiCode='BAD_REQUEST'; throw e; }
    if (code) seenCodes[code] = true;
    if (type === 'session' && code) { const e=new Error('Custom sessions must not use a standard role code'); e.apiCode='BAD_REQUEST'; throw e; }
    const memberId = String(item.member_id || '');
    if (memberId && !members[memberId]) { const e=new Error('Unknown or inactive member_id: '+memberId); e.apiCode='BAD_REQUEST'; throw e; }
    if (memberId && !String(item.member_name || '').trim()) item.member_name = members[memberId].name;
  });
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

  return {enabled:true,synced:true,planner_meeting_no:result.planner_meeting_no};
}

function apiSaveMeeting_(clubId,payload) {
  const meetingId = String(payload.meeting_id || '');
  const saveMode = String(payload.save_mode || 'DRAFT').toUpperCase();
  if (['DRAFT','CONFIRMED'].indexOf(saveMode) < 0) {
    const e = new Error('save_mode must be DRAFT or CONFIRMED'); e.apiCode='BAD_REQUEST'; throw e;
  }

  const sheets = apiEnsureSheets_();
  const rec = apiFindMeetingRecord_(sheets,clubId,meetingId);
  const mr = rec.values;
  if (String(mr[3]) === 'ACTIVE') throw new Error('Active meeting cannot be structurally edited in API V1');

  apiValidateMeetingItems_(payload.items || [],clubId);
  const normalized = apiReplaceMeetingItems_(sheets.items,clubId,meetingId,payload.items || []);

  const nextStatus = saveMode === 'CONFIRMED' ? 'PLANNED' : 'DRAFT';
  const now = new Date();
  sheets.meetings.getRange(rec.row_number,4,1,3)
    .setValues([[nextStatus,mr[4],now]]);

  apiInvalidateMeetingCaches_(clubId,meetingId);

  let plannerMeetingNo = String(mr[6] || '');
  if (saveMode === 'CONFIRMED') {
    const bridge = apiBridgePlannedMeetingIfEnabled_(clubId,meetingId);
    if (bridge && bridge.synced) plannerMeetingNo = String(bridge.planner_meeting_no || '');
  }

  const meetingDate = mr[2] instanceof Date
    ? Utilities.formatDate(mr[2], configValue_('TIMEZONE','Asia/Kolkata'), 'yyyy-MM-dd')
    : String(mr[2] || '').slice(0,10);

  return apiCacheMeeting_({
    meeting_id:meetingId,
    club_id:clubId,
    meeting_date:meetingDate,
    status:nextStatus,
    planner_meeting_no:plannerMeetingNo,
    items:normalized
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

  // Safety: API V1 activation does NOT send WhatsApp. It only marks the web plan
  // ready for the legacy-planner adapter. Messaging is enabled after adapter tests.
  const sheets = apiEnsureSheets_();
  for (let r=2;r<=sheets.meetings.getLastRow();r++) {
    if (String(sheets.meetings.getRange(r,1).getValue())===meetingId &&
        String(sheets.meetings.getRange(r,2).getValue())===clubId) {
      sheets.meetings.getRange(r,4,1,3).setValues([['READY_FOR_SYNC',sheets.meetings.getRange(r,5).getValue(),new Date()]]);
      apiInvalidateMeetingCaches_(clubId,meetingId);
      break;
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
  const allowed = String(PropertiesService.getScriptProperties().getProperty(API_ALLOW_LEGACY_SYNC_PROPERTY_) || '').toUpperCase() === 'TRUE';
  if (!allowed) {
    const err = new Error('Legacy sync is disabled in this environment');
    err.apiCode = 'LEGACY_SYNC_DISABLED';
    throw err;
  }
  if (clubId !== API_DEFAULT_CLUB_ID_) throw new Error('Legacy adapter currently supports ATK only');
  const meeting = apiGetMeeting_(clubId, meetingId);
  const ss = apiSpreadsheet_();
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
  const projection = apiBuildLegacyPlannerProjection_(meeting);

  // Speakers and their status columns.
  [['S1',7,9],['S2',11,13],['S3',15,17]].forEach(x => {
    planner.getRange(rowNo,x[1]).setValue(projection.speakers[x[0]] || '');
    planner.getRange(rowNo,x[2]).setValue(projection.speakers[x[0]] ? 'Planned' : '');
  });

  // Standard role columns.
  ROLE_DEFS_.forEach(def => planner.getRange(rowNo,def.column).setValue(projection.roles[def.code] || ''));
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

function apiDashboard_(clubId) {
  if (clubId !== API_DEFAULT_CLUB_ID_) throw new Error('Club not provisioned: '+clubId);
  const ss = apiSpreadsheet_();
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
