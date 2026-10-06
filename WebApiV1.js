/**
 * Web API V1 for the VPE Assistant.
 *
 * This is deliberately an adapter layer. It gives the Flutter/Web client an
 * ordered meeting model without changing the proven ATK automation tables.
 * Existing automation continues to read VPE Weekly Planner.
 */

const API_V1_VERSION_ = '2026-10-06';
const API_MEETING_ITEMS_SHEET_ = 'API Meeting Items';
const API_MEETINGS_SHEET_ = 'API Meetings';
const API_DEFAULT_CLUB_ID_ = 'ATK';

function apiJson_(status, data, error) {
  return ContentService.createTextOutput(JSON.stringify({
    ok: status >= 200 && status < 300,
    status: status,
    version: API_V1_VERSION_,
    data: data === undefined ? null : data,
    error: error || null
  })).setMimeType(ContentService.MimeType.JSON);
}

function apiHandleRequest_(payload) {
  payload = payload || {};
  const action = String(payload.action || '').trim();
  const clubId = String(payload.club_id || API_DEFAULT_CLUB_ID_).trim();

  if (!action) return {status:400,error:{code:'MISSING_ACTION',message:'action is required'}};

  try {
    switch (action) {
      case 'health':
        return {status:200,data:{service:'vpe-api',status:'ok'}};
      case 'members.list':
        return {status:200,data:{members:apiListMembers_(clubId)}};
      case 'meeting_template.get':
        return {status:200,data:apiDefaultMeetingTemplate_(clubId)};
      case 'meeting.create':
        return {status:201,data:apiCreateMeeting_(clubId,payload)};
      case 'meeting.get':
        return {status:200,data:apiGetMeeting_(clubId,String(payload.meeting_id || ''))};
      case 'meeting.save':
        return {status:200,data:apiSaveMeeting_(clubId,payload)};
      case 'meeting.activate':
        return {status:200,data:apiActivateMeeting_(clubId,String(payload.meeting_id || ''))};
      default:
        return {status:404,error:{code:'UNKNOWN_ACTION',message:'Unknown API action: '+action}};
    }
  } catch (err) {
    return {status:500,error:{code:'INTERNAL_ERROR',message:String(err && err.message || err)}};
  }
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
    item_id:String(item.item_id || apiNewId_('ITM')),
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
  if (!meetingId) throw new Error('meeting_id is required');
  const sheets = apiEnsureSheets_();
  const mr = sheets.meetings.getLastRow() >= 2
    ? sheets.meetings.getRange(2,1,sheets.meetings.getLastRow()-1,7).getValues()
      .find(r => String(r[0])===meetingId && String(r[1])===clubId)
    : null;
  if (!mr) throw new Error('Meeting not found');

  const rows = sheets.items.getLastRow() >= 2
    ? sheets.items.getRange(2,1,sheets.items.getLastRow()-1,11).getValues()
      .filter(r => String(r[1])===meetingId && String(r[2])===clubId)
    : [];
  const items = rows.map(r=>({
    item_id:String(r[0]), position:Number(r[3]), type:String(r[4]), code:String(r[5]),
    label:String(r[6]), member_id:String(r[7]), member_name:String(r[8])
  })).sort((a,b)=>a.position-b.position);

  return {
    meeting_id:String(mr[0]), club_id:String(mr[1]), meeting_date:String(mr[2]),
    status:String(mr[3]), planner_meeting_no:String(mr[6] || ''), items:items
  };
}

function apiSaveMeeting_(clubId,payload) {
  const meetingId = String(payload.meeting_id || '');
  const sheets = apiEnsureSheets_();
  const meeting = apiGetMeeting_(clubId,meetingId);
  if (meeting.status === 'ACTIVE') throw new Error('Active meeting cannot be structurally edited in API V1');
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

function apiActivateMeeting_(clubId,meetingId) {
  const meeting = apiGetMeeting_(clubId,meetingId);
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
