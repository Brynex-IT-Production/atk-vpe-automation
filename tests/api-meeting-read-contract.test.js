'use strict';
const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');

const source = fs.readFileSync('WebApiV1.js','utf8');
const start = source.indexOf('function apiGetMeeting_(clubId,meetingId) {');
const end = source.indexOf('\nfunction apiListMeetings_',start);
assert.ok(start>=0 && end>start,'meeting reader exists');
const block = source.slice(start,end);
assert.ok(!block.includes('const beforeItems = apiGetMeeting_('),
  'Meeting reader must not call itself when cache misses.');
const meetingId='MTG_45520261102A001';
let reads=0;
const row=['ITM_45520261102A004',meetingId,'ATK',4,'speech','S1','Speaker 1','ATK024','Riyas AF','',''];
const context={
 apiCacheKey_:()=>meetingId,
 apiCacheGetJson_:()=>null,
 apiEnsureSheets_:()=>({items:{
  getLastRow:()=>2,
  getRange:()=>({getValues:()=>{reads++;return [row];}})
 }}),
 apiFindMeetingRecord_:()=>({values:[meetingId,'ATK','2026-11-02','ACTIVE','','',455]}),
 configValue_:()=> 'Asia/Kolkata',
 apiCacheMeeting_:x=>x
};
vm.createContext(context);
vm.runInContext(block,context);
const result=context.apiGetMeeting_('ATK',meetingId);
assert.equal(result.meeting_id,meetingId);
assert.equal(result.meeting_date,'2026-11-02');
assert.equal(result.planner_meeting_no,'455');
assert.equal(result.status,'ACTIVE');
assert.equal(result.items.length,1);
assert.equal(result.items[0].member_name,'Riyas AF');
assert.equal(reads,1,'Read should be performed once without self-recursion.');
console.log('Meeting 455 uncached detail read contract PASS');
