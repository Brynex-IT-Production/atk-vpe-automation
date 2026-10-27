'use strict';
const fs=require('node:fs');
const vm=require('node:vm');
const assert=require('node:assert/strict');

const file=fs.readFileSync('Workflow.js','utf8');
const start=file.indexOf('function sendSpeakerDetailsToTmod_(');
const end=file.indexOf('\nfunction storedSubmissionText_(',start);
assert.ok(start>=0&&end>start,'handoff function should exist');
let assigned='Priyesh',attempts=0;
const sheet={
  getLastRow:()=>8,
  getRange:(row,col)=>({
    getValues:()=>col===1&&row===3
      ? [[450],[451],[452],[453],[454],[455]]
      : [[...Array(6).fill(''),assigned,...Array(11).fill(''),'Dileep Choyappally',...Array(9).fill('')]],
    getValue:()=> 'Dileep Choyappally'
  })
};
const ctx={
  meetingAutomationEligible_:()=>true,
  SPREADSHEET_ID:'test-sheet',
  SpreadsheetApp:{openById:()=>({getSheetByName:()=>sheet})},
  normalizePhone_:x=>x,
  phoneForMember_:()=> '919999999999',
  isWhatsAppRecipientAllowed_:()=>true,
  hasReminderBeenSent_:()=>false,
  sendAutomatedText_:()=>{attempts++;return {ok:true};},
  recordWorkflowOutbound_:()=>{}
};
vm.createContext(ctx);
vm.runInContext(file.slice(start,end),ctx);
const input=['Talk','Path', '1', '2', '7 minutes'];
assert.equal(ctx.sendSpeakerDetailsToTmod_(455,'Riyas AF','S1',input),false);
assert.equal(attempts,0,'Withdrawn speaker must not be sent');
assert.equal(ctx.sendSpeakerDetailsToTmod_(455,'Priyesh','S1',input),true);
assert.equal(attempts,1,'Correct assigned speaker must still be sent');
assert.equal(ctx.sendSpeakerDetailsToTmod_(455,'Priyesh','S4',input),false);
assert.equal(attempts,1,'Unknown speaker slot must not be sent');
console.log('TMOD handoff current-speaker validation PASS');
