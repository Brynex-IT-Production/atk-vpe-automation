'use strict';
const fs=require('node:fs');
const vm=require('node:vm');
const assert=require('node:assert/strict');

const source=fs.readFileSync('Speaking.js','utf8');
const api=fs.readFileSync('WebApiV1.js','utf8');
const extract=(start,end)=>{
  const i=source.indexOf('function '+start+'(');
  const j=end?source.indexOf('function '+end+'(',i+1):source.length;
  assert.ok(i>=0&&j>i,'Must locate '+start);
  return source.slice(i,j);
};

const replies=[],unexpected=[];
const context={
  configBool_:key=>{
    assert.equal(key,'SPEAKING_SLOT_WORKFLOW_ENABLED');
    return false;
  },
  SpreadsheetApp:{openById:()=>{unexpected.push('sheet opened');throw Error('Must not load sheet')}},
  LockService:{getScriptLock:()=>{unexpected.push('lock');throw Error('Must not reserve')}},
  availableSpeechMeetings_:()=>{unexpected.push('list dates');throw Error('Must not list dates')},
  templateNameForKind_:()=>{unexpected.push('template read');throw Error('Must not inspect template')},
  sendWhatsAppText_:(phone,message)=>replies.push({phone,message}),
  sendAutomatedQuickReply_:()=>{unexpected.push('invitation sent');throw Error('Must not invite')}
};
vm.createContext(context);

vm.runInContext(extract('speechBookingClosedMessage_','reserveSpeechSlot_'),context);
vm.runInContext(extract('reserveSpeechSlot_','sendSpeakerDetailsPrompt_'),context);
vm.runInContext(extract('processSpeechDateInvites_','speechDateInviteBody_'),context);
vm.runInContext(extract('handleSpeechDateMenu_',null),context);
vm.runInContext(extract('processSpeakingSlotTargeting_','sendSpeechSlotInvite_'),context);

const reserved=context.reserveSpeechSlot_(456,'Riyas AF');
assert.equal(reserved.status,'CLOSED','No reservation possible directly');
context.handleSpeechSlotSelection_('Riyas AF','919999999999',{meetingNo:456});
context.handleSpeechDateMenu_('Riyas AF','919999999999');
context.processSpeechDateInvites_();
context.processSpeakingSlotTargeting_();
assert.equal(replies.length,2,'Old buttons get a polite reply');
for(const x of replies) {
  assert.match(x.message,/booking is currently closed/);
  assert.match(x.message,/VPE/);
}
assert.deepEqual(unexpected,[],'No interaction may reach a booking, planner or invite when disabled');
assert.ok(api.includes("'closed_speech_booking_rejected'"),
  'Lifecycle self-test must verify closed booking behaviour');
assert.ok(api.includes("'closed_booking_does_not_modify_api'"),
  'Lifecycle self-test must verify no API mutation');
console.log('Speech booking closure: old button taps blocked, invitations stopped, no sheet write PASS');
