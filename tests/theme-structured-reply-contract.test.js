'use strict';

const fs=require('node:fs');
const vm=require('node:vm');
const assert=require('node:assert/strict');

function functionSource(path, name) {
  const source=fs.readFileSync(path,'utf8');
  const begin=source.indexOf('function '+name+'(');
  assert.ok(begin>=0,'Function exists: '+name);
  const next=source.indexOf('\nfunction ',begin+1);
  return source.slice(begin,next===-1?source.length:next);
}

let conversation={kind:'THEME',meetingNo:455,related:'TMOD',step:'THEME',data:{}};
let clearCalls=0,fallbackCalls=0,submissions=[];
const context={
  activeConversation_:()=>conversation,
  parseSpeakerDetailsBlock_:()=>({recognized:0}),
  clearConversation_:()=>{conversation=null;clearCalls++;},
  sendStructuredInputFallbackOnce_:()=>{fallbackCalls++;},
  themeCopyMessage_:()=> 'Theme:',
  handleThemeSubmission_:(member,phone,no,theme,messageId)=>{
    submissions.push({member,phone,no,theme,messageId});
  }
};
vm.createContext(context);
vm.runInContext(functionSource('Conversation.js','handleGuidedReply_'),context);
vm.runInContext(functionSource('Workflow.js','handleStructuredSubmission_'),context);

const message='Theme: Movie dialogues for every occasion';
assert.equal(context.handleGuidedReply_('Dileep Choyappally','919447350954',message,'wamid-test'),false,
  'Theme: must fall through guided handler to structured submission parser');
assert.equal(clearCalls,1,'clear guided theme state');
assert.equal(fallbackCalls,0,'never instruct member to re-send correctly formatted theme');
assert.equal(context.handleStructuredSubmission_('Dileep Choyappally','919447350954',message,'wamid-test'),true);
assert.deepEqual(submissions,[{
  member:'Dileep Choyappally',phone:'919447350954',no:0,
  theme:'Movie dialogues for every occasion',messageId:'wamid-test'
}], 'structured parser should receive clean meeting theme');

conversation={kind:'THEME',meetingNo:455,related:'TMOD',step:'THEME',data:{}};
assert.equal(context.handleGuidedReply_('Dileep Choyappally','919447350954',
  'THEME 455: Movie dialogues for every occasion','wamid-test-2'),false,
  'optional explicit meeting number accepted');
assert.equal(context.handleStructuredSubmission_('Dileep Choyappally','919447350954',
  'THEME 455: Movie dialogues for every occasion','wamid-test-2'),true);
assert.equal(submissions[1].no,455);

conversation={kind:'THEME',meetingNo:455,related:'TMOD',step:'THEME',data:{}};
assert.equal(context.handleGuidedReply_('Dileep Choyappally','919447350954',
  'Hello, can we discuss?', 'wamid-test-3'),true);
assert.equal(fallbackCalls,1,'ordinary chat still gets a single format guidance');
assert.equal(submissions.length,2,'ordinary chat does not save a theme');

console.log('PASS: copied Theme: and THEME 455: replies reach structured parser, ordinary chat does not');
