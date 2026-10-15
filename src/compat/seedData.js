const fs = require('fs');
const path = require('path');

function seedDefaultSpreadsheet(storageDir, spreadsheetId = '1qZnZov1l9DfvhTzruqlLubGhydcOOFKLf2kG8xplNLM') {
  if (!fs.existsSync(storageDir)) {
    fs.mkdirSync(storageDir, { recursive: true });
  }

  const filePath = path.join(storageDir, `${spreadsheetId}.json`);
  if (fs.existsSync(filePath)) {
    return; // Already initialized
  }

  const today = new Date();
  const nextWeek = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000);
  const weekAfter = new Date(today.getTime() + 14 * 24 * 60 * 60 * 1000);

  const seed = {
    sheets: {
      'Automation Config': [
        ['Key', 'Value'],
        ['SYSTEM_MODE', 'DEV'],
        ['TIMEZONE', 'Asia/Kolkata'],
        ['ROLE_SELECTION_ENABLED', 'TRUE'],
        ['TASK_ENGINE_ENABLED', 'TRUE'],
        ['VPE_MEMBER_NAME', 'Riyas AF'],
        ['WA_SEND_MODE', 'TEST'],
        ['WA_TEST_RECIPIENTS', '919846000000,919999999999,918888888888,917777777777,916666666666,915555555555'],
        ['INTRO_WORKFLOW_ENABLED', 'TRUE'],
        ['TMOD_WORKFLOW_ENABLED', 'TRUE'],
        ['GRAMMARIAN_WORKFLOW_ENABLED', 'TRUE'],
        ['TIMER_WORKFLOW_ENABLED', 'TRUE'],
        ['AH_COUNTER_WORKFLOW_ENABLED', 'TRUE'],
        ['GE_WORKFLOW_ENABLED', 'TRUE'],
        ['SPEAKER_WORKFLOW_ENABLED', 'TRUE'],
        ['EVALUATOR_WORKFLOW_ENABLED', 'TRUE'],
        ['TTM_WORKFLOW_ENABLED', 'TRUE'],
        ['SAA_WORKFLOW_ENABLED', 'TRUE'],
        ['API_DEFAULT_CLUB_ID', 'atk_default']
      ],
      'Member Master': [
        ['Member ID', 'Name', 'WhatsApp', 'Status'],
        ['M001', 'Riyas AF', '919846000000', 'Active'],
        ['M002', 'Member Two', '919999999999', 'Active'],
        ['M003', 'Member Three', '918888888888', 'Active'],
        ['M004', 'Member Four', '917777777777', 'Active'],
        ['M005', 'Member Five', '916666666666', 'Active'],
        ['M006', 'Member Six', '915555555555', 'Active']
      ],
      'VPE Weekly Planner': [
        ['Meeting No', 'Date', 'Theme', 'TMOD', 'General Evaluator', 'Timer', 'Ah-Counter', 'Grammarian', 'Table Topics Master', 'Speaker 1', 'Speaker 2', 'Speaker 3', 'Speaker 4', 'Evaluator 1', 'Evaluator 2', 'Evaluator 3', 'Evaluator 4', 'Status', 'Notes', 'VPE Hold', 'Lock', 'Col22', 'Col23', 'Col24', 'Col25', 'Col26', 'Col27', 'Col28'],
        ['Header', 'Subheader'],
        [101, nextWeek.toISOString(), 'Leadership Horizons', 'Member Two', 'Member Three', 'Member Four', 'Member Five', 'Member Six', '', 'Member Three', 'Member Four', '', '', 'Member Five', 'Member Six', '', '', 'SCHEDULED', '', '', '', '', '', '', '', '', '', ''],
        [102, weekAfter.toISOString(), 'Voice of Resilience', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'DRAFT', '', '', '', '', '', '', '', '', '', '']
      ],
      'Automation Tasks': [
        ['Task ID', 'Type', 'Target Date', 'Status', 'Payload', 'Created At', 'Updated At', 'Scheduled For', 'Attempt', 'Last Error']
      ],
      'WA Outbox': [
        ['Outbox ID', 'Member', 'Phone', 'Kind', 'Status', 'Text', 'State', 'Result', 'Meeting No', 'Key', 'Timestamp']
      ],
      'WA Submissions': [
        ['Submission ID', 'Phone', 'Member', 'Type', 'Content', 'Timestamp', 'Processed', 'Action Taken', 'Notes']
      ],
      'Automation Log': [
        ['Timestamp', 'Module', 'Meeting No', 'Member/Item', 'Action', 'Result', 'Notes']
      ],
      'WA Webhook Log': [
        ['Timestamp', 'Event Type', 'Sender', 'Raw Payload']
      ],
      'Agenda Preview': [
        ['Preview ID', 'Meeting No', 'HTML Content', 'Updated At']
      ],
      'Agenda Approvals': [
        ['Approval ID', 'Meeting No', 'Status', 'Approved By', 'Timestamp']
      ],
      'Programme Slots': [
        ['Slot ID', 'Meeting No', 'Slot No', 'Type', 'Speaker/Role', 'Status', 'Created At', 'Updated At']
      ],
      'Programme Reviews': [
        ['Token', 'Period Key', 'Created At', 'Expires At', 'Status']
      ],
      'API Club Settings': [
        ['Club ID', 'Club Name', 'Timezone', 'Meeting Time', 'Updated At'],
        ['atk_default', 'Advanced Toastmasters of Kerala', 'Asia/Kolkata', '10:00 AM', new Date().toISOString()]
      ],
      'API ExCom': [
        ['Club ID', 'Officer Role', 'Member ID', 'Member Name', 'Updated At'],
        ['atk_default', 'VPE', 'M001', 'Riyas AF', new Date().toISOString()]
      ],
      'API Meetings': [
        ['Meeting ID', 'Club ID', 'Meeting Date', 'Status', 'Created At', 'Updated At', 'Planner Meeting No']
      ],
      'API Meeting Items': [
        ['Item ID', 'Meeting ID', 'Club ID', 'Position', 'Type', 'Code', 'Label', 'Member ID', 'Member Name', 'Created At', 'Updated At']
      ]
    }
  };

  fs.writeFileSync(filePath, JSON.stringify(seed, null, 2), 'utf8');
}

module.exports = { seedDefaultSpreadsheet };
