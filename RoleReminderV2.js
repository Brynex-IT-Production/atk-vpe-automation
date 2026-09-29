/** One-time submission of the reviewed role invitation. Does not switch sending. */
function submitRoleReminderV2_() {
  if (configValue_('SYSTEM_MODE', '') !== 'DEV' ||
      configValue_('WA_SEND_MODE', '') !== 'TEST') {
    throw new Error('Role template submission requires DEV TEST mode.');
  }
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('META_ACCESS_TOKEN');
  const wabaId = props.getProperty('WA_WABA_ID');
  if (!token || !wabaId) throw new Error('Meta template credentials are missing.');
  const spec = roleReminderV2Spec_();
  const result = createOrReadTemplate_(token, wabaId, spec);
  logAutomation_('WhatsApp Templates', '', '', 'Role reminder v2 submission',
    result.ok ? 'Success' : 'Error', JSON.stringify(result));
  return result;
}

function roleReminderV2Spec_() {
  return {
    name: 'atk_role_reminder_v2',
    language: 'en_US',
    category: 'UTILITY',
    body: 'Hi,\n\nHope you are attending the meeting this week.\n\n' +
      'Please tap below to see the available roles for ATK Meeting No. {{1}}.\n\nThank you.',
    examples: ['451'],
    buttons: [{type: 'QUICK_REPLY', text: 'View roles'}]
  };
}
