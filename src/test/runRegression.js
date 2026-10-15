require('dotenv').config();
require('../compat');

console.log('====================================================');
console.log('  🧪 Running ATK VPE Automation Master Regression   ');
console.log('====================================================\n');

const started = Date.now();

try {
  const result = globalThis.apiDevMasterRegressionSuite();
  const duration = Date.now() - started;

  console.log(`⏱ Duration: ${duration}ms\n`);

  const suites = [
    { name: 'Grammarian Parser', data: result.grammarian_parser },
    { name: 'Planner API Lifecycle', data: result.planner_api_lifecycle },
    { name: 'Core End-to-End Engine', data: result.core },
    { name: 'Configured Role Workflows', data: result.configured_workflows },
    { name: 'Forced-Enabled Role Workflows', data: result.forced_enabled_workflows }
  ];

  let totalPassed = 0;
  let totalChecks = 0;

  for (const s of suites) {
    if (s.data) {
      const passed = s.data.checks_passed || 0;
      const total = s.data.checks_total || 0;
      totalPassed += passed;
      totalChecks += total;
      const icon = s.data.ok ? '✅' : '❌';
      console.log(`${icon} ${s.name.padEnd(30)}: ${passed}/${total} checks passed`);
    } else {
      console.log(`⚠️ ${s.name.padEnd(30)}: No result returned`);
    }
  }

  console.log('\n----------------------------------------------------');
  console.log(`Total Checks Passed: ${totalPassed} / ${totalChecks}`);
  console.log('----------------------------------------------------');

  if (result.ok) {
    console.log('\n🎉 ALL MASTER REGRESSION TESTS PASSED! (100% Exact Logic Parity)\n');
    process.exit(0);
  } else {
    console.error('\n❌ MASTER REGRESSION FAILED!\n');
    console.error(JSON.stringify(result, null, 2));
    process.exit(1);
  }
} catch (err) {
  console.error('\n💥 Unhandled exception during regression testing:');
  console.error(err);
  process.exit(1);
}
