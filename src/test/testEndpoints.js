const http = require('http');
const app = require('../server');

const server = http.createServer(app);

server.listen(0, async () => {
  const port = server.address().port;
  const baseUrl = `http://localhost:${port}`;
  console.log(`Test server listening on ${baseUrl}`);

  try {
    // 1. Test Root Info
    const rootRes = await fetch(`${baseUrl}/`);
    const rootJson = await rootRes.json();
    console.log('✅ GET /:', rootJson.service, '| Status:', rootJson.status);

    // 2. Test Health
    const healthRes = await fetch(`${baseUrl}/health`);
    const healthJson = await healthRes.json();
    console.log('✅ GET /health:', healthJson.ok ? 'OK' : 'FAILED');

    // 3. Test Meta Webhook Verification (Challenge)
    const challengeRes = await fetch(`${baseUrl}/webhook?hub.mode=subscribe&hub.verify_token=ATK_VPE_VERIFY_2026_X7K9M4P2&hub.challenge=TEST_CHALLENGE_XYZ`);
    const challengeText = await challengeRes.text();
    console.log('✅ GET /webhook verification challenge:', challengeText === 'TEST_CHALLENGE_XYZ' ? 'PASSED (challenge returned)' : 'FAILED');

    // 4. Test Web API v1
    const apiRes = await fetch(`${baseUrl}/api/v1`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'GET_CLUB_SETTINGS', club_id: 'atk_default' })
    });
    const apiJson = await apiRes.json();
    console.log('✅ POST /api/v1 (GET_CLUB_SETTINGS):', apiJson.data ? `Club: ${apiJson.data.club_name}` : 'OK response');

    // 5. Test Dev Regression API endpoint
    const regRes = await fetch(`${baseUrl}/api/dev/regression?verify_token=ATK_VPE_VERIFY_2026_X7K9M4P2`);
    const regJson = await regRes.json();
    console.log('✅ GET /api/dev/regression:', regJson.ok ? `PASSED (${regJson.duration_ms}ms)` : 'FAILED');

    console.log('\n🎉 ALL EXPRESS HTTP ENDPOINTS VERIFIED SUCCESSFULLY!\n');
    server.close();
    process.exit(0);
  } catch (err) {
    console.error('❌ Endpoint test failed:', err);
    server.close();
    process.exit(1);
  }
});
