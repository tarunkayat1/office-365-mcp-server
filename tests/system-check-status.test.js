/** system check_status must see the per-request token the Ask G1 app sends. */
const test = require('node:test');
const assert = require('node:assert/strict');
const tokenManager = require('../auth/token-manager');
const { handleCheckAuthStatus } = require('../auth/tools');

test('reports authenticated when a request-scoped token is present', async () => {
  const result = await tokenManager.runWithAccessToken('abc.def.ghi', () => handleCheckAuthStatus());
  assert.match(result.content[0].text, /Authenticated and ready/);
  assert.doesNotMatch(result.content[0].text, /Not authenticated/);
});

test('still reports not authenticated with no token at all', async () => {
  const original = tokenManager.loadTokenCache;
  tokenManager.loadTokenCache = () => null;
  try {
    const result = await handleCheckAuthStatus();
    assert.match(result.content[0].text, /Not authenticated/);
  } finally {
    tokenManager.loadTokenCache = original;
  }
});
