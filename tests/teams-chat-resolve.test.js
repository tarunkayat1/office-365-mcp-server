/**
 * teams_chat: garbled chat IDs, name-based sends and ID-less list_messages.
 * Graph is stubbed; the chat data mirrors a real tenant's `me/chats` listing.
 * Run with: node --test tests/teams-chat-resolve.test.js
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

const calls = [];
let graphImpl = async () => {
  throw new Error('graph stub not set');
};

// Stub the auth + Graph modules before the handler loads them.
const stub = (relative, exports) => {
  const resolved = require.resolve(path.join('..', relative));
  require.cache[resolved] = { id: resolved, filename: resolved, loaded: true, exports };
};
stub('auth', { ensureAuthenticated: async () => 'token' });
stub('utils/graph-api', {
  callGraphAPI: async (token, method, p, body, query) => {
    calls.push([token, method, p, body, query]);
    return graphImpl(token, method, p, body, query);
  },
});

const handleTeamsChat = require('../teams/consolidated/teams_chat');

const KARAN_ID =
  '19:9444193b-2d01-4de1-a636-99fa6874382a_b72d146f-863e-4e98-8391-f656709bec29@unq.gbl.spaces';
const SHIVAM_ID =
  '19:b72d146f-863e-4e98-8391-f656709bec29_bd2883ce-fc19-498b-9c56-b50bb12fd6d1@unq.gbl.spaces';
const NAVEEN_ID =
  '19:8498b45d-bc1b-4349-bbe4-b6ea9daaf650_b72d146f-863e-4e98-8391-f656709bec29@unq.gbl.spaces';
const GROUP_ID = '19:b1f917148dcd4344a5d07cc0dfc3568f@thread.v2';

const chats = [
  {
    id: KARAN_ID,
    chatType: 'oneOnOne',
    lastUpdatedDateTime: '2026-10-12T17:38:01Z',
    members: [{ displayName: 'Karan Devda' }, { displayName: 'Tarun Kayat' }],
  },
  {
    id: SHIVAM_ID,
    chatType: 'oneOnOne',
    lastUpdatedDateTime: '2026-07-17T12:41:55Z',
    members: [{ displayName: 'Tarun Kayat' }, { displayName: 'Shivam Dubey' }],
  },
  {
    id: NAVEEN_ID,
    chatType: 'oneOnOne',
    lastUpdatedDateTime: '2026-10-06T11:04:19Z',
    members: [{ displayName: 'Naveen  Vashist' }, { displayName: 'Tarun Kayat' }],
  },
  {
    id: GROUP_ID,
    chatType: 'group',
    topic: 'Internal Tech Discussion',
    lastUpdatedDateTime: '2026-10-13T16:18:51Z',
    members: Array.from({ length: 12 }, (_, i) => ({ displayName: `Member ${i}` })),
  },
];

const posts = () => calls.filter((c) => c[1] === 'POST');

test.beforeEach(() => {
  calls.length = 0;
  graphImpl = async (_token, method, p) => {
    if (method === 'GET' && p === 'me/chats') return { value: chats };
    if (method === 'POST' && /^chats\/.+\/messages$/.test(p)) return { id: 'msg-1' };
    if (method === 'GET' && /^chats\/.+\/messages$/.test(p)) {
      return {
        value: [
          {
            id: 'm1',
            messageType: 'message',
            createdDateTime: p.includes('9444193b') ? '2026-10-19T10:00:00Z' : '2026-10-18T10:00:00Z',
            from: { user: { displayName: 'Someone' } },
            body: { content: 'hello there' },
          },
        ],
      };
    }
    throw new Error(`unexpected call ${method} ${p}`);
  };
});

test('sends to the right chat when the ID is truncated and garbled', async () => {
  const garbled = `${KARAN_ID.slice(0, 80)}...29e60e9634...`;
  const result = await handleTeamsChat({
    operation: 'send_message',
    chatId: garbled,
    content: 'hii from g1desk',
  });
  assert.match(result.content[0].text, /Message sent successfully/);
  assert.equal(posts().length, 1);
  assert.equal(posts()[0][2], `chats/${KARAN_ID}/messages`);
  assert.equal(posts()[0][3].body.content, 'hii from g1desk');
});

test('leaves an exact, valid ID untouched', async () => {
  await handleTeamsChat({ operation: 'send_message', chatId: SHIVAM_ID, content: 'hi' });
  assert.equal(posts()[0][2], `chats/${SHIVAM_ID}/messages`);
});

test('sends by recipient name and prefers the 1:1 chat', async () => {
  const result = await handleTeamsChat({
    operation: 'send_message',
    recipient: 'karan devda',
    content: 'hii from g1desk',
  });
  assert.match(result.content[0].text, /Message sent successfully/);
  assert.equal(posts()[0][2], `chats/${KARAN_ID}/messages`);
});

test('does not send when the recipient matches no chat', async () => {
  const result = await handleTeamsChat({
    operation: 'send_message',
    recipient: 'nobody here',
    content: 'hi',
  });
  assert.match(result.content[0].text, /No existing chat found/);
  assert.equal(posts().length, 0);
});

test('list_messages without chatId returns newest messages across recent chats', async () => {
  const result = await handleTeamsChat({ operation: 'list_messages', maxResults: 5 });
  const text = result.content[0].text;
  assert.match(text, /Latest/);
  assert.match(text, /Karan Devda, Tarun Kayat/);
  assert.ok(text.indexOf('Karan Devda') < text.indexOf('Internal Tech Discussion'));
});

test('finds a chat whose display name has extra spaces (Naveen  Vashist)', async () => {
  const result = await handleTeamsChat({
    operation: 'send_message',
    recipient: 'Naveen Vashist',
    content: 'hello bro this is g1desk',
  });
  assert.match(result.content[0].text, /Message sent successfully/);
  assert.equal(posts()[0][2], `chats/${NAVEEN_ID}/messages`);
});

test('matches names in any order, case and spacing', async () => {
  const result = await handleTeamsChat({
    operation: 'send_message',
    recipient: '  vashist   NAVEEN ',
    content: 'hi',
  });
  assert.match(result.content[0].text, /Message sent successfully/);
  assert.equal(posts()[0][2], `chats/${NAVEEN_ID}/messages`);
});
