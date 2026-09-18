import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createHmac } from 'node:crypto';
import { once } from 'node:events';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { createApp } from '../src/index.js';
import { createExcelStore } from '../src/excel.js';
import { extractPhoneNumbers, normalizePhoneNumber } from '../src/phone.js';
import { parseMessages } from '../src/instagram.js';
import type { Lead } from '../src/types.js';

test('Turkish mobile formats and invalid boundaries', () => {
  for (const phone of ['05321234567', '5321234567', '+905321234567', '0532 123 45 67', '0532-123-45-67', '(0532) 123 45 67']) {
    assert.equal(normalizePhoneNumber(phone), '+905321234567');
    assert.deepEqual(extractPhoneNumbers(`Merhaba ${phone} beni arayın`), ['+905321234567']);
  }
  for (const phone of ['02121234567', '+495321234567', '15321234567', '053212345678', '532123456', '00905321234567']) {
    assert.deepEqual(extractPhoneNumbers(phone), []);
  }
  assert.deepEqual(extractPhoneNumbers('Numaram 5551234567 dönüş yapabilir misiniz?'), ['+905551234567']);
});

const payload = { object: 'instagram', entry: [{ id: '123', messaging: [{ sender: { id: '456' }, recipient: { id: '123' }, message: { mid: 'm1', text: 'Merhaba 0532 123 45 67' } }] }] };

test('local mode accepts the minimal unsigned fixture, deduplicates, and never calls Meta', async () => {
  const logs: string[] = [];
  const originalLog = console.log;
  console.log = (...values) => { logs.push(values.join(' ')); };
  let metaCalls = 0;
  const seen = new Set<string>();
  const server = createApp({ localTestMode: true, verifyToken: '', appSecret: '', accountId: '', store: { async insertIfNew(lead, messageId) {
    const key = `${lead.phone}:${messageId}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  } },
    instagram: { async username() { metaCalls++; return ''; }, async reply() { metaCalls++; } },
  }).listen(0, '127.0.0.1');
  try {
    await once(server, 'listening');
    const address = server.address();
    assert.ok(address && typeof address !== 'string');
    const fixture = { object: 'instagram', entry: [{ messaging: [{ sender: { id: '123456789' }, message: { mid: 'HYS_TEST_001', text: 'Merhaba, 0532 123 45 67 numarasından beni arar mısınız?' } }] }] };
    assert.deepEqual(parseMessages(fixture, '123'), []);
    for (let i = 0; i < 2; i++) {
      const response = await fetch(`http://127.0.0.1:${address.port}/webhook`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(fixture) });
      assert.equal(response.status, 200);
      assert.equal(response.headers.get('x-hys-local-test'), 'true');
    }
    assert.equal(metaCalls, 0);
    assert.equal(logs.filter(line => line.includes('Instagram reply would be sent.')).length, 1);
    assert.ok(logs.some(line => line.includes('Phone: +905321234567')));
  } finally { console.log = originalLog; server.close(); await once(server, 'close'); }
});

test('Excel store creates a workbook and prevents message and phone duplicates', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'hys-excel-test-'));
  const filePath = join(directory, 'HYS_Instagram_Leads.xlsx');
  try {
    const store = createExcelStore(filePath);
    const lead: Lead = { created_at: '2026-09-18T06:45:00.000Z', instagram_user_id: '123456789', instagram_username: 'ahmet', phone: '+905321234567', original_message: 'Beni arar mısınız?', status: 'NEW' };
    assert.equal(await store.insertIfNew(lead, 'HYS_TEST_001'), true);
    assert.equal(await store.insertIfNew(lead, 'HYS_TEST_001'), false);
    assert.equal(await store.insertIfNew({ ...lead, phone: '+905551234567' }, 'HYS_TEST_001'), false);
    assert.equal(await store.insertIfNew({ ...lead, instagram_user_id: '999' }, 'HYS_TEST_002'), false);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet('Instagram Leads');
    assert.ok(sheet);
    assert.equal(sheet.getRow(1).getCell(11).value, 'Meta Message ID');
    assert.equal(sheet.getColumn(11).hidden, true);
    assert.equal(sheet.getRow(2).getCell(1).value, 1);
    assert.equal(sheet.getRow(2).getCell(6).value, '+905321234567');
    assert.equal(sheet.getRow(2).getCell(8).value, 'YENİ');
    assert.equal(sheet.getRow(2).getCell(9).value, 'HAYIR');
    assert.equal(sheet.getRow(2).getCell(11).value, 'HYS_TEST_001');
  } finally { await rm(directory, { recursive: true, force: true }); }
});

test('unsupported and malformed events are ignored', () => {
  for (const value of [null, [], {}, { object: 'instagram', entry: [null, {}, { id: '123', messaging: [null, { read: {} }] }] }]) assert.deepEqual(parseMessages(value, '123'), []);
  assert.equal(parseMessages(payload, '123').length, 1);
  assert.equal(parseMessages(payload, '999').length, 0);
  const echo = structuredClone(payload);
  Object.assign(echo.entry[0].messaging[0].message, { is_echo: true });
  assert.deepEqual(parseMessages(echo, '123'), []);
});

test('HTTP verification, signature rejection, concurrent duplicates, restart and failure', async () => {
  const rows: Lead[] = [];
  let replies = 0;
  let fail = false;
  const deps = { verifyToken: 'test-verify', appSecret: 'test-secret', accountId: '123',
    store: { async insertIfNew(lead: Lead) { if (fail) throw new Error('simulated'); if (rows.some(row => row.phone === lead.phone)) return false; await new Promise(resolve => setTimeout(resolve, 10)); rows.push(lead); return true; } },
    instagram: { async username() { return 'customer'; }, async reply() { replies++; } } };
  async function run(check: (base: string) => Promise<void>) {
    const server = createApp(deps).listen(0, '127.0.0.1');
    await once(server, 'listening');
    const address = server.address();
    assert.ok(address && typeof address !== 'string');
    try { await check(`http://127.0.0.1:${address.port}`); } finally { server.close(); await once(server, 'close'); }
  }
  const body = JSON.stringify(payload);
  const signature = 'sha256=' + createHmac('sha256', deps.appSecret).update(body).digest('hex');
  const post = (base: string, sig = signature) => fetch(`${base}/webhook`, { method: 'POST', headers: { 'content-type': 'application/json', 'x-hub-signature-256': sig }, body });
  await run(async base => {
    assert.deepEqual(await (await fetch(`${base}/health`)).json(), { status: 'ok' });
    assert.equal(await (await fetch(`${base}/webhook?hub.mode=subscribe&hub.verify_token=test-verify&hub.challenge=abc`)).text(), 'abc');
    assert.equal((await fetch(`${base}/webhook?hub.mode=subscribe&hub.verify_token=bad&hub.challenge=abc`)).status, 403);
    assert.equal((await post(base, 'sha256=' + '0'.repeat(64))).status, 401);
    fail = true;
    assert.equal((await post(base)).status, 503);
    assert.equal(replies, 0);
    fail = false;
    const results = await Promise.all([post(base), post(base), post(base)]);
    assert.ok(results.every(result => result.status === 200));
    assert.equal(rows.length, 1);
    assert.equal(replies, 1);
    assert.equal(rows[0].status, 'NEW');
  });
  await run(async base => { assert.equal((await post(base)).status, 200); assert.equal(replies, 1); });
});
