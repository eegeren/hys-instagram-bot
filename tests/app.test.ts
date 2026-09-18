import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createHmac } from 'node:crypto';
import { once } from 'node:events';
import { mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import ExcelJS from 'exceljs';
import { createApp } from '../src/index.js';
import { createExcelStore, HEADERS, resolveExcelPath, SHEET_NAME } from '../src/excel.js';
import { WebhookQueue } from '../src/queue.js';
import { extractPhoneNumbers, normalizePhoneNumber } from '../src/phone.js';
import { parseMessages } from '../src/instagram.js';
import type { IncomingMessage, LeadStore, ProcessResult } from '../src/types.js';

async function temporary<T>(work: (directory: string) => Promise<T>): Promise<T> {
  const directory = await mkdtemp(join(tmpdir(), 'hys-test-'));
  try { return await work(directory); } finally { await rm(directory, { recursive: true, force: true }); }
}
const message = (id = 'message-1', phone = '0532 123 45 67', timestamp = 1789713900000): IncomingMessage => ({ id, senderId: '123456789', text: `Merhaba ${phone} beni arayın`, timestamp });
const payload = (m = message()) => ({ object: 'instagram', entry: [{ id: '123', messaging: [{ sender: { id: m.senderId }, recipient: { id: '123' }, timestamp: m.timestamp, message: { mid: m.id, text: m.text } }] }] });
async function readWorkbook(path: string) { const workbook = new ExcelJS.Workbook(); await workbook.xlsx.readFile(path); return workbook; }
async function eventually<T>(get: () => Promise<T | undefined>): Promise<T> {
  const deadline = Date.now() + 5000;
  do { const value = await get(); if (value !== undefined) return value; await new Promise(resolve => setTimeout(resolve, 10)); } while (Date.now() < deadline);
  throw new Error('Timed out waiting for background work');
}

test('phone parser accepts all requested forms and rejects invalid boundaries', () => {
  for (const phone of ['05321234567', '5321234567', '+905321234567', '+90 532 123 45 67', '0532 123 45 67', '0532-123-45-67', '(0532) 123 45 67']) {
    assert.equal(normalizePhoneNumber(phone), '+905321234567');
    assert.deepEqual(extractPhoneNumbers(`Merhaba ${phone} beni arayın`), ['+905321234567']);
  }
  for (const phone of ['02121234567', '+495321234567', '15321234567', '053212345678', '532123456', '00905321234567']) assert.deepEqual(extractPhoneNumbers(phone), []);
});

test('storage path defaults and override', () => {
  assert.equal(resolveExcelPath({}, 'win32'), resolve('./data/HYS_Instagram_Leads.xlsx'));
  assert.equal(resolveExcelPath({}, 'linux'), resolve('/data/HYS_Instagram_Leads.xlsx'));
  assert.equal(resolveExcelPath({ LOCAL_TEST_MODE: 'true' }, 'linux'), resolve('./data/HYS_Instagram_Leads.xlsx'));
  assert.equal(resolveExcelPath({ EXCEL_FILE_PATH: './custom/leads.xlsx' }, 'linux'), resolve('./custom/leads.xlsx'));
});

test('Excel creation, persistent message history, repeat contact, older event and multi-phone message', async () => temporary(async directory => {
  const path = join(directory, 'data', 'leads.xlsx');
  let store = createExcelStore(path);
  await store.initialize();
  assert.equal((await readWorkbook(path)).getWorksheet(SHEET_NAME)!.rowCount, 1);
  const first = await store.processMessage(message(), 'ahmet');
  assert.equal(first.leads[0].action, 'inserted');
  let workbook = await readWorkbook(path);
  let sheet = workbook.getWorksheet(SHEET_NAME)!;
  assert.deepEqual(HEADERS.map((_, i) => sheet.getCell(1, i + 1).text), HEADERS);
  assert.equal(sheet.getColumn(10).hidden, true);
  assert.deepEqual([sheet.getCell(2, 7).text, sheet.getCell(2, 8).text], ['YENİ', 'HAYIR']);
  assert.equal(sheet.getCell(2, 4).text, '@ahmet');
  // Staff edits must survive repeat contacts.
  sheet.getCell(2, 7).value = 'TAKİP'; sheet.getCell(2, 8).value = 'EVET'; sheet.getCell(2, 9).value = 'Call tomorrow';
  await workbook.xlsx.writeFile(path);
  const second = { ...message('message-2', '5321234567', 1789713960000), text: 'Tekrar arayın 5321234567' };
  assert.equal((await store.processMessage(second, '')).leads[0].action, 'updated');
  store = createExcelStore(path);
  assert.equal((await store.processMessage(message(), '')).duplicate, true);
  assert.equal((await store.processMessage(second, '')).duplicate, true);
  assert.equal((await store.processMessage(message('older', '5321234567', 1789713800000), '')).leads[0].action, 'existing');
  workbook = await readWorkbook(path); sheet = workbook.getWorksheet(SHEET_NAME)!;
  assert.equal(sheet.rowCount, 2);
  assert.equal(sheet.getCell(2, 6).text, second.text);
  assert.deepEqual([sheet.getCell(2, 7).text, sheet.getCell(2, 8).text, sheet.getCell(2, 9).text], ['TAKİP', 'EVET', 'Call tomorrow']);
  assert.equal(workbook.getWorksheet('_Processed Messages')!.state, 'veryHidden');
  const multi = { ...message('multi'), text: 'Numaralar: 5551234567; 5561234567' };
  assert.equal((await store.processMessage(multi, '')).leads.length, 2);
  assert.equal((await store.processMessage(multi, '')).duplicate, true);
}));

test('Excel serializes concurrent writes across store instances and keeps monotonic IDs', async () => temporary(async directory => {
  const path = join(directory, 'leads.xlsx');
  const stores = [createExcelStore(path), createExcelStore(path)];
  const results = await Promise.all(Array.from({ length: 12 }, (_, i) => stores[i % 2].processMessage(message(`m${i}`, `55${String(i).padStart(8, '0')}`), '')));
  assert.equal(results.filter(r => r.leads[0].action === 'inserted').length, 12);
  const sheet = (await readWorkbook(path)).getWorksheet(SHEET_NAME)!;
  assert.equal(sheet.rowCount, 13);
  assert.deepEqual(Array.from({ length: 12 }, (_, i) => sheet.getCell(i + 2, 1).value), Array.from({ length: 12 }, (_, i) => i + 1));
}));

test('legacy migration preserves backup and idempotency; corrupt file is never overwritten', async () => temporary(async directory => {
  const path = join(directory, 'leads.xlsx');
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet(SHEET_NAME);
  sheet.addRow([...HEADERS.slice(0, 3), 'Instagram Kullanıcı ID', ...HEADERS.slice(3)]);
  sheet.addRow([1, '18.09.2026', '09:45', '123456789', '@ahmet', '+905321234567', 'Merhaba', 'YENİ', 'HAYIR', 'Keep', 'old-id']);
  await workbook.xlsx.writeFile(path);
  const original = await readFile(path);
  const store = createExcelStore(path);
  await store.initialize();
  assert.deepEqual(await readFile(`${path}.pre-railway.bak.xlsx`), original);
  assert.equal((await store.processMessage(message('old-id'), '')).duplicate, true);
  assert.equal((await readWorkbook(path)).getWorksheet(SHEET_NAME)!.getCell(2, 9).text, 'Keep');
  const corrupt = join(directory, 'corrupt.xlsx');
  await writeFile(corrupt, 'not an xlsx');
  await assert.rejects(createExcelStore(corrupt).initialize());
  assert.equal(await readFile(corrupt, 'utf8'), 'not an xlsx');
}));

test('unsupported and malformed events remain ignored', () => {
  for (const value of [null, [], {}, { object: 'instagram', entry: [null, {}, { id: '123', messaging: [null, { read: {} }] }] }]) assert.deepEqual(parseMessages(value, '123'), []);
  assert.equal(parseMessages(payload(), '123').length, 1);
  assert.equal(parseMessages(payload(), '999').length, 0);
  const echo = payload(); Object.assign(echo.entry[0].messaging[0].message, { is_echo: true });
  assert.deepEqual(parseMessages(echo, '123'), []);
});

test('local HTTP acknowledges before slow Excel finishes; duplicate reply suppressed; Meta never called', async () => temporary(async directory => {
  const realStore = createExcelStore(join(directory, 'leads.xlsx'));
  let release!: () => void;
  const blocked = new Promise<void>(resolve => { release = resolve; });
  let metaCalls = 0;
  const store: LeadStore = { initialize: () => realStore.initialize(), async processMessage(m, name) { await blocked; return realStore.processMessage(m, name); } };
  const app = await createApp({ localTestMode: true, verifyToken: '', appSecret: '', accountId: '', store,
    queueDirectory: join(directory, 'queue'), instagram: { async username() { metaCalls++; return ''; }, async reply() { metaCalls++; } } });
  const server = app.listen(0, '127.0.0.1'); await once(server, 'listening');
  const address = server.address(); assert.ok(address && typeof address !== 'string');
  const base = `http://127.0.0.1:${address.port}`;
  const fixture = { object: 'instagram', entry: [{ messaging: [{ sender: { id: '123456789' }, message: { mid: 'test-local', text: 'Merhaba 0532 123 45 67' } }] }] };
  const post = async () => {
    const response = await fetch(`${base}/webhook`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(fixture), signal: AbortSignal.timeout(1000) });
    assert.equal(response.status, 200); return await response.json() as { jobId: string };
  };
  const completed = (id: string) => eventually(async () => {
    const status = await (await fetch(`${base}/local-test/jobs/${id}`)).json() as { state: string; results: ProcessResult[] };
    return status.state === 'completed' ? status.results : undefined;
  });
  try {
    assert.deepEqual(await (await fetch(`${base}/health`)).json(), { status: 'ok' });
    const first = await post(); // This must return even though Excel is still blocked.
    assert.equal((await readdir(join(directory, 'queue'))).filter(f => f.endsWith('.json')).length, 1);
    release();
    assert.equal((await completed(first.jobId))[0].reply, 'simulated');
    const repeated = await post();
    assert.equal((await completed(repeated.jobId))[0].duplicate, true);
    assert.equal((await completed(repeated.jobId))[0].reply, 'none');
    assert.equal(metaCalls, 0);
    assert.equal((await readWorkbook(join(directory, 'leads.xlsx'))).getWorksheet(SHEET_NAME)!.rowCount, 2);
  } finally { release(); await app.locals.stop(); server.close(); await once(server, 'close'); }
}));

test('production webhook verification and signatures remain required', async () => temporary(async directory => {
  const app = await createApp({ verifyToken: 'verify', appSecret: 'secret', accountId: '123', store: createExcelStore(join(directory, 'leads.xlsx')),
    queueDirectory: join(directory, 'queue'), instagram: { async username() { return ''; }, async reply() {} } });
  const server = app.listen(0, '127.0.0.1'); await once(server, 'listening');
  const address = server.address(); assert.ok(address && typeof address !== 'string'); const base = `http://127.0.0.1:${address.port}`;
  try {
    assert.equal(await (await fetch(`${base}/webhook?hub.mode=subscribe&hub.verify_token=verify&hub.challenge=abc`)).text(), 'abc');
    assert.equal((await fetch(`${base}/webhook?hub.mode=subscribe&hub.verify_token=wrong&hub.challenge=abc`)).status, 403);
    const body = JSON.stringify(payload());
    const post = (signature: string) => fetch(`${base}/webhook`, { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Hub-Signature-256': signature }, body });
    assert.equal((await post('sha256=' + '0'.repeat(64))).status, 401);
    assert.equal((await post('sha256=' + createHmac('sha256', 'secret').update(body).digest('hex'))).status, 200);
    assert.equal((await fetch(`${base}/local-test/jobs/test`)).status, 404);
  } finally { await app.locals.stop(); server.close(); await once(server, 'close'); }
}));

test('failed accepted job survives restart and is processed from durable spool', async () => temporary(async directory => {
  let attempts = 0;
  const first = new WebhookQueue(directory, async () => { attempts++; throw new Error('disk unavailable'); });
  await first.initialize();
  const id = await first.enqueue([message()]);
  await eventually(async () => attempts > 0 ? true : undefined);
  await first.stop();
  assert.equal((await readdir(directory)).filter(name => name.endsWith('.json')).length, 1);
  let recovered = 0;
  const second = new WebhookQueue(directory, async m => { recovered++; return { messageId: m.id, duplicate: false, leads: [], reply: 'none' }; });
  await second.initialize();
  await eventually(async () => second.status(id)?.state === 'completed' ? true : undefined);
  await second.stop();
  assert.equal(recovered, 1);
  assert.equal((await readdir(directory)).filter(name => name.endsWith('.json')).length, 0);
}));
