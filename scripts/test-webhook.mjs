import dotenv from 'dotenv';
import ExcelJS from 'exceljs';
import { randomUUID } from 'node:crypto';
import { resolve } from 'node:path';

dotenv.config({ quiet: true });
const base = `http://127.0.0.1:${process.env.PORT || 3000}`;
const path = resolve(process.env.EXCEL_FILE_PATH || './data/HYS_Instagram_Leads.xlsx');
let failures = 0;
function check(label, passed, detail = '') {
  console.log(`${passed ? 'PASS' : 'FAIL'} ${label}${detail ? ` — ${detail}` : ''}`);
  if (!passed) failures++;
}
async function workbook() { const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(path); return wb; }
async function send(id, text) {
  const response = await fetch(`${base}/webhook`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ object: 'instagram', entry: [{ messaging: [{ sender: { id: '123456789' }, message: { mid: id, text } }] }] }), signal: AbortSignal.timeout(5000) });
  if (response.status !== 200) throw new Error(`Webhook HTTP ${response.status}`);
  const { jobId } = await response.json();
  if (!jobId) throw new Error('Expected durable queue receipt; restart the updated local server.');
  const deadline = Date.now() + 30000;
  while (Date.now() < deadline) {
    const status = await (await fetch(`${base}/local-test/jobs/${jobId}`, { signal: AbortSignal.timeout(5000) })).json();
    if (status.state === 'completed') return status.results[0];
    await new Promise(resolve => setTimeout(resolve, 100));
  }
  throw new Error('Timed out waiting for persisted lead');
}

try {
  const health = await fetch(`${base}/health`, { signal: AbortSignal.timeout(5000) });
  if (health.headers.get('x-hys-local-test') !== 'true') throw new Error('Start npm run dev with LOCAL_TEST_MODE=true.');
  check('Local server health', health.status === 200);
  const id = `HYS_TEST_${randomUUID()}`;
  const text = 'Merhaba, 0532 123 45 67 numarasından beni arar mısınız?';
  const first = await send(id, text);
  check('Fake webhook acknowledged and processed', true, 'HTTP 200');
  check('Phone normalized', first.leads.some(lead => lead.phone === '+905321234567'), '+905321234567');
  let wb = await workbook();
  let sheet = wb.getWorksheet('Instagram Leads');
  let matching = [];
  sheet.eachRow((row, index) => { if (index > 1 && row.getCell(5).text === '+905321234567') matching.push(row); });
  check('Real Excel lead exists with latest message', matching.length === 1 && matching[0].getCell(6).text === text && matching[0].getCell(10).text === id, path);
  check('Meta message column hidden', sheet.getColumn(10).hidden);
  const count = sheet.rowCount;
  if (first.leads.some(lead => lead.action === 'inserted')) {
    check('New-lead automatic response simulated', first.reply === 'simulated');
  } else {
    check('Existing phone updated without another reply', first.reply === 'none');
    // Preserve existing data. A separate synthetic new phone proves the reply path on repeat test runs.
    const used = new Set(); sheet.eachRow(row => used.add(row.getCell(5).text));
    let national;
    do { national = `5${String(Math.floor(Math.random() * 1e9)).padStart(9, '0')}`; } while (used.has(`+90${national}`));
    const probe = await send(`HYS_REPLY_PROBE_${randomUUID()}`, `Yerel otomatik yanıt testi ${national}`);
    check('New-lead automatic response simulated', probe.reply === 'simulated' && probe.leads[0]?.action === 'inserted', 'separate synthetic lead retained');
  }
  const beforeReplay = await workbook();
  const rowsBeforeReplay = beforeReplay.getWorksheet('Instagram Leads').rowCount;
  const replay = await send(id, text);
  check('Exact same Meta message ID ignored', replay.duplicate === true && replay.leads.length === 0);
  wb = await workbook(); sheet = wb.getWorksheet('Instagram Leads');
  check('Replay creates no duplicate Excel row', sheet.rowCount === rowsBeforeReplay && count >= 2);
  check('Replay sends no automatic response', replay.reply === 'none');
  check('Original message ID retained in workbook ledger', wb.getWorksheet('_Processed Messages').getColumn(1).values.includes(id));
} catch (error) { check('Webhook test completed', false, error instanceof Error ? error.message : 'Unknown failure'); }
console.log(failures ? `FAIL: ${failures} check(s)` : 'PASS: all webhook checks');
process.exitCode = failures ? 1 : 0;
