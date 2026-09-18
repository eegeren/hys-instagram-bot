import ExcelJS from 'exceljs';
import { constants } from 'node:fs';
import { copyFile, readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { atomicWrite } from './files.js';
import { extractPhoneNumbers, normalizePhoneNumber } from './phone.js';
import type { IncomingMessage, LeadStore, ProcessResult } from './types.js';

export function resolveExcelPath(env = process.env, platform = process.platform): string {
  return resolve(env.EXCEL_FILE_PATH?.trim() || (platform === 'win32' || env.LOCAL_TEST_MODE === 'true'
    ? './data/HYS_Instagram_Leads.xlsx' : '/data/HYS_Instagram_Leads.xlsx'));
}
export const SHEET_NAME = 'Instagram Leads';
export const HEADERS = ['ID', 'Tarih', 'Saat', 'Instagram Kullanıcı Adı', 'Telefon', 'Mesaj', 'Durum', 'Arandı mı', 'Not', 'Meta Message ID'];
const LEDGER = '_Processed Messages';
const LEDGER_HEADERS = ['Meta Message ID', 'Timestamp', 'Phones'];
// All instances targeting the same file in this process share one queue.
const queues = new Map<string, Promise<unknown>>();

function dateTime(timestamp: number): [string, string] {
  const parts = new Intl.DateTimeFormat('tr-TR', { timeZone: 'Europe/Istanbul', day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit', hourCycle: 'h23' }).formatToParts(timestamp);
  const get = (type: string) => parts.find(part => part.type === type)?.value;
  return [`${get('day')}.${get('month')}.${get('year')}`, `${get('hour')}:${get('minute')}`];
}

export class ExcelLeadStore implements LeadStore {
  readonly filePath: string;
  constructor(filePath = resolveExcelPath()) { this.filePath = resolve(filePath); }

  private serialized<T>(work: () => Promise<T>): Promise<T> {
    const task = (queues.get(this.filePath) ?? Promise.resolve()).then(work);
    queues.set(this.filePath, task.then(() => undefined, () => undefined));
    return task;
  }

  private async load(): Promise<{ workbook: ExcelJS.Workbook; changed: boolean }> {
    const workbook = new ExcelJS.Workbook();
    let bytes: Buffer | undefined;
    try { bytes = await readFile(this.filePath); }
    catch (error) { if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error; }
    if (bytes) await workbook.xlsx.load(bytes as unknown as ExcelJS.Buffer);
    let changed = !bytes;
    let sheet = workbook.getWorksheet(SHEET_NAME);
    if (!sheet) {
      if (bytes) throw new Error('Expected Instagram Leads worksheet is missing');
      sheet = workbook.addWorksheet(SHEET_NAME);
      sheet.getRow(1).values = HEADERS;
    }
    // Migrate the earlier 11-column format, preserving all data in a backup first.
    if (sheet.getCell(1, 4).text === 'Instagram Kullanıcı ID') {
      const legacy = [...HEADERS.slice(0, 3), 'Instagram Kullanıcı ID', ...HEADERS.slice(3)];
      if (!legacy.every((header, i) => sheet!.getCell(1, i + 1).text === header)) throw new Error('Unexpected legacy workbook headers');
      await copyFile(this.filePath, `${this.filePath}.pre-railway.bak.xlsx`, constants.COPYFILE_EXCL)
        .catch((error: NodeJS.ErrnoException) => { if (error.code !== 'EEXIST') throw error; });
      sheet.spliceColumns(4, 1);
      changed = true;
    }
    if (!HEADERS.every((header, i) => sheet!.getCell(1, i + 1).text === header)) throw new Error('Unexpected Excel headers');
    sheet.columns.forEach((column, index) => { column.width = [8, 14, 8, 28, 18, 65, 12, 12, 40, 38][index] ?? 15; column.hidden = index === 9; });
    sheet.getRow(1).font = { bold: true };
    sheet.views = [{ state: 'frozen', ySplit: 1 }];
    let ledger = workbook.getWorksheet(LEDGER);
    if (!ledger) {
      ledger = workbook.addWorksheet(LEDGER, { state: 'veryHidden' });
      ledger.getRow(1).values = LEDGER_HEADERS;
      const ids = new Set<string>();
      sheet.eachRow((row, index) => {
        const id = row.getCell(10).text;
        if (index > 1 && id && !ids.has(id)) { ledger!.addRow([id, 0, row.getCell(5).text]); ids.add(id); }
      });
      changed = true;
    }
    if (!LEDGER_HEADERS.every((header, i) => ledger!.getCell(1, i + 1).text === header)) throw new Error('Unexpected message ledger headers');
    ledger.state = 'veryHidden';
    return { workbook, changed };
  }

  private async save(workbook: ExcelJS.Workbook): Promise<void> {
    await atomicWrite(this.filePath, Buffer.from(await workbook.xlsx.writeBuffer()));
  }

  initialize(): Promise<void> {
    return this.serialized(async () => {
      const { workbook, changed } = await this.load();
      if (changed) await this.save(workbook);
    });
  }

  processMessage(message: IncomingMessage, username: string): Promise<ProcessResult> {
    return this.serialized(async () => {
      const { workbook, changed } = await this.load();
      const sheet = workbook.getWorksheet(SHEET_NAME)!;
      const ledger = workbook.getWorksheet(LEDGER)!;
      const result: ProcessResult = { messageId: message.id, duplicate: false, leads: [], reply: 'none' };
      const latest = new Map<string, number>();
      let duplicate = false;
      ledger.eachRow((row, index) => {
        if (index === 1) return;
        if (row.getCell(1).text === message.id) duplicate = true;
        for (const phone of row.getCell(3).text.split(',')) latest.set(phone, Math.max(latest.get(phone) ?? 0, Number(row.getCell(2).value) || 0));
      });
      // Also inspect visible lead rows in case a workbook was imported manually.
      sheet.eachRow((row, index) => { if (index > 1 && row.getCell(10).text === message.id) duplicate = true; });
      if (duplicate) { if (changed) await this.save(workbook); return { ...result, duplicate: true }; }
      const phones = extractPhoneNumbers(message.text);
      const timestamp = message.timestamp ?? Date.now();
      const [date, time] = dateTime(timestamp);
      let maxId = 0;
      sheet.eachRow((row, index) => { if (index > 1) maxId = Math.max(maxId, Number(row.getCell(1).value) || 0); });
      for (const phone of phones) {
        let existing: ExcelJS.Row | undefined;
        sheet.eachRow((row, index) => { if (index > 1 && normalizePhoneNumber(row.getCell(5).text) === phone) existing = row; });
        if (existing) {
          const update = timestamp >= (latest.get(phone) ?? 0);
          if (update) {
            existing.getCell(6).value = message.text;
            existing.getCell(10).value = message.id;
            if (username) existing.getCell(4).value = `@${username.replace(/^@/, '')}`;
          }
          result.leads.push({ phone, action: update ? 'updated' : 'existing' });
        } else {
          sheet.addRow([++maxId, date, time, username ? `@${username.replace(/^@/, '')}` : '', phone, message.text, 'YENİ', 'HAYIR', '', message.id]);
          result.leads.push({ phone, action: 'inserted' });
        }
      }
      // Every message ID is committed together with all its phone rows, including repeat contacts.
      ledger.addRow([message.id, timestamp, phones.join(',')]);
      await this.save(workbook);
      if (result.leads.some(lead => lead.action !== 'inserted')) console.log('Existing lead contacted HYS again; contact recorded.');
      return result;
    });
  }
}

export function createExcelStore(filePath = resolveExcelPath()): ExcelLeadStore { return new ExcelLeadStore(filePath); }
