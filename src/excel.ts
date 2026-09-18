import ExcelJS from 'exceljs';
import { access, mkdir, rename } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import type { Lead, LeadStore } from './types.js';

export const WORKBOOK_PATH = resolve('data', 'HYS_Instagram_Leads.xlsx');
const SHEET_NAME = 'Instagram Leads';
const HEADERS = ['ID', 'Tarih', 'Saat', 'Instagram Kullanıcı ID', 'Instagram Kullanıcı Adı', 'Telefon', 'Mesaj', 'Durum', 'Arandı mı', 'Not', 'Meta Message ID'];

function formatDateTime(isoDate: string): { date: string; time: string } {
  const parts = new Intl.DateTimeFormat('tr-TR', {
    timeZone: 'Europe/Istanbul', day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit', hourCycle: 'h23',
  }).formatToParts(new Date(isoDate));
  const value = (type: Intl.DateTimeFormatPartTypes) => parts.find(part => part.type === type)?.value ?? '';
  return { date: `${value('day')}.${value('month')}.${value('year')}`, time: `${value('hour')}:${value('minute')}` };
}

function getWorksheet(workbook: ExcelJS.Workbook): ExcelJS.Worksheet {
  const worksheet = workbook.getWorksheet(SHEET_NAME) ?? workbook.addWorksheet(SHEET_NAME);
  const headers = HEADERS.map((_, index) => worksheet.getRow(1).getCell(index + 1).value);
  if (headers.every(header => header === null)) {
    const headerRow = worksheet.getRow(1);
    headerRow.values = HEADERS;
    headerRow.font = { bold: true };
    worksheet.views = [{ state: 'frozen', ySplit: 1 }];
    worksheet.columns = [
      { width: 8 }, { width: 14 }, { width: 8 }, { width: 24 }, { width: 28 }, { width: 18 }, { width: 60 }, { width: 12 }, { width: 12 }, { width: 40 }, { width: 38, hidden: true },
    ];
  } else if (!HEADERS.every((header, index) => headers[index] === header)) {
    throw new Error('Unexpected Excel worksheet headers');
  }
  worksheet.getColumn(11).hidden = true;
  return worksheet;
}

export class ExcelLeadStore implements LeadStore {
  private queue: Promise<unknown> = Promise.resolve();

  constructor(private readonly filePath = WORKBOOK_PATH) {}

  insertIfNew(lead: Lead, metaMessageId: string): Promise<boolean> {
    const task = this.queue.then(() => this.writeLead(lead, metaMessageId));
    this.queue = task.then(() => undefined, () => undefined);
    return task;
  }

  private async writeLead(lead: Lead, metaMessageId: string): Promise<boolean> {
    await mkdir(dirname(this.filePath), { recursive: true });
    const workbook = new ExcelJS.Workbook();
    try { await access(this.filePath); await workbook.xlsx.readFile(this.filePath); }
    catch (error: unknown) {
      if (!(error instanceof Error && 'code' in error && error.code === 'ENOENT')) throw error;
    }
    const worksheet = getWorksheet(workbook);
    for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      if (String(row.getCell(11).value ?? '') === metaMessageId || String(row.getCell(6).value ?? '') === lead.phone) return false;
    }
    const { date, time } = formatDateTime(lead.created_at);
    worksheet.addRow([
      worksheet.rowCount, date, time, lead.instagram_user_id, lead.instagram_username ? `@${lead.instagram_username.replace(/^@/, '')}` : '', lead.phone,
      lead.original_message, 'YENİ', 'HAYIR', '', metaMessageId,
    ]);
    // ExcelJS writes to a complete temporary workbook first. The serialized queue prevents concurrent readers/writers.
    const temporaryPath = `${this.filePath}.${process.pid}.${Date.now()}.tmp`;
    await workbook.xlsx.writeFile(temporaryPath);
    await rename(temporaryPath, this.filePath);
    return true;
  }
}

export function createExcelStore(filePath = WORKBOOK_PATH): LeadStore {
  return new ExcelLeadStore(filePath);
}
