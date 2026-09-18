import { randomUUID } from 'node:crypto';
import { mkdir, readFile, readdir, unlink } from 'node:fs/promises';
import { join } from 'node:path';
import { atomicWrite } from './files.js';
import { isRecord, type IncomingMessage, type ProcessResult } from './types.js';

interface Job { id: string; messages: IncomingMessage[] }
export type JobStatus = { state: 'pending' | 'retrying' } | { state: 'completed'; results: ProcessResult[] };

/** Small durable file spool. Leads and processed IDs remain exclusively in Excel. */
export class WebhookQueue {
  private jobs: Job[] = [];
  private results = new Map<string, JobStatus>();
  private running?: Promise<void>;
  private timer?: NodeJS.Timeout;
  private stopping = false;
  private enqueueTail: Promise<unknown> = Promise.resolve();

  constructor(private directory: string, private processMessage: (message: IncomingMessage) => Promise<ProcessResult>) {}

  async initialize(): Promise<void> {
    await mkdir(this.directory, { recursive: true });
    for (const name of (await readdir(this.directory)).filter(name => name.endsWith('.json')).sort()) {
      const value: unknown = JSON.parse(await readFile(join(this.directory, name), 'utf8'));
      if (!isRecord(value) || typeof value.id !== 'string' || `${value.id}.json` !== name || !Array.isArray(value.messages)
        || !value.messages.every(m => isRecord(m) && typeof m.id === 'string' && typeof m.senderId === 'string' && typeof m.text === 'string'
          && typeof m.timestamp === 'number' && Number.isFinite(m.timestamp))) throw new Error('Invalid saved webhook job; inspect volume');
      this.jobs.push(value as unknown as Job);
    }
    this.kick();
  }

  async enqueue(messages: IncomingMessage[]): Promise<string> {
    const operation = this.enqueueTail.then(async () => {
      if (this.stopping || this.jobs.length >= 500) throw new Error('Webhook queue unavailable');
      const job = { id: `${Date.now()}-${randomUUID()}`, messages: messages.map(message => ({ ...message, timestamp: message.timestamp ?? Date.now() })) };
      await atomicWrite(join(this.directory, `${job.id}.json`), JSON.stringify(job));
      this.jobs.push(job);
      this.kick();
      return job.id;
    });
    this.enqueueTail = operation.catch(() => {});
    return operation;
  }

  status(id: string): JobStatus | undefined {
    return this.results.get(id) ?? (this.jobs.some(job => job.id === id) ? { state: 'pending' } : undefined);
  }

  private kick(): void {
    if (this.running || this.stopping || this.timer) return;
    this.running = this.drain().finally(() => { this.running = undefined; });
  }

  private async drain(): Promise<void> {
    while (!this.stopping && this.jobs.length) {
      const job = this.jobs[0];
      try {
        const results: ProcessResult[] = [];
        for (const message of job.messages) results.push(await this.processMessage(message));
        await unlink(join(this.directory, `${job.id}.json`));
        this.jobs.shift();
        this.results.set(job.id, { state: 'completed', results });
        if (this.results.size > 500) this.results.delete(this.results.keys().next().value!);
      } catch {
        console.error('Webhook processing failed; durable job retained for retry.');
        this.results.set(job.id, { state: 'retrying' });
        this.timer = setTimeout(() => { this.timer = undefined; this.kick(); }, 10000);
        this.timer.unref();
        return;
      }
    }
  }

  async stop(): Promise<void> {
    this.stopping = true;
    clearTimeout(this.timer);
    await this.enqueueTail;
    await this.running;
    // Queued jobs remain on the volume and resume on next startup.
  }
}
