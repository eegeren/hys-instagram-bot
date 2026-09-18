import dotenv from 'dotenv';
import express, { type ErrorRequestHandler } from 'express';
import { createHmac, timingSafeEqual } from 'node:crypto';
import { pathToFileURL } from 'node:url';
import { dirname, join } from 'node:path';
import { createInstagramClient, parseMessages } from './instagram.js';
import { extractPhoneNumbers } from './phone.js';
import { createExcelStore, resolveExcelPath } from './excel.js';
import { WebhookQueue } from './queue.js';
import type { IncomingMessage, InstagramClient, LeadStore, ProcessResult } from './types.js';

interface Dependencies {
  verifyToken: string; appSecret: string; accountId: string; store: LeadStore;
  instagram: InstagramClient; localTestMode?: boolean; queueDirectory: string;
}

export async function createApp(deps?: Dependencies) {
  const app = express();
  if (deps?.localTestMode && process.env.NODE_ENV === 'production') throw new Error('LOCAL_TEST_MODE is forbidden in production');
  app.disable('x-powered-by');
  if (deps?.localTestMode) app.use((req, res, next) => {
    if (!['127.0.0.1', '::1', '::ffff:127.0.0.1'].includes(req.socket.remoteAddress ?? '') || req.get('x-forwarded-for')) { res.sendStatus(403); return; }
    res.setHeader('X-HYS-Local-Test', 'true');
    next();
  });
  let queue: WebhookQueue | undefined;
  if (deps) {
    await deps.store.initialize();
    const processMessage = async (message: IncomingMessage): Promise<ProcessResult> => {
      const phones = extractPhoneNumbers(message.text);
      const username = phones.length && !deps.localTestMode ? await deps.instagram.username(message.senderId) : '';
      const result = await deps.store.processMessage(message, username);
      if (deps.localTestMode && !result.duplicate) for (const lead of result.leads) {
        const cleanText = message.text.replace(/[\u0000-\u001f\u007f-\u009f]/g, ' ');
        console.log(`--- HYS LEAD ---\nInstagram ID: ${message.senderId}\nPhone: ${lead.phone}\nMessage: ${cleanText}\nStatus: ${lead.action}\n-----------`);
      }
      // Message IDs are already persisted. Replays never send another reply.
      if (result.leads.some(lead => lead.action === 'inserted')) {
        if (deps.localTestMode) {
          console.log('[LOCAL TEST] Instagram reply would be sent.');
          result.reply = 'simulated';
        } else {
          try { await deps.instagram.reply(message.senderId); result.reply = 'sent'; }
          catch { result.reply = 'failed'; console.error('Lead saved but automatic reply failed; manual follow-up required.'); }
        }
      }
      return result;
    };
    queue = new WebhookQueue(deps.queueDirectory, processMessage);
    await queue.initialize();
  }
  app.locals.stop = async () => { await queue?.stop(); };
  app.get('/health', (_req, res) => { res.json({ status: 'ok' }); });
  app.get('/webhook', (req, res) => {
    if (!deps) { res.sendStatus(503); return; }
    if (deps.verifyToken && req.query['hub.mode'] === 'subscribe' && req.query['hub.verify_token'] === deps.verifyToken && typeof req.query['hub.challenge'] === 'string') {
      res.status(200).type('text/plain').send(req.query['hub.challenge']);
    } else res.sendStatus(403);
  });
  // Local-only completion receipts let tests observe asynchronous persistence and simulated replies.
  if (deps?.localTestMode) app.get('/local-test/jobs/:id', (req, res) => {
    const status = queue?.status(req.params.id);
    if (!status) { res.sendStatus(404); return; }
    res.json(status);
  });
  app.post('/webhook', express.raw({ type: 'application/json', limit: '256kb', inflate: false }), async (req, res) => {
    if (!deps || !queue) { res.sendStatus(503); return; }
    if (!Buffer.isBuffer(req.body)) { res.sendStatus(415); return; }
    if (!deps.localTestMode) {
      const signature = req.get('x-hub-signature-256');
      if (!signature || !/^sha256=[a-f\d]{64}$/i.test(signature)) { res.sendStatus(401); return; }
      const expected = createHmac('sha256', deps.appSecret).update(req.body).digest();
      if (!timingSafeEqual(expected, Buffer.from(signature.slice(7), 'hex'))) { res.sendStatus(401); return; }
    }
    let payload: unknown;
    try { payload = JSON.parse(req.body.toString('utf8')); } catch { res.sendStatus(400); return; }
    const messages = parseMessages(payload, deps.accountId, deps.localTestMode);
    if (!messages.length) { res.sendStatus(200); return; }
    try {
      // Only wait for a small durable job file, not Excel or the Meta API.
      const jobId = await queue.enqueue(messages);
      if (deps.localTestMode) res.json({ accepted: true, jobId });
      else res.sendStatus(200);
    } catch { console.error('Unable to persist webhook job; Meta should retry.'); res.sendStatus(503); }
  });
  const errorHandler: ErrorRequestHandler = (error: unknown, _req, res, _next) => {
    const status = typeof error === 'object' && error !== null && 'status' in error ? error.status : 500;
    res.sendStatus(status === 413 ? 413 : status === 415 ? 415 : 500);
  };
  app.use(errorHandler);
  return app;
}

export async function start() {
  dotenv.config({ quiet: true });
  const localTestMode = process.env.LOCAL_TEST_MODE === 'true';
  if (localTestMode && process.env.NODE_ENV === 'production') throw new Error('LOCAL_TEST_MODE is forbidden in production');
  const required = ['META_VERIFY_TOKEN', 'META_APP_SECRET', 'META_ACCESS_TOKEN', 'META_INSTAGRAM_ACCOUNT_ID'] as const;
  const missing = required.filter(key => !process.env[key]?.trim());
  if (missing.length && process.env.NODE_ENV === 'production') throw new Error(`Missing configuration: ${missing.join(', ')}`);
  const port = Number(process.env.PORT ?? 3000);
  if (!Number.isInteger(port) || port < 1 || port > 65535) throw new Error('Invalid PORT');
  const version = process.env.META_API_VERSION ?? 'v22.0';
  if (!/^v\d+\.0$/.test(version)) throw new Error('Invalid META_API_VERSION');
  const filePath = resolveExcelPath();
  const store = createExcelStore(filePath);
  // Initialize/migrate only at runtime, when the Railway volume is mounted.
  await store.initialize();
  let deps: Dependencies | undefined;
  const queueDirectory = join(dirname(filePath), '.hys-webhook-queue');
  if (localTestMode) {
    deps = { localTestMode: true, verifyToken: '', appSecret: '', accountId: '', store, queueDirectory,
      instagram: { async username() { return ''; }, async reply() {} } };
    console.log(`[LOCAL TEST] Meta calls disabled. Excel file: ${filePath}`);
  } else if (!missing.length) {
    const env = process.env as Record<typeof required[number], string>;
    if (!/^\d+$/.test(env.META_INSTAGRAM_ACCOUNT_ID)) throw new Error('Invalid META_INSTAGRAM_ACCOUNT_ID');
    deps = { verifyToken: env.META_VERIFY_TOKEN, appSecret: env.META_APP_SECRET, accountId: env.META_INSTAGRAM_ACCOUNT_ID,
      store, queueDirectory, instagram: createInstagramClient(env.META_ACCESS_TOKEN, env.META_INSTAGRAM_ACCOUNT_ID, version) };
  } else console.warn(`Setup mode: webhooks disabled. Configure: ${missing.join(', ')}`);
  const app = await createApp(deps);
  const server = app.listen(port, '0.0.0.0', () => console.log(`HYS server listening on port ${port}`));
  for (const signal of ['SIGINT', 'SIGTERM'] as const) process.once(signal, () => {
    const timeout = setTimeout(() => process.exit(1), 25000);
    timeout.unref();
    server.close(() => {
      void app.locals.stop().then(() => process.exit(0), () => process.exit(1));
    });
  });
  return server;
}
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  start().catch((error: unknown) => {
    const code = error && typeof error === 'object' && 'code' in error ? error.code : undefined;
    if (code === 'EPERM' || code === 'EACCES' || code === 'EBUSY') {
      console.error('Startup blocked by file permissions or a workbook lock. Close Excel and verify volume write access.');
    } else console.error('Startup failed. Check configuration, workbook format and volume permissions.');
    process.exitCode = 1;
  });
}
