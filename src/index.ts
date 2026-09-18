import dotenv from 'dotenv';
import express, { type ErrorRequestHandler } from 'express';
import { createHmac, timingSafeEqual } from 'node:crypto';
import { pathToFileURL } from 'node:url';
import { createInstagramClient, parseMessages } from './instagram.js';
import { extractPhoneNumbers } from './phone.js';
import { createExcelStore } from './excel.js';
import type { InstagramClient, LeadStore } from './types.js';

interface Dependencies { verifyToken: string; appSecret: string; accountId: string; store: LeadStore; instagram: InstagramClient; localTestMode?: boolean }

export function createApp(deps?: Dependencies) {
  const app = express();
  if (deps?.localTestMode && process.env.NODE_ENV === 'production') throw new Error('LOCAL_TEST_MODE is forbidden in production');
  app.disable('x-powered-by');
  if (deps?.localTestMode) app.use((req, res, next) => {
    if (!['127.0.0.1', '::1', '::ffff:127.0.0.1'].includes(req.socket.remoteAddress ?? '')) { res.sendStatus(403); return; }
    res.setHeader('X-HYS-Local-Test', 'true');
    next();
  });
  app.get('/health', (_req, res) => { res.json({ status: 'ok' }); });
  app.get('/webhook', (req, res) => {
    if (!deps) { res.sendStatus(503); return; }
    if (req.query['hub.mode'] === 'subscribe' && req.query['hub.verify_token'] === deps.verifyToken && typeof req.query['hub.challenge'] === 'string') {
      res.status(200).type('text/plain').send(req.query['hub.challenge']);
    } else res.sendStatus(403);
  });

  // One process only: serialize the complete read/check/append/send transaction.
  let tail = Promise.resolve();
  let pending = 0;
  app.post('/webhook', express.raw({ type: 'application/json', limit: '256kb', inflate: false }), async (req, res) => {
    if (!deps) { res.sendStatus(503); return; }
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
    if (pending >= 50) { res.sendStatus(503); return; }
    pending++;
    const localResults: { phone: string; inserted: boolean; replyWouldBeSent: boolean }[] = [];
    const work = tail.then(async () => {
      for (const message of messages) {
        const phones = extractPhoneNumbers(message.text);
        if (!phones.length) continue;
        const username = deps.localTestMode ? '' : await deps.instagram.username(message.senderId);
        let inserted = false;
        for (const phone of phones) {
          if (deps.localTestMode) {
            const cleanText = message.text.replace(/[\u0000-\u001f\u007f-\u009f]/g, ' ');
            console.log(`--- HYS LEAD ---\nInstagram ID: ${message.senderId}\nPhone: ${phone}\nMessage: ${cleanText}\nStatus: NEW\n-----------`);
          }
          const created = await deps.store.insertIfNew({ created_at: new Date().toISOString(), instagram_user_id: message.senderId, instagram_username: username, phone, original_message: message.text, status: 'NEW' }, message.id);
          if (deps.localTestMode) localResults.push({ phone, inserted: created, replyWouldBeSent: created });
          inserted = created || inserted;
        }
        if (inserted) {
          try {
            if (deps.localTestMode) console.log('[LOCAL TEST] Instagram reply would be sent.');
            else await deps.instagram.reply(message.senderId);
          }
          catch { console.error('Lead saved but automatic reply failed; manual follow-up required.'); }
        }
      }
    });
    tail = work.catch(() => {});
    try {
      await work;
      if (deps.localTestMode) res.status(200).json({ mode: 'local-test', leads: localResults });
      else res.sendStatus(200);
    }
    catch { console.error('Lead persistence failed; webhook may be retried.'); res.sendStatus(503); }
    finally { pending--; }
  });
  const errorHandler: ErrorRequestHandler = (error: unknown, _req, res, _next) => {
    const status = typeof error === 'object' && error !== null && 'status' in error ? error.status : 500;
    res.sendStatus(status === 413 ? 413 : status === 415 ? 415 : 500);
  };
  app.use(errorHandler);
  return app;
}

export function start() {
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
  let deps: Dependencies | undefined;
  if (localTestMode) {
    deps = { localTestMode: true, verifyToken: '', appSecret: '', accountId: '',
      store: createExcelStore(),
      instagram: { async username() { return ''; }, async reply() {} } };
    console.log('[LOCAL TEST] Meta calls disabled. Leads will be written to data/HYS_Instagram_Leads.xlsx. Do not expose through a tunnel or proxy.');
  } else if (!missing.length) {
    const env = process.env as Record<typeof required[number], string>;
    if (!/^\d+$/.test(env.META_INSTAGRAM_ACCOUNT_ID)) throw new Error('Invalid META_INSTAGRAM_ACCOUNT_ID');
    deps = { verifyToken: env.META_VERIFY_TOKEN, appSecret: env.META_APP_SECRET, accountId: env.META_INSTAGRAM_ACCOUNT_ID,
      store: createExcelStore(),
      instagram: createInstagramClient(env.META_ACCESS_TOKEN, env.META_INSTAGRAM_ACCOUNT_ID, version) };
  } else console.warn(`Setup mode: webhooks disabled. Configure: ${missing.join(', ')}`);
  const server = createApp(deps).listen(port, localTestMode ? '127.0.0.1' : '0.0.0.0', () => console.log(`HYS server listening on port ${port}`));
  for (const signal of ['SIGINT', 'SIGTERM'] as const) process.once(signal, () => {
    server.close(() => process.exit(0));
    setTimeout(() => process.exit(1), 30000).unref();
  });
  return server;
}
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) start();
