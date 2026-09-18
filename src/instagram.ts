import { isRecord, type IncomingMessage, type InstagramClient } from './types.js';

export const AUTO_REPLY = "Merhaba 👋 İletişim bilgilerinizi aldık. Ekibimiz en kısa sürede sizinle iletişime geçecektir. HYS'yi tercih ettiğiniz için teşekkür ederiz.";

export function parseMessages(payload: unknown, accountId: string, localTestMode = false): IncomingMessage[] {
  if (!isRecord(payload) || payload.object !== 'instagram' || !Array.isArray(payload.entry)) return [];
  const messages: IncomingMessage[] = [];
  for (const entry of payload.entry) {
    if (!isRecord(entry) || (!localTestMode && entry.id !== accountId) || !Array.isArray(entry.messaging)) continue;
    for (const event of entry.messaging) {
      if (!isRecord(event) || !isRecord(event.sender) || !isRecord(event.message)) continue;
      const { message, sender, recipient } = event;
      if ((!localTestMode && (!isRecord(recipient) || recipient.id !== accountId)) || sender.id === accountId || message.is_echo || message.is_deleted) continue;
      if (typeof sender.id !== 'string' || !/^\d+$/.test(sender.id)) continue;
      if (typeof message.mid !== 'string' || !message.mid || typeof message.text !== 'string' || !message.text.trim()) continue;
      messages.push({ id: message.mid, senderId: sender.id, text: message.text });
    }
  }
  return messages;
}

export function createInstagramClient(token: string, accountId: string, version: string): InstagramClient {
  const base = `https://graph.instagram.com/${version}`;
  async function request(path: string, body?: unknown): Promise<unknown> {
    const response = await fetch(`${base}/${path}`, {
      method: body ? 'POST' : 'GET',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      ...(body ? { body: JSON.stringify(body) } : {}),
      signal: AbortSignal.timeout(8000),
    });
    if (!response.ok) throw new Error('Instagram request failed');
    return response.json();
  }
  return {
    async username(userId) {
      try {
        const data = await request(`${encodeURIComponent(userId)}?fields=username`);
        return isRecord(data) && typeof data.username === 'string' ? data.username : '';
      } catch {
        console.warn('Instagram profile lookup unavailable; saving lead without username.');
        return '';
      }
    },
    async reply(userId) {
      await request(`${accountId}/messages`, { recipient: { id: userId }, message: { text: AUTO_REPLY } });
    },
  };
}
