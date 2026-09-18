export interface Lead {
  created_at: string;
  instagram_user_id: string;
  instagram_username: string;
  phone: string;
  original_message: string;
  status: 'NEW';
}

export interface IncomingMessage { id: string; senderId: string; text: string }
export interface LeadStore { insertIfNew(lead: Lead, metaMessageId: string): Promise<boolean> }
export interface InstagramClient {
  username(userId: string): Promise<string>;
  reply(userId: string): Promise<void>;
}
export function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
