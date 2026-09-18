export interface IncomingMessage { id: string; senderId: string; text: string; timestamp?: number }
export interface LeadResult { phone: string; action: 'inserted' | 'updated' | 'existing' }
export interface ProcessResult { messageId: string; duplicate: boolean; leads: LeadResult[]; reply: 'none' | 'simulated' | 'sent' | 'failed' }
export interface LeadStore {
  initialize(): Promise<void>;
  processMessage(message: IncomingMessage, username: string): Promise<ProcessResult>;
}
export interface InstagramClient {
  username(userId: string): Promise<string>;
  reply(userId: string): Promise<void>;
}
export function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
