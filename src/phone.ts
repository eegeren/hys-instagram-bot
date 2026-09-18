/** Full candidate validation prevents matching the tail of an invalid longer number. */
export function normalizePhoneNumber(value: string): string | null {
  if (!/^[+\d\s()-]+$/.test(value)) return null;
  const compact = value.replace(/[\s()-]/g, '');
  const match = /^(?:\+90|0)?(5\d{9})$/.exec(compact);
  return match ? `+90${match[1]}` : null;
}

export function extractPhoneNumbers(text: string): string[] {
  const phones = new Set<string>();
  // Separators other than spaces, hyphens and parentheses end a candidate.
  for (const match of text.matchAll(/\+?\d[\d ()-]*\d/g)) {
    const phone = normalizePhoneNumber(match[0]);
    if (phone) phones.add(phone);
  }
  return [...phones];
}
