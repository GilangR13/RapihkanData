export const RULES = [
  { value: 'space', label: 'Rapikan spasi' },
  { value: 'title', label: 'Kapital awal kata' },
  { value: 'upper', label: 'HURUF BESAR' },
  { value: 'lower', label: 'huruf kecil' },
  { value: 'keep', label: 'Biarkan isi' },
];

const PRESERVE_UPPER = new Set(['PT', 'CV', 'RT', 'RW', 'DKI', 'SD', 'SMP', 'SMA', 'SMK', 'NIK', 'NISN']);

export function normalizeSpace(value) {
  return value.replace(/[\u00a0\u2007\u202f\t\r\n ]+/g, ' ').trim();
}

export function titleCase(value) {
  return value.toLocaleLowerCase('id-ID').replace(/(^|[\s/([{-])([^\s/()[\]{}-]+)/gu, (match, prefix, word) => {
    const upper = word.toLocaleUpperCase('id-ID');
    if (PRESERVE_UPPER.has(upper)) return prefix + upper;
    return prefix + word.charAt(0).toLocaleUpperCase('id-ID') + word.slice(1);
  });
}

export function cleanString(value, rule = 'space') {
  if (rule === 'keep') return value;
  const spaced = normalizeSpace(value);
  if (rule === 'upper') return spaced.toLocaleUpperCase('id-ID');
  if (rule === 'lower') return spaced.toLocaleLowerCase('id-ID');
  if (rule === 'title') return titleCase(spaced);
  return spaced;
}

export function isLeadingZeroText(value) {
  return typeof value === 'string' && /^0\d/.test(value);
}

export function paddedNumberText(value, format) {
  if (!Number.isSafeInteger(value) || value < 0 || typeof format !== 'string') return null;
  if (!/^0{2,}$/.test(format)) return null;
  const result = String(value).padStart(format.length, '0');
  return result.startsWith('0') ? result : null;
}

export function transformCell(value, rule = 'space', numFmt = '') {
  if (typeof value === 'string') {
    const next = cleanString(value, rule);
    return { next, changed: next !== value, textFormat: isLeadingZeroText(next), reason: next !== value ? 'Format teks dirapikan' : '' };
  }
  if (typeof value === 'number') {
    const padded = paddedNumberText(value, numFmt);
    if (padded !== null) {
      return { next: padded, changed: true, textFormat: true, reason: 'Nol awal dari format Excel disimpan sebagai teks' };
    }
  }
  return { next: value, changed: false, textFormat: false, reason: '' };
}
