import regions from './data/regions.json' with { type: 'json' };
import { cleanString, normalizeSpace, titleCase } from './cleaner.js';

const MONTHS = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
const REGIONS = [...new Set([...regions.provinsi, ...regions.kabupatenKota].map(item => item.nama.trim()))]
  .map(name => ({ key: name.toLocaleLowerCase('id-ID'), name: titleCase(name) }));
const EXACT = new Map(REGIONS.map(region => [region.key, region.name]));
const REGION_CACHE = new Map();

export function parseCorrectionDictionary(text) {
  const entries = [];
  for (const [index, line] of text.split(/\r?\n/).entries()) {
    if (!line.trim()) continue;
    const separator = line.includes('=>') ? '=>' : '=';
    const parts = line.split(separator);
    if (parts.length !== 2 || !parts[0].trim() || !parts[1].trim()) {
      throw new Error(`Baris kamus ${index + 1} harus berbentuk salah = benar.`);
    }
    const wrong = parts[0].trim();
    const right = parts[1].trim();
    if (wrong.length > 80 || right.length > 80) throw new Error(`Baris kamus ${index + 1} terlalu panjang.`);
    entries.push({ wrong, right });
  }
  if (entries.length > 200) throw new Error('Kamus maksimal 200 pasangan.');
  return entries;
}

export function applyCustomCorrections(value, entries = []) {
  let result = value;
  for (const { wrong, right } of entries) {
    const escaped = wrong.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const expression = new RegExp(`(?<![\\p{L}\\p{N}])${escaped}(?![\\p{L}\\p{N}])`, 'giu');
    result = result.replace(expression, () => right);
  }
  return result;
}

function distanceAtMostTwo(a, b) {
  if (Math.abs(a.length - b.length) > 2) return 3;
  let previous = Array.from({ length: b.length + 1 }, (_, index) => index);
  for (let row = 1; row <= a.length; row++) {
    const current = [row];
    for (let column = 1; column <= b.length; column++) {
      current[column] = Math.min(
        current[column - 1] + 1,
        previous[column] + 1,
        previous[column - 1] + (a[row - 1] === b[column - 1] ? 0 : 1),
      );
    }
    previous = current;
  }
  return previous[b.length];
}

export function correctRegionName(input) {
  const spaced = normalizeSpace(input);
  const prefixMatch = spaced.match(/^(Kabupaten|Kab\.|Kota|Provinsi|Prov\.)\s+(.+)$/i);
  const prefix = prefixMatch ? prefixMatch[1] : '';
  const place = prefixMatch ? prefixMatch[2] : spaced;
  const key = place.toLocaleLowerCase('id-ID');
  const formattedPrefix = prefix ? (prefix.toLocaleLowerCase('id-ID').startsWith('kab') ? 'Kab. ' : prefix.toLocaleLowerCase('id-ID').startsWith('prov') ? 'Provinsi ' : 'Kota ') : '';
  if (EXACT.has(key)) return { value: formattedPrefix + EXACT.get(key), corrected: false, suggestion: '' };
  if (REGION_CACHE.has(key)) {
    const cached = REGION_CACHE.get(key);
    return { ...cached, value: formattedPrefix + cached.value };
  }
  const matches = [];
  if (key.length >= 5) {
    for (const region of REGIONS) {
      if (Math.abs(key.length - region.key.length) > 2) continue;
      if (key[0] !== region.key[0]) continue;
      const distance = distanceAtMostTwo(key, region.key);
      if (distance <= 2) matches.push({ ...region, distance });
    }
  }
  matches.sort((a, b) => a.distance - b.distance || a.name.localeCompare(b.name, 'id-ID'));
  const best = matches[0];
  const uniqueBest = best && !matches.some((candidate, index) => index > 0 && candidate.distance === best.distance);
  let result;
  if (best && best.distance === 1 && uniqueBest) {
    result = { value: best.name, corrected: true, suggestion: '' };
  } else {
    result = { value: titleCase(place), corrected: false, suggestion: best && uniqueBest ? best.name : '' };
  }
  REGION_CACHE.set(key, result);
  return { ...result, value: formattedPrefix + result.value };
}

export function correctRegionText(input) {
  let corrected = false;
  let suggestion = '';
  const check = candidate => {
    const result = correctRegionName(candidate);
    if (result.corrected) {
      corrected = true;
      return result.value;
    }
    if (result.suggestion && !suggestion) suggestion = `Nama wilayah “${candidate}” mungkin “${result.suggestion}”; perlu diperiksa.`;
    return candidate;
  };
  // Nama tempat dalam TTL atau alamat berkoma diperiksa terpisah dari teks setelah koma.
  let value = input.split(/(,\s*)/).map((part, index) => index % 2 ? part : check(part)).join('');
  // Nama wilayah setelah penanda administrasi juga dapat muncul di tengah alamat tanpa koma.
  value = value.replace(/\b(Kabupaten|Kab\.|Kota|Provinsi|Prov\.)\s+([\p{L}]+(?:\s+[\p{L}]+){0,3})/giu, (whole, prefix, tail) => {
    const words = tail.split(/\s+/);
    for (let length = words.length; length >= 1; length--) {
      const candidate = `${prefix} ${words.slice(0, length).join(' ')}`;
      const result = correctRegionName(candidate);
      if (result.corrected) {
        corrected = true;
        return result.value + (words.length > length ? ` ${words.slice(length).join(' ')}` : '');
      }
    }
    return whole;
  });
  return { value, corrected, warning: suggestion };
}

export function formatIndonesianAddress(input, caseRule = 'title', corrections = []) {
  let value = normalizeSpace(applyCustomCorrections(input, corrections));
  value = value
    .replace(/\b(kp)\s*\.?\s*(?=[a-z])/gi, 'Kp. ')
    .replace(/\b(jl|jln)\s*\.?\s*(?=[a-z])/gi, 'Jl. ')
    .replace(/\b(gg)\s*\.?\s*(?=[a-z])/gi, 'Gg. ')
    .replace(/\b(no)\s*\.?\s*(?=\d)/gi, 'No. ')
    .replace(/\bRT\s*\.?\s*(\d+)\s*\/\s*(?:RW\s*\.?\s*)?(\d+)\b/gi, 'RT $1/RW $2')
    .replace(/\bRT\s*\.?\s*(\d+)\b/gi, 'RT $1')
    .replace(/\bRW\s*\.?\s*(\d+)\b/gi, 'RW $1')
    .replace(/\bRT\s+(\d+)\s+RW\s+(\d+)\b/gi, 'RT $1/RW $2');
  return cleanString(value, caseRule);
}

export function cleanConfiguredText(input, { caseRule = 'space', address = false, region = false, corrections = [] } = {}) {
  let value = input;
  const reasons = [];
  let warning = '';
  if (address) {
    value = formatIndonesianAddress(value, caseRule, corrections);
    if (value !== input) reasons.push('Tata alamat Indonesia dan kapitalisasi dirapikan');
  } else {
    if (region) {
      value = applyCustomCorrections(value, corrections);
      if (value !== input) reasons.push('Kamus koreksi lokal diterapkan');
    }
    value = cleanString(value, caseRule);
    if (value !== input) reasons.push('Format teks dirapikan');
  }
  if (region) {
    const checked = correctRegionText(value);
    if (checked.corrected) {
      value = cleanString(checked.value, caseRule);
      reasons.push('Nama wilayah dicocokkan dengan daftar BPS');
    }
    warning = checked.warning;
  }
  return { next: value, changed: value !== input, reason: reasons.join('; '), warning };
}

export function splitNisNisn(input) {
  const value = normalizeSpace(input);
  const match = value.match(/^(.+?)\s*\/\s*(.+)$/);
  if (!match) return { parts: [value, ''], warning: 'Tidak ditemukan pemisah NIS/NISN berupa garis miring (/).' };
  return { parts: [match[1].trim(), match[2].trim()], warning: '' };
}

function formatBirthDate(input) {
  const value = normalizeSpace(input);
  const match = value.match(/^(\d{1,2})[\s/-]+([A-Za-z]+|\d{1,2})[\s/-]+(\d{4})$/i);
  if (!match) return null;
  const day = Number(match[1]);
  const monthIndex = /^\d+$/.test(match[2]) ? Number(match[2]) - 1 : MONTHS.findIndex(month => month.toLocaleLowerCase('id-ID') === match[2].toLocaleLowerCase('id-ID'));
  const year = Number(match[3]);
  if (monthIndex < 0 || monthIndex > 11 || year < 1800 || year > 2200) return null;
  const date = new Date(Date.UTC(year, monthIndex, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== monthIndex || date.getUTCDate() !== day) return null;
  return `${String(day).padStart(2, '0')} ${MONTHS[monthIndex]} ${year}`;
}

export function splitBirthPlaceDate(input, corrections = []) {
  const value = normalizeSpace(input);
  const match = value.match(/^(.+?),\s*(.+)$/);
  if (!match) return { parts: [value, ''], warning: 'Tidak ditemukan koma antara tempat dan tanggal lahir.' };
  const date = formatBirthDate(match[2]);
  if (!date) return { parts: [value, ''], warning: 'Tanggal lahir tidak dikenali atau tidak valid.' };
  const customPlace = applyCustomCorrections(match[1], corrections);
  const place = correctRegionName(customPlace);
  return {
    parts: [place.value, date],
    warning: place.suggestion ? `Nama wilayah “${match[1]}” mungkin “${place.suggestion}”; perlu diperiksa.` : '',
    correction: place.corrected ? `Nama wilayah “${match[1]}” dikoreksi menjadi “${place.value}”.` :
      customPlace !== match[1] ? `Kamus koreksi lokal: “${match[1]}” menjadi “${place.value}”.` : '',
  };
}
