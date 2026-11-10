import ExcelJS from 'exceljs';
import { read, utils } from 'xlsx';

function safeSheetName(name, used) {
  const cleaned = name.replace(/[\\/*?:\[\]]/g, '_').slice(0, 31) || 'Lembar';
  let candidate = cleaned;
  let suffix = 2;
  while (used.has(candidate.toLocaleLowerCase('id-ID'))) {
    const tail = `_${suffix++}`;
    candidate = cleaned.slice(0, 31 - tail.length) + tail;
  }
  used.add(candidate.toLocaleLowerCase('id-ID'));
  return candidate;
}

function convertCell(source) {
  if (source.f) {
    return { value: { formula: source.f, result: source.v }, numFmt: source.z };
  }
  if (source.t === 'n') {
    const shown = typeof source.w === 'string' ? source.w.trim() : '';
    if (/^0\d+$/.test(shown) && shown !== String(source.v)) {
      return { value: shown, numFmt: '@' };
    }
    return { value: source.v, numFmt: source.z };
  }
  if (source.t === 'd') {
    const date = source.v instanceof Date ? source.v : new Date(source.v);
    return { value: Number.isNaN(date.getTime()) ? String(source.v) : date, numFmt: source.z };
  }
  if (source.t === 'b') return { value: Boolean(source.v) };
  if (source.t === 'e') return { value: source.w || String(source.v ?? '') };
  return { value: source.v == null ? '' : String(source.v) };
}

export function odsToExcelWorkbook(data) {
  const source = read(data, { type: 'array', cellDates: true, cellNF: true, cellText: true });
  const workbook = new ExcelJS.Workbook();
  const warnings = ['File ODS dikonversi ke XLSX; periksa kembali tata letak dan fitur lembar kerja khusus.'];
  const usedNames = new Set();

  for (const originalName of source.SheetNames) {
    const name = safeSheetName(originalName, usedNames);
    if (name !== originalName) warnings.push(`Nama lembar “${originalName}” disesuaikan menjadi “${name}” untuk Excel.`);
    const target = workbook.addWorksheet(name);
    const input = source.Sheets[originalName];
    for (const address of Object.keys(input)) {
      if (address.startsWith('!')) continue;
      const cell = input[address];
      if (!cell || cell.t === 'z') continue;
      const converted = convertCell(cell);
      const output = target.getCell(address);
      output.value = converted.value;
      if (converted.numFmt) output.numFmt = converted.numFmt;
    }
    for (const merge of input['!merges'] || []) {
      target.mergeCells(utils.encode_range(merge));
    }
  }
  return { workbook, warnings };
}
