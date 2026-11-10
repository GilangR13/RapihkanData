import ExcelJS from 'exceljs';
import { read, utils, write } from 'xlsx';

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
  const warnings = ['Periksa kembali tata letak dan fitur lembar kerja khusus setelah merapikan file ODS.'];
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

function toSheetJsCell(cell) {
  const value = cell.value;
  if (value == null) return null;
  let result;
  if (typeof value === 'object' && 'formula' in value) {
    const calculated = value.result;
    const type = typeof calculated === 'string' ? 's' : typeof calculated === 'boolean' ? 'b' : 'n';
    result = { t: type, f: value.formula };
    if (calculated !== undefined && calculated !== null) result.v = calculated;
  } else if (value instanceof Date) {
    result = { t: 'd', v: value };
  } else if (typeof value === 'number') {
    result = { t: 'n', v: value };
  } else if (typeof value === 'boolean') {
    result = { t: 'b', v: value };
  } else {
    result = { t: 's', v: typeof value === 'string' ? value : cell.text || String(value) };
  }
  if (cell.numFmt && cell.numFmt !== 'General') result.z = cell.numFmt;
  return result;
}

export function excelWorkbookToOds(workbook) {
  const output = utils.book_new();
  for (const source of workbook.worksheets) {
    const sheet = {};
    let maxRow = 0;
    let maxColumn = 0;
    source.eachRow((row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, cell => {
        const converted = toSheetJsCell(cell);
        if (!converted) return;
        sheet[cell.address] = converted;
        maxRow = Math.max(maxRow, rowNumber);
        maxColumn = Math.max(maxColumn, cell.col);
      });
    });
    sheet['!ref'] = maxRow && maxColumn ? utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxRow - 1, c: maxColumn - 1 } }) : 'A1';
    const merges = source.model.merges || [];
    if (merges.length) sheet['!merges'] = merges.map(range => utils.decode_range(range));
    utils.book_append_sheet(output, sheet, source.name);
  }
  return write(output, { bookType: 'ods', type: 'array', cellDates: true });
}
