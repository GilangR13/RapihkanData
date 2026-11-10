import ExcelJS from 'exceljs';
import Papa from 'papaparse';
import JSZip from 'jszip';
import { RULES, transformCell } from './cleaner.js';
import './style.css';

const els = {
  dropzone: document.querySelector('#dropzone'),
  input: document.querySelector('#file-input'),
  editor: document.querySelector('#editor'),
  name: document.querySelector('#file-name'),
  details: document.querySelector('#file-details'),
  replace: document.querySelector('#replace-file'),
  sheet: document.querySelector('#sheet-select'),
  header: document.querySelector('#header-checkbox'),
  rules: document.querySelector('#rule-list'),
  previewHead: document.querySelector('#preview-table thead'),
  previewBody: document.querySelector('#preview-table tbody'),
  previewCount: document.querySelector('#preview-count'),
  previewNote: document.querySelector('#preview-note'),
  export: document.querySelector('#export-button'),
  status: document.querySelector('#status'),
};

const state = { file: null, workbook: null, rules: new Map(), sheetIndex: 0, hasHeader: true, importWarnings: [] };
let previewTimer;

function setStatus(message, error = false) {
  els.status.textContent = message;
  els.status.classList.toggle('error', error);
}

function fileSize(size) {
  return size < 1024 * 1024 ? `${Math.max(1, Math.round(size / 1024))} KB` : `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

async function readFile(file) {
  let workbook = new ExcelJS.Workbook();
  const warnings = [];
  if (/\.xlsx$/i.test(file.name)) {
    await workbook.xlsx.load(await file.arrayBuffer());
  } else if (/\.ods$/i.test(file.name)) {
    const { odsToExcelWorkbook } = await import('./ods.js');
    const converted = odsToExcelWorkbook(await file.arrayBuffer());
    workbook = converted.workbook;
    warnings.push(...converted.warnings);
  } else if (/\.csv$/i.test(file.name)) {
    const content = (await file.text()).replace(/^\uFEFF/, '');
    const parsed = Papa.parse(content, { skipEmptyLines: false, dynamicTyping: false });
    if (parsed.errors.length) warnings.push(`${parsed.errors.length} peringatan saat membaca CSV`);
    if (content.includes('\uFFFD')) warnings.push('Ada karakter yang tidak terbaca; periksa encoding CSV');
    const sheet = workbook.addWorksheet('Data CSV');
    const data = parsed.data;
    while (data.length && data[data.length - 1].every(value => value === '')) data.pop();
    data.forEach(row => sheet.addRow(row));
  } else {
    throw new Error('Pilih file .xlsx, .ods, atau .csv.');
  }
  if (!workbook.worksheets.length) throw new Error('File tidak berisi lembar data.');
  return { workbook, warnings };
}

function currentSheet() { return state.workbook.worksheets[state.sheetIndex]; }
function rulesFor(sheetIndex) {
  if (!state.rules.has(sheetIndex)) state.rules.set(sheetIndex, []);
  return state.rules.get(sheetIndex);
}
function ruleFor(sheetIndex, columnIndex, rowIndex) {
  return state.hasHeader && rowIndex === 1 ? 'space' : rulesFor(sheetIndex)[columnIndex - 1] || 'space';
}

function cellLabel(cell) {
  if (cell.value == null) return '';
  if (typeof cell.value === 'object' && 'formula' in cell.value) return `=${cell.value.formula}`;
  if (cell.value instanceof Date) return cell.text || cell.value.toLocaleDateString('id-ID');
  if (typeof cell.value === 'object') return cell.text || '[data khusus]';
  return String(cell.value);
}

function makeOption(value, label) {
  const option = document.createElement('option');
  option.value = value;
  option.textContent = label;
  return option;
}

function renderRules() {
  const sheet = currentSheet();
  const count = Math.max(1, sheet.columnCount);
  const fragment = document.createDocumentFragment();
  for (let index = 1; index <= Math.min(count, 120); index++) {
    const columnName = sheet.getColumn(index).letter;
    const heading = cellLabel(sheet.getRow(1).getCell(index)).trim();
    const item = document.createElement('label');
    item.className = 'rule-item';
    const small = document.createElement('span');
    small.textContent = `KOLOM ${columnName}`;
    const strong = document.createElement('strong');
    strong.textContent = heading || `Kolom ${columnName}`;
    strong.title = strong.textContent;
    const select = document.createElement('select');
    select.setAttribute('aria-label', `Aturan kolom ${columnName}: ${strong.textContent}`);
    RULES.forEach(rule => select.append(makeOption(rule.value, rule.label)));
    select.value = rulesFor(state.sheetIndex)[index - 1] || 'space';
    select.addEventListener('change', () => {
      rulesFor(state.sheetIndex)[index - 1] = select.value;
      schedulePreview();
    });
    item.append(small, strong, select);
    fragment.append(item);
  }
  els.rules.replaceChildren(fragment);
  if (count > 120) {
    const note = document.createElement('p');
    note.className = 'muted-note';
    note.textContent = `${count - 120} kolom berikutnya tetap diproses dengan aturan “Rapikan spasi”.`;
    els.rules.append(note);
  }
}

function schedulePreview() {
  clearTimeout(previewTimer);
  previewTimer = setTimeout(renderPreview, 100);
}

function renderPreview() {
  const sheet = currentSheet();
  const maxColumn = Math.min(Math.max(1, sheet.columnCount), 20);
  const headRow = document.createElement('tr');
  headRow.append(document.createElement('th'));
  headRow.firstChild.textContent = 'BARIS';
  for (let column = 1; column <= maxColumn; column++) {
    const th = document.createElement('th');
    th.textContent = state.hasHeader ? cellLabel(sheet.getRow(1).getCell(column)).trim() || sheet.getColumn(column).letter : sheet.getColumn(column).letter;
    headRow.append(th);
  }
  els.previewHead.replaceChildren(headRow);

  const fragment = document.createDocumentFragment();
  const maxRow = Math.min(sheet.rowCount, 31);
  let previewChanges = 0;
  for (let rowIndex = 1; rowIndex <= maxRow; rowIndex++) {
    const tr = document.createElement('tr');
    const rowNumber = document.createElement('td');
    rowNumber.textContent = String(rowIndex);
    tr.append(rowNumber);
    for (let column = 1; column <= maxColumn; column++) {
      const cell = sheet.getRow(rowIndex).getCell(column);
      const td = document.createElement('td');
      const original = cellLabel(cell);
      const transformed = transformCell(cell.value, ruleFor(state.sheetIndex, column, rowIndex), cell.numFmt);
      td.textContent = transformed.changed ? String(transformed.next) : original;
      if (transformed.changed) {
        td.className = 'changed';
        td.title = `Sebelum: ${original}`;
        previewChanges++;
      } else {
        td.title = original;
      }
      tr.append(td);
    }
    fragment.append(tr);
  }
  els.previewBody.replaceChildren(fragment);
  els.previewCount.textContent = `${previewChanges} perubahan di pratinjau`;
  els.previewNote.textContent = `Menampilkan ${maxRow} dari ${sheet.rowCount} baris dan ${maxColumn} dari ${sheet.columnCount} kolom. Semua baris dan kolom diproses saat mengunduh.`;
}

async function handleFile(file) {
  if (!file) return;
  setStatus('Membaca file…');
  els.editor.classList.add('hidden');
  try {
    const { workbook, warnings } = await readFile(file);
    state.file = file;
    state.workbook = workbook;
    state.rules = new Map();
    state.sheetIndex = 0;
    state.importWarnings = warnings;
    els.name.textContent = file.name;
    els.details.textContent = `${fileSize(file.size)} · ${workbook.worksheets.length} lembar`;
    els.sheet.replaceChildren(...workbook.worksheets.map((sheet, index) => makeOption(String(index), sheet.name)));
    els.sheet.value = '0';
    renderRules();
    renderPreview();
    els.editor.classList.remove('hidden');
    setStatus(warnings.length ? warnings.join(' · ') : 'File siap ditinjau.');
    els.editor.scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (error) {
    state.file = null;
    state.workbook = null;
    setStatus(`File tidak dapat dibaca: ${error.message}`, true);
  }
}

function buildTextReport(rows, sourceName, warningCount) {
  const changedCount = rows.filter(row => row[2] === 'Diubah').length;
  const lines = [
    'LAPORAN RAPIHKANDATA',
    `File sumber : ${sourceName}`,
    `Sel diubah  : ${changedCount}`,
    `Peringatan : ${warningCount}`,
    '',
    'File asli tidak diubah. Nilai yang sudah kehilangan nol awal sebelum dibaca tidak dapat dipulihkan tanpa acuan.',
    '',
    'RINCIAN',
  ];
  if (!rows.length) lines.push('Tidak ada perubahan atau peringatan.');
  rows.forEach(([sheet, address, status, before, after, reason], index) => {
    lines.push('', `${index + 1}. [${status}] ${sheet}${address ? `!${address}` : ''}`);
    if (status === 'Diubah') {
      lines.push(`   Sebelum : ${JSON.stringify(String(before))}`);
      lines.push(`   Sesudah : ${JSON.stringify(String(after))}`);
    }
    lines.push(`   Catatan : ${reason}`);
  });
  return '\uFEFF' + lines.join('\r\n') + '\r\n';
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}

async function exportFiles() {
  if (!state.file) return;
  els.export.disabled = true;
  setStatus('Menyiapkan file rapi dan laporan…');
  try {
    const { workbook, warnings } = await readFile(state.file);
    const rows = [];
    let warningCount = warnings.length;
    for (let sheetIndex = 0; sheetIndex < workbook.worksheets.length; sheetIndex++) {
      const sheet = workbook.worksheets[sheetIndex];
      sheet.eachRow((row, rowIndex) => {
        row.eachCell({ includeEmpty: false }, cell => {
          if (cell.value && typeof cell.value === 'object' && 'formula' in cell.value) {
            rows.push([sheet.name, cell.address, 'Peringatan', cellLabel(cell), cellLabel(cell), 'Rumus dibiarkan; periksa hasil hitung saat dibuka di Excel']);
            warningCount++;
            return;
          }
          const original = cellLabel(cell);
          const result = transformCell(cell.value, ruleFor(sheetIndex, cell.col, rowIndex), cell.numFmt);
          const needsTextFormat = result.textFormat && cell.numFmt !== '@';
          if (result.changed) cell.value = result.next;
          if (needsTextFormat) cell.numFmt = '@';
          if (result.changed || needsTextFormat) {
            rows.push([sheet.name, cell.address, 'Diubah', original, cellLabel(cell), result.reason || 'Nilai berawalan nol disimpan sebagai teks']);
          }
        });
      });
    }
    warnings.forEach(message => rows.push(['File sumber', '', 'Peringatan', '', '', message]));
    const { excelWorkbookToOds } = await import('./ods.js');
    const cleanBuffer = excelWorkbookToOds(workbook);
    const reportText = buildTextReport(rows, state.file.name, warningCount);
    const base = state.file.name.replace(/\.(xlsx|ods|csv)$/i, '').replace(/[\\/:*?"<>|]/g, '_');
    const zip = new JSZip();
    zip.file(`${base}_rapi.ods`, cleanBuffer);
    zip.file(`${base}_laporan.txt`, reportText);
    downloadBlob(await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' }), `${base}_hasil.zip`);
    const changedCount = rows.filter(row => row[2] === 'Diubah').length;
    setStatus(`Selesai. ${changedCount} perubahan dan ${warningCount} peringatan tercatat dalam ZIP.`);
  } catch (error) {
    setStatus(`Gagal membuat hasil: ${error.message}`, true);
  } finally {
    els.export.disabled = false;
  }
}

els.dropzone.addEventListener('click', () => els.input.click());
els.dropzone.addEventListener('keydown', event => {
  if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); els.input.click(); }
});
els.dropzone.addEventListener('dragover', event => { event.preventDefault(); els.dropzone.classList.add('dragover'); });
els.dropzone.addEventListener('dragleave', () => els.dropzone.classList.remove('dragover'));
els.dropzone.addEventListener('drop', event => {
  event.preventDefault();
  els.dropzone.classList.remove('dragover');
  handleFile(event.dataTransfer.files[0]);
});
els.input.addEventListener('change', () => { handleFile(els.input.files[0]); els.input.value = ''; });
els.replace.addEventListener('click', () => els.input.click());
els.sheet.addEventListener('change', () => { state.sheetIndex = Number(els.sheet.value); renderRules(); renderPreview(); });
els.header.addEventListener('change', () => { state.hasHeader = els.header.checked; renderRules(); renderPreview(); });
els.export.addEventListener('click', exportFiles);
