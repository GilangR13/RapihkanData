import ExcelJS from 'exceljs';
import Papa from 'papaparse';
import { RULES, transformCell } from './cleaner.js';
import { cleanConfiguredText, splitNisNisn, splitBirthPlaceDate, parseCorrectionDictionary } from './indonesia.js';
import { shiftedColumn, groupColumns, finalMergeGroups, mergeGroupValue } from './merge.js';
import './style.css';
import './columns.css';

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
  mergeList: document.querySelector('#merge-list'),
  addMerge: document.querySelector('#add-merge'),
  dictionaryInput: document.querySelector('#dictionary-input'),
  dictionarySave: document.querySelector('#dictionary-save'),
  dictionaryStatus: document.querySelector('#dictionary-status'),
  previewHead: document.querySelector('#preview-table thead'),
  previewBody: document.querySelector('#preview-table tbody'),
  previewCount: document.querySelector('#preview-count'),
  previewNote: document.querySelector('#preview-note'),
  previewPrev: document.querySelector('#preview-prev'),
  previewNext: document.querySelector('#preview-next'),
  previewRange: document.querySelector('#preview-range'),
  previewColumnPrev: document.querySelector('#preview-column-prev'),
  previewColumnNext: document.querySelector('#preview-column-next'),
  previewColumnRange: document.querySelector('#preview-column-range'),
  export: document.querySelector('#export-button'),
  status: document.querySelector('#status'),
};

const state = {
  file: null, workbook: null, rules: new Map(), addressColumns: new Map(),
  regionColumns: new Map(), splitColumns: new Map(), mergeGroups: new Map(), sheetIndex: 0,
  hasHeader: true, importWarnings: [],
  customCorrections: [],
  previewPage: 0,
  previewColumnPage: 0,
  nextMergeId: 1,
};
const PREVIEW_PAGE_SIZE = 50;
const PREVIEW_COLUMN_PAGE_SIZE = 20;
const SPLIT_OPTIONS = [
  { value: 'none', label: 'Tidak dipisah' },
  { value: 'nis', label: 'NIS / NISN' },
  { value: 'birth', label: 'Tempat / tanggal lahir' },
];
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
function settingFor(map, sheetIndex, columnIndex, fallback = false) {
  return map.get(sheetIndex)?.[columnIndex - 1] ?? fallback;
}
function splitFor(sheetIndex, columnIndex) {
  return settingFor(state.splitColumns, sheetIndex, columnIndex, 'none');
}
function mergeGroupsFor(sheetIndex) {
  if (!state.mergeGroups.has(sheetIndex)) state.mergeGroups.set(sheetIndex, []);
  return state.mergeGroups.get(sheetIndex);
}
function mergeError(sheetIndex, group, groupIndex) {
  if (group.columns.length + group.groupIds.length < 2) return 'Pilih minimal dua sumber: kolom atau hasil gabungan sebelumnya.';
  if (group.columns.some(column => splitFor(sheetIndex, column) !== 'none')) return 'Kolom yang digabung tidak boleh sekaligus memakai aturan Pisahkan kolom.';
  const groups = mergeGroupsFor(sheetIndex);
  if (groups.some((other, index) => index !== groupIndex && other.columns.some(column => group.columns.includes(column)))) return 'Satu kolom hanya boleh masuk ke satu gabungan.';
  for (const id of group.groupIds) {
    const previousIndex = groups.findIndex(item => item.id === id);
    if (previousIndex < 0 || previousIndex >= groupIndex) return 'Hasil gabungan harus berasal dari langkah sebelumnya.';
    if (mergeError(sheetIndex, groups[previousIndex], previousIndex)) return `Gabungan ${previousIndex + 1} perlu dilengkapi lebih dulu.`;
  }
  return '';
}
function finalMergeError(sheetIndex) {
  const groups = mergeGroupsFor(sheetIndex);
  const ready = groups.filter((group, index) => !mergeError(sheetIndex, group, index));
  const finals = finalMergeGroups(ready);
  for (let index = 0; index < finals.length; index++) {
    const firstColumns = groupColumns(finals[index], groups);
    for (let other = index + 1; other < finals.length; other++) {
      if (groupColumns(finals[other], groups).some(column => firstColumns.includes(column))) {
        return 'Dua hasil akhir memakai kolom sumber yang sama. Satukan kedua hasil itu pada gabungan berikutnya.';
      }
    }
  }
  return '';
}
function mergedCellText(cell) {
  if (cell.value && typeof cell.value === 'object' && 'formula' in cell.value) return cell.value.result == null ? '' : String(cell.value.result);
  return cellLabel(cell);
}
function mergedRowValue(sheet, sheetIndex, rowIndex, group, groups) {
  return mergeGroupValue(group, groups, column => {
    const cell = sheet.getRow(rowIndex).getCell(column);
    if (cell.value && typeof cell.value === 'object' && 'formula' in cell.value) return mergedCellText(cell);
    const transformed = transformConfiguredCell(cell.value, sheetIndex, column, rowIndex, cell.numFmt);
    return transformed.changed ? transformed.next : mergedCellText(cell);
  });
}
function isHeaderRow(rowIndex) { return state.hasHeader && rowIndex === 1; }
function splitHeaders(mode) {
  return mode === 'nis' ? ['NIS', 'NISN'] : ['Tempat Lahir', 'Tanggal Lahir'];
}
function splitValue(mode, value) {
  if (value == null || value === '') return { parts: ['', ''], warning: '', correction: '' };
  if (typeof value !== 'string') return { parts: [value, ''], warning: 'Isi bukan teks; pemisahan perlu diperiksa.', correction: '' };
  return mode === 'nis' ? splitNisNisn(value) : splitBirthPlaceDate(value, state.customCorrections);
}

function initializeColumnSettings(workbook) {
  state.rules = new Map();
  state.addressColumns = new Map();
  state.regionColumns = new Map();
  state.splitColumns = new Map();
  state.mergeGroups = new Map();
  state.nextMergeId = 1;
  workbook.worksheets.forEach((sheet, sheetIndex) => {
    const rules = [];
    const addresses = [];
    const regions = [];
    const splits = [];
    for (let column = 1; column <= sheet.columnCount; column++) {
      const header = cellLabel(sheet.getRow(1).getCell(column)).toLocaleLowerCase('id-ID');
      const sample = cellLabel(sheet.getRow(2).getCell(column));
      const address = /\balamat\b|domisili/.test(header);
      const birth = /\bttl\b|tempat\s*(?:dan\s*)?tanggal\s*lahir|tempat\s*lahir/.test(header);
      const nis = /\bnis\b\s*[/&-]\s*\bnisn\b|\bnisn\b\s*[/&-]\s*\bnis\b/.test(header);
      rules.push(address ? 'title' : 'space');
      addresses.push(address);
      regions.push(birth || /\bkabupaten\b|\bkota\b|\bprovinsi\b/.test(header));
      splits.push(nis ? 'nis' : birth && /,\s*\d{1,2}\s+\S+\s+\d{4}/.test(sample) ? 'birth' : 'none');
    }
    state.rules.set(sheetIndex, rules);
    state.addressColumns.set(sheetIndex, addresses);
    state.regionColumns.set(sheetIndex, regions);
    state.splitColumns.set(sheetIndex, splits);
  });
}

function transformConfiguredCell(value, sheetIndex, column, rowIndex, numFmt = '') {
  if (typeof value !== 'string' || isHeaderRow(rowIndex)) {
    return { ...transformCell(value, ruleFor(sheetIndex, column, rowIndex), numFmt), warning: '' };
  }
  const result = cleanConfiguredText(value, {
    caseRule: ruleFor(sheetIndex, column, rowIndex),
    address: settingFor(state.addressColumns, sheetIndex, column),
    region: settingFor(state.regionColumns, sheetIndex, column),
    corrections: state.customCorrections,
  });
  return { ...result, textFormat: /^0\d/.test(result.next) };
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
    const addressLabel = document.createElement('label');
    addressLabel.className = 'rule-toggle';
    const addressCheckbox = document.createElement('input');
    addressCheckbox.type = 'checkbox';
    addressCheckbox.checked = settingFor(state.addressColumns, state.sheetIndex, index);
    addressCheckbox.setAttribute('aria-label', `Tata alamat kolom ${columnName}`);
    addressCheckbox.addEventListener('change', () => {
      state.addressColumns.get(state.sheetIndex)[index - 1] = addressCheckbox.checked;
      schedulePreview();
    });
    addressLabel.append(addressCheckbox, ' Tata alamat Indonesia');
    const regionLabel = document.createElement('label');
    regionLabel.className = 'rule-toggle';
    const regionCheckbox = document.createElement('input');
    regionCheckbox.type = 'checkbox';
    regionCheckbox.checked = settingFor(state.regionColumns, state.sheetIndex, index);
    regionCheckbox.setAttribute('aria-label', `Koreksi wilayah kolom ${columnName}`);
    regionCheckbox.addEventListener('change', () => {
      state.regionColumns.get(state.sheetIndex)[index - 1] = regionCheckbox.checked;
      schedulePreview();
    });
    regionLabel.append(regionCheckbox, ' Koreksi nama wilayah');
    const splitCaption = document.createElement('span');
    splitCaption.className = 'rule-caption';
    splitCaption.textContent = 'PISAHKAN KOLOM';
    const splitSelect = document.createElement('select');
    splitSelect.setAttribute('aria-label', `Pisahkan kolom ${columnName}`);
    SPLIT_OPTIONS.forEach(option => splitSelect.append(makeOption(option.value, option.label)));
    splitSelect.value = splitFor(state.sheetIndex, index);
    splitSelect.addEventListener('change', () => {
      state.splitColumns.get(state.sheetIndex)[index - 1] = splitSelect.value;
      renderMergeGroups();
      schedulePreview();
    });
    item.append(small, strong, select, addressLabel, regionLabel, splitCaption, splitSelect);
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

function renderMergeGroups() {
  const sheet = currentSheet();
  const groups = mergeGroupsFor(state.sheetIndex);
  const fragment = document.createDocumentFragment();
  groups.forEach((group, groupIndex) => {
    const card = document.createElement('div');
    card.className = 'merge-card';
    const head = document.createElement('div');
    head.className = 'merge-card-head';
    const title = document.createElement('strong');
    title.textContent = `Gabungan ${groupIndex + 1}`;
    const remove = document.createElement('button');
    remove.type = 'button';
    remove.className = 'merge-remove';
    remove.textContent = 'Hapus gabungan';
    remove.addEventListener('click', () => {
      groups.splice(groupIndex, 1);
      groups.forEach(other => { other.groupIds = other.groupIds.filter(id => id !== group.id); });
      renderMergeGroups();
      renderPreview();
    });
    head.append(title, remove);

    const fields = document.createElement('div');
    fields.className = 'merge-fields';
    for (const [caption, key] of [['Judul kolom hasil', 'header'], ['Teks penghubung (spasi ikut dihitung)', 'separator']]) {
      const label = document.createElement('label');
      label.textContent = caption;
      const input = document.createElement('input');
      input.type = 'text';
      input.value = group[key];
      input.setAttribute('aria-label', `${caption} gabungan ${groupIndex + 1}`);
      input.addEventListener('input', () => { group[key] = input.value; schedulePreview(); });
      label.append(input);
      fields.append(label);
    }

    const choices = document.createElement('div');
    choices.className = 'merge-columns';
    groups.slice(0, groupIndex).forEach((previous, previousIndex) => {
      const option = document.createElement('label');
      option.className = 'merge-prior';
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.checked = group.groupIds.includes(previous.id);
      checkbox.disabled = Boolean(mergeError(state.sheetIndex, previous, previousIndex));
      checkbox.setAttribute('aria-label', `Gabungan ${groupIndex + 1}, hasil Gabungan ${previousIndex + 1}`);
      checkbox.addEventListener('change', () => {
        group.groupIds = checkbox.checked ? [...group.groupIds, previous.id] : group.groupIds.filter(id => id !== previous.id);
        renderMergeGroups();
        renderPreview();
      });
      option.append(checkbox, `Hasil Gabungan ${previousIndex + 1} · ${previous.header.trim() || 'Hasil gabungan'}`);
      choices.append(option);
    });
    for (let column = 1; column <= sheet.columnCount; column++) {
      const option = document.createElement('label');
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.checked = group.columns.includes(column);
      checkbox.disabled = groups.some((other, index) => index !== groupIndex && other.columns.includes(column)) || splitFor(state.sheetIndex, column) !== 'none';
      checkbox.setAttribute('aria-label', `Gabungan ${groupIndex + 1}, kolom ${sheet.getColumn(column).letter}`);
      checkbox.addEventListener('change', () => {
        group.columns = checkbox.checked ? [...group.columns, column].sort((a, b) => a - b) : group.columns.filter(value => value !== column);
        renderMergeGroups();
        renderPreview();
      });
      option.append(checkbox, `${sheet.getColumn(column).letter} · ${cellLabel(sheet.getRow(1).getCell(column)) || `Kolom ${column}`}`);
      choices.append(option);
    }
    const error = mergeError(state.sheetIndex, group, groupIndex);
    card.append(head, fields, choices);
    const consumers = groups.flatMap((other, index) => other.groupIds.includes(group.id) ? [index + 1] : []);
    if (consumers.length) {
      const note = document.createElement('p');
      note.className = 'merge-note';
      note.textContent = `Hasil ini dipakai lagi pada Gabungan ${consumers.join(', ')}.`;
      card.append(note);
    }
    const selectedResults = group.groupIds.map(id => groups.find(previous => previous.id === id)).filter(Boolean);
    if (selectedResults.some((first, index) => selectedResults.slice(index + 1).some(second => groupColumns(first, groups).some(column => groupColumns(second, groups).includes(column))))) {
      const note = document.createElement('p');
      note.className = 'merge-warning';
      note.textContent = 'Dua hasil yang dipilih memakai sumber yang sama; bagian tersebut akan muncul dua kali dalam gabungan ini.';
      card.append(note);
    }
    if (error) {
      const message = document.createElement('p');
      message.className = 'merge-error';
      message.textContent = error;
      card.append(message);
    }
    fragment.append(card);
  });
  els.mergeList.replaceChildren(fragment);
  const planError = finalMergeError(state.sheetIndex);
  if (planError) {
    const message = document.createElement('p');
    message.className = 'merge-error';
    message.textContent = planError;
    els.mergeList.append(message);
  }
}

function schedulePreview() {
  clearTimeout(previewTimer);
  previewTimer = setTimeout(renderPreview, 100);
}

function renderPreview() {
  const sheet = currentSheet();
  const pageCount = Math.max(1, Math.ceil(sheet.rowCount / PREVIEW_PAGE_SIZE));
  state.previewPage = Math.min(state.previewPage, pageCount - 1);
  const firstRow = state.previewPage * PREVIEW_PAGE_SIZE + 1;
  const lastRow = Math.min(sheet.rowCount, firstRow + PREVIEW_PAGE_SIZE - 1);
  const groups = mergeGroupsFor(state.sheetIndex);
  const validGroups = groups.filter((group, index) => !mergeError(state.sheetIndex, group, index));
  const finalGroups = finalMergeError(state.sheetIndex) ? [] : finalMergeGroups(validGroups);
  const mergedSources = new Map(finalGroups.flatMap(group => groupColumns(group, groups).map(column => [column, group])));
  const allColumns = [];
  for (let column = 1; column <= Math.max(1, sheet.columnCount); column++) {
    const group = mergedSources.get(column);
    if (group) {
      if (column === groupColumns(group, groups)[0]) allColumns.push({ column, part: -1, mergeGroup: group, label: group.header.trim() || 'Hasil gabungan' });
      continue;
    }
    const mode = splitFor(state.sheetIndex, column);
    if (mode === 'none') allColumns.push({ column, part: -1, label: state.hasHeader ? cellLabel(sheet.getRow(1).getCell(column)).trim() || sheet.getColumn(column).letter : sheet.getColumn(column).letter });
    else splitHeaders(mode).forEach((label, part) => allColumns.push({ column, part, label, mode }));
  }
  const totalColumns = allColumns.length;
  const columnPageCount = Math.max(1, Math.ceil(totalColumns / PREVIEW_COLUMN_PAGE_SIZE));
  state.previewColumnPage = Math.min(state.previewColumnPage, columnPageCount - 1);
  const firstColumn = state.previewColumnPage * PREVIEW_COLUMN_PAGE_SIZE + 1;
  const lastColumn = Math.min(totalColumns, firstColumn + PREVIEW_COLUMN_PAGE_SIZE - 1);
  const columns = allColumns.slice(firstColumn - 1, lastColumn);
  const headRow = document.createElement('tr');
  headRow.append(document.createElement('th'));
  headRow.firstChild.textContent = 'BARIS';
  for (const column of columns) {
    const th = document.createElement('th');
    th.textContent = column.label;
    headRow.append(th);
  }
  els.previewHead.replaceChildren(headRow);

  const fragment = document.createDocumentFragment();
  let previewChanges = 0;
  let previewWarnings = 0;
  for (let rowIndex = firstRow; rowIndex <= lastRow; rowIndex++) {
    const tr = document.createElement('tr');
    const rowNumber = document.createElement('td');
    rowNumber.textContent = String(rowIndex);
    tr.append(rowNumber);
    for (const column of columns) {
      const cell = sheet.getRow(rowIndex).getCell(column.column);
      const td = document.createElement('td');
      const original = cellLabel(cell);
      if (column.mergeGroup) {
        const sourceColumns = groupColumns(column.mergeGroup, groups);
        const originalValues = sourceColumns.map(source => cellLabel(sheet.getRow(rowIndex).getCell(source)));
        td.textContent = isHeaderRow(rowIndex) ? column.label : mergedRowValue(sheet, state.sheetIndex, rowIndex, column.mergeGroup, groups);
        if (!isHeaderRow(rowIndex) && originalValues.some(Boolean)) {
          const formula = sourceColumns.some(source => {
            const value = sheet.getRow(rowIndex).getCell(source).value;
            return value && typeof value === 'object' && 'formula' in value;
          });
          td.className = formula ? 'needs-review' : 'changed';
          td.title = formula ? 'Gabungan memakai hasil hitung rumus yang tersimpan; periksa nilainya.' : `Sebelum: ${originalValues.join(' | ')}`;
          if (formula) previewWarnings++;
          else previewChanges++;
        }
      } else if (column.part >= 0) {
        const result = isHeaderRow(rowIndex) ? { parts: splitHeaders(column.mode) } : splitValue(column.mode, cell.value);
        td.textContent = String(result.parts[column.part] ?? '');
        if (!isHeaderRow(rowIndex) && !result.warning && original) {
          td.className = 'changed';
          td.title = `Sebelum: ${original}${result.correction ? ` · ${result.correction}` : ''}`;
          if (column.part === 0) previewChanges++;
        } else if (result.warning && column.part === 0) {
          td.className = 'needs-review';
          td.title = result.warning;
          previewWarnings++;
        }
      } else {
        const transformed = transformConfiguredCell(cell.value, state.sheetIndex, column.column, rowIndex, cell.numFmt);
        td.textContent = transformed.changed ? String(transformed.next) : original;
        if (transformed.warning) {
          td.className = 'needs-review';
          td.title = `${transformed.changed ? `Sebelum: ${original} · ` : ''}${transformed.warning}`;
          previewWarnings++;
          if (transformed.changed) previewChanges++;
        } else if (transformed.changed) {
          td.className = 'changed';
          td.title = `Sebelum: ${original}`;
          previewChanges++;
        } else td.title = original;
      }
      tr.append(td);
    }
    fragment.append(tr);
  }
  els.previewBody.replaceChildren(fragment);
  els.previewCount.textContent = `${previewChanges} perubahan · ${previewWarnings} perlu diperiksa`;
  els.previewRange.textContent = sheet.rowCount ? `Baris ${firstRow}–${lastRow} dari ${sheet.rowCount}` : 'Tidak ada baris';
  els.previewPrev.disabled = state.previewPage === 0;
  els.previewNext.disabled = state.previewPage >= pageCount - 1;
  els.previewColumnRange.textContent = `Kolom ${firstColumn}–${lastColumn} dari ${totalColumns}`;
  els.previewColumnPrev.disabled = state.previewColumnPage === 0;
  els.previewColumnNext.disabled = state.previewColumnPage >= columnPageCount - 1;
  els.previewNote.textContent = 'Pratinjau menampilkan sebagian baris dan kolom per halaman. Semua baris dan kolom diproses saat mengunduh.';
}

async function handleFile(file) {
  if (!file) return;
  setStatus('Membaca file…');
  els.editor.classList.add('hidden');
  try {
    const { workbook, warnings } = await readFile(file);
    state.file = file;
    state.workbook = workbook;
    initializeColumnSettings(workbook);
    state.sheetIndex = 0;
    state.previewPage = 0;
    state.previewColumnPage = 0;
    state.importWarnings = warnings;
    els.name.textContent = file.name;
    els.details.textContent = `${fileSize(file.size)} · ${workbook.worksheets.length} lembar`;
    els.sheet.replaceChildren(...workbook.worksheets.map((sheet, index) => makeOption(String(index), sheet.name)));
    els.sheet.value = '0';
    renderRules();
    renderMergeGroups();
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
  for (let sheetIndex = 0; sheetIndex < state.workbook.worksheets.length; sheetIndex++) {
    const groups = mergeGroupsFor(sheetIndex);
    for (let index = 0; index < groups.length; index++) {
      const error = mergeError(sheetIndex, groups[index], index);
      if (error) {
        setStatus(`Gabungan ${index + 1} pada lembar ${state.workbook.worksheets[sheetIndex].name}: ${error}`, true);
        return;
      }
    }
    const planError = finalMergeError(sheetIndex);
    if (planError) {
      setStatus(`Lembar ${state.workbook.worksheets[sheetIndex].name}: ${planError}`, true);
      return;
    }
  }
  els.export.disabled = true;
  setStatus('Menyiapkan file ODS…');
  try {
    const { workbook, warnings } = await readFile(state.file);
    let changedCount = 0;
    let warningCount = warnings.length;
    for (let sheetIndex = 0; sheetIndex < workbook.worksheets.length; sheetIndex++) {
      const sheet = workbook.worksheets[sheetIndex];
      const originalColumnCount = sheet.columnCount;
      sheet.eachRow((row, rowIndex) => {
        row.eachCell({ includeEmpty: false }, cell => {
          if (cell.value && typeof cell.value === 'object' && 'formula' in cell.value) {
            warningCount++;
            return;
          }
          if (!isHeaderRow(rowIndex) && splitFor(sheetIndex, cell.col) !== 'none') return;
          const result = transformConfiguredCell(cell.value, sheetIndex, cell.col, rowIndex, cell.numFmt);
          const needsTextFormat = result.textFormat && cell.numFmt !== '@';
          if (result.changed) cell.value = result.next;
          if (needsTextFormat) cell.numFmt = '@';
          if (result.changed || needsTextFormat) changedCount++;
          if (result.warning) warningCount++;
        });
      });
      const groups = mergeGroupsFor(sheetIndex);
      const outputs = finalMergeGroups(groups).map(group => ({
        group,
        columns: groupColumns(group, groups),
        values: Array.from({ length: sheet.rowCount }, (_, index) => {
          const rowIndex = index + 1;
          return isHeaderRow(rowIndex) ? group.header.trim() || 'Hasil gabungan' : mergeGroupValue(group, groups, column => mergedCellText(sheet.getRow(rowIndex).getCell(column)));
        }),
      }));
      for (const { columns, values } of outputs) {
        const anchor = columns[0];
        for (let rowIndex = 1; rowIndex <= sheet.rowCount; rowIndex++) {
          const target = sheet.getRow(rowIndex).getCell(anchor);
          const joined = values[rowIndex - 1];
          if (!isHeaderRow(rowIndex) && joined) changedCount++;
          target.value = joined;
          target.numFmt = '@';
        }
      }
      const removedColumns = [...new Set(outputs.flatMap(output => output.columns.slice(1)))].sort((a, b) => b - a);
      for (const column of removedColumns) sheet.spliceColumns(column, 1);
      for (let column = originalColumnCount; column >= 1; column--) {
        const mode = splitFor(sheetIndex, column);
        if (mode === 'none') continue;
        const currentColumn = shiftedColumn(column, removedColumns);
        const first = [];
        const second = [];
        for (let rowIndex = 1; rowIndex <= sheet.rowCount; rowIndex++) {
          const cell = sheet.getRow(rowIndex).getCell(currentColumn);
          const hasValue = cell.value != null && cell.value !== '';
          const result = isHeaderRow(rowIndex) ? { parts: splitHeaders(mode), warning: '' } : splitValue(mode, cell.value);
          first.push(result.parts[0]);
          second.push(result.parts[1]);
          if (!isHeaderRow(rowIndex) && hasValue) {
            if (result.warning) warningCount++;
            else changedCount++;
          }
        }
        sheet.spliceColumns(currentColumn, 1, first, second);
      }
    }
    const { excelWorkbookToOds } = await import('./ods.js');
    const cleanBuffer = excelWorkbookToOds(workbook);
    const base = state.file.name.replace(/\.(xlsx|ods|csv)$/i, '').replace(/[\\/:*?"<>|]/g, '_');
    downloadBlob(new Blob([cleanBuffer], { type: 'application/vnd.oasis.opendocument.spreadsheet' }), `${base}_rapi.ods`);
    setStatus(`Selesai. File ODS diunduh dengan ${changedCount} perubahan dan ${warningCount} peringatan untuk ditinjau.`);
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
els.sheet.addEventListener('change', () => { state.sheetIndex = Number(els.sheet.value); state.previewPage = 0; state.previewColumnPage = 0; renderRules(); renderMergeGroups(); renderPreview(); });
els.header.addEventListener('change', () => { state.hasHeader = els.header.checked; state.previewPage = 0; state.previewColumnPage = 0; renderRules(); renderPreview(); });
els.addMerge.addEventListener('click', () => {
  const groups = mergeGroupsFor(state.sheetIndex);
  groups.push({ id: state.nextMergeId++, columns: [], groupIds: [], separator: ', ', header: `Gabungan ${groups.length + 1}` });
  renderMergeGroups();
  renderPreview();
});
els.previewPrev.addEventListener('click', () => { state.previewPage--; renderPreview(); });
els.previewNext.addEventListener('click', () => { state.previewPage++; renderPreview(); });
els.previewColumnPrev.addEventListener('click', () => { state.previewColumnPage--; renderPreview(); });
els.previewColumnNext.addEventListener('click', () => { state.previewColumnPage++; renderPreview(); });
els.export.addEventListener('click', exportFiles);
try {
  const saved = localStorage.getItem('rapihkandata-kamus') || '';
  state.customCorrections = parseCorrectionDictionary(saved);
  els.dictionaryInput.value = saved;
} catch {
  els.dictionaryStatus.textContent = 'Kamus tersimpan tidak dapat dibaca di browser ini.';
}
els.dictionarySave.addEventListener('click', () => {
  try {
    const entries = parseCorrectionDictionary(els.dictionaryInput.value);
    state.customCorrections = entries;
    let persisted = true;
    try { localStorage.setItem('rapihkandata-kamus', els.dictionaryInput.value); }
    catch { persisted = false; }
    if (state.workbook) renderPreview();
    els.dictionaryStatus.textContent = persisted ? `${entries.length} pasangan diterapkan.` : 'Kamus diterapkan, tetapi penyimpanan browser tidak tersedia.';
  } catch (error) {
    els.dictionaryStatus.textContent = error.message;
  }
});
