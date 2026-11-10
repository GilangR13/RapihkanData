import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { read } from 'xlsx';
import { odsToExcelWorkbook, excelWorkbookToOds } from './ods.js';

test('membaca ODS LibreOffice dengan dua lembar, nol awal, dan rumus', async () => {
  const fixture = readFileSync(new URL('../test/fixtures/sample.ods', import.meta.url));
  const { workbook } = odsToExcelWorkbook(fixture);
  assert.deepEqual(workbook.worksheets.map(sheet => sheet.name), ['Peserta', 'Alamat']);
  assert.equal(workbook.getWorksheet('Peserta').getCell('B2').value, '00127');
  assert.equal(workbook.getWorksheet('Peserta').getCell('B2').numFmt, '@');
  assert.equal(workbook.getWorksheet('Alamat').getCell('B2').value, '0812345');
  assert.deepEqual(workbook.getWorksheet('Peserta').getCell('C2').value, { formula: '1+2', result: 3 });

  const output = read(excelWorkbookToOds(workbook), { type: 'array' });
  assert.deepEqual(output.SheetNames, ['Peserta', 'Alamat']);
  assert.equal(output.Sheets.Peserta.B2.v, '00127');
  assert.equal(output.Sheets.Peserta.B2.t, 's');
  assert.equal(output.Sheets.Alamat.B2.v, '0812345');
  assert.equal(output.Sheets.Alamat.B2.t, 's');
  assert.equal(output.Sheets.Peserta.C2.f, '1+2');
});
