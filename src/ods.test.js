import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import ExcelJS from 'exceljs';
import { odsToExcelWorkbook } from './ods.js';

test('membaca ODS LibreOffice dengan dua lembar, nol awal, dan rumus', async () => {
  const fixture = readFileSync(new URL('../test/fixtures/sample.ods', import.meta.url));
  const { workbook } = odsToExcelWorkbook(fixture);
  assert.deepEqual(workbook.worksheets.map(sheet => sheet.name), ['Peserta', 'Alamat']);
  assert.equal(workbook.getWorksheet('Peserta').getCell('B2').value, '00127');
  assert.equal(workbook.getWorksheet('Peserta').getCell('B2').numFmt, '@');
  assert.equal(workbook.getWorksheet('Alamat').getCell('B2').value, '0812345');
  assert.deepEqual(workbook.getWorksheet('Peserta').getCell('C2').value, { formula: '1+2', result: 3 });

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(await workbook.xlsx.writeBuffer());
  assert.equal(output.getWorksheet('Peserta').getCell('B2').value, '00127');
  assert.equal(output.getWorksheet('Peserta').getCell('B2').numFmt, '@');
});
