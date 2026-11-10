import test from 'node:test';
import assert from 'node:assert/strict';
import { cleanString, transformCell } from './cleaner.js';

test('rapikan spasi tanpa mengubah angka dalam teks', () => {
  assert.equal(cleanString('  Jl.  Melati\t  No.  07\n'), 'Jl. Melati No. 07');
});

test('nilai berawalan nol tetap teks', () => {
  assert.deepEqual(transformCell('  00127  ', 'space'), {
    next: '00127', changed: true, textFormat: true, reason: 'Format teks dirapikan',
  });
});

test('format Excel 00000 dipulihkan dari nilai angka', () => {
  assert.equal(transformCell(127, 'space', '00000').next, '00127');
  assert.equal(transformCell(127, 'space', 'General').next, 127);
});

test('kapital awal mempertahankan singkatan umum', () => {
  assert.equal(cleanString('  pt  maju  RT  02 ', 'title'), 'PT Maju RT 02');
});

test('aturan biarkan isi tidak merapikan spasi', () => {
  assert.equal(cleanString('  007  ', 'keep'), '  007  ');
});
