import test from 'node:test';
import assert from 'node:assert/strict';
import { joinMergeValues, shiftedColumn, groupColumns, finalMergeGroups, mergeGroupValue } from './merge.js';

test('gabung beberapa nilai memakai penghubung manual dan melewati sel kosong', () => {
  assert.equal(joinMergeValues(['Kp. Sekartoya', 'Ds. Pasirhuni', 'Kec. Cimaung'], ', '), 'Kp. Sekartoya, Ds. Pasirhuni, Kec. Cimaung');
  assert.equal(joinMergeValues(['1234', '', '0012345678'], ' / '), '1234 / 0012345678');
  assert.equal(joinMergeValues(['A', 'B', 'C'], ''), 'ABC');
});

test('posisi kolom pemisahan dihitung setelah sumber gabungan dihapus', () => {
  assert.equal(shiftedColumn(5, [2, 4]), 3);
  assert.equal(shiftedColumn(1, [2, 4]), 1);
});

test('hasil gabungan dapat menjadi sumber gabungan berikutnya dengan penghubung berbeda', () => {
  const groups = [
    { id: 1, columns: [1, 2, 3], groupIds: [], separator: ', ', header: 'Alamat awal' },
    { id: 2, columns: [4], groupIds: [1], separator: ' - ', header: 'Alamat lengkap' },
  ];
  const values = ['Kp. Sekartoya', 'Ds. Pasirhuni', 'Kec. Cimaung', 'Kab. Bandung'];
  assert.deepEqual(groupColumns(groups[1], groups), [1, 2, 3, 4]);
  assert.deepEqual(finalMergeGroups(groups).map(group => group.id), [2]);
  assert.equal(mergeGroupValue(groups[1], groups, column => values[column - 1]), 'Kp. Sekartoya, Ds. Pasirhuni, Kec. Cimaung - Kab. Bandung');
});

test('Gabungan 3 menyatukan hasil Gabungan 1 dan 2 tanpa mengganti penghubung masing-masing', () => {
  const groups = [
    { id: 1, columns: [1, 2, 3], groupIds: [], separator: ', ', header: 'Alamat' },
    { id: 2, columns: [4, 5], groupIds: [], separator: ' / ', header: 'NIS / NISN' },
    { id: 3, columns: [], groupIds: [1, 2], separator: ' | ', header: 'Data gabungan' },
  ];
  const values = ['Kp. Sekartoya', 'Ds. Pasirhuni', 'Kec. Cimaung', '1234', '001234'];
  assert.deepEqual(finalMergeGroups(groups).map(group => group.id), [3]);
  assert.deepEqual(groupColumns(groups[2], groups), [1, 2, 3, 4, 5]);
  assert.equal(mergeGroupValue(groups[2], groups, column => values[column - 1]), 'Kp. Sekartoya, Ds. Pasirhuni, Kec. Cimaung | 1234 / 001234');
});
