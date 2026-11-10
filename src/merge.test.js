import test from 'node:test';
import assert from 'node:assert/strict';
import { joinMergeValues, shiftedColumn } from './merge.js';

test('gabung beberapa nilai memakai penghubung manual dan melewati sel kosong', () => {
  assert.equal(joinMergeValues(['Kp. Sekartoya', 'Ds. Pasirhuni', 'Kec. Cimaung'], ', '), 'Kp. Sekartoya, Ds. Pasirhuni, Kec. Cimaung');
  assert.equal(joinMergeValues(['1234', '', '0012345678'], ' / '), '1234 / 0012345678');
  assert.equal(joinMergeValues(['A', 'B', 'C'], ''), 'ABC');
});

test('posisi kolom pemisahan dihitung setelah sumber gabungan dihapus', () => {
  assert.equal(shiftedColumn(5, [2, 4]), 3);
  assert.equal(shiftedColumn(1, [2, 4]), 1);
});
