import test from 'node:test';
import assert from 'node:assert/strict';
import {
  correctRegionName, formatIndonesianAddress, splitBirthPlaceDate, splitNisNisn,
  parseCorrectionDictionary,
} from './indonesia.js';

test('alamat RT/RW dan kampung ditata sesuai contoh konsumen', () => {
  assert.equal(formatIndonesianAddress('Kp Cijengkol Rt01/09'), 'Kp. Cijengkol RT 01/RW 09');
  assert.equal(formatIndonesianAddress('Kp.babakan RT 2 RW 5 Desa Majasetra'), 'Kp. Babakan RT 2/RW 5 Desa Majasetra');
});

test('typo wilayah Garuta dan Cirebong dikoreksi dengan daftar BPS', () => {
  assert.equal(correctRegionName('Garuta').value, 'Garut');
  assert.equal(correctRegionName('Cirebong').value, 'Cirebon');
  assert.equal(correctRegionName('Bandung').value, 'Bandung');
});

test('NIS dan NISN dipisah tanpa menghilangkan nol awal', () => {
  assert.deepEqual(splitNisNisn('1234.56.007 / 0012345678').parts, ['1234.56.007', '0012345678']);
});

test('TTL dipisah, tanggal diperiksa, dan tempat lahir dikoreksi', () => {
  assert.deepEqual(splitBirthPlaceDate('Garuta, 05 Mei 2006').parts, ['Garut', '05 Mei 2006']);
  assert.deepEqual(splitBirthPlaceDate('Cirebong, 07 Agustus 2006').parts, ['Cirebon', '07 Agustus 2006']);
  assert.match(splitBirthPlaceDate('Bandung, 31 Februari 2006').warning, /tidak valid/);
});

test('kamus lokal memperbaiki nama jalan dan tempat lahir sesuai pasangan eksplisit', () => {
  const dictionary = parseCorrectionDictionary('Mawarh = Mawar\nCirebong = Cirebon');
  assert.equal(formatIndonesianAddress('Jl.Mawarh No07', 'title', dictionary), 'Jl. Mawar No. 07');
  assert.deepEqual(splitBirthPlaceDate('Cirebong, 07 Agustus 2006', dictionary).parts, ['Cirebon', '07 Agustus 2006']);
  assert.throws(() => parseCorrectionDictionary('salah benar'), /Baris kamus 1/);
});
