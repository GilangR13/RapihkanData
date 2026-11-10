import { mkdir, writeFile } from 'node:fs/promises';

const base = 'https://sig.bps.go.id/rest-drop-down/getwilayah';

async function getRegions(level, parent) {
  const response = await fetch(`${base}?level=${level}&parent=${parent}`, {
    headers: { accept: 'application/json' },
  });
  if (!response.ok) throw new Error(`${level}/${parent}: HTTP ${response.status}`);
  const data = await response.json();
  if (!Array.isArray(data)) throw new Error(`${level}/${parent}: respons bukan array`);
  return data.map(({ kode, nama }) => ({ kode: String(kode), nama: String(nama).trim() }));
}

const legacyProvinces = await getRegions('provinsi', '0');
const regencies = [];
for (let index = 0; index < legacyProvinces.length; index += 5) {
  const batch = await Promise.all(legacyProvinces.slice(index, index + 5).map(async province =>
    (await getRegions('kabupaten', province.kode)).map(region => ({ ...region }))
  ));
  regencies.push(...batch.flat());
}
// Endpoint dropdown BPS masih mengelompokkan Papua dalam 34 provinsi lama.
// Empat nama provinsi baru dicocokkan dengan tabel provinsi BPS tahun 2025.
const additionalProvinces = [
  { kode: '93', nama: 'PAPUA SELATAN' },
  { kode: '94', nama: 'PAPUA TENGAH' },
  { kode: '95', nama: 'PAPUA PEGUNUNGAN' },
  { kode: '96', nama: 'PAPUA BARAT DAYA' },
];
const provinces = [...legacyProvinces, ...additionalProvinces];
if (provinces.length < 38 || regencies.length < 500) {
  throw new Error(`Daftar wilayah tidak lengkap: ${provinces.length} provinsi, ${regencies.length} kabupaten/kota`);
}
const output = {
  sumber: base,
  sumberProvinsiTambahan: 'https://www.bps.go.id/id/statistics-table/2/ODY5IzI%3D/kapasitas-terpasang-pembangkit-listrik-menurut-provinsi-mega-watt.html',
  diambil: new Date().toISOString().slice(0, 10),
  provinsi: provinces,
  kabupatenKota: regencies,
  catatan: 'Nama kabupaten/kota dari endpoint BPS berjumlah 514. Pengelompokan Papua dalam endpoint ini masih mengikuti 34 provinsi lama; data dipakai sebagai daftar nama, bukan validasi hierarki provinsi.',
};
await mkdir(new URL('../src/data/', import.meta.url), { recursive: true });
await writeFile(new URL('../src/data/regions.json', import.meta.url), JSON.stringify(output, null, 2) + '\n');
console.log(`${provinces.length} provinsi, ${regencies.length} kabupaten/kota`);
