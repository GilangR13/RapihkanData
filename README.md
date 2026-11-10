# RapihkanData

Web untuk merapikan **seluruh kolom data** dalam file Excel (`.xlsx`), OpenDocument Spreadsheet (`.ods`), atau CSV dari konsumen. Tidak ada pemeriksaan duplikat. File diproses di browser; web tidak mengirim isi file ke server.

## Cara memakai

1. Buka web dan pilih satu file `.xlsx`, `.ods`, atau `.csv`.
2. Tinjau tiap lembar. Pilih aturan untuk setiap kolom: rapikan spasi (default), kapital awal kata, huruf besar, huruf kecil, atau biarkan isi. Aktifkan tata alamat, koreksi wilayah, atau pemisahan kolom sesuai isi.
3. Tentukan apakah baris pertama adalah judul kolom. Bila ya, aturan huruf tidak diterapkan pada judul, tetapi spasinya tetap dirapikan.
4. Bila ada nama jalan atau kampung yang salah ketik, masukkan pasangan `salah = benar` pada kamus koreksi sendiri, lalu simpan. Periksa pratinjau; sel yang berubah diberi warna hijau dan yang perlu diperiksa diberi warna kuning.
5. Klik **Unduh hasil & laporan**. ZIP berisi `<nama>_rapi.ods` dan `<nama>_laporan.txt` untuk dibuka di Notepad.

## Aturan pengolahan

- Semua kolom teks memakai aturan **rapikan spasi** secara default: spasi berlebih, tab, dan baris baru menjadi satu spasi; spasi di awal/akhir dibuang. Jenis kapitalisasi lain hanya berjalan bila dipilih.
- Kolom alamat yang dikenali dari judul otomatis menata `Kp.babakan` menjadi `Kp. Babakan` serta `Rt01/09` menjadi `RT 01/RW 09`. Pilihan ini dapat diubah per kolom.
- Koreksi wilayah memakai daftar nama 38 provinsi dan 514 kabupaten/kota dari [BPS](https://sig.bps.go.id/rest-drop-down/getwilayah) dengan pelengkap provinsi baru dari [tabel BPS](https://www.bps.go.id/id/statistics-table/2/ODY5IzI%3D/kapasitas-terpasang-pembangkit-listrik-menurut-provinsi-mega-watt.html). Salah ketik satu huruf hanya diperbaiki otomatis jika cocok ke satu nama wilayah saja. Contoh `Garuta` menjadi `Garut`, `Cirebong` menjadi `Cirebon`. Kandidat yang meragukan hanya diberi peringatan.
- Kolom `NIS / NISN` dapat dipisahkan menjadi `NIS` dan `NISN`. Kolom tempat/tanggal lahir seperti `Cirebong, 07 Agustus 2006` dapat dipisahkan menjadi `Cirebon` dan `07 Agustus 2006`. Tanggal tidak valid atau pola yang tidak dikenali dilaporkan untuk diperiksa.
- Kamus koreksi sendiri dipakai pada alamat, wilayah, dan tempat lahir. Pasangan disimpan di `localStorage` browser perangkat yang sama; file sumber dan kamus tidak dikirim ke server.
- Sel teks yang diawali `0` lalu angka disimpan dengan format teks Excel (`@`), termasuk kode, nomor telepon, NIK, NISN, atau kolom lain.
- Jika sel Excel berisi angka `127` tetapi formatnya `00000`, hasil disimpan sebagai teks `00127`.
- Angka biasa, tanggal, dan sel kosong tidak diubah. Rumus tidak diubah dan dicatat sebagai peringatan di laporan.
- Urutan baris, urutan kolom, dan nama lembar dipertahankan. File sumber tidak ditimpa.
- Laporan mencatat lembar, alamat sel, nilai awal, nilai hasil, dan alasan perubahan atau peringatan.

**Batas penting:** Bila nol awal sudah hilang di file sumber dan tidak ada petunjuk dari format sel, web tidak bisa menebaknya. Daftar BPS mencakup provinsi dan kabupaten/kota, bukan seluruh jalan, desa, atau kampung; ejaan nama jalan yang belum masuk kamus sendiri tidak ditebak otomatis. Aturan kapital awal kata memakai pengecualian singkatan umum, tetapi nama, gelar, dan alamat tetap perlu ditinjau. Hasil akhir selalu ODS; file sumber `.xlsx` atau `.ods` yang berisi fitur kompleks seperti grafik, pivot, macro, atau koneksi data perlu diperiksa kembali karena konversi berfokus pada isi sel, lembar, dan merge, bukan seluruh tampilan atau fitur khusus. CSV yang memakai encoding lama dapat menunjukkan karakter pengganti; bila itu terjadi, ubah sumber menjadi UTF-8 lebih dulu.

## Teknis

- **Frontend:** HTML, CSS, JavaScript, Vite.
- **Baca Excel:** ExcelJS. **Baca/tulis ODS:** SheetJS 0.20.3 dari rilis resminya. **Baca CSV:** PapaParse. **Paket hasil:** JSZip. **Laporan:** teks UTF-8 dengan BOM agar Notepad mengenali karakter Indonesia.
- **Hosting:** GitHub Pages melalui GitHub Actions pada setiap push ke branch `main`.
- **Privasi:** browser membaca file lokal memakai File API; kode ini tidak memiliki endpoint upload, database, atau penyimpanan cloud. Jangan commit file data konsumen ke repositori.
- **Hasil:** Semua input diekspor sebagai `.ods`. Nilai berawalan nol ditulis sebagai sel teks agar tetap utuh saat dibuka di LibreOffice Calc.

### Menjalankan lokal

```bash
npm ci
npm run dev
```

### Pemeriksaan

```bash
npm test
npm run build
```

### GitHub Pages

GitHub Pages menggunakan sumber **GitHub Actions**. Workflow di `.github/workflows/pages.yml` membangun dan memublikasikan web setiap ada push ke branch `main`. `vite.config.js` memakai path aset relatif agar situs bekerja di subpath `/RapihkanData/`.
