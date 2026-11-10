# RapihkanData

Web untuk merapikan **seluruh kolom data** dalam file Excel (`.xlsx`) atau CSV dari konsumen. Tidak ada pemeriksaan duplikat. File diproses di browser; web tidak mengirim isi file ke server.

## Cara memakai

1. Buka web dan pilih satu file `.xlsx` atau `.csv`.
2. Tinjau tiap lembar. Pilih aturan untuk setiap kolom: rapikan spasi (default), kapital awal kata, huruf besar, huruf kecil, atau biarkan isi.
3. Tentukan apakah baris pertama adalah judul kolom. Bila ya, aturan huruf tidak diterapkan pada judul, tetapi spasinya tetap dirapikan.
4. Periksa pratinjau. Sel yang berubah diberi warna hijau; arahkan kursor untuk melihat nilai awal.
5. Klik **Unduh hasil & laporan**. ZIP berisi `<nama>_rapi.xlsx` dan `<nama>_laporan.xlsx`.

## Aturan pengolahan

- Semua kolom teks memakai aturan **rapikan spasi** secara default: spasi berlebih, tab, dan baris baru menjadi satu spasi; spasi di awal/akhir dibuang. Jenis kapitalisasi lain hanya berjalan bila dipilih.
- Sel teks yang diawali `0` lalu angka disimpan dengan format teks Excel (`@`), termasuk kode, nomor telepon, NIK, NISN, atau kolom lain.
- Jika sel Excel berisi angka `127` tetapi formatnya `00000`, hasil disimpan sebagai teks `00127`.
- Angka biasa, tanggal, dan sel kosong tidak diubah. Rumus tidak diubah dan dicatat sebagai peringatan di laporan.
- Urutan baris, urutan kolom, dan nama lembar dipertahankan. File sumber tidak ditimpa.
- Laporan mencatat lembar, alamat sel, nilai awal, nilai hasil, dan alasan perubahan atau peringatan.

**Batas penting:** Bila nol awal sudah hilang di file sumber dan tidak ada petunjuk dari format sel, web tidak bisa menebaknya. Aturan kapital awal kata memakai pengecualian singkatan umum, tetapi nama, gelar, dan alamat tetap perlu ditinjau. File `.xlsx` yang berisi fitur Excel kompleks seperti grafik, pivot, macro, atau koneksi data perlu diperiksa kembali karena pustaka ExcelJS berfokus pada data lembar kerja. CSV yang memakai encoding lama dapat menunjukkan karakter pengganti; bila itu terjadi, ubah sumber menjadi UTF-8 lebih dulu.

## Teknis

- **Frontend:** HTML, CSS, JavaScript, Vite.
- **Baca/tulis Excel:** ExcelJS. **Baca CSV:** PapaParse. **Paket hasil:** JSZip.
- **Hosting:** GitHub Pages melalui GitHub Actions pada setiap push ke branch `main`.
- **Privasi:** browser membaca file lokal memakai File API; kode ini tidak memiliki endpoint upload, database, atau penyimpanan cloud. Jangan commit file data konsumen ke repositori.
- **Hasil:** CSV juga diekspor sebagai `.xlsx` supaya nilai berawalan nol tetap bertipe teks saat dibuka di Excel.

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

### Mengaktifkan GitHub Pages

Di repositori GitHub, buka **Settings → Pages → Build and deployment → Source** dan pilih **GitHub Actions**. Workflow di `.github/workflows/pages.yml` kemudian membangun dan memublikasikan web dari branch `main`.
