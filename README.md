# Sistem Web Traffic Counting (TC) Surveyor ITDP CLEAR Jabodetabek v2.0

Sistem aplikasi web otomatis untuk menghitung volume kendaraan per interval **15 menit** dari rekaman video CCTV dan menginput hasil counting terverifikasi langsung ke template Microsoft Excel resmi ITDP (`.xlsx`).

---

## ⚡ Fitur Terbaru (Versi 2.0)

1. **Card Slot Drag & Drop Berkas**:
   - Sistem **hanya membaca dan mengedit file yang sudah di-drag & drop / diunggah** ke slot.
   - Dilengkapi slot kartu interaktif di halaman utama dan opsi pilih cepat template dari folder kerja.
   - Tombol *Ganti File* di bagian atas memungkinkan pergantian file kapan saja.

2. **Tampilan View Excel (Live Spreadsheet Viewer)**:
   - Tab khusus **"Tampilan View Excel"** untuk melihat langsung visual spreadsheet Excel otentik.
   - Menampilkan sel dengan arsir warna asli (**Biru** untuk jam peak/EV dan **Kuning** untuk 16 jam).
   - Menghitung kolom total rumus (R–Z) secara live.
   - Klik pada baris manapun di tabel Excel untuk langsung lompat ke sesi 15 menit tersebut di Mode Counter.

3. **Auto-Load Data Existing (Lanjut Counting Tanpa Dari 0)**:
   - Jika berkas Excel sudah sempat diisi sebagian atau separuh jalan, saat di-drag sistem akan otomatis membaca seluruh angka hitungan yang sudah ada.
   - Angka langsung tampil di kartu pencacah dan rekapitulasi, sehingga surveyor bisa langsung melanjutkan (*resume*) tanpa harus menghitung ulang dari 0.

4. **Deteksi Otomatis Ruas 8/2 D (8 Jalur) & Kunci Paket Otomatis**:
   - Pembagian Kelompok **K1 & K2 HANYA berlaku untuk ruas 8/2 D (8 jalur)**. Jika ruas adalah `4/2 D`, `2/2 UD`, atau `6/2 D`, sistem otomatis mengunci ke *"Semua Kategori (Ruas Non-8 Jalur)"* (1 counter menghitung semua kendaraan sesuai panduan).
   - Pemilihan paket (`_PG`, `_SR`, atau 16 Jam) **otomatis terkunci sesuai judul file dan header template** agar tidak terjadi salah pengisian rentang waktu.

5. **Penguncian Ketat Kendaraan Listrik (EV Hanya Jam 06.00 – 07.00)**:
   - Kolom EV (L–Q) hanya aktif pada baris 11–14 (pukul 06.00–07.00).
   - Di luar jam 06.00–07.00 (mulai pukul 07.00 ke atas), bagian EV **dinonaktifkan total dan dikunci**. Hotkey EV (`Shift+1..9`) otomatis diblokir dengan peringatan bahwa kendaraan listrik wajib dihitung ke kategori konvensional di atas (C–K) sesuai Slide 16 panduan ITDP.

6. **Jaminan Keamanan Sel Excel 100%**:
   - Sistem **HANYA** mengubah sel berarsir biru dan kuning di sheet pertama.
   - Rumus teks biru di header baris 1–3, rumus formula di kolom R–AI, dan sheet `index` **100% utuh dan tidak tersentuh**.

---

## 🚀 Cara Menjalankan Aplikasi

Cukup **klik ganda (double-click)** file:
```
d:\surveyor\run_app.bat
```
Atau lewat PowerShell:
```powershell
cd d:\surveyor
python app.py
```
Akses di peramban: `http://127.0.0.1:5000`.

---

## 📋 Alur Penggunaan

1. **Tarik & Lepaskan File**:
   - Drag & drop file Excel (misal: `TC_1_Arah-A_@HK_x_NAMA YANG HITUNG.xlsx`) ke Card Slot.
2. **Cek Tampilan Excel**:
   - Klik tab **Tampilan View Excel** untuk memastikan file sudah terbaca dan melihat data yang sudah ada sebelumnya.
3. **Mulai / Lanjutkan Pencacahan**:
   - Beralih ke **Mode Counter**. Angka yang sudah terisi akan otomatis muncul.
   - Tekan tombol `+1` / `+5` atau gunakan hotkey keyboard (`1`–`9` untuk non-EV; `Shift+1`–`9` untuk EV di jam 06–07).
4. **Verifikasi Sesi**:
   - Tekan `Space` atau `Ctrl + Enter` (otomatis isi 0 pada kategori yang nihil dan maju ke sesi 15 menit berikutnya).
5. **Simpan ke Excel**:
   - Klik **Simpan ke Excel** untuk menulis data secara aman ke file `.xlsx`.
