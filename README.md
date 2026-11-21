# Traffic Counting (TC) Surveyor ITDP CLEAR Jabodetabek v2.2

Sistem web otomatis untuk menghitung volume kendaraan per interval **15 menit** dari rekaman video CCTV dan menginput hasil counting terverifikasi langsung ke template Microsoft Excel resmi ITDP (`.xlsx`).

---

## 🌐 Deploy Online Gratis (Aktif 24 Jam Nonstop Tanpa Laptop Hidup)

Aplikasi ini dapat di-deploy ke cloud gratis sehingga **berjalan terus 24 jam nonstop di internet tanpa perlu laptop Anda dinyalakan**. Anda bisa memilih salah satu platform favorit berikut:

### Opsi A: Deploy di Vercel (Gratis & Cepat)
1. Klik tombol deploy Vercel di bawah ini:
   
   👉 **[Deploy ke Vercel (Gratis)](https://vercel.com/new/clone?repository-url=https://github.com/novaldtym/surveyorcount)**

2. Masuk menggunakan akun GitHub Anda (`novaldtym`).
3. Klik **Create** / **Deploy**.
4. Dalam 1 menit, Anda akan mendapatkan URL resmi Vercel (contoh: `https://surveyorcount.vercel.app`) yang bisa dibuka langsung oleh tim surveyor kapan saja!

---

### Opsi B: Deploy di Render (Gratis)
1. Klik tombol deploy Render di bawah ini:
   
   👉 **[Deploy ke Render (Gratis)](https://render.com/deploy?repo=https://github.com/novaldtym/surveyorcount)**

2. Masuk menggunakan akun GitHub Anda.
3. Klik **Apply** / **Deploy Web Service**.
4. Selesai! Web akan online di domain `https://surveyorcount.onrender.com`.

---

## ⚡ Fitur Utama

1. **Online 24/7 di Cloud**:
   - Berjalan di server cloud. Laptop Anda **bisa dimatikan total** dan web tetap aktif diakses oleh siapa saja.
2. **Multi-User Session Isolation**:
   - Tiap surveyor yang membuka link web memiliki ruang penyimpanan privat tersendiri (`/tmp` / `uploads/<session_id>/`).
   - File Excel yang di-drag & drop tidak akan saling tertukar atau menimpa antar-surveyor.
3. **Card Slot Drag & Drop Berkas**:
   - Hanya membaca dan mengedit berkas template Excel yang dimasukkan ke slot.
4. **Tampilan View Excel (Spreadsheet Viewer)**:
   - Menampilkan visual asli spreadsheet dengan arsir warna Biru (Peak/EV) dan Kuning (16 Jam), serta total rumus R–Z terkalkulasi live.
5. **15 Kategori Kendaraan Lengkap (Kolom C s/d Q)**:
   - 9 Kategori Nonlistrik (Kolom C–K): MCY, CAR, LGV, MGV, HGV, MIKROBUS, LBUS, MBUS, HBUS.
   - 6 Kategori Listrik EV (Kolom L–Q): Aktif pada baris 11–14 (jam 06.00–07.00).
6. **Penguncian Ketat EV (Hanya Jam 06.00–07.00)**:
   - Di luar jam 06.00–07.00, tombol dan hotkey EV otomatis dikunci dengan panduan resmi ITDP bahwa kendaraan listrik wajib dicatat ke kategori konvensional di atas (C–K).
7. **Jaminan Keamanan Formula Excel 100%**:
   - Hanya mengubah sel berarsir biru dan kuning. Rumus header, kolom R–AI, dan sheet index 100% utuh.

---

## 💻 Menjalankan di Komputer Lokal

Jika ingin menjalankan secara offline di komputer lokal tanpa koneksi internet:
1. Klik ganda file:
   ```
   run_app.bat
   ```
2. Buka di peramban: `http://127.0.0.1:5000`.
