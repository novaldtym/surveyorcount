# Traffic Counting (TC) Surveyor ITDP CLEAR Jabodetabek v2.2

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/novaldtym/surveyorcount)

Sistem aplikasi web otomatis untuk menghitung volume kendaraan per interval **15 menit** dari rekaman video CCTV dan menginput hasil counting terverifikasi langsung ke template Microsoft Excel resmi ITDP (`.xlsx`).

---

## 🚀 Cara Membuka Web Secara Online (Gratis)

### 1. Deploy Otomatis 1-Klik di Render (Gratis)
Klik tombol di bawah ini untuk langsung mendeploy aplikasi ke Render secara gratis tanpa konfigurasi rumit:

👉 **[Deploy Web App ke Render Gratis](https://render.com/deploy?repo=https://github.com/novaldtym/surveyorcount)**

- Render akan otomatis membaca konfigurasi `render.yaml` dan `requirements.txt`.
- Pilih nama dan klik **Apply** / **Deploy**.
- Dalam 1–2 menit, Anda akan mendapatkan URL HTTPS publik resmi (contoh: `https://surveyorcount.onrender.com`) yang bisa langsung dibuka oleh Anda dan tim surveyor lain dari perangkat manapun.

---

## ⚡ Fitur Utama

1. **Multi-User Session Isolation**:
   - Tiap surveyor yang membuka link web memiliki ruang penyimpanan sementara tersendiri (`uploads/<session_id>/`).
   - File Excel yang di-drag & drop tidak akan saling tertukar atau menimpa antar-surveyor.

2. **Card Slot Drag & Drop Berkas**:
   - Hanya membaca dan mengedit berkas template Excel yang dimasukkan ke slot.

3. **Tampilan View Excel (Spreadsheet Viewer)**:
   - Menampilkan visual asli spreadsheet dengan arsir warna Biru (Peak/EV) dan Kuning (16 Jam), serta total rumus R–Z terkalkulasi live.

4. **15 Kategori Kendaraan Lengkap (Kolom C s/d Q)**:
   - 9 Kategori Nonlistrik (Kolom C–K): MCY, CAR, LGV, MGV, HGV, MIKROBUS, LBUS, MBUS, HBUS.
   - 6 Kategori Listrik EV (Kolom L–Q): Aktif pada baris 11–14 (jam 06.00–07.00).

5. **Penguncian Ketat EV (Hanya Jam 06.00–07.00)**:
   - Di luar jam 06.00–07.00, tombol dan hotkey EV otomatis dikunci dengan panduan bahwa kendaraan listrik wajib dicatat ke kategori konvensional di atas (C–K).

6. **Jaminan Keamanan Formula Excel 100%**:
   - Hanya mengubah sel berarsir biru dan kuning. Rumus header, kolom R–AI, dan sheet index 100% utuh.

---

## 💻 Menjalankan di Komputer Lokal

Cukup klik ganda:
```
run_app.bat
```
Atau jalankan:
```bash
python app.py
```
Buka di browser: `http://127.0.0.1:5000`.
