# Panduan Deploy Web Traffic Counting ke Hostingan Gratis

Aplikasi ini berbasis **Python Flask** dengan penyimpanan file Excel terisolasi per-sesi pengguna (*Multi-User Session Isolation*). Artinya, banyak surveyor dapat membuka link web ini secara bersamaan, masing-masing melakukan drag & drop file Excelnya sendiri, menghitung, dan mengunduh hasilnya tanpa saling menimpa data.

Berikut adalah **3 pilihan terbaik hosting gratis** beserta langkah-langkah mudahnya:

---

## Opsi 1: Render.com (Paling Direkomendasikan & Populer)

Render menyediakan tier gratis untuk Web Service Python dengan URL publik HTTPS (contoh: `https://tc-surveyor.onrender.com`).

### Langkah-langkah:
1. **Upload folder proyek ke GitHub**:
   - Buat akun di [github.com](https://github.com) (jika belum punya).
   - Buat repository baru (bisa *Public* atau *Private*), misal beri nama `tc-surveyor`.
   - Buka terminal di folder `d:\surveyor` lalu jalankan:
     ```bash
     git init
     git add .
     git commit -m "Initial commit for traffic counting"
     git branch -M main
     git remote add origin https://github.com/USERNAME_ANDA/tc-surveyor.git
     git push -u origin main
     ```
2. **Deploy di Render**:
   - Buka [render.com](https://render.com) dan daftar/login menggunakan akun GitHub Anda.
   - Klik tombol **New +** di pojok kanan atas ➔ Pilih **Web Service**.
   - Pilih repository `tc-surveyor` yang baru dibuat.
   - Isi form konfigurasi:
     - **Name**: `tc-surveyor` (atau nama lain yang Anda inginkan)
     - **Region**: Singapore (paling dekat dan cepat dari Indonesia)
     - **Runtime**: `Python 3`
     - **Build Command**: `pip install -r requirements.txt`
     - **Start Command**: `gunicorn app:app`
     - **Instance Type**: `Free`
   - Klik **Create Web Service**.
3. **Selesai**:
   - Tunggu sekitar 1–2 menit sampai status menjadi *Live*.
   - Render akan memberikan link publik HTTPS yang bisa Anda bagikan ke tim surveyor lain!

---

## Opsi 2: Hugging Face Spaces (100% Gratis & Tanpa Sleep/Tidur)

Hugging Face Spaces sangat stabil dan **tidak pernah mati/tidur** (*no cold-start delay*).

### Langkah-langkah:
1. Buka [huggingface.co](https://huggingface.co) dan buat akun gratis.
2. Masuk ke menu **Spaces** ➔ Klik **Create new Space**.
3. Isi konfigurasi:
   - **Space name**: `traffic-counting-itdp`
   - **License**: `mit` (atau bebas)
   - **Select the Space SDK**: Pilih **Docker** ➔ Pilih template **Blank**.
   - **Space Hardware**: Free CPU (16 GB RAM).
   - Klik **Create Space**.
4. **Upload File**:
   - Masuk ke tab **Files and versions** di Space Anda ➔ Klik **Add file** ➔ **Upload files**.
   - Drag & drop seluruh file dari folder `d:\surveyor` (terutama `Dockerfile`, `app.py`, `excel_engine.py`, `requirements.txt`, folder `templates`, dan folder `static`).
   - Klik **Commit changes**.
5. **Selesai**:
   - Hugging Face akan otomatis mem-build Docker container dan web app Anda langsung aktif di link publik HTTPS resmi.

---

## Opsi 3: Cloudflare Tunnel (Instan dari Laptop Sendiri Tanpa Perlu Akun Cloud)

Jika tim Anda ingin segera menghitung sekarang juga dan aplikasi sedang jalan di laptop Anda, Anda bisa membuat link publik instan dalam 10 detik tanpa upload ke GitHub:

1. Pastikan server lokal sedang jalan (`python app.py` atau klik `run_app.bat`).
2. Buka terminal PowerShell baru, lalu jalankan:
   ```powershell
   winget install --id Cloudflare.cloudflared
   ```
3. Buat tunnel publik gratis:
   ```powershell
   cloudflared tunnel --url http://localhost:5000
   ```
4. Di layar terminal akan muncul link acak HTTPS, misalnya:
   `https://random-words.trycloudflare.com`
5. Bagikan link tersebut ke surveyor lain. Mereka langsung bisa membuka web di browser HP/laptop mereka dari mana saja!

---

## 🔒 Keamanan File Sementara (Multi-User Isolation)

Kode `app.py` telah diperbarui dengan fitur **User Session Isolation**:
- Setiap surveyor yang membuka link web otomatis diberi ID sesi unik (UUID).
- File yang di-drag & drop oleh Surveyor A disimpan di folder privat `uploads/<ID_A>/`, dan Surveyor B di folder `uploads/<ID_B>/`.
- Hasil hitungan tidak akan tertukar atau saling menimpa data satu sama lain.
