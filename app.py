"""
Web Server Flask untuk Traffic Counting ITDP Jabodetabek
Versi Production / Cloud-Ready dengan Multi-User Session Isolation.
Mendukung hosting gratis (Render, Hugging Face, PythonAnywhere, Railway, dsb).
"""

import os
import glob
import uuid
import tempfile
from flask import Flask, render_template, request, jsonify, send_file, session
from werkzeug.utils import secure_filename
from excel_engine import (
    parse_filename,
    inspect_template,
    write_verified_data_to_excel,
    perform_qc_check,
    INTERVALS_METADATA,
    NON_EV_CATEGORIES,
    EV_CATEGORIES
)

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "itdp-traffic-counting-secret-key-2026-production")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Gunakan /tmp pada serverless cloud seperti Vercel agar tidak error read-only filesystem
if os.environ.get("VERCEL") or os.environ.get("AWS_LAMBDA_FUNCTION_NAME"):
    UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "surveyor_uploads")
else:
    UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# Simpan state sementara per user session agar tidak bertabrakan saat diakses banyak surveyor
user_sessions_cache = {}

def get_current_user_id():
    """Mengambil atau membuat ID unik untuk tiap surveyor/browser yang mengakses"""
    if "user_id" not in session:
        session["user_id"] = str(uuid.uuid4())
    return session["user_id"]

def get_user_storage_dir():
    """Direktori penyimpanan file sementara khusus untuk user/surveyor aktif"""
    uid = get_current_user_id()
    user_dir = os.path.join(UPLOAD_DIR, uid)
    os.makedirs(user_dir, exist_ok=True)
    return user_dir

def get_user_cache():
    """Cache data hitungan khusus untuk user aktif"""
    uid = get_current_user_id()
    if uid not in user_sessions_cache:
        user_sessions_cache[uid] = {}
    return user_sessions_cache[uid]

@app.route("/")
def index():
    # Pastikan user_id tercipta saat pertama kali membuka halaman
    get_current_user_id()
    return render_template("index.html")

@app.route("/api/files", methods=["GET"])
def list_excel_files():
    """
    Mendata file template yang tersedia:
    1. File template global default di direktori utama
    2. File yang sebelumnya diunggah oleh user aktif
    """
    user_dir = get_user_storage_dir()
    
    # Ambil file template bawaan sistem dan file user
    system_files = glob.glob(os.path.join(BASE_DIR, "TC_*.xlsx"))
    user_files = glob.glob(os.path.join(user_dir, "*.xlsx"))
    
    files_info = []
    seen = set()
    
    # Prioritaskan file user
    for f in user_files + system_files:
        fn = os.path.basename(f)
        if fn.startswith("~$") or fn.startswith("test_") or fn.startswith("temp_") or fn in seen:
            continue
        seen.add(fn)
        parsed = parse_filename(fn)
        files_info.append({
            "filename": fn,
            "filepath": f,
            "parsed": parsed,
            "size": os.path.getsize(f)
        })
        
    cache = get_user_cache()
    return jsonify({
        "success": True, 
        "files": files_info,
        "has_active_file": bool(cache.get("current_filepath"))
    })

@app.route("/api/load-file", methods=["POST"])
def load_file():
    """Membaca file template Excel khusus untuk sesi surveyor aktif"""
    data = request.json or {}
    filepath = data.get("filepath")
    
    if not filepath or not os.path.exists(filepath):
        return jsonify({"success": False, "message": "File tidak ditemukan di server."}), 404
            
    try:
        template_data = inspect_template(filepath)
        cache = get_user_cache()
        cache["current_filepath"] = filepath
        cache["counts_data"] = template_data["counts_data"]
        cache["catatan"] = template_data["header_info"]["catatan"]
        
        return jsonify({
            "success": True,
            "data": template_data
        })
    except Exception as e:
        return jsonify({"success": False, "message": f"Gagal membaca file Excel: {str(e)}"}), 500

@app.route("/api/upload", methods=["POST"])
def upload_file():
    """
    Mengunggah berkas Excel via Drag & Drop.
    Tersimpan di direktori privat per-surveyor sehingga tidak saling menimpa.
    """
    if "file" not in request.files:
        return jsonify({"success": False, "message": "Tidak ada berkas yang dikirim."}), 400
        
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"success": False, "message": "Nama berkas kosong."}), 400
        
    if not file.filename.lower().endswith(".xlsx"):
        return jsonify({"success": False, "message": "Berkas harus berformat .xlsx"}), 400
        
    orig_name = os.path.basename(file.filename)
    user_dir = get_user_storage_dir()
    dest_path = os.path.join(user_dir, orig_name)
    file.save(dest_path)
    
    try:
        template_data = inspect_template(dest_path)
        cache = get_user_cache()
        cache["current_filepath"] = dest_path
        cache["counts_data"] = template_data["counts_data"]
        cache["catatan"] = template_data["header_info"]["catatan"]
        
        return jsonify({
            "success": True,
            "message": f"File '{orig_name}' berhasil diunggah!",
            "data": template_data
        })
    except Exception as e:
        return jsonify({"success": False, "message": f"Berkas diunggah tapi gagal diproses: {str(e)}"}), 500

@app.route("/api/close-file", methods=["POST"])
def close_file():
    """Menutup sesi file aktif user saat ini"""
    cache = get_user_cache()
    cache.clear()
    return jsonify({"success": True})

@app.route("/api/sync-to-excel", methods=["POST"])
def sync_to_excel():
    """Menulis hasil counting terverifikasi ke sel biru & kuning file Excel user"""
    payload = request.json or {}
    cache = get_user_cache()
    filepath = payload.get("filepath") or cache.get("current_filepath")
    counts_data = payload.get("counts_data") or cache.get("counts_data")
    note = payload.get("catatan")
    
    if not filepath or not os.path.exists(filepath):
        return jsonify({"success": False, "message": "Berkas Excel target tidak ditemukan."}), 400
        
    if not counts_data:
        return jsonify({"success": False, "message": "Data hitungan kosong."}), 400
        
    try:
        res = write_verified_data_to_excel(
            template_filepath=filepath,
            session_data_dict=counts_data,
            note=note,
            output_filepath=filepath
        )
        cache["counts_data"] = counts_data
        if note is not None:
            cache["catatan"] = note
            
        fresh_data = inspect_template(filepath)
            
        return jsonify({
            "success": True,
            "message": f"Berhasil memperbarui {res['cells_updated_count']} sel di file Excel!",
            "details": res,
            "excel_view_rows": fresh_data["excel_view_rows"],
            "existing_stats": fresh_data["existing_stats"]
        })
    except Exception as e:
        return jsonify({"success": False, "message": f"Gagal menulis ke Excel: {str(e)}"}), 500

@app.route("/api/download", methods=["GET"])
def download_excel():
    """Mengunduh file Excel hasil counting user aktif"""
    cache = get_user_cache()
    filepath = request.args.get("filepath") or cache.get("current_filepath")
    if not filepath or not os.path.exists(filepath):
        return jsonify({"success": False, "message": "File tidak ditemukan."}), 404
        
    return send_file(
        filepath,
        as_attachment=True,
        download_name=os.path.basename(filepath)
    )

@app.route("/api/qc-check", methods=["POST"])
def qc_check():
    """Validasi QC otomatis 8 poin Slide 21 ITDP"""
    payload = request.json or {}
    counts_data = payload.get("counts_data") or {}
    package = payload.get("package", "ALL")
    kelompok = payload.get("kelompok", "ALL")
    
    result = perform_qc_check(counts_data, package=package, kelompok=kelompok)
    return jsonify({"success": True, "qc": result})

if __name__ == "__main__":
    # Support dynamic PORT untuk platform cloud seperti Render, Heroku, Railway
    port = int(os.environ.get("PORT", 5000))
    print(f"Menjalankan Traffic Counting ITDP di port {port}...")
    app.run(host="0.0.0.0", port=port, debug=False)
