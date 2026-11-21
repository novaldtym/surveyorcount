"""
Excel Engine untuk Sistem Web Traffic Counting ITDP CLEAR Jabodetabek
Memastikan HANYA sel berarsir biru dan kuning di sheet pertama yang dimodifikasi.
Semua sel rumus (teks biru, kolom R-AI, header baris 1-3) dan sheet index dilindungi 100%.
"""

import os
import re
import datetime
import openpyxl

# Metadata Kategori Kendaraan Nonlistrik (Kolom C s/d K · 9 Kolom)
NON_EV_CATEGORIES = [
    {"col": 3, "col_letter": "C", "code": "MCY", "display_code": "MCY", "name": "Sepeda motor, bajaj", "group": "SM", "k_type": "K1", "hotkey": "1"},
    {"col": 4, "col_letter": "D", "code": "CAR", "display_code": "CAR", "name": "Sedan, taksi, van, mobil penumpang lainnya", "group": "MP", "k_type": "K2", "hotkey": "2"},
    {"col": 5, "col_letter": "E", "code": "LGV", "display_code": "LGV", "name": "Mobil box, pick-up, kend. barang nontruk lainnya", "group": "KS/MP", "k_type": "K1", "hotkey": "3"},
    {"col": 6, "col_letter": "F", "code": "MGV", "display_code": "MGV", "name": "Truk 2-3 gandar", "group": "KS/TB", "k_type": "K1", "hotkey": "4"},
    {"col": 7, "col_letter": "G", "code": "HGV", "display_code": "HGV", "name": "Truk 4 gandar atau lebih", "group": "TB", "k_type": "K1", "hotkey": "5"},
    {"col": 8, "col_letter": "H", "code": "MIKROBUS", "display_code": "MIKROBUS", "name": "JAK, KWK, Mikrolet", "group": "MP", "k_type": "K2", "hotkey": "6"},
    {"col": 9, "col_letter": "I", "code": "LBUS", "display_code": "LBUS", "name": "Bus kecil", "group": "KS/BB", "k_type": "K2", "hotkey": "7"},
    {"col": 10, "col_letter": "J", "code": "MBUS", "display_code": "MBUS", "name": "Bus medium/normal", "group": "KS/BB", "k_type": "K2", "hotkey": "8"},
    {"col": 11, "col_letter": "K", "code": "HBUS", "display_code": "HBUS", "name": "Bus gandeng/tingkat", "group": "KS/BB", "k_type": "K2", "hotkey": "9"},
]

# Metadata Kategori Kendaraan Listrik / EV (Kolom L s/d Q · 6 Kolom) - Hanya 06.00-07.00
EV_CATEGORIES = [
    {"col": 12, "col_letter": "L", "code": "MCY_EV", "display_code": "MCY (Listrik)", "name": "Sepeda motor, bajaj", "group": "SM", "k_type": "K1", "hotkey": "Shift+1"},
    {"col": 13, "col_letter": "M", "code": "CAR_EV", "display_code": "CAR (Listrik)", "name": "Sedan, taksi, van, mobil penumpang lainnya", "group": "MP", "k_type": "K2", "hotkey": "Shift+2"},
    {"col": 14, "col_letter": "N", "code": "MIKROBUS_EV", "display_code": "MIKROBUS (Listrik)", "name": "JAK, KWK, Mikrolet", "group": "MP", "k_type": "K2", "hotkey": "Shift+3"},
    {"col": 15, "col_letter": "O", "code": "LBUS_EV", "display_code": "LBUS (Listrik)", "name": "Bus kecil", "group": "KS/BB", "k_type": "K2", "hotkey": "Shift+4"},
    {"col": 16, "col_letter": "P", "code": "MBUS_EV", "display_code": "MBUS (Listrik)", "name": "Bus medium/normal", "group": "KS/BB", "k_type": "K2", "hotkey": "Shift+5"},
    {"col": 17, "col_letter": "Q", "code": "HBUS_EV", "display_code": "HBUS (Listrik)", "name": "Bus gandeng/tingkat", "group": "KS/BB", "k_type": "K2", "hotkey": "Shift+6"},
]

# Definisi Rentang Waktu (64 interval per 15 menit: Baris 11 s/d 74)
def generate_intervals_metadata():
    intervals = []
    start_hour = 6
    start_minute = 0
    
    for row in range(11, 75):
        h1 = start_hour
        m1 = start_minute
        
        m2 = m1 + 15
        h2 = h1
        if m2 >= 60:
            m2 = 0
            h2 = h1 + 1
            
        time_label = f"{h1:02d}.{m1:02d}-{h2:02d}.{m2:02d}"
        
        if 11 <= row <= 14:
            period = "Peak Pagi (EV Aktif)"
            package_group = "PG"
            color = "blue"
            is_ev_active = True
        elif 15 <= row <= 26:
            period = "Peak Pagi"
            package_group = "PG"
            color = "blue"
            is_ev_active = False
        elif 27 <= row <= 50:
            period = "Siang (16 Jam)"
            package_group = "OFF"
            color = "yellow"
            is_ev_active = False
        elif 51 <= row <= 66:
            period = "Peak Sore"
            package_group = "SR"
            color = "blue"
            is_ev_active = False
        else: # 67 <= row <= 74
            period = "Malam (16 Jam)"
            package_group = "OFF"
            color = "yellow"
            is_ev_active = False
            
        intervals.append({
            "row": row,
            "time_label": time_label,
            "period": period,
            "package_group": package_group,
            "color": color,
            "is_ev_active": is_ev_active
        })
        
        start_hour = h2
        start_minute = m2

    return intervals

INTERVALS_METADATA = generate_intervals_metadata()
INTERVAL_BY_ROW = {item["row"]: item for item in INTERVALS_METADATA}


def parse_filename(filename):
    """
    Mengurai nama file template sesuai standar ITDP:
    TC_<titik>_Arah-<A/B>_@<HK/HL>[_PG/_SR][_K1/_K2]_x_<NAMA>.xlsx
    """
    clean_name = os.path.basename(filename)
    if clean_name.lower().endswith(".xlsx"):
        name_no_ext = clean_name[:-5]
    else:
        name_no_ext = clean_name
        
    result = {
        "raw_filename": clean_name,
        "titik": "1",
        "arah": "A",
        "hari": "HK",
        "package": "ALL",       # PG, SR, or ALL (16 jam)
        "package_label": "16 Jam Penuh (06.00 - 22.00)",
        "kelompok": "ALL",      # K1, K2, or ALL
        "pencacah": "Surveyor",
        "valid": False
    }
    
    parts_x = name_no_ext.split("_x_")
    if len(parts_x) == 2:
        result["pencacah"] = parts_x[1].strip()
        pkg_part = parts_x[0]
    else:
        pkg_part = name_no_ext
        
    m_titik = re.search(r"TC_([0-9A-Za-z]+)", pkg_part, re.IGNORECASE)
    if m_titik:
        result["titik"] = m_titik.group(1).upper()
        
    m_arah = re.search(r"Arah-([AB])", pkg_part, re.IGNORECASE)
    if m_arah:
        result["arah"] = m_arah.group(1).upper()
        
    m_hari = re.search(r"@([Hh][KkLl])", pkg_part)
    if m_hari:
        result["hari"] = m_hari.group(1).upper()
        
    if "_PG" in pkg_part.upper():
        result["package"] = "PG"
        result["package_label"] = "Peak Pagi (_PG, 06.00 - 10.00)"
    elif "_SR" in pkg_part.upper():
        result["package"] = "SR"
        result["package_label"] = "Peak Sore (_SR, 16.00 - 20.00)"
    else:
        result["package"] = "ALL"
        result["package_label"] = "16 Jam Penuh (06.00 - 22.00)"
        
    if "_K1" in pkg_part.upper():
        result["kelompok"] = "K1"
    elif "_K2" in pkg_part.upper():
        result["kelompok"] = "K2"
    else:
        result["kelompok"] = "ALL"
        
    result["valid"] = True
    return result


def lookup_index_info(ws_index, titik_target, arah="A", hari="HK"):
    """
    Mencocokkan informasi header dari sheet 'index' berdasarkan No. titik.
    Meniru rumus Excel: XLOOKUP(C2, index!H:H, index!...)
    """
    info = {
        "ruas": "Tidak Diketahui",
        "hari_tgl": "-",
        "dari": "a",
        "ke": "b",
        "tol_non": "Nontol",
        "tipe": "4/2 D",
        "lebar": 7,
        "kapasitas": 1,
        "perekam": "-"
    }
    
    if ws_index is None:
        return info

    titik_clean = str(titik_target).strip().upper()

    for r in range(2, ws_index.max_row + 1):
        cell_val = ws_index.cell(r, 8).value
        if cell_val is not None and str(cell_val).strip().upper() == titik_clean:
            info["ruas"] = ws_index.cell(r, 9).value or info["ruas"]
            
            tgl_raw = ws_index.cell(r, 11 if hari.upper() == "HK" else 13).value
            if isinstance(tgl_raw, (datetime.date, datetime.datetime)):
                info["hari_tgl"] = tgl_raw.strftime("%d/%m/%Y")
            elif tgl_raw:
                info["hari_tgl"] = str(tgl_raw)
                
            if arah.upper() == "A":
                info["dari"] = str(ws_index.cell(r, 18).value or "a")
                info["ke"] = str(ws_index.cell(r, 19).value or "b")
            else:
                info["dari"] = str(ws_index.cell(r, 20).value or "b")
                info["ke"] = str(ws_index.cell(r, 21).value or "a")
                
            info["tol_non"] = str(ws_index.cell(r, 22).value or "Nontol")
            info["tipe"] = str(ws_index.cell(r, 23).value or "4/2 D")
            info["lebar"] = ws_index.cell(r, 26).value or 7
            info["kapasitas"] = ws_index.cell(r, 27).value or 1
            break
            
    return info


def inspect_template(filepath):
    """
    Membaca struktur file template Excel tanpa merusak rumus apapun.
    Mengambil informasi header otomatis, catatan, dan seluruh data hitungan (existing counts).
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"File tidak ditemukan: {filepath}")
        
    wb = openpyxl.load_workbook(filepath, data_only=False)
    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]
    
    ws_index = wb["index"] if "index" in wb.sheetnames else None
    file_info = parse_filename(filepath)
    index_info = lookup_index_info(ws_index, file_info["titik"], file_info["arah"], file_info["hari"])
    
    catatan_val = ws.cell(4, 25).value
    catatan_str = str(catatan_val).strip() if catatan_val is not None else ""

    is_8_lanes = ("8/2 D" in index_info["tipe"]) or ("8" in index_info["tipe"] and "D" in index_info["tipe"])
    
    # Kelompok status: Hanya dikunci ke K1 atau K2 bila nama file secara eksplisit memuat _K1 atau _K2
    # Jika tidak ada _K1 atau _K2 di nama file, default adalah ALL (menampilkan SEMUA 15 kategori C s/d Q)
    if file_info["kelompok"] in ["K1", "K2"]:
        kelompok_status = file_info["kelompok"]
        kelompok_locked = True
    else:
        kelompok_status = "ALL"
        kelompok_locked = False

    header_info = {
        "sheet_name": sheet_name,
        "all_sheets": wb.sheetnames,
        "titik": file_info["titik"],
        "ruas": index_info["ruas"],
        "hari_tgl": f"{file_info['hari']} ({index_info['hari_tgl']})",
        "hk_hl": file_info["hari"],
        "arah": f"Arah {file_info['arah']}",
        "dari": index_info["dari"],
        "ke": index_info["ke"],
        "tipe": index_info["tipe"],
        "is_8_lanes": is_8_lanes,
        "tol_non": index_info["tol_non"],
        "pencacah": file_info["pencacah"],
        "catatan": catatan_str,
        "package": file_info["package"],
        "package_label": file_info["package_label"],
        "kelompok": kelompok_status,
        "kelompok_locked": kelompok_locked
    }

    counts_data = {}
    excel_view_rows = []
    
    total_existing_vehicles = 0
    total_existing_filled_sessions = 0

    for item in INTERVALS_METADATA:
        row = item["row"]
        row_counts = {}
        row_has_counts = False
        
        # Non-EV (Kolom C s/d K · Kolom 3 s/d 11)
        for cat in NON_EV_CATEGORIES:
            val = ws.cell(row, cat["col"]).value
            if val is not None:
                try:
                    num = int(round(float(val)))
                    row_counts[cat["code"]] = num
                    row_has_counts = True
                    total_existing_vehicles += num
                except (ValueError, TypeError):
                    row_counts[cat["code"]] = 0
            else:
                row_counts[cat["code"]] = 0
                
        # EV (Kolom L s/d Q · Kolom 12 s/d 17) - Hanya baris 11 s/d 14
        for cat in EV_CATEGORIES:
            if item["is_ev_active"]:
                val = ws.cell(row, cat["col"]).value
                if val is not None:
                    try:
                        num = int(round(float(val)))
                        row_counts[cat["code"]] = num
                        row_has_counts = True
                        total_existing_vehicles += num
                    except (ValueError, TypeError):
                        row_counts[cat["code"]] = 0
                else:
                    row_counts[cat["code"]] = 0
            else:
                row_counts[cat["code"]] = None
                
        status = "empty"
        if row_has_counts:
            sum_vals = sum(v for v in row_counts.values() if v is not None)
            if sum_vals > 0 or all(v is not None for v in row_counts.values()):
                status = "verified"
                total_existing_filled_sessions += 1
            else:
                status = "draft"

        non_ev_subtotal = sum(row_counts[c["code"]] for c in NON_EV_CATEGORIES if row_counts.get(c["code"]) is not None)
        ev_subtotal = sum(row_counts[c["code"]] for c in EV_CATEGORIES if row_counts.get(c["code"]) is not None) if item["is_ev_active"] else 0

        counts_data[str(row)] = {
            "row": row,
            "counts": row_counts,
            "status": status,
            "total_non_ev": non_ev_subtotal,
            "total_ev": ev_subtotal,
            "total": non_ev_subtotal + ev_subtotal
        }

        # Live calculation untuk tampilan Excel View
        mcy_tot = row_counts.get("MCY", 0) + (row_counts.get("MCY_EV", 0) if item["is_ev_active"] else 0)
        car_tot = row_counts.get("CAR", 0) + (row_counts.get("CAR_EV", 0) if item["is_ev_active"] else 0)
        lgv_tot = row_counts.get("LGV", 0)
        mgv_tot = row_counts.get("MGV", 0)
        hgv_tot = row_counts.get("HGV", 0)
        mikro_tot = row_counts.get("MIKROBUS", 0) + (row_counts.get("MIKROBUS_EV", 0) if item["is_ev_active"] else 0)
        bus_tot = (
            row_counts.get("LBUS", 0) + row_counts.get("MBUS", 0) + row_counts.get("HBUS", 0) +
            ((row_counts.get("LBUS_EV", 0) + row_counts.get("MBUS_EV", 0) + row_counts.get("HBUS_EV", 0)) if item["is_ev_active"] else 0)
        )
        total_15mnt_non_mcy = car_tot + lgv_tot + mgv_tot + hgv_tot + mikro_tot + bus_tot
        total_15mnt_all = total_15mnt_non_mcy + mcy_tot

        excel_view_rows.append({
            "row": row,
            "time_label": item["time_label"],
            "color": item["color"],
            "is_ev_active": item["is_ev_active"],
            "period": item["period"],
            "counts": row_counts,
            "formula_totals": {
                "R_MCY": mcy_tot,
                "S_CAR": car_tot,
                "T_LGV": lgv_tot,
                "U_MGV": mgv_tot,
                "V_HGV": hgv_tot,
                "W_MIKRO": mikro_tot,
                "X_BUS": bus_tot,
                "Y_NON_MCY": total_15mnt_non_mcy,
                "Z_TOTAL": total_15mnt_all
            }
        })

    wb.close()

    return {
        "filepath": filepath,
        "filename": os.path.basename(filepath),
        "file_info": file_info,
        "header_info": header_info,
        "intervals": INTERVALS_METADATA,
        "categories_non_ev": NON_EV_CATEGORIES,
        "categories_ev": EV_CATEGORIES,
        "counts_data": counts_data,
        "excel_view_rows": excel_view_rows,
        "existing_stats": {
            "total_vehicles": total_existing_vehicles,
            "filled_sessions": total_existing_filled_sessions
        }
    }


def write_verified_data_to_excel(template_filepath, session_data_dict, note=None, output_filepath=None):
    """
    Menulis hasil counting terverifikasi ke file Excel dengan AMAN dan PRESISI.
    
    JAMINAN KEAMANAN:
    - HANYA menulis ke sel biru & kuning di sheet pertama:
      - Baris 11-14 (06.00-07.00): Kolom 3-17 (C s/d Q)
      - Baris 15-74 (07.00-22.00): Kolom 3-11 (C s/d K saja)
      - Baris 4: Kolom 25 (Y4 untuk Catatan)
    - TIDAK PERNAH menyentuh kolom 18 ke atas (R s/d AI berisi rumus)
    - TIDAK PERNAH menyentuh baris 1-3 (rumus header)
    - TIDAK PERNAH menyentuh sheet index
    """
    if not os.path.exists(template_filepath):
        raise FileNotFoundError(f"File template tidak ditemukan: {template_filepath}")
        
    target_path = output_filepath if output_filepath else template_filepath
    
    wb = openpyxl.load_workbook(template_filepath, data_only=False)
    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]
    
    written_cells_log = []
    
    # 1. Update Catatan jika ada
    if note is not None:
        ws.cell(4, 25).value = str(note).strip()
        written_cells_log.append({"cell": "Y4", "val": str(note).strip(), "desc": "Catatan pengamatan"})
        
    # 2. Tulis data tiap sesi 15 menit
    for row_str, session in session_data_dict.items():
        try:
            row = int(row_str)
        except ValueError:
            continue
            
        if not (11 <= row <= 74):
            continue
            
        counts = session.get("counts", {})
        is_ev_active = (11 <= row <= 14)
        
        # Tulis Non-EV (Kolom C s/d K -> Kolom 3 s/d 11)
        for cat in NON_EV_CATEGORIES:
            col = cat["col"]
            code = cat["code"]
            raw_val = counts.get(code)
            
            val = int(raw_val) if raw_val is not None and str(raw_val).strip() != "" else 0
            
            if 3 <= col <= 11:
                ws.cell(row=row, column=col).value = val
                written_cells_log.append({
                    "cell": f"{cat['col_letter']}{row}",
                    "val": val,
                    "desc": f"{cat['code']} ({INTERVAL_BY_ROW[row]['time_label']})"
                })
                
        # Tulis EV (Kolom L s/d Q -> Kolom 12 s/d 17) HANYA untuk baris 11 s/d 14
        if is_ev_active:
            for cat in EV_CATEGORIES:
                col = cat["col"]
                code = cat["code"]
                raw_val = counts.get(code)
                val = int(raw_val) if raw_val is not None and str(raw_val).strip() != "" else 0
                
                if 12 <= col <= 17:
                    ws.cell(row=row, column=col).value = val
                    written_cells_log.append({
                        "cell": f"{cat['col_letter']}{row}",
                        "val": val,
                        "desc": f"{cat['code']} ({INTERVAL_BY_ROW[row]['time_label']})"
                    })
                    
    wb.save(target_path)
    wb.close()
    
    return {
        "success": True,
        "saved_path": target_path,
        "cells_updated_count": len(written_cells_log),
        "log_sample": written_cells_log[:15]
    }


def perform_qc_check(counts_data, package="ALL", kelompok="ALL"):
    """
    Pemeriksaan Quality Control (QC) otomatis sesuai checklist slide 21 ITDP
    """
    issues = []
    passed = []
    
    if package == "PG":
        required_rows = list(range(11, 27))
    elif package == "SR":
        required_rows = list(range(51, 67))
    else:
        required_rows = list(range(11, 75))
        
    total_required = len(required_rows)
    filled_sessions = 0
    total_vehicles = 0
    interval_totals = {}
    
    for r in required_rows:
        sess = counts_data.get(str(r), {})
        counts = sess.get("counts", {})
        
        missing_in_row = []
        row_tot = 0
        for cat in NON_EV_CATEGORIES:
            if kelompok == "K1" and cat["k_type"] != "K1":
                continue
            if kelompok == "K2" and cat["k_type"] != "K2":
                continue
                
            val = counts.get(cat["code"])
            if val is None:
                missing_in_row.append(cat["code"])
            else:
                row_tot += val
                
        if 11 <= r <= 14:
            for cat in EV_CATEGORIES:
                if kelompok == "K1" and cat["k_type"] != "K1":
                    continue
                if kelompok == "K2" and cat["k_type"] != "K2":
                    continue
                val = counts.get(cat["code"])
                if val is None:
                    missing_in_row.append(cat["code"])
                else:
                    row_tot += val
                    
        interval_totals[r] = row_tot
        total_vehicles += row_tot
        
        if not missing_in_row:
            filled_sessions += 1
        else:
            issues.append({
                "type": "incomplete_row",
                "row": r,
                "time": INTERVAL_BY_ROW[r]["time_label"],
                "message": f"Interval {INTERVAL_BY_ROW[r]['time_label']} belum lengkap: {', '.join(missing_in_row)} belum diisi."
            })
            
    if filled_sessions == total_required:
        passed.append(f"Semua {total_required} interval paket {package} telah terisi lengkap ({filled_sessions}/{total_required}).")
    else:
        issues.append({
            "type": "completeness",
            "message": f"Kelengkapan paket: baru {filled_sessions} dari {total_required} interval yang selesai."
        })
        
    for i in range(len(required_rows) - 1):
        r1 = required_rows[i]
        r2 = required_rows[i+1]
        tot1 = interval_totals.get(r1, 0)
        tot2 = interval_totals.get(r2, 0)
        
        if tot1 > 10 and tot2 > tot1 * 4 and (tot2 - tot1) > 100:
            issues.append({
                "type": "anomaly_jump",
                "row": r2,
                "time": INTERVAL_BY_ROW[r2]["time_label"],
                "message": f"Peringatan lonjakan tajam: dari {tot1} kendaraan ({INTERVAL_BY_ROW[r1]['time_label']}) melonjak ke {tot2} kendaraan ({INTERVAL_BY_ROW[r2]['time_label']}). Mohon cek kembali video."
            })
            
    return {
        "status": "PASS" if not [i for i in issues if i["type"] in ["incomplete_row", "completeness"]] else "WARN",
        "passed_checks": passed,
        "issues": issues,
        "total_vehicles": total_vehicles,
        "progress_percent": round((filled_sessions / total_required) * 100, 1) if total_required else 0
    }
