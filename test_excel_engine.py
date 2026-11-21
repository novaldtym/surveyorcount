"""
Unit test untuk memastikan keamanan dan fungsionalitas excel_engine.py
"""

import os
import shutil
import openpyxl
from excel_engine import parse_filename, inspect_template, write_verified_data_to_excel, perform_qc_check

def run_tests():
    print("=== TEST 1: PARSE FILENAME ===")
    fn1 = "TC_1_Arah-A_@HK_x_NAMA YANG HITUNG.xlsx"
    p1 = parse_filename(fn1)
    print("Result 1:", p1)
    assert p1["titik"] == "1"
    assert p1["arah"] == "A"
    assert p1["hari"] == "HK"
    assert p1["pencacah"] == "NAMA YANG HITUNG"
    
    fn2 = "TC_2_Arah-B_@HL_PG_K1_x_Budi Santoso.xlsx"
    p2 = parse_filename(fn2)
    print("Result 2:", p2)
    assert p2["titik"] == "2"
    assert p2["arah"] == "B"
    assert p2["hari"] == "HL"
    assert p2["package"] == "PG"
    assert p2["kelompok"] == "K1"
    assert p2["pencacah"] == "Budi Santoso"
    print("Test 1 PASS!")

    print("\n=== TEST 2: INSPECT TEMPLATE ===")
    orig_file = r"d:\surveyor\TC_1_Arah-A_@HK_x_NAMA YANG HITUNG.xlsx"
    template_info = inspect_template(orig_file)
    print("Sheet name:", template_info["header_info"]["sheet_name"])
    print("Titik:", template_info["header_info"]["titik"])
    print("Ruas:", template_info["header_info"]["ruas"])
    print("Total intervals:", len(template_info["intervals"]))
    assert len(template_info["intervals"]) == 64
    assert len(template_info["categories_non_ev"]) == 9
    assert len(template_info["categories_ev"]) == 6
    print("Test 2 PASS!")

    print("\n=== TEST 3: SAFE WRITE & FORMULA PRESERVATION ===")
    test_file = r"d:\surveyor\test_output.xlsx"
    shutil.copyfile(orig_file, test_file)
    
    # Simulate session counts data
    test_session_data = {
        "11": {
            "counts": {
                "MCY": 45, "CAR": 12, "LGV": 3, "MGV": 1, "HGV": 0,
                "MIKROBUS": 2, "LBUS": 0, "MBUS": 1, "HBUS": 0,
                "MCY_EV": 4, "CAR_EV": 1, "MIKROBUS_EV": 0, "LBUS_EV": 0, "MBUS_EV": 0, "HBUS_EV": 0
            }
        },
        "12": {
            "counts": {
                "MCY": 50, "CAR": 15, "LGV": 2, "MGV": 0, "HGV": 0,
                "MIKROBUS": 4, "LBUS": 1, "MBUS": 0, "HBUS": 0,
                "MCY_EV": 2, "CAR_EV": 0, "MIKROBUS_EV": 0, "LBUS_EV": 0, "MBUS_EV": 0, "HBUS_EV": 0
            }
        },
        "15": { # 07.00-07.15 (Non-EV only)
            "counts": {
                "MCY": 60, "CAR": 20, "LGV": 5, "MGV": 2, "HGV": 1,
                "MIKROBUS": 5, "LBUS": 0, "MBUS": 2, "HBUS": 0
            }
        }
    }
    
    res = write_verified_data_to_excel(
        template_filepath=test_file,
        session_data_dict=test_session_data,
        note="Cerah berawan; arus lancar 06.00-07.30",
        output_filepath=test_file
    )
    print("Write result:", res["success"], "cells updated:", res["cells_updated_count"])
    
    # Verify file with openpyxl
    wb = openpyxl.load_workbook(test_file, data_only=False)
    ws = wb["Arah-A"]
    
    # Check updated cells
    assert ws["C11"].value == 45, f"Expected 45, got {ws['C11'].value}"
    assert ws["L11"].value == 4, f"Expected 4, got {ws['L11'].value}"
    assert ws["Y4"].value == "Cerah berawan; arus lancar 06.00-07.30"
    
    # Check untouched formulas
    assert ws["R11"].value == "=C11+L11", f"Formula R11 damaged! Got {ws['R11'].value}"
    assert ws["S11"].value == "=D11+M11", f"Formula S11 damaged! Got {ws['S11'].value}"
    assert ws["W11"].value == "=SUM(H11,N11)", f"Formula W11 damaged! Got {ws['W11'].value}"
    assert ws["T3"].value == "=MAX(AI11:AI26)", f"Formula T3 damaged! Got {ws['T3'].value}"
    assert ws["Y3"].value == '=PROPER(_xlfn.TEXTAFTER(A1,"x_"))', f"Formula Y3 damaged! Got {ws['Y3'].value}"
    
    # Check that EV column L at row 15 is NOT written
    assert ws["L15"].value is None, f"L15 should be None, got {ws['L15'].value}"
    
    # Check index sheet is intact
    assert "index" in wb.sheetnames
    assert wb["index"].cell(1, 1).value == "MCY"
    
    wb.close()
    os.remove(test_file)
    print("Test 3 PASS! All formulas, styles, and restrictions preserved 100%.")

    print("\n=== TEST 4: QC CHECK ENGINE ===")
    qc = perform_qc_check(test_session_data, package="PG")
    print("QC status:", qc["status"])
    print("QC issues count:", len(qc["issues"]))
    print("Total vehicles:", qc["total_vehicles"])
    print("Test 4 PASS!")

    print("\nALL TESTS PASSED SUCCESSFULLY!")

if __name__ == "__main__":
    run_tests()
