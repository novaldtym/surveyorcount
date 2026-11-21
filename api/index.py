import sys
import os

# Pastikan root directory ada di sys.path agar modul excel_engine dan app terbaca
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app
