import streamlit as st
import gspread
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
import pandas as pd
import re
import io
import warnings

# Mengabaikan warning kecil dari Pandas
warnings.filterwarnings('ignore')

# Konfigurasi Halaman Web Streamlit
st.set_page_config(
    page_title="Rekap Absensi Vendor",
    page_icon="📊",
    layout="centered"
)

# Judul Utama
st.title("📊 Aplikasi Otomatisasi Absensi")
st.markdown("Rekap Absensi Vendor Warehouse")
st.divider()

# Fungsi Bantuan
def extract_spreadsheet_id(input_text):
    """Mengekstrak Spreadsheet ID dari link URL"""
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", input_text)
    if match:
        return match.group(1)
    return input_text.strip()

def parse_shift(val):
    """Mengekstrak format Shift"""
    if pd.isna(val):
        return "S1"
    val_str = str(val).strip()
    match = re.search(r'(\d+)', val_str)
    if match:
        return f"S{match.group(1)}"
    return "S1"

# --- UI FORM INPUT ---
link_input = st.text_input("Masukkan Link Google Sheets disini", placeholder="https://docs.google.com/spreadsheets/d/...")
nama_input = st.text_input("Nama File Karyawan", placeholder="Cth: John Doe")

col1, col2 = st.columns(2)
with col1:
    tahun_input = st.selectbox("Pilih Tahun", ["2025", "2026", "2027", "2028", "2029", "2030"], index=1)
with col2:
    bulan_input = st.selectbox("Pilih Bulan", ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"], index=6)

st.divider()

# --- LOGIKA TOMBOL GENERATE ---
if st.button("🚀 Tarik Data & Generate Excel", use_container_width=True):
    if not link_input:
        st.warning("⚠️ Silakan masukkan link Google Sheets terlebih dahulu!")
    elif not nama_input:
        st.warning("⚠️ Ketik dulu nama karyawannya untuk nama file Excel!")
    else:
        # Menampilkan animasi loading saat memproses
        with st.spinner("Menghubungi Google Server dan Memproses Data..."):
            spreadsheet_id = extract_spreadsheet_id(link_input)
            
            try:
                # 1. TARIK DATA DARI GOOGLE SHEETS
                gc = gspread.service_account(filename="credentials.json")
                sheet = gc.open_by_key(spreadsheet_id).sheet1
import os
import json

def get_gspread_client():
    """Get gspread client from credentials (env var or file)"""
    cred_path = "credentials.json"
    
    # Try env var first
    if os.getenv("GOOGLE_CREDENTIALS"):
        import tempfile
        creds = json.loads(os.getenv("GOOGLE_CREDENTIALS"))
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            json.dump(creds, f)
            temp_cred_path = f.name
        return gspread.service_account(filename=temp_cred_path)
    
    # Fallback to file
    if os.path.exists(cred_path):
        return gspread.service_account(filename=cred_path)
    
    raise FileNotFoundError("credentials.json not found and GOOGLE_CREDENTIALS env var not set")

# Replace the original line with this:
# gc = gspread.service_account(filename="credentials.json")
