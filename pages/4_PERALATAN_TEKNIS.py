import streamlit as st
import pandas as pd
import os

# --- KONFIGURASI HALAMAN ---
# Hapus st.set_page_config jika sudah dideklarasikan di halaman utama (main.py)
# st.set_page_config(page_title="Peralatan Teknis", page_icon="🛠️", layout="wide")

st.title("🛠️ Laporan Bulanan Peralatan Teknis")
st.markdown("""
Halaman ini digunakan untuk memonitoring Jadwal Pemeliharaan, Logbook Alat, dan OLA/SLA Peralatan Meteorologi Utama.
""")

# --- FUNGSI UNTUK MEMBACA & MEMBERSIHKAN DATA EXCEL ---
@st.cache_data
def load_data(uploaded_file):
    try:
        # Membaca seluruh sheet
        xls = pd.ExcelFile(uploaded_file)
        
        # Dictionary untuk menyimpan dataframe yang sudah di-load
        sheets_data = {}
        
        for sheet in xls.sheet_names:
            # Baca sheet, biarkan header None agar kita bisa membersihkan baris kosong di atas
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            
            # Cleaning Dasar: Buang baris & kolom yang 100% kosong (NaN)
            df_cleaned = df.dropna(how='all').dropna(axis=1, how='all')
            
            # Reset index agar rapi
            df_cleaned = df_cleaned.reset_index(drop=True)
            
            # Simpan ke dictionary
            sheets_data[sheet] = df_cleaned
            
        return sheets_data
    except Exception as e:
        st.error(f"Terjadi kesalahan saat membaca file: {e}")
        return None

# --- BAGIAN INPUT/UPLOAD DATA ---
st.sidebar.header("📁 Pengaturan Data")
# Cek apakah file sudah ada secara lokal untuk efisiensi
default_file_path = "Peralatan Teknis.xlsx"

if os.path.exists(default_file_path):
    st.sidebar.success(f"File sumber ditemukan: {default_file_path}")
    data = load_data(default_file_path)
else:
    st.sidebar.warning("File default tidak ditemukan. Silakan upload manual.")
    uploaded_file = st.sidebar.file_uploader("Upload File Peralatan Teknis (Excel)", type=["xlsx"])
    if uploaded_file is not None:
        data = load_data(uploaded_file)
    else:
        data = None

# --- MENAMPILKAN DATA DI WEB ---
if data:
    # Mengambil nama-nama sheet yang ada
    sheet_names = list(data.keys())
    
    # Membuat Tab untuk masing-masing bagian laporan
    tabs = st.tabs(["📅 Jadwal Pemeliharaan", "📝 Logbook Kondisi Alat", "📊 Monitoring OLA & SLA"])
    
    # TAB 1: JADWAL
    with tabs[0]:
        st.subheader("Jadwal Pemeliharaan Teknisi")
        if 'JADWAL 2026' in sheet_names:
            df_jadwal = data['JADWAL 2026']
            # Menggunakan dataframe bawaan streamlit dengan height disesuaikan
            st.dataframe(df_jadwal, use_container_width=True, height=500)
            st.info("💡 **Tips:** Anda bisa mendownload tabel ini langsung dengan mengklik tombol ikon unduh yang muncul saat kursor diarahkan ke tabel.")
        else:
            st.warning("Sheet 'JADWAL 2026' tidak ditemukan.")

    # TAB 2: LOGBOOK
    with tabs[1]:
        st.subheader("Logbook Monitoring Alat Utama")
        if 'LOGBOOK' in sheet_names:
            df_logbook = data['LOGBOOK']
            st.dataframe(df_logbook, use_container_width=True, height=500)
        else:
            st.warning("Sheet 'LOGBOOK' tidak ditemukan.")

    # TAB 3: OLA & SLA
    with tabs[2]:
        st.subheader("Monitoring Harian OLA & SLA")
        if 'OLA SLA 2026' in sheet_names:
            df_olasla = data['OLA SLA 2026']
            st.dataframe(df_olasla, use_container_width=True, height=500)
            
            # --- CONTOH TAMBAHAN VISUALISASI SEDERHANA ---
            st.markdown("---")
            st.markdown("### Ringkasan Status (Eksperimental)")
            st.caption("Menghitung kemunculan status 'ON' dan 'OFF' secara cepat dari tabel.")
            
            # Konversi semua data menjadi string dan hitung kemunculan kata 'ON'
            total_on = df_olasla.astype(str).apply(lambda x: x.str.contains('ON', case=True, na=False)).sum().sum()
            total_off = df_olasla.astype(str).apply(lambda x: x.str.contains('OFF', case=True, na=False)).sum().sum()
            
            col1, col2, col3 = st.columns(3)
            col1.metric("Status Peralatan ON 🟢", total_on)
            col2.metric("Status Peralatan OFF 🔴", total_off)
        else:
            st.warning("Sheet 'OLA SLA 2026' tidak ditemukan.")
else:
    st.info("Silakan pastikan file Excel tersedia untuk memuat laporan.")