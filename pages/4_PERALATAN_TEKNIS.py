import streamlit as st
import pandas as pd
import sqlite3
import os
from datetime import datetime
from fpdf import FPDF

# --- KONFIGURASI AWAL ---
UPLOAD_DIR = "uploads"
if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

# --- FUNGSI DATABASE SQLITE ---
def init_db():
    conn = sqlite3.connect('kinerja_teknisi.db')
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS e_kinerja (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tanggal TEXT,
            nama_teknisi TEXT,
            penjelasan_kegiatan TEXT,
            foto_path TEXT
        )
    ''')
    conn.commit()
    return conn

conn = init_db()

# --- FUNGSI GENERATOR PDF (E-KINERJA) ---
class PDFKinerja(FPDF):
    def header(self):
        # Header Kop Surat BMKG
        self.set_font('helvetica', 'B', 12)
        self.cell(0, 5, 'BADAN METEOROLOGI, KLIMATOLOGI, DAN GEOFISIKA', align='C', new_x='LMARGIN', new_y='NEXT')
        self.cell(0, 5, 'STASIUN METEOROLOGI KELAS III UMBU MEHANG KUNDA', align='C', new_x='LMARGIN', new_y='NEXT')
        self.set_font('helvetica', '', 9)
        self.cell(0, 5, 'Jl. Adi Sucipto, Waingapu, Sumba Timur', align='C', new_x='LMARGIN', new_y='NEXT')
        self.cell(0, 5, 'Email: stamet.sumbatimur@bmkg.go.id', align='C', new_x='LMARGIN', new_y='NEXT')
        self.line(10, 30, 200, 30)
        self.ln(10)

def generate_pdf(nama_teknisi, bulan_tahun, df_kegiatan):
    pdf = PDFKinerja()
    pdf.add_page()
    
    # Judul Dokumen
    pdf.set_font('helvetica', 'B', 11)
    pdf.cell(0, 10, f'LAPORAN E-KINERJA TEKNISI', align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.cell(0, 5, f'BULAN {bulan_tahun.upper()}', align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.ln(5)
    
    # Narasi Kinerja & Formula OLA/SLA
    pdf.set_font('helvetica', '', 10)
    teks_narasi = (
        "Persentase Alat Operasional Utama Meteorologi yang Laik Operasi dengan "
        "Target 97% di Stasiun Meteorologi Kelas III Umbu Mehang Kunda telah dihitung. "
        "Berikut adalah lampiran dokumentasi pemeliharaan yang telah dilaksanakan:"
    )
    pdf.multi_cell(0, 5, teks_narasi)
    pdf.ln(5)
    
    # Tabel Lampiran Kegiatan
    pdf.set_font('helvetica', 'B', 10)
    pdf.cell(0, 8, f'LAMPIRAN KEGIATAN TEKNISI: {nama_teknisi.upper()}', align='L', new_x='LMARGIN', new_y='NEXT')
    pdf.set_font('helvetica', '', 10)
    
    # Looping Data dari Database untuk Dimasukkan ke PDF
    for index, row in df_kegiatan.iterrows():
        pdf.set_font('helvetica', 'B', 10)
        pdf.cell(0, 8, f"Tanggal: {row['tanggal']}", new_x='LMARGIN', new_y='NEXT')
        pdf.set_font('helvetica', '', 10)
        pdf.multi_cell(0, 6, f"Kegiatan: {row['penjelasan_kegiatan']}")
        
        # Insert Foto jika ada
        if pd.notna(row['foto_path']) and os.path.exists(row['foto_path']):
            try:
                # Mengatur ukuran gambar agar muat di PDF
                pdf.image(row['foto_path'], w=100)
            except:
                pdf.cell(0, 5, "[Gambar gagal dimuat]", new_x='LMARGIN', new_y='NEXT')
        pdf.ln(5)
        
    nama_file = f"E_Kinerja_{nama_teknisi.replace(' ', '_')}_{bulan_tahun.replace(' ', '')}.pdf"
    pdf.output(nama_file)
    return nama_file

# --- ANTARMUKA STREAMLIT ---
st.title("🛠️ Laporan & E-Kinerja Teknisi")

# Menggunakan 3 Tab
tab1, tab2, tab3 = st.tabs(["📝 Input Kegiatan Harian", "📅 Data Kinerja (DB)", "🖨️ Cetak PDF E-Kinerja"])

# TAB 1: INPUT KEGIATAN HARIAN
with tab1:
    st.subheader("Form Input Dokumentasi Kegiatan")
    with st.form("form_kinerja", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            tanggal = st.date_input("Tanggal Kegiatan", datetime.today())
            nama = st.selectbox("Nama Teknisi", ["Zulqha Ariandi Al Zikri, S.Tr.Inst.", "Adi Junaidi Rachman, S.Kom.", "Luqmanul Hakim, S.Tr.", "Mohammad Hasyim Hanif, S.Tr.Inst."])
        with col2:
            kegiatan = st.text_area("Penjelasan Kegiatan (Misal: Pemeliharaan Taman Alat)", height=100)
            
        foto = st.file_uploader("Upload Foto Dokumentasi", type=['jpg', 'jpeg', 'png'])
        
        submit = st.form_submit_button("Simpan Data")
        
        if submit:
            if foto is not None:
                # Simpan foto ke folder lokal
                file_path = os.path.join(UPLOAD_DIR, foto.name)
                with open(file_path, "wb") as f:
                    f.write(foto.getbuffer())
            else:
                file_path = None
                
            # Simpan data ke SQLite
            c = conn.cursor()
            c.execute("INSERT INTO e_kinerja (tanggal, nama_teknisi, penjelasan_kegiatan, foto_path) VALUES (?, ?, ?, ?)",
                      (tanggal.strftime("%Y-%m-%d"), nama, kegiatan, file_path))
            conn.commit()
            st.success("Data kegiatan berhasil disimpan ke database!")

# TAB 2: DATA KINERJA (DB)
with tab2:
    st.subheader("Arsip Kegiatan Teknisi")
    df_db = pd.read_sql("SELECT * FROM e_kinerja", conn)
    if not df_db.empty:
        st.dataframe(df_db, use_container_width=True)
    else:
        st.info("Belum ada data kegiatan yang diinput.")

# TAB 3: CETAK PDF E-KINERJA
with tab3:
    st.subheader("Generate Dokumen E-Kinerja")
    st.markdown("Pilih teknisi dan bulan untuk menghasilkan dokumen PDF siap cetak seperti format laporan BMKG.")
    
    col_a, col_b = st.columns(2)
    with col_a:
        filter_nama = st.selectbox("Pilih Teknisi untuk Dicetak", ["Zulqha Ariandi Al Zikri, S.Tr.Inst.", "Adi Junaidi Rachman, S.Kom.", "Luqmanul Hakim, S.Tr.", "Mohammad Hasyim Hanif, S.Tr.Inst."])
    with col_b:
        filter_bulan = st.selectbox("Pilih Bulan & Tahun", ["Juli 2026", "Agustus 2026", "September 2026"])
    
    if st.button("Tarik Data & Generate PDF 🚀"):
        # Ambil data spesifik dari DB berdasarkan nama (Untuk akurasi sebaiknya filter bulan juga diterapkan di SQL)
        df_filter = pd.read_sql(f"SELECT * FROM e_kinerja WHERE nama_teknisi = '{filter_nama}'", conn)
        
        if not df_filter.empty:
            with st.spinner('Menyusun dokumen PDF...'):
                pdf_file = generate_pdf(filter_nama, filter_bulan, df_filter)
                
            st.success(f"File PDF berhasil dibuat: {pdf_file}")
            
            # Tombol Unduh
            with open(pdf_file, "rb") as pdf_data:
                st.download_button(
                    label="⬇️ Download PDF E-Kinerja",
                    data=pdf_data,
                    file_name=pdf_file,
                    mime="application/pdf"
                )
        else:
            st.warning("Tidak ada data kegiatan untuk teknisi tersebut di database.")
