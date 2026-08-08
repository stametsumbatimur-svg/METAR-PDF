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

# Path file sumber
EXCEL_FILE = "Peralatan Teknis.xlsx"
LOGO_FILE = "logo_bmkg.png" # Pastikan file logo BMKG ada di folder yang sama

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
        if os.path.exists(LOGO_FILE):
            self.image(LOGO_FILE, 10, 8, 25) # Menyesuaikan ukuran logo
        
        self.set_font('helvetica', 'B', 12)
        self.set_left_margin(40) 
        self.cell(0, 5, 'BADAN METEOROLOGI, KLIMATOLOGI, DAN GEOFISIKA', align='C', new_x='LMARGIN', new_y='NEXT')
        self.cell(0, 5, 'STASIUN METEOROLOGI KELAS III UMBU MEHANG KUNDA', align='C', new_x='LMARGIN', new_y='NEXT')
        
        self.set_font('helvetica', '', 10)
        self.cell(0, 5, 'Jl. Adi Sucipto, Waingapu, Sumba Timur', align='C', new_x='LMARGIN', new_y='NEXT')
        self.cell(0, 5, 'Telp. (0387) 61227 | Fax: (0387) 61228 | Kode Pos 87114', align='C', new_x='LMARGIN', new_y='NEXT')
        self.cell(0, 5, 'Email: stamet.sumbatimur@bmkg.go.id | Website: http://ntt.bmkg.go.id', align='C', new_x='LMARGIN', new_y='NEXT')
        
        # Garis Kop Surat
        self.set_left_margin(10)
        self.set_line_width(0.8)
        self.line(10, 34, 200, 34)
        self.set_line_width(0.2)
        self.line(10, 35, 200, 35)
        self.ln(15)

def generate_pdf(nama_teknisi, bulan, tahun, df_kegiatan, excel_data):
    pdf = PDFKinerja()
    
    # --- HALAMAN 1: NARASI RUMUS (PORTRAIT) ---
    pdf.add_page(orientation='portrait')
    pdf.set_font('helvetica', 'B', 11)
    pdf.cell(0, 6, 'MENINGKATNYA LAYANAN OPERASIONAL', align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.cell(0, 6, 'ALOPTAMA METEOROLOGI YANG PRIMA', align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.cell(0, 6, 'STASIUN METEOROLOGI KELAS III UMBU MEHANG KUNDA', align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.cell(0, 6, f'BULAN {bulan.upper()} TAHUN {tahun}', align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.ln(8)
    
    pdf.set_font('helvetica', '', 11)
    teks_pembuka = (
        "Persentase Alat Operasional Utama Meteorologi yang Laik Operasi dengan "
        "Target 97% di Stasiun Meteorologi Kelas III Umbu Mehang Kunda diperoleh "
        "menggunakan formula perhitungan sebagai berikut:"
    )
    pdf.multi_cell(0, 6, teks_pembuka)
    pdf.ln(5)
    
    # Perbaikan Error FPDF Unicode: Mengganti simbol 'Sigma' dengan kata 'Jumlah'
    pdf.set_font('helvetica', 'I', 11)
    pdf.cell(0, 6, "Laik Operasi Aloptama MET = (Jumlah aloptama meteorologi yg terpelihara / Jumlah Aloptama meteorologi) x 100%", align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.ln(3)
    pdf.set_font('helvetica', 'B', 11)
    pdf.cell(0, 6, "Laik Operasi Aloptama MET = 24/24 x 100%", align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.ln(5)
    
    pdf.set_font('helvetica', '', 11)
    teks_penutup = (
        f"Sehingga hasil persentase untuk peralatan operasional pada Triwulan terkait "
        f"menghasilkan nilai akurasi sebesar 100%. Adapun capaian hasil tersebut "
        "disusun dan didukung oleh laporan yang menjadi dasar pemantauan, pemeliharaan, "
        "serta evaluasi kinerja peralatan operasional utama meteorologi, yaitu sebagai berikut:"
    )
    pdf.multi_cell(0, 6, teks_penutup)
    pdf.ln(5)
    
    pdf.cell(10, 6, "1.")
    pdf.cell(0, 6, "Laporan Hasil Monitoring Kondisi Peralatan Operasional Utama Meteorologi", new_x='LMARGIN', new_y='NEXT')
    pdf.set_x(10)
    pdf.cell(10, 6, "2.")
    pdf.cell(0, 6, "Laporan Hasil Pemeliharaan Preventif Peralatan Operasional Utama Meteorologi", new_x='LMARGIN', new_y='NEXT')
    pdf.set_x(10)
    pdf.cell(10, 6, "3.")
    pdf.cell(0, 6, "Laporan Ketersediaan Data Peralatan Otomatis", new_x='LMARGIN', new_y='NEXT')
    
    # --- HALAMAN 2: LOGBOOK (LANDSCAPE) ---
    pdf.add_page(orientation='landscape')
    pdf.set_font('helvetica', 'B', 10)
    pdf.cell(0, 8, 'LAPORAN HASIL MONITORING KONDISI PERALATAN (LOGBOOK)', align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.ln(3)
    
    if excel_data and 'LOGBOOK' in excel_data:
        # Membersihkan tabel (hilangkan NaN kosong total)
        df_log = excel_data['LOGBOOK'].dropna(how='all').dropna(axis=1, how='all').fillna("-")
        pdf.set_font('helvetica', '', 6) # Font diperkecil agar muat ke landscape
        
        try:
            with pdf.table() as table:
                for idx, row_data in df_log.iterrows():
                    row = table.row()
                    for item in row_data:
                        # Dibatasi 40 karakter agar tabel tidak pecah/meluap
                        row.cell(str(item)[:40]) 
        except Exception as e:
            pdf.cell(0, 6, f"Gagal merender tabel LOGBOOK: {e}", new_x='LMARGIN', new_y='NEXT')
    else:
        pdf.set_font('helvetica', '', 10)
        pdf.cell(0, 6, "Data LOGBOOK tidak ditemukan di file Excel.", align='C', new_x='LMARGIN', new_y='NEXT')

    # --- HALAMAN 3: JADWAL (LANDSCAPE) ---
    pdf.add_page(orientation='landscape')
    pdf.set_font('helvetica', 'B', 10)
    pdf.cell(0, 8, 'JADWAL PEMELIHARAAN', align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.ln(3)
    
    if excel_data and 'JADWAL 2026' in excel_data:
        # Karena tabel Jadwal sangat lebar (30+ kolom), ukuran font harus ekstra kecil
        df_jadwal = excel_data['JADWAL 2026'].dropna(how='all').dropna(axis=1, how='all').fillna("-")
        pdf.set_font('helvetica', '', 5) 
        
        try:
            with pdf.table() as table:
                for idx, row_data in df_jadwal.iterrows():
                    row = table.row()
                    for item in row_data:
                        # Dibatasi 15 karakter per sel agar kolom 1-31 tidak meluap
                        row.cell(str(item)[:15])
        except Exception as e:
            pdf.cell(0, 6, f"Gagal merender tabel JADWAL: {e}", new_x='LMARGIN', new_y='NEXT')
    else:
        pdf.set_font('helvetica', '', 10)
        pdf.cell(0, 6, "Data JADWAL 2026 tidak ditemukan di file Excel.", align='C', new_x='LMARGIN', new_y='NEXT')

    # --- HALAMAN 4: DOKUMENTASI FOTO (PORTRAIT) ---
    pdf.add_page(orientation='portrait')
    pdf.set_font('helvetica', 'B', 11)
    pdf.cell(0, 10, f'LAMPIRAN KEGIATAN TEKNISI REGULER: {nama_teknisi.upper()}', align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.ln(5)
    
    y_start = pdf.get_y()
    x_start = 10
    col_width = 90
    row_height = 80
    
    col_index = 0
    for index, row in df_kegiatan.iterrows():
        # Buat page baru otomatis jika foto terlalu banyak ke bawah
        if pdf.get_y() + row_height > 270:
            pdf.add_page(orientation='portrait')
            y_start = pdf.get_y()
            col_index = 0
            
        x_pos = x_start + (col_index * col_width)
        
        pdf.set_xy(x_pos, y_start)
        pdf.set_font('helvetica', 'B', 9)
        pdf.multi_cell(col_width - 5, 5, f"Tanggal: {row['tanggal']}\n{row['penjelasan_kegiatan']}", align='C')
        
        y_img = pdf.get_y() + 2
        
        if pd.notna(row['foto_path']) and os.path.exists(row['foto_path']):
            try:
                # Lebar foto di set proporsional
                pdf.image(row['foto_path'], x=x_pos + 10, y=y_img, w=70)
            except:
                pdf.set_xy(x_pos + 5, y_img)
                pdf.cell(col_width - 10, 10, "[Format Gambar Tidak Didukung]", border=1, align='C')
        else:
            pdf.set_xy(x_pos + 5, y_img)
            pdf.cell(col_width - 10, 10, "[Tidak Ada Gambar]", border=1, align='C')
            
        col_index += 1
        # Geser baris ke bawah jika 2 kolom sudah terisi
        if col_index > 1:
            col_index = 0
            y_start += row_height
            
    nama_file = f"E_Kinerja_{nama_teknisi.replace(' ', '_')}_{bulan}_{tahun}.pdf"
    pdf.output(nama_file)
    return nama_file

# --- ANTARMUKA STREAMLIT ---
st.title("🛠️ Laporan & E-Kinerja Teknisi")

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
                file_path = os.path.join(UPLOAD_DIR, foto.name)
                with open(file_path, "wb") as f:
                    f.write(foto.getbuffer())
            else:
                file_path = None
                
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
    
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        filter_nama = st.selectbox("Pilih Teknisi untuk Dicetak", ["Zulqha Ariandi Al Zikri, S.Tr.Inst.", "Adi Junaidi Rachman, S.Kom.", "Luqmanul Hakim, S.Tr.", "Mohammad Hasyim Hanif, S.Tr.Inst."])
    with col_b:
        filter_bulan = st.selectbox("Pilih Bulan", ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"])
    with col_c:
        filter_tahun = st.selectbox("Pilih Tahun", ["2026", "2027", "2028"])
    
    if st.button("Tarik Data & Generate PDF 🚀"):
        bulan_dict = {"Januari":"01", "Februari":"02", "Maret":"03", "April":"04", "Mei":"05", "Juni":"06", "Juli":"07", "Agustus":"08", "September":"09", "Oktober":"10", "November":"11", "Desember":"12"}
        bulan_angka = bulan_dict[filter_bulan]
        
        query = f"SELECT * FROM e_kinerja WHERE nama_teknisi = '{filter_nama}' AND strftime('%m', tanggal) = '{bulan_angka}' AND strftime('%Y', tanggal) = '{filter_tahun}'"
        df_filter = pd.read_sql(query, conn)
        
        # Load File Excel untuk dilampirkan tabelnya
        excel_data = None
        if os.path.exists(EXCEL_FILE):
             try:
                 excel_data = pd.read_excel(EXCEL_FILE, sheet_name=None, header=None)
             except Exception as e:
                 st.error(f"Gagal membaca Excel: {e}")

        if not df_filter.empty:
            with st.spinner('Menyusun dokumen PDF (Memuat tabel dan foto)...'):
                pdf_file = generate_pdf(filter_nama, filter_bulan, filter_tahun, df_filter, excel_data)
                
            st.success(f"File PDF berhasil dibuat: {pdf_file}")
            
            with open(pdf_file, "rb") as pdf_data:
                st.download_button(
                    label="⬇️ Download PDF E-Kinerja",
                    data=pdf_data,
                    file_name=pdf_file,
                    mime="application/pdf"
                )
        else:
            st.warning("Tidak ada data dokumentasi kegiatan untuk teknisi dan bulan tersebut di database. Silakan isi form di tab pertama terlebih dahulu.")
