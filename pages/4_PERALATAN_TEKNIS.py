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

# --- HELPER PARSING EXCEL LOGBOOK ---
def parse_logbook(df_log, bulan):
    """Mengekstrak tabel Logbook dari Excel berdasarkan Bulan"""
    bulan_str = f"BULAN {bulan.upper()}"
    indices = df_log[df_log[0].astype(str).str.strip().str.upper() == bulan_str].index.tolist()
    
    stamet_data, posmet_data = [], []
    
    # 1. Ekstrak Data Stamet UMK (Bulan match pertama)
    if len(indices) >= 1:
        stamet_start = indices[0] + 3 
        for i in range(stamet_start, len(df_log)):
            row = df_log.iloc[i]
            if str(row[0]).strip().upper() in ["TOTAL", "POS METEOROLOGI TAMBOLAKA", "NAMA PEGAWAI"] or str(row[0]).startswith("LAPORAN"):
                break
            if pd.notna(row[1]) or pd.notna(row[4]):
                stamet_data.append(row)
                
    # 2. Ekstrak Data Posmet Tambolaka (Bulan match kedua)
    if len(indices) >= 2:
        posmet_start = indices[1] + 1 
        if str(df_log.iloc[posmet_start][0]).strip() == "No": 
            posmet_start += 2
            
        for i in range(posmet_start, len(df_log)):
            row = df_log.iloc[i]
            if str(row[0]).strip().upper() in ["TOTAL", "NAMA PEGAWAI"] or str(row[0]).startswith("LAPORAN"):
                break
            if pd.notna(row[1]) or pd.notna(row[4]):
                posmet_data.append(row)
                
    return stamet_data, posmet_data


# --- FUNGSI GENERATOR PDF (E-KINERJA) ---
class PDFKinerja(FPDF):
    # Kita TIDAK menggunakan fungsi header() bawaan agar Kop Surat tidak muncul di setiap halaman.
    # Kita buat fungsi custom cetak_kop_surat() untuk dipanggil secara manual.
    
    def cetak_kop_surat(self):
        # Menyimpan posisi awal
        y_awal = self.get_y()
        
        # Logo BMKG
        if os.path.exists(LOGO_FILE):
            self.image(LOGO_FILE, 15, y_awal, 18) # Logo diperkecil (Width = 18)
        
        # Teks Kop
        self.set_left_margin(40) # Geser margin kiri agar teks tidak menabrak logo
        self.set_y(y_awal + 2)
        self.set_font('helvetica', 'B', 14)
        self.cell(0, 6, 'BADAN METEOROLOGI, KLIMATOLOGI, DAN GEOFISIKA', align='C', new_x='LMARGIN', new_y='NEXT')
        self.set_font('helvetica', 'B', 12)
        self.cell(0, 6, 'STASIUN METEOROLOGI KELAS III UMBU MEHANG KUNDA', align='C', new_x='LMARGIN', new_y='NEXT')
        
        self.set_font('helvetica', '', 10)
        self.cell(0, 5, 'Jl. Adi Sucipto, Waingapu, Sumba Timur', align='C', new_x='LMARGIN', new_y='NEXT')
        self.cell(0, 5, 'Telp. (0387) 61227 | Fax: (0387) 61228 | Kode Pos 87114', align='C', new_x='LMARGIN', new_y='NEXT')
        self.cell(0, 5, 'Email: stamet.sumbatimur@bmkg.go.id | Website: http://ntt.bmkg.go.id', align='C', new_x='LMARGIN', new_y='NEXT')
        
        # Garis Kop Surat
        self.set_left_margin(10) # Kembalikan margin kiri ke default
        y_line = self.get_y() + 2
        self.set_line_width(1.0)
        self.line(10, y_line, self.w - 10, y_line)
        self.set_line_width(0.3)
        self.line(10, y_line + 1.5, self.w - 10, y_line + 1.5)
        
        # Jarak setelah kop surat ke konten
        self.set_y(y_line + 8)


def draw_logbook_page(pdf, title, data_rows, is_posmet=False):
    """Fungsi pembantu menggambar tabel logbook"""
    pdf.set_font('helvetica', 'B', 10)
    pdf.cell(0, 6, 'LAPORAN HASIL MONITORING KONDISI PERALATAN OPERASIONAL UTAMA METEOROLOGI', align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.cell(0, 6, title, align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.ln(5)
    
    if not data_rows:
        pdf.set_font('helvetica', '', 10)
        pdf.cell(0, 6, "Data tidak ditemukan.", align='C')
        return

    pdf.set_fill_color(220, 220, 220)
    pdf.set_font('helvetica', 'B', 8)
    
    # Lebar masing-masing kolom
    w = [10, 70, 45, 45, 12, 12, 12, 12, 30, 24]
    
    # Header Baris 1
    x_start = pdf.get_x()
    y_start = pdf.get_y()
    pdf.cell(w[0], 10, 'No', border=1, align='C', fill=True)
    pdf.cell(w[1], 10, 'Nama Alat', border=1, align='C', fill=True)
    pdf.cell(w[2], 10, 'Lokasi', border=1, align='C', fill=True)
    pdf.cell(w[3], 10, 'Merk/Type', border=1, align='C', fill=True)
    
    pdf.cell(w[4]*4, 5, 'KONDISI', border=1, align='C', fill=True)
    
    pdf.set_xy(x_start + sum(w[:4]) + (w[4]*4), y_start)
    pdf.cell(w[8], 10, 'Tahun Kalibrasi', border=1, align='C', fill=True)
    pdf.cell(w[9], 10, 'Pengadaan', border=1, align='C', fill=True, new_x='LMARGIN', new_y='NEXT')
    
    # Header Baris 2 (Pekan)
    pdf.set_xy(x_start + sum(w[:4]), y_start + 5)
    pdf.cell(w[4], 5, 'P I', border=1, align='C', fill=True)
    pdf.cell(w[5], 5, 'P II', border=1, align='C', fill=True)
    pdf.cell(w[6], 5, 'P III', border=1, align='C', fill=True)
    pdf.cell(w[7], 5, 'P IV', border=1, align='C', fill=True)
    
    pdf.set_xy(x_start, y_start + 10)
    
    # Isi Data
    pdf.set_font('helvetica', '', 8)
    total_alat = 0
    for row in data_rows:
        cols = [str(x) if pd.notna(x) else "" for x in row]
        cols = cols + [""] * (10 - len(cols))
        
        # Bersihkan timestamp jika terbaca dari excel
        if '00:00:00' in cols[8]: cols[8] = cols[8].split(' ')[0]
        
        pdf.cell(w[0], 6, cols[0][:5], border=1, align='C')
        pdf.cell(w[1], 6, cols[1][:45], border=1, align='L')
        pdf.cell(w[2], 6, cols[2][:25], border=1, align='C')
        pdf.cell(w[3], 6, cols[3][:25], border=1, align='C')
        pdf.cell(w[4], 6, cols[4][:3], border=1, align='C')
        pdf.cell(w[5], 6, cols[5][:3], border=1, align='C')
        pdf.cell(w[6], 6, cols[6][:3], border=1, align='C')
        pdf.cell(w[7], 6, cols[7][:3], border=1, align='C')
        pdf.cell(w[8], 6, cols[8][:15], border=1, align='C')
        pdf.cell(w[9], 6, cols[9][:10], border=1, align='C', new_x='LMARGIN', new_y='NEXT')
        total_alat += 1

    # Jika ini Posmet, tambahkan baris TOTAL di bawah
    if is_posmet:
        pdf.set_font('helvetica', 'B', 8)
        lebar_gabungan = sum(w[:4])
        pdf.cell(lebar_gabungan, 6, 'TOTAL', border=1, align='C')
        pdf.cell(w[4], 6, str(total_alat), border=1, align='C')
        pdf.cell(sum(w[5:]), 6, '', border=1, new_x='LMARGIN', new_y='NEXT')


def generate_pdf(nama_teknisi, bulan, tahun, df_kegiatan, uploaded_excel):
    pdf = PDFKinerja()
    
    # ================= HALAMAN 1: NARASI (PORTRAIT) =================
    pdf.add_page(orientation='portrait')
    pdf.cetak_kop_surat() # Panggil kop surat manual
    
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
    
    # ================= HALAMAN 2 & 3: LOGBOOK (LANDSCAPE) =================
    if uploaded_excel is not None:
        try:
            df_log = pd.read_excel(uploaded_excel, sheet_name='LOGBOOK', header=None)
            stamet_data, posmet_data = parse_logbook(df_log, bulan)
            
            # Halaman 2: Stamet UMK (Dengan Kop Surat)
            pdf.add_page(orientation='landscape')
            pdf.cetak_kop_surat()
            draw_logbook_page(pdf, "STASIUN METEOROLOGI UMBU MEHANG KUNDA", stamet_data, is_posmet=False)
            
            # Halaman 3: Posmet Tambolaka (Tanpa Kop Surat, Ada Baris Total)
            pdf.add_page(orientation='landscape')
            # Perhatikan: pdf.cetak_kop_surat() TIDAK dipanggil di sini!
            pdf.set_y(15) # Memberi jarak aman dari atas karena tidak ada kop
            draw_logbook_page(pdf, "POS METEOROLOGI TAMBOLAKA", posmet_data, is_posmet=True)
            
        except Exception as e:
            pdf.add_page()
            pdf.cell(0, 10, f"Error membaca sheet LOGBOOK: {e}", align='C')
    else:
        pdf.add_page()
        pdf.cell(0, 10, "File Excel tidak diunggah. Data Logbook kosong.", align='C')


    # ================= HALAMAN 4: DOKUMENTASI FOTO (PORTRAIT) =================
    pdf.add_page(orientation='portrait')
    pdf.cetak_kop_surat() # Pasang Kop Surat untuk Lampiran
    
    pdf.set_font('helvetica', 'B', 11)
    pdf.cell(0, 10, f'LAMPIRAN KEGIATAN TEKNISI REGULER: {nama_teknisi.upper()}', align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.ln(5)
    
    y_start = pdf.get_y()
    x_start = 10
    col_width = 90
    row_height = 80
    
    col_index = 0
    for index, row in df_kegiatan.iterrows():
        # Auto New Page
        if pdf.get_y() + row_height > 270:
            pdf.add_page(orientation='portrait')
            pdf.cetak_kop_surat() # Pasang Kop Surat jika ada lembar foto baru
            y_start = pdf.get_y()
            col_index = 0
            
        x_pos = x_start + (col_index * col_width)
        
        pdf.set_xy(x_pos, y_start)
        pdf.set_font('helvetica', 'B', 9)
        pdf.multi_cell(col_width - 5, 5, f"Tanggal: {row['tanggal']}\n{row['penjelasan_kegiatan']}", align='C')
        
        y_img = pdf.get_y() + 2
        
        if pd.notna(row['foto_path']) and os.path.exists(row['foto_path']):
            try:
                # Menyesuaikan foto dalam grid
                pdf.image(row['foto_path'], x=x_pos + 10, y=y_img, w=70)
            except:
                pdf.set_xy(x_pos + 5, y_img)
                pdf.cell(col_width - 10, 10, "[Format Gambar Error]", border=1, align='C')
        else:
            pdf.set_xy(x_pos + 5, y_img)
            pdf.cell(col_width - 10, 10, "[Tidak Ada Gambar]", border=1, align='C')
            
        col_index += 1
        if col_index > 1:
            col_index = 0
            y_start += row_height
            
    nama_file = f"E_Kinerja_{nama_teknisi.replace(' ', '_')}_{bulan}_{tahun}.pdf"
    pdf.output(nama_file)
    return nama_file

# --- ANTARMUKA STREAMLIT ---
st.title("🛠️ Laporan & E-Kinerja Teknisi")

tab1, tab2, tab3 = st.tabs(["📝 Input Kegiatan Harian", "📅 Data Kinerja (DB)", "🖨️ Cetak PDF E-Kinerja"])

# OPSI DROPDOWN SESUAI EXCEL JADWAL PEMELIHARAAN
opsi_kegiatan = [
    "Pemeliharaan Taman Alat",
    "Pemeliharaan Display Bandara",
    "Pemeliharaan AWOS",
    "Pengolahan OLA dan SLA",
    "Pembuatan Gas"
]

# TAB 1: INPUT KEGIATAN HARIAN
with tab1:
    st.subheader("Form Input Dokumentasi Kegiatan")
    with st.form("form_kinerja", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            tanggal = st.date_input("Tanggal Kegiatan", datetime.today())
            nama = st.selectbox("Nama Teknisi", ["Zulqha Ariandi Al Zikri, S.Tr.Inst.", "Adi Junaidi Rachman, S.Kom.", "Luqmanul Hakim, S.Tr.", "Mohammad Hasyim Hanif, S.Tr.Inst."])
        with col2:
            # Menggunakan multiselect sesuai pilihan dari excel
            kegiatan_dipilih = st.multiselect("Penjelasan Kegiatan", opsi_kegiatan)
            
        foto = st.file_uploader("Upload Foto Dokumentasi", type=['jpg', 'jpeg', 'png'])
        
        submit = st.form_submit_button("Simpan Data")
        
        if submit:
            if not kegiatan_dipilih:
                st.error("Silakan pilih minimal satu Penjelasan Kegiatan!")
            else:
                kegiatan_str = ", ".join(kegiatan_dipilih) # Menggabungkan jika memilih >1
                
                if foto is not None:
                    file_path = os.path.join(UPLOAD_DIR, foto.name)
                    with open(file_path, "wb") as f:
                        f.write(foto.getbuffer())
                else:
                    file_path = None
                    
                c = conn.cursor()
                c.execute("INSERT INTO e_kinerja (tanggal, nama_teknisi, penjelasan_kegiatan, foto_path) VALUES (?, ?, ?, ?)",
                          (tanggal.strftime("%Y-%m-%d"), nama, kegiatan_str, file_path))
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
    st.subheader("Generate Dokumen E-Kinerja & Logbook")
    
    st.markdown("⚠️ **Wajib: Upload File Excel (Peralatan Teknis.xlsx) untuk menarik tabel Logbook secara otomatis.**")
    uploaded_excel = st.file_uploader("Pilih File Excel", type=['xlsx', 'xls'])
    
    st.markdown("---")
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        filter_nama = st.selectbox("Pilih Teknisi", ["Zulqha Ariandi Al Zikri, S.Tr.Inst.", "Adi Junaidi Rachman, S.Kom.", "Luqmanul Hakim, S.Tr.", "Mohammad Hasyim Hanif, S.Tr.Inst."])
    with col_b:
        filter_bulan = st.selectbox("Pilih Bulan", ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"])
    with col_c:
        filter_tahun = st.selectbox("Pilih Tahun", ["2026", "2027", "2028"])
    
    if st.button("Tarik Data & Generate PDF 🚀"):
        if uploaded_excel is None:
            st.error("Silakan unggah File Excel terlebih dahulu pada form di atas!")
        else:
            bulan_dict = {"Januari":"01", "Februari":"02", "Maret":"03", "April":"04", "Mei":"05", "Juni":"06", "Juli":"07", "Agustus":"08", "September":"09", "Oktober":"10", "November":"11", "Desember":"12"}
            bulan_angka = bulan_dict[filter_bulan]
            
            query = f"SELECT * FROM e_kinerja WHERE nama_teknisi = '{filter_nama}' AND strftime('%m', tanggal) = '{bulan_angka}' AND strftime('%Y', tanggal) = '{filter_tahun}'"
            df_filter = pd.read_sql(query, conn)
            
            if not df_filter.empty:
                with st.spinner('Menyusun PDF & Mengekstrak Logbook 2 Lembar...'):
                    pdf_file = generate_pdf(filter_nama, filter_bulan, filter_tahun, df_filter, uploaded_excel)
                    
                st.success(f"File PDF berhasil dibuat: {pdf_file}")
                
                with open(pdf_file, "rb") as pdf_data:
                    st.download_button(
                        label="⬇️ Download PDF E-Kinerja",
                        data=pdf_data,
                        file_name=pdf_file,
                        mime="application/pdf"
                    )
            else:
                st.warning("Tidak ada data dokumentasi foto/kegiatan untuk teknisi tersebut di bulan ini. Silakan isi form di Tab 1.")
