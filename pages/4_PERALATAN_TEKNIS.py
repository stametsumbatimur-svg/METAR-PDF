import streamlit as st
import pandas as pd
import sqlite3
import os
from datetime import datetime
from fpdf import FPDF
try:
    from pypdf import PdfWriter
    PYPDF_INSTALLED = True
except ImportError:
    PYPDF_INSTALLED = False

# --- KONFIGURASI AWAL ---
UPLOAD_DIR = "uploads"
if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

LOGO_FILE = "logo_bmkg.png"

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

# --- HELPER FORMAT TANGGAL ---
def format_tanggal_indo(tanggal_str):
    """Mengubah 2026-07-02 menjadi 2 Juli 2026"""
    try:
        dt = datetime.strptime(tanggal_str, "%Y-%m-%d")
        bulan_indo = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
        return f"{dt.day} {bulan_indo[dt.month-1]} {dt.year}"
    except:
        return tanggal_str

# --- PARSER EXCEL LOGBOOK (ANTI ERROR) ---
def parse_logbook(df_log, bulan):
    """Membaca Logbook dengan cerdas mengabaikan baris kosong yang acak"""
    bulan_str = f"BULAN{bulan.upper()}"
    stamet_data, posmet_data = [], []
    state = 0 
    
    for index, row in df_log.iterrows():
        col0_raw = str(row[0]).strip().upper()
        col0_clean = col0_raw.replace(" ", "")
        col1 = str(row[1]).strip() if pd.notna(row[1]) else ""
        col4 = str(row[4]).strip() if pd.notna(row[4]) else ""
        
        # Transisi Status Pembacaan
        if bulan_str in col0_clean and state == 0:
            state = 1
            continue
        elif "POSMETEOROLOGI" in col0_clean and state == 1:
            state = 2
            continue
        elif bulan_str in col0_clean and state == 2:
            state = 3
            continue
            
        # Perekaman Data Stamet UMK
        if state == 1:
            if col0_raw in ["TOTAL", "NAMA PEGAWAI"] or col0_raw.startswith("LAPORAN"):
                state = 2 # Berhenti mencari UMK
                continue
            if col1 or col4:
                # Abaikan baris header tabel yang tersangkut
                if "NAMA ALAT" not in col1.upper() and "PEKAN I" not in col4.upper() and col0_raw != "NO":
                    stamet_data.append(row)
                    
        # Perekaman Data Posmet Tambolaka
        elif state == 3:
            if col0_raw in ["TOTAL", "NAMA PEGAWAI"] or col0_raw.startswith("LAPORAN"):
                break # Berhenti total
            if col1 or col4:
                if "NAMA ALAT" not in col1.upper() and "PEKAN I" not in col4.upper() and col0_raw != "NO":
                    posmet_data.append(row)
                    
    return stamet_data, posmet_data

# --- FUNGSI GENERATOR PDF ---
class PDFKinerja(FPDF):
    def cetak_kop_surat(self):
        """Kop Surat Presisi seperti BMKG"""
        y_awal = self.get_y()
        if os.path.exists(LOGO_FILE):
            self.image(LOGO_FILE, 12, y_awal, 25) # Logo di kiri
        
        self.set_y(y_awal)
        self.set_font('helvetica', 'B', 13)
        self.cell(0, 5, 'BADAN METEOROLOGI, KLIMATOLOGI, DAN GEOFISIKA', align='C', new_x='LMARGIN', new_y='NEXT')
        self.set_font('helvetica', 'B', 12)
        self.cell(0, 5, 'STASIUN METEOROLOGI KELAS III UMBU MEHANG KUNDA', align='C', new_x='LMARGIN', new_y='NEXT')
        
        self.set_font('helvetica', '', 9)
        self.cell(0, 4, 'Jl. Adi Sucipto, Waingapu, Sumba Timur', align='C', new_x='LMARGIN', new_y='NEXT')
        self.cell(0, 4, 'Telp. (0387) 61227 | Fax: (0387) 61228 | Kode Pos 87114', align='C', new_x='LMARGIN', new_y='NEXT')
        self.cell(0, 4, 'Email: stamet.sumbatimur@bmkg.go.id | Website: http://ntt.bmkg.go.id', align='C', new_x='LMARGIN', new_y='NEXT')
        
        y_line = self.get_y() + 2
        self.set_line_width(1.0)
        self.line(10, y_line, self.w - 10, y_line)
        self.set_line_width(0.3)
        self.line(10, y_line + 1.5, self.w - 10, y_line + 1.5)
        self.set_y(y_line + 8)

def draw_logbook_page(pdf, title, data_rows, is_posmet=False):
    pdf.set_font('helvetica', 'B', 10)
    pdf.cell(0, 6, 'LAPORAN HASIL MONITORING KONDISI PERALATAN OPERASIONAL UTAMA METEOROLOGI', align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.cell(0, 6, title, align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.ln(5)
    
    if not data_rows:
        pdf.set_font('helvetica', '', 10)
        pdf.cell(0, 6, "Data tidak ditemukan. Silakan cek format Excel.", align='C')
        return

    pdf.set_fill_color(220, 220, 220)
    pdf.set_font('helvetica', 'B', 8)
    
    # Kolom: No | Alat | Lokasi | Merk | PI | PII | PIII | PIV | Kalibrasi | Pengadaan
    w = [10, 60, 30, 35, 18, 18, 18, 18, 26, 22] 
    
    x_start = pdf.get_x()
    y_start = pdf.get_y()
    
    pdf.rect(x_start, y_start, w[0], 12, style='DF')
    pdf.rect(x_start+w[0], y_start, w[1], 12, style='DF')
    pdf.rect(x_start+sum(w[:2]), y_start, w[2], 12, style='DF')
    pdf.rect(x_start+sum(w[:3]), y_start, w[3], 12, style='DF')
    
    x_kondisi = x_start+sum(w[:4])
    pdf.rect(x_kondisi, y_start, sum(w[4:8]), 6, style='DF')
    
    pdf.rect(x_kondisi, y_start+6, w[4], 6, style='DF')
    pdf.rect(x_kondisi+w[4], y_start+6, w[5], 6, style='DF')
    pdf.rect(x_kondisi+sum(w[4:6]), y_start+6, w[6], 6, style='DF')
    pdf.rect(x_kondisi+sum(w[4:7]), y_start+6, w[7], 6, style='DF')
    
    x_sisa = x_start+sum(w[:8])
    pdf.rect(x_sisa, y_start, w[8], 12, style='DF')
    pdf.rect(x_sisa+w[8], y_start, w[9], 12, style='DF')
    
    pdf.set_xy(x_start, y_start+3)
    pdf.cell(w[0], 6, 'No', align='C')
    pdf.cell(w[1], 6, 'Nama Alat', align='C')
    pdf.cell(w[2], 6, 'Lokasi', align='C')
    pdf.cell(w[3], 6, 'Merk/Type', align='C')
    
    pdf.set_xy(x_kondisi, y_start)
    pdf.cell(sum(w[4:8]), 6, 'KONDISI', align='C')
    pdf.set_xy(x_kondisi, y_start+6)
    pdf.cell(w[4], 6, 'PEKAN I', align='C')
    pdf.cell(w[5], 6, 'PEKAN II', align='C')
    pdf.cell(w[6], 6, 'PEKAN III', align='C')
    pdf.cell(w[7], 6, 'PEKAN IV', align='C')
    
    pdf.set_xy(x_sisa, y_start + 1.5)
    pdf.multi_cell(w[8], 4, 'Tahun\nKalibrasi', align='C')
    pdf.set_xy(x_sisa+w[8], y_start+3)
    pdf.cell(w[9], 6, 'Pengadaan', align='C')
    
    pdf.set_y(y_start + 12)
    
    pdf.set_font('helvetica', '', 8)
    total_alat = 0
    for row in data_rows:
        cols = [str(x) if pd.notna(x) else "" for x in row]
        cols = cols + [""] * (10 - len(cols))
        
        if '00:00:00' in cols[8]: cols[8] = cols[8].split(' ')[0]
        # Auto Wrap Lokasi
        cols[2] = cols[2].replace('Bandara Umbu Mehang Kunda', 'Bandara Umbu\nMehang Kunda')
        
        lines = 1
        if '\n' in cols[2] or len(cols[1]) > 40 or len(cols[3]) > 20:
            lines = 2
        row_h = lines * 5
        
        x = pdf.get_x()
        y = pdf.get_y()
        for idx_w, width in enumerate(w):
            pdf.rect(x, y, width, row_h)
            x += width
            
        x = pdf.get_x() - sum(w)
        for i, text in enumerate(cols):
            pdf.set_xy(x, y + (0.5 if lines==1 else 0))
            pdf.multi_cell(w[i], 5, text, border=0, align='C' if i!=1 else 'L')
            x += w[i]
            
        pdf.set_xy(10, y + row_h)
        total_alat += 1

    if is_posmet:
        pdf.set_font('helvetica', 'B', 8)
        lebar_gabungan = sum(w[:4])
        y = pdf.get_y()
        pdf.rect(10, y, lebar_gabungan, 6)
        pdf.rect(10+lebar_gabungan, y, w[4], 6)
        pdf.set_xy(10, y)
        pdf.cell(lebar_gabungan, 6, 'TOTAL', border=0, align='C')
        pdf.cell(w[4], 6, str(total_alat), border=0, align='C')
        pdf.set_y(y + 6)


def generate_pdf(nama_teknisi, bulan, tahun, df_kegiatan, uploaded_excel, poin_korektif):
    pdf = PDFKinerja()
    
    # === HALAMAN 1: NARASI (PORTRAIT) ===
    pdf.add_page(orientation='portrait')
    pdf.cetak_kop_surat()
    
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
    
    if poin_korektif:
        pdf.set_x(10)
        pdf.cell(10, 6, "4.")
        pdf.cell(0, 6, "Laporan Hasil Pemeliharaan Korektif Peralatan Operasional Utama Meteorologi", new_x='LMARGIN', new_y='NEXT')

    
    # === HALAMAN 2 & 3: LOGBOOK DARI EXCEL (LANDSCAPE) ===
    if uploaded_excel is not None:
        try:
            df_log = pd.read_excel(uploaded_excel, sheet_name='LOGBOOK', header=None)
            stamet_data, posmet_data = parse_logbook(df_log, bulan)
            
            # Hal 2: Stamet UMK (Punya Kop Surat)
            pdf.add_page(orientation='landscape')
            pdf.cetak_kop_surat()
            draw_logbook_page(pdf, "STASIUN METEOROLOGI UMBU MEHANG KUNDA", stamet_data, is_posmet=False)
            
            # Hal 3: Posmet (TIDAK Punya Kop Surat)
            pdf.add_page(orientation='landscape')
            pdf.set_y(15) 
            draw_logbook_page(pdf, "POS METEOROLOGI TAMBOLAKA", posmet_data, is_posmet=True)
            
        except Exception as e:
            pdf.add_page()
            pdf.cell(0, 10, f"Error membaca sheet LOGBOOK: {e}", align='C')
    else:
        pdf.add_page()
        pdf.cell(0, 10, "File Excel tidak diunggah. Data Logbook kosong.", align='C')

    # === HALAMAN 4: TABEL DOKUMENTASI FOTO (PORTRAIT) ===
    pdf.add_page(orientation='portrait')
    # Sesuai instruksi: LAMPIRAN TIDAK PERLU KOP SURAT.
    pdf.set_y(15)
    pdf.set_font('helvetica', 'B', 11)
    pdf.cell(0, 10, f'LAMPIRAN KEGIATAN TEKNISI REGULER: {nama_teknisi.upper()}', align='L', new_x='LMARGIN', new_y='NEXT')
    pdf.ln(2)
    
    # Render Tabel Header
    pdf.set_fill_color(180, 200, 255) # Warna biru muda tabel
    pdf.set_font('helvetica', 'B', 10)
    pdf.cell(15, 8, 'NO', border=1, align='C', fill=True)
    pdf.cell(175, 8, 'PENJELASAN', border=1, align='C', fill=True, new_x='LMARGIN', new_y='NEXT')
    
    pdf.set_font('helvetica', '', 10)
    nomor_urut = 1
    
    for index, row in df_kegiatan.iterrows():
        tgl_indo = format_tanggal_indo(row['tanggal'])
        teks = f"{row['penjelasan_kegiatan']} Tanggal {tgl_indo}"
        
        row_h = 75 # Tinggi paksa per baris foto
        
        # Cek jika butuh halaman baru
        if pdf.get_y() + row_h > 280:
            pdf.add_page(orientation='portrait')
            pdf.set_y(15)
            # Cetak ulang header tabel di halaman baru
            pdf.set_font('helvetica', 'B', 10)
            pdf.cell(15, 8, 'NO', border=1, align='C', fill=True)
            pdf.cell(175, 8, 'PENJELASAN', border=1, align='C', fill=True, new_x='LMARGIN', new_y='NEXT')
            pdf.set_font('helvetica', '', 10)
            
        y_start = pdf.get_y()
        
        # Buat Kotak Tabel
        pdf.rect(10, y_start, 15, row_h)
        pdf.rect(25, y_start, 175, row_h)
        
        # Tulis Nomor
        pdf.set_xy(10, y_start + (row_h/2) - 3)
        pdf.cell(15, 6, str(nomor_urut), align='C')
        
        # Tulis Gambar (Memaksa masuk batas secara proporsional)
        img_y = y_start + 3
        if pd.notna(row['foto_path']) and os.path.exists(row['foto_path']):
            try:
                # FPDF2 keep_aspect_ratio menjaga agar foto tidak penyok
                # X di set tengah tabel (25 + (175 - 120)/2 = 52.5)
                pdf.image(row['foto_path'], x=52.5, y=img_y, w=120, h=60, keep_aspect_ratio=True)
            except:
                pdf.set_xy(25, img_y + 25)
                pdf.cell(175, 6, "[Format Gambar Error / Corrupt]", align='C')
        else:
            pdf.set_xy(25, img_y + 25)
            pdf.cell(175, 6, "[Tidak Ada Gambar Diunggah]", align='C')
            
        # Tulis Keterangan Text
        pdf.set_xy(25, y_start + row_h - 10)
        pdf.cell(175, 6, teks, align='C')
        
        pdf.set_y(y_start + row_h)
        nomor_urut += 1
            
    temp_pdf = f"temp_kinerja.pdf"
    pdf.output(temp_pdf)
    return temp_pdf


# --- ANTARMUKA STREAMLIT ---
st.title("🛠️ Laporan & E-Kinerja Teknisi")

tab1, tab2, tab3 = st.tabs(["📝 Input Kegiatan", "📅 Arsip Data (DB)", "🖨️ Cetak PDF E-Kinerja"])

opsi_kegiatan = [
    "Pemeliharaan Taman Alat",
    "Pemeliharaan Display Bandara",
    "Pemeliharaan AWOS",
    "Pengolahan OLA dan SLA",
    "Pembuatan Gas"
]

with tab1:
    st.subheader("Form Input Dokumentasi Kegiatan")
    with st.form("form_kinerja", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            tanggal = st.date_input("Tanggal Kegiatan", datetime.today())
            nama = st.selectbox("Nama Teknisi", ["Zulqha Ariandi Al Zikri, S.Tr.Inst.", "Adi Junaidi Rachman, S.Kom.", "Luqmanul Hakim, S.Tr.", "Mohammad Hasyim Hanif, S.Tr.Inst."])
        with col2:
            kegiatan_dipilih = st.multiselect("Penjelasan Kegiatan", opsi_kegiatan)
            
        foto = st.file_uploader("Upload Foto Dokumentasi", type=['jpg', 'jpeg', 'png'])
        submit = st.form_submit_button("Simpan Data")
        
        if submit:
            if not kegiatan_dipilih:
                st.error("Silakan pilih minimal satu Penjelasan Kegiatan!")
            else:
                kegiatan_str = ", ".join(kegiatan_dipilih) 
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

with tab2:
    st.subheader("Arsip Kegiatan Teknisi")
    df_db = pd.read_sql("SELECT * FROM e_kinerja", conn)
    if not df_db.empty:
        st.dataframe(df_db, use_container_width=True)
    else:
        st.info("Belum ada data kegiatan yang diinput.")

with tab3:
    st.subheader("Generate Dokumen E-Kinerja & Logbook")
    
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        filter_nama = st.selectbox("Pilih Teknisi", ["Zulqha Ariandi Al Zikri, S.Tr.Inst.", "Adi Junaidi Rachman, S.Kom.", "Luqmanul Hakim, S.Tr.", "Mohammad Hasyim Hanif, S.Tr.Inst."])
    with col_b:
        filter_bulan = st.selectbox("Pilih Bulan", ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"])
    with col_c:
        filter_tahun = st.selectbox("Pilih Tahun", ["2026", "2027", "2028"])
        
    st.markdown("---")
    st.markdown("#### 🖇️ Konfigurasi Lampiran")
    
    uploaded_excel = st.file_uploader("1. Wajib: File Excel Logbook (Peralatan Teknis.xlsx)", type=['xlsx', 'xls'])
    poin_korektif = st.checkbox("Tambahkan Poin Ke-4 (Laporan Korektif/Kalibrasi) di Narasi Hal. 1")
    
    if PYPDF_INSTALLED:
        pdf_kalibrasi = st.file_uploader("2. Opsional: Upload PDF Laporan Kalibrasi/Korektif (Digabungkan di halaman akhir)", type=['pdf'])
    else:
        st.warning("⚠️ Modul 'pypdf' tidak ditemukan. Penggabungan PDF eksternal dinonaktifkan.")
        pdf_kalibrasi = None
    
    if st.button("Tarik Data & Generate PDF Lengkap 🚀"):
        if uploaded_excel is None:
            st.error("Silakan unggah File Excel terlebih dahulu!")
        else:
            bulan_dict = {"Januari":"01", "Februari":"02", "Maret":"03", "April":"04", "Mei":"05", "Juni":"06", "Juli":"07", "Agustus":"08", "September":"09", "Oktober":"10", "November":"11", "Desember":"12"}
            bulan_angka = bulan_dict[filter_bulan]
            
            query = f"SELECT * FROM e_kinerja WHERE nama_teknisi = '{filter_nama}' AND strftime('%m', tanggal) = '{bulan_angka}' AND strftime('%Y', tanggal) = '{filter_tahun}'"
            df_filter = pd.read_sql(query, conn)
            
            with st.spinner('Menyusun Logbook & Memformat Tabel Lampiran...'):
                temp_pdf = generate_pdf(filter_nama, filter_bulan, filter_tahun, df_filter, uploaded_excel, poin_korektif)
                final_filename = f"E_Kinerja_{filter_nama.replace(' ', '_')}_{filter_bulan}_{filter_tahun}.pdf"
                
                # Proses Merge Jika Ada PDF Eksternal
                if PYPDF_INSTALLED:
                    merger = PdfWriter()
                    merger.append(temp_pdf) 
                    
                    if pdf_kalibrasi is not None:
                        kalibrasi_path = os.path.join(UPLOAD_DIR, "temp_kalibrasi.pdf")
                        with open(kalibrasi_path, "wb") as f:
                            f.write(pdf_kalibrasi.getbuffer())
                        merger.append(kalibrasi_path)
                        
                    with open(final_filename, "wb") as f_out:
                        merger.write(f_out)
                else:
                    os.rename(temp_pdf, final_filename)
                
            st.success("✅ Dokumen PDF E-Kinerja berhasil dibuat sempurna!")
            
            with open(final_filename, "rb") as pdf_data:
                st.download_button(
                    label="⬇️ Download Hasil PDF E-Kinerja",
                    data=pdf_data,
                    file_name=final_filename,
                    mime="application/pdf"
                )
