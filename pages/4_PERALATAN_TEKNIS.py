import streamlit as st
import pandas as pd
import sqlite3
import os
import tempfile
from datetime import datetime
from PIL import Image
from fpdf import FPDF

try:
    from pypdf import PdfWriter
    PYPDF_INSTALLED = True
except ImportError:
    PYPDF_INSTALLED = False

# --- KONFIGURASI AWAL ---
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)
LOGO_FILE = "logo_bmkg.png"
DB_FILE = "kinerja_teknisi.db"

MONTH_MAP = {
    "Januari": "01", "Februari": "02", "Maret": "03", "April": "04",
    "Mei": "05", "Juni": "06", "Juli": "07", "Agustus": "08",
    "September": "09", "Oktober": "10", "November": "11", "Desember": "12"
}

QUARTER_MAP = {
    "Triwulan I (Jan - Mar)": {
        "label": "TRIWULAN I",
        "months": ["01", "02", "03"],
        "month_names": ["Januari", "Februari", "Maret"]
    },
    "Triwulan II (Apr - Jun)": {
        "label": "TRIWULAN II",
        "months": ["04", "05", "06"],
        "month_names": ["April", "Mei", "Juni"]
    },
    "Triwulan III (Jul - Sep)": {
        "label": "TRIWULAN III",
        "months": ["07", "08", "09"],
        "month_names": ["Juli", "Agustus", "September"]
    },
    "Triwulan IV (Okt - Des)": {
        "label": "TRIWULAN IV",
        "months": ["10", "11", "12"],
        "month_names": ["Oktober", "November", "Desember"]
    }
}

# --- HELPER DATABASE ---
def get_db_connection():
    return sqlite3.connect(DB_FILE)

def init_db():
    with get_db_connection() as conn:
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

init_db()

# --- FORMATTER TANGGAL INDONESIA ---
def format_tanggal_indo(tanggal_str):
    try:
        dt = datetime.strptime(tanggal_str, "%Y-%m-%d")
        bulan_indo = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", 
                      "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"]
        return f"{dt.day} {bulan_indo[dt.month-1]} {dt.year}"
    except Exception:
        return tanggal_str

# --- PARSER LOGBOOK EXCEL ---
def parse_logbook(df_log, bulan):
    bulan_str = f"BULAN {bulan.upper()}"
    stamet_data, posmet_data = [], []
    
    idx_bulan = df_log[df_log[0].astype(str).str.contains(bulan_str, case=False, na=False)].index.tolist()
    
    if len(idx_bulan) >= 1:
        start_umk = idx_bulan[0] + 1
        for i in range(start_umk, len(df_log)):
            row_val = str(df_log.iloc[i, 0]).strip().upper()
            if row_val in ["TOTAL", "NAMA PEGAWAI"] or row_val.startswith("LAPORAN") or "POS METEOROLOGI" in row_val:
                break
            col1 = str(df_log.iloc[i, 1]).strip() if pd.notna(df_log.iloc[i, 1]) else ""
            if col1 and "NAMA ALAT" not in col1.upper() and row_val != "NO":
                row_list = df_log.iloc[i].values.tolist() + [""] * 10
                stamet_data.append(row_list[:10])
                
    if len(idx_bulan) >= 2:
        start_posmet = idx_bulan[1] + 1
        for i in range(start_posmet, len(df_log)):
            row_val = str(df_log.iloc[i, 0]).strip().upper()
            if row_val in ["TOTAL", "NAMA PEGAWAI"] or row_val.startswith("LAPORAN"):
                break
            col1 = str(df_log.iloc[i, 1]).strip() if pd.notna(df_log.iloc[i, 1]) else ""
            if col1 and "NAMA ALAT" not in col1.upper() and row_val != "NO":
                row_list = df_log.iloc[i].values.tolist() + [""] * 10
                posmet_data.append(row_list[:10])
                
    return stamet_data, posmet_data

# --- PARSER JADWAL PEMELIHARAAN EXCEL ---
def parse_jadwal(df_jadwal, nama_teknisi, bulan_nama):
    bulan_target = f"BULAN : {bulan_nama.upper()}"
    idx_bulan = df_jadwal[df_jadwal[2].astype(str).str.strip().str.upper() == bulan_target].index.tolist()
    
    if not idx_bulan:
        for col in range(3):
            matches = df_jadwal[df_jadwal[col].astype(str).str.strip().str.upper().str.contains(bulan_nama.upper())].index.tolist()
            if matches:
                idx_bulan = matches
                break
                
    if not idx_bulan:
        return nama_teknisi, "", {}

    start_idx = idx_bulan[0]
    nip = ""
    jadwal_days = {}
    
    for i in range(start_idx, min(start_idx + 35, len(df_jadwal))):
        val_col1 = str(df_jadwal.iloc[i, 1]).strip()
        if nama_teknisi.upper() in val_col1.upper() or val_col1.upper() in nama_teknisi.upper():
            nip_val = str(df_jadwal.iloc[i+1, 1]).strip() if i+1 < len(df_jadwal) else ""
            if nip_val.isdigit():
                nip = nip_val
            
            for day in range(1, 32):
                col_idx = day + 1
                if col_idx < df_jadwal.shape[1]:
                    val = str(df_jadwal.iloc[i, col_idx]).strip()
                    jadwal_days[day] = val if val.lower() != 'nan' else ""
            break
            
    return nama_teknisi, nip, jadwal_days

# --- CLASS GENERATOR PDF ---
class PDFKinerja(FPDF):
    def cetak_kop_surat(self):
        y_awal = self.get_y()
        if os.path.exists(LOGO_FILE):
            self.image(LOGO_FILE, x=14, y=y_awal + 1, w=17)
            
        self.set_y(y_awal)
        self.set_font('helvetica', 'B', 13)
        self.cell(0, 5, 'BADAN METEOROLOGI, KLIMATOLOGI, DAN GEOFISIKA', align='C', new_x='LMARGIN', new_y='NEXT')
        self.set_font('helvetica', 'B', 11)
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

def draw_header_table(pdf, x_start, y_start, w):
    pdf.set_fill_color(220, 220, 220)
    pdf.set_font('helvetica', 'B', 8)
    
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

def draw_logbook_page(pdf, title, data_rows, total_alat_keseluruhan=24, is_posmet=False):
    pdf.set_font('helvetica', 'B', 10)
    pdf.cell(0, 6, 'LAPORAN HASIL MONITORING KONDISI PERALATAN OPERASIONAL UTAMA METEOROLOGI', align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.cell(0, 6, title, align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.ln(5)
    
    if not data_rows:
        pdf.set_font('helvetica', '', 10)
        pdf.cell(0, 6, "Data tidak ditemukan. Silakan cek format Excel / Bulan yang dipilih.", align='C')
        return

    w = [8, 65, 30, 35, 16, 16, 16, 16, 25, 23]
    draw_header_table(pdf, 10, pdf.get_y(), w)
    
    pdf.set_y(pdf.get_y() + 12)
    pdf.set_font('helvetica', '', 8)
    
    for row in data_rows:
        cols = [str(x).replace("nan", "") for x in row]
        if '00:00:00' in cols[8]: cols[8] = cols[8].split(' ')[0]
        cols[2] = cols[2].replace('Bandara Umbu Mehang Kunda', 'Bandara Umbu\nMehang Kunda')
        cols[1] = cols[1][:60]
        
        lines = 1
        if '\n' in cols[2] or len(cols[1]) > 35 or len(cols[3]) > 18:
            lines = 2
        if len(cols[1]) > 70:
            lines = 3
        row_h = lines * 5
        
        if pdf.get_y() + row_h > 185:
            pdf.add_page(orientation='landscape')
            draw_header_table(pdf, 10, 15, w)
            pdf.set_y(27)
            
        x = 10
        y = pdf.get_y()
        
        for width in w:
            pdf.rect(x, y, width, row_h)
            x += width
            
        x = 10
        for i, text in enumerate(cols):
            y_offset = y + (0.5 if lines == 1 else 1)
            pdf.set_xy(x, y_offset)
            align = 'C' if i != 1 else 'L'
            pdf.multi_cell(w[i], 4.5, text, border=0, align=align)
            x += w[i]
            
        pdf.set_xy(10, y + row_h)

    if is_posmet:
        pdf.set_font('helvetica', 'B', 8)
        lebar_gabungan = sum(w[:4])
        y = pdf.get_y()
        
        if y + 6 > 185:
            pdf.add_page(orientation='landscape')
            y = 15
            
        pdf.rect(10, y, lebar_gabungan, 6)
        pdf.rect(10+lebar_gabungan, y, w[4], 6)
        pdf.set_xy(10, y)
        pdf.cell(lebar_gabungan, 6, 'TOTAL', border=0, align='C')
        pdf.cell(w[4], 6, str(total_alat_keseluruhan), border=0, align='C')
        pdf.set_y(y + 6)

# --- 3. DRAW JADWAL PEMELIHARAAN (RATA KIRI, TANPA KOP, MERGE TANGGAL 1-31) ---
def draw_jadwal_page(pdf, nama_teknisi, nip, bulan_nama, tahun, triwulan_label, days_dict):
    # Judul Rata Kiri
    pdf.set_font('helvetica', 'B', 10)
    pdf.cell(0, 5, 'JADWAL PEMELIHARAAN', align='L', new_x='LMARGIN', new_y='NEXT')
    pdf.cell(0, 5, 'STASIUN METEOROLOGI UMBU MEHANG KUNDA', align='L', new_x='LMARGIN', new_y='NEXT')
    pdf.cell(0, 5, f'TAHUN : {tahun}', align='L', new_x='LMARGIN', new_y='NEXT')
    if triwulan_label:
        pdf.cell(0, 5, f'{triwulan_label.upper()}', align='L', new_x='LMARGIN', new_y='NEXT')
    pdf.cell(0, 5, f'BULAN : {bulan_nama.upper()}', align='L', new_x='LMARGIN', new_y='NEXT')
    pdf.ln(4)
    
    w_name = 60
    w_day = 7
    
    y_hdr = pdf.get_y()
    pdf.set_font('helvetica', 'B', 8)
    pdf.set_fill_color(220, 220, 220)
    
    # Header Tabel: NAMA PEGAWAI | TANGGAL
    pdf.rect(10, y_hdr, w_name, 10, style='DF')
    pdf.rect(10 + w_name, y_hdr, w_day * 31, 5, style='DF')
    
    pdf.set_xy(10, y_hdr + 3)
    pdf.cell(w_name, 4, 'NAMA PEGAWAI', align='C')
    pdf.set_xy(10 + w_name, y_hdr + 0.5)
    pdf.cell(w_day * 31, 4, 'TANGGAL', align='C')
    
    # Header Tabel: Angka 1..31
    for d in range(1, 32):
        x_d = 10 + w_name + (d - 1) * w_day
        pdf.rect(x_d, y_hdr + 5, w_day, 5, style='DF')
        pdf.set_xy(x_d, y_hdr + 5.5)
        pdf.cell(w_day, 4, str(d), align='C')
        
    y_data = y_hdr + 10
    row_h = 6
    total_row_h = row_h * 2 # Merge tinggi sel tanggal 1-31 (12 mm)
    
    # Kolom Kiri: Nama Pegawai
    pdf.rect(10, y_data, w_name, row_h)
    pdf.set_xy(10, y_data + 1)
    pdf.set_font('helvetica', 'B', 7)
    pdf.cell(w_name, 4, nama_teknisi, align='L')
    
    # Kolom Kiri: NIP
    pdf.rect(10, y_data + row_h, w_name, row_h)
    pdf.set_xy(10, y_data + row_h + 1)
    pdf.set_font('helvetica', '', 7)
    pdf.cell(w_name, 4, f"NIP. {nip}" if nip else "", align='L')
    
    # Kolom Tanggal (1..31): DI-MERGE DENGAN BARIS NIP (Tinggi 12mm)
    pdf.set_font('helvetica', '', 6)
    for d in range(1, 32):
        x_d = 10 + w_name + (d - 1) * w_day
        pdf.rect(x_d, y_data, w_day, total_row_h) # Kotak gabungan rata atas hingga bawah
        val = days_dict.get(d, '')
        pdf.set_xy(x_d, y_data + (total_row_h / 2) - 2)
        pdf.cell(w_day, 4, val, align='C')
        
    pdf.set_y(y_data + total_row_h + 8)
    y_sec = pdf.get_y()
    
    # Legenda Keterangan
    pdf.set_font('helvetica', 'B', 8)
    pdf.set_xy(10, y_sec)
    pdf.cell(40, 4, 'KETERANGAN:', new_x='LMARGIN', new_y='NEXT')
    
    legenda = [
        ("O", "= Pemeliharaan Taman Alat"),
        ("D", "= Pemeliharaan Display Bandara"),
        ("A", "= Pemeliharaan AWOS"),
        ("DL", "= Pengolahan OLA dan SLA"),
        ("G", "= Pembuatan Gas")
    ]
    
    for kode, ket in legenda:
        pdf.set_x(10)
        pdf.set_font('helvetica', 'B', 8)
        pdf.cell(8, 4, kode, align='L')
        pdf.set_font('helvetica', '', 8)
        pdf.cell(60, 4, ket, new_x='LMARGIN', new_y='NEXT')
        
    # Tanda Tangan
    pdf.set_xy(200, y_sec)
    pdf.set_font('helvetica', '', 9)
    pdf.cell(80, 4, f'Waingapu, 31 {bulan_nama} {tahun}', align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.set_x(200)
    pdf.cell(80, 4, 'Kepala Stasiun Meteorologi', align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.set_x(200)
    pdf.cell(80, 4, 'Umbu Mehang Kunda', align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.ln(18)
    pdf.set_x(200)
    pdf.set_font('helvetica', 'BU', 9)
    pdf.cell(80, 4, 'Carles Alexander Tari, S.TP', align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.set_x(200)
    pdf.set_font('helvetica', '', 9)
    pdf.cell(80, 4, 'NIP. 197712082001121001', align='C', new_x='LMARGIN', new_y='NEXT')

def generate_pdf_bytes(nama_teknisi, label_periode, triwulan_label, tahun, df_kegiatan, uploaded_excel, poin_korektif, list_bulan_logbook):
    pdf = PDFKinerja()
    
    # --- 1. KOVER BERKOP REKAPITULASI ---
    pdf.add_page(orientation='portrait')
    pdf.cetak_kop_surat()
    pdf.set_font('helvetica', 'B', 11)
    pdf.cell(0, 6, 'MENINGKATNYA LAYANAN OPERASIONAL', align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.cell(0, 6, 'ALOPTAMA METEOROLOGI YANG PRIMA', align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.cell(0, 6, 'STASIUN METEOROLOGI KELAS III UMBU MEHANG KUNDA', align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.cell(0, 6, f'{label_periode.upper()} TAHUN {tahun}', align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.ln(8)
    
    pdf.set_font('helvetica', '', 11)
    pdf.multi_cell(0, 6, "Persentase Alat Operasional Utama Meteorologi yang Laik Operasi dengan Target 97% di Stasiun Meteorologi Kelas III Umbu Mehang Kunda diperoleh menggunakan formula perhitungan sebagai berikut:")
    pdf.ln(5)
    
    pdf.set_font('helvetica', 'I', 11)
    pdf.cell(0, 6, "Laik Operasi Aloptama MET = (Jumlah aloptama meteorologi yg terpelihara / Jumlah Aloptama meteorologi) x 100%", align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.ln(3)
    pdf.set_font('helvetica', 'B', 11)
    pdf.cell(0, 6, "Laik Operasi Aloptama MET = 24/24 x 100%", align='C', new_x='LMARGIN', new_y='NEXT')
    pdf.ln(5)
    
    pdf.set_font('helvetica', '', 11)
    pdf.multi_cell(0, 6, f"Sehingga hasil persentase untuk peralatan operasional pada periode terkait menghasilkan nilai akurasi sebesar 100%. Adapun capaian hasil tersebut disusun dan didukung oleh laporan yang menjadi dasar pemantauan, pemeliharaan, serta evaluasi kinerja peralatan operasional utama meteorologi, yaitu sebagai berikut:")
    pdf.ln(5)
    
    items = [
        "Laporan Hasil Monitoring Kondisi Peralatan Operasional Utama Meteorologi",
        "Laporan Hasil Pemeliharaan Preventif Peralatan Operasional Utama Meteorologi",
        "Laporan Ketersediaan Data Peralatan Otomatis"
    ]
    if poin_korektif:
        items.append("Laporan Hasil Pemeliharaan Korektif / Kalibrasi Peralatan")
        
    for idx, item in enumerate(items, 1):
        pdf.cell(10, 6, f"{idx}.")
        pdf.cell(0, 6, item, new_x='LMARGIN', new_y='NEXT')
        pdf.set_x(10)

    # --- 2. LOGBOOK EXCEL ---
    df_jadwal_sheet = None
    if uploaded_excel is not None:
        try:
            df_log = pd.read_excel(uploaded_excel, sheet_name='LOGBOOK', header=None)
            try:
                df_jadwal_sheet = pd.read_excel(uploaded_excel, sheet_name='JADWAL 2026', header=None)
            except Exception:
                df_jadwal_sheet = None

            for bulan_item in list_bulan_logbook:
                stamet_data, posmet_data = parse_logbook(df_log, bulan_item)
                total_aloptama = sum(1 for r in stamet_data + posmet_data if str(r[0]).strip().isdigit())
                if total_aloptama == 0:
                    total_aloptama = 24
                
                pdf.add_page(orientation='landscape')
                pdf.cetak_kop_surat()
                draw_logbook_page(pdf, f"STASIUN METEOROLOGI UMBU MEHANG KUNDA - BULAN {bulan_item.upper()}", stamet_data, total_aloptama, is_posmet=False)
                
                pdf.add_page(orientation='landscape')
                pdf.set_y(15) 
                draw_logbook_page(pdf, f"POS METEOROLOGI TAMBOLAKA - BULAN {bulan_item.upper()}", posmet_data, total_aloptama, is_posmet=True)
        except Exception as e:
            pdf.add_page()
            pdf.cell(0, 10, f"Error membaca sheet LOGBOOK: {e}", align='C')

    # --- 3. JADWAL PEMELIHARAAN (TANPA KOP SURAT) ---
    if df_jadwal_sheet is not None:
        for bulan_item in list_bulan_logbook:
            name_res, nip_res, days_dict = parse_jadwal(df_jadwal_sheet, nama_teknisi, bulan_item)
            pdf.add_page(orientation='landscape')
            pdf.set_y(15) # Mulai dari margin atas tanpa kop surat
            draw_jadwal_page(pdf, name_res, nip_res, bulan_item, tahun, triwulan_label, days_dict)

    # --- 4. LAMPIRAN DOKUMENTASI FOTO ---
    pdf.add_page(orientation='portrait')
    pdf.set_y(15)
    pdf.set_font('helvetica', 'B', 11)
    pdf.cell(0, 10, f'LAMPIRAN KEGIATAN TEKNISI REGULER: {nama_teknisi.upper()}', align='L', new_x='LMARGIN', new_y='NEXT')
    pdf.ln(2)
    
    pdf.set_fill_color(180, 200, 255)
    pdf.set_font('helvetica', 'B', 10)
    pdf.cell(15, 8, 'NO', border=1, align='C', fill=True)
    pdf.cell(175, 8, 'PENJELASAN', border=1, align='C', fill=True, new_x='LMARGIN', new_y='NEXT')
    
    if df_kegiatan.empty:
        pdf.set_font('helvetica', '', 10)
        y_start = pdf.get_y()
        pdf.rect(10, y_start, 15, 15)
        pdf.rect(25, y_start, 175, 15)
        pdf.set_xy(25, y_start + 4)
        pdf.cell(175, 6, "Belum ada data dokumentasi untuk periode ini.", align='C')
    else:
        pdf.set_font('helvetica', '', 10)
        y_start = pdf.get_y()
        
        for idx, (_, row) in enumerate(df_kegiatan.iterrows(), 1):
            tgl_indo = format_tanggal_indo(row['tanggal'])
            teks = f"Tanggal: {tgl_indo}\n{row['penjelasan_kegiatan']}"
            row_h = 75
            
            if y_start + row_h > 270:
                pdf.add_page(orientation='portrait')
                pdf.set_y(15)
                pdf.set_font('helvetica', 'B', 10)
                pdf.cell(15, 8, 'NO', border=1, align='C', fill=True)
                pdf.cell(175, 8, 'PENJELASAN', border=1, align='C', fill=True, new_x='LMARGIN', new_y='NEXT')
                pdf.set_font('helvetica', '', 10)
                y_start = pdf.get_y()
                
            pdf.rect(10, y_start, 15, row_h)
            pdf.rect(25, y_start, 175, row_h)
            
            pdf.set_xy(10, y_start + (row_h/2) - 3)
            pdf.cell(15, 6, str(idx), align='C')
            
            pdf.set_xy(25, y_start + 4)
            pdf.multi_cell(175, 5, teks, align='C')
            
            img_y = y_start + 16
            foto_path = row['foto_path']
            
            if pd.notna(foto_path) and os.path.exists(foto_path):
                try:
                    with Image.open(foto_path) as img:
                        img_w, img_h = img.size
                    
                    max_w, max_h = 130.0, 54.0
                    ratio = min(max_w / img_w, max_h / img_h)
                    fit_w = img_w * ratio
                    fit_h = img_h * ratio
                    
                    fit_x = 25.0 + (175.0 - fit_w) / 2.0
                    fit_y = img_y + (max_h - fit_h) / 2.0
                    
                    pdf.image(foto_path, x=fit_x, y=fit_y, w=fit_w, h=fit_h)
                except Exception:
                    pdf.set_xy(25, img_y + 20)
                    pdf.cell(175, 6, "[Format Gambar Error / Corrupt]", align='C')
            else:
                pdf.set_xy(25, img_y + 20)
                pdf.cell(175, 6, "[Tidak Ada Gambar Diunggah]", align='C')
                
            y_start += row_h
            
    return bytes(pdf.output())

# --- STREAMLIT UI ---
st.title("🛠️ Laporan & E-Kinerja Teknisi")

tab1, tab2, tab3 = st.tabs(["📝 Input Kegiatan Harian", "📅 Data Kinerja (DB)", "🖨️ Cetak PDF E-Kinerja"])

opsi_kegiatan = [
    "Pemeliharaan Taman Alat",
    "Pemeliharaan Display Bandara",
    "Pemeliharaan AWOS",
    "Pengolahan OLA dan SLA",
    "Pembuatan Gas"
]
teknisi_list = [
    "Zulqha Ariandi Al Zikri, S.Tr.Inst.", 
    "Adi Junaidi Rachman, S.Kom.", 
    "Luqmanul Hakim, S.Tr.", 
    "Mohammad Hasyim Hanif, S.Tr.Inst."
]

with tab1:
    st.subheader("Form Input Dokumentasi Kegiatan")
    with st.form("form_kinerja", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            tanggal = st.date_input("Tanggal Kegiatan", datetime.today())
            nama = st.selectbox("Nama Teknisi", teknisi_list)
        with col2:
            kegiatan_dipilih = st.multiselect("Penjelasan Kegiatan", opsi_kegiatan)
            
        foto = st.file_uploader("Upload Foto Dokumentasi", type=['jpg', 'jpeg', 'png'])
        submit = st.form_submit_button("Simpan Data")
        
        if submit:
            if not kegiatan_dipilih:
                st.error("Silakan pilih minimal satu Penjelasan Kegiatan!")
            else:
                kegiatan_str = ", ".join(kegiatan_dipilih) 
                file_path = None
                if foto is not None:
                    file_path = os.path.join(UPLOAD_DIR, f"{datetime.now().timestamp()}_{foto.name}")
                    with open(file_path, "wb") as f:
                        f.write(foto.getbuffer())
                    
                with get_db_connection() as conn:
                    c = conn.cursor()
                    c.execute(
                        "INSERT INTO e_kinerja (tanggal, nama_teknisi, penjelasan_kegiatan, foto_path) VALUES (?, ?, ?, ?)",
                        (tanggal.strftime("%Y-%m-%d"), nama, kegiatan_str, file_path)
                    )
                    conn.commit()
                st.success("Data kegiatan berhasil disimpan ke database!")

with tab2:
    st.subheader("Arsip Kegiatan Teknisi")
    with get_db_connection() as conn:
        df_db = pd.read_sql("SELECT * FROM e_kinerja", conn)
    if not df_db.empty:
        st.dataframe(df_db, use_container_width=True)
    else:
        st.info("Belum ada data kegiatan yang diinput.")

with tab3:
    st.subheader("Generate Dokumen E-Kinerja & Logbook")
    
    jenis_laporan = st.radio("Pilih Tipe Laporan:", ["Bulanan", "Triwulan"], horizontal=True)
    
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        filter_nama = st.selectbox("Pilih Teknisi", teknisi_list)
    
    with col_b:
        if jenis_laporan == "Bulanan":
            filter_bulan = st.selectbox("Pilih Bulan", list(MONTH_MAP.keys()))
            filter_triwulan = None
        else:
            filter_triwulan = st.selectbox("Pilih Triwulan", list(QUARTER_MAP.keys()))
            filter_bulan = None

    with col_c:
        filter_tahun = st.selectbox("Pilih Tahun", ["2026", "2027", "2028"])
        
    st.markdown("---")
    st.markdown("#### 🖇️ Konfigurasi Lampiran")
    
    uploaded_excel = st.file_uploader("1. Wajib: File Excel Logbook (Peralatan Teknis.xlsx)", type=['xlsx', 'xls'])
    poin_korektif = st.checkbox("Tambahkan Poin Ke-4 (Laporan Korektif/Kalibrasi) di Narasi Hal. 1")
    
    pdf_kalibrasi = None
    if PYPDF_INSTALLED:
        pdf_kalibrasi = st.file_uploader("2. Opsional: Upload PDF Laporan Kalibrasi/Korektif", type=['pdf'])
    else:
        st.caption("📝 *Install `pypdf` (`pip install pypdf`) untuk mengaktifkan penggabungan file PDF.*")
    
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Tarik Data & Generate PDF Lengkap 🚀"):
        if uploaded_excel is None:
            st.error("Silakan unggah File Excel terlebih dahulu untuk menarik Logbook!")
        else:
            if jenis_laporan == "Bulanan":
                label_periode = f"BULAN {filter_bulan.upper()}"
                triwulan_label = ""
                list_bulan_logbook = [filter_bulan]
                bulan_code = MONTH_MAP[filter_bulan]
                query = "SELECT * FROM e_kinerja WHERE nama_teknisi = ? AND strftime('%m', tanggal) = ? AND strftime('%Y', tanggal) = ?"
                params = [filter_nama, bulan_code, filter_tahun]
            else:
                q_info = QUARTER_MAP[filter_triwulan]
                label_periode = f"{q_info['label']} ({q_info['month_names'][0].upper()} - {q_info['month_names'][-1].upper()})"
                triwulan_label = q_info['label']
                list_bulan_logbook = q_info["month_names"]
                months_code = q_info["months"]
                query = f"SELECT * FROM e_kinerja WHERE nama_teknisi = ? AND strftime('%m', tanggal) IN ({','.join(['?']*len(months_code))}) AND strftime('%Y', tanggal) = ?"
                params = [filter_nama] + months_code + [filter_tahun]
            
            with get_db_connection() as conn:
                df_filter = pd.read_sql(query, conn, params=params)
            
            with st.spinner('Menyusun Logbook, Jadwal, & Memformat PDF...'):
                pdf_bytes = generate_pdf_bytes(
                    filter_nama, 
                    label_periode, 
                    triwulan_label,
                    filter_tahun, 
                    df_filter, 
                    uploaded_excel, 
                    poin_korektif, 
                    list_bulan_logbook
                )
                
                if PYPDF_INSTALLED and pdf_kalibrasi is not None:
                    merger = PdfWriter()
                    
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as f_base:
                        f_base.write(pdf_bytes)
                        base_path = f_base.name

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as f_extra:
                        f_extra.write(pdf_kalibrasi.getbuffer())
                        extra_path = f_extra.name

                    merger.append(base_path)
                    merger.append(extra_path)
                    
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as f_out:
                        merger.write(f_out)
                        final_path = f_out.name

                    with open(final_path, "rb") as f_final:
                        output_data = f_final.read()

                    os.remove(base_path)
                    os.remove(extra_path)
                    os.remove(final_path)
                else:
                    output_data = pdf_bytes

            filename = f"E_Kinerja_{filter_nama.replace(' ', '_')}_{jenis_laporan}_{filter_tahun}.pdf"
            st.success("✅ Dokumen PDF E-Kinerja & Jadwal Pemeliharaan berhasil dibuat!")
            st.download_button(
                label="⬇️ Download Hasil PDF E-Kinerja",
                data=output_data,
                file_name=filename,
                mime="application/pdf"
            )
