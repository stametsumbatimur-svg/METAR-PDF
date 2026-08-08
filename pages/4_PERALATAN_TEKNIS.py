import streamlit as st
import pandas as pd
import sqlite3
import os
import tempfile
import re
import json
import calendar
from datetime import datetime
from PIL import Image
from fpdf import FPDF

try:
    from pypdf import PdfWriter
    PYPDF_INSTALLED = True
except ImportError:
    PYPDF_INSTALLED = False

# --- KONFIGURASI METADATA STASIUN & PEJABAT ---
STATION_NAME = "STASIUN METEOROLOGI KELAS III UMBU MEHANG KUNDA"
HEAD_OF_STATION_NAME = "Carles Alexander Tari, S.TP"
HEAD_OF_STATION_NIP = "197712082001121001"
STATION_CITY = "Waingapu"
DEFAULT_ALAT_COUNT = 24

# --- KONFIGURASI FILE & FOLDER ---
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

# --- PALET WARNA JADWAL PEMELIHARAAN ---
COLOR_MAP_JADWAL = {
    "O": (169, 223, 191),   # Hijau Muda
    "D": (174, 214, 241),   # Biru Muda
    "A": (249, 231, 159),   # Kuning Muda
    "DL": (210, 180, 222),  # Ungu Muda
    "G": (248, 196, 113)    # Oranye Muda
}

# --- HELPER PEWARNAAN KATEGORI OLA SLA ---
def get_percentage_color(val):
    try:
        if isinstance(val, str):
            val = val.replace('%', '').strip()
        v = float(val)
        if 0 < v <= 1.0:
            v = v * 100.0
        v = round(v, 2)
    except (ValueError, TypeError):
        return (255, 255, 255)

    if v >= 100.0:
        return (169, 223, 191)
    elif v >= 80.0:
        return (249, 231, 159)
    else:
        return (241, 148, 138)

# --- HELPER PENANGGALAN DINAMIS ---
def get_last_day_of_month(bulan_nama, tahun):
    try:
        m = int(MONTH_MAP.get(bulan_nama, "01"))
        y = int(tahun)
        _, last_day = calendar.monthrange(y, m)
        return last_day
    except Exception:
        return 31

def parse_foto_paths(foto_path_val):
    """Mendukung format lama (string tunggal/koma) & format baru (JSON list)"""
    if not foto_path_val or pd.isna(foto_path_val):
        return []
    fstr = str(foto_path_val).strip()
    if fstr.startswith("["):
        try:
            return json.loads(fstr)
        except Exception:
            pass
    if "," in fstr:
        return [p.strip() for p in fstr.split(",") if p.strip()]
    return [fstr]

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
    def clean_text(s):
        return re.sub(r'[,.\-\s]+', ' ', str(s)).upper().strip()

    def get_name_keywords(n):
        clean = clean_text(n)
        titles = ['S TR INST', 'S TR', 'S KOM', 'S TP', 'M T', 'ST', 'A MD']
        for t in titles:
            clean = clean.replace(t, '')
        return [w for w in clean.split() if len(w) > 2]

    target_words = get_name_keywords(nama_teknisi)

    month_row = -1
    for r in range(len(df_jadwal)):
        row_str = " ".join([str(val).strip().upper() for val in df_jadwal.iloc[r].values if pd.notna(val)])
        if "BULAN" in row_str and bulan_nama.upper() in row_str and len(row_str) < 60:
            month_row = r
            break

    if month_row == -1:
        for r in range(len(df_jadwal)):
            row_str = " ".join([str(val).strip().upper() for val in df_jadwal.iloc[r].values if pd.notna(val)])
            if bulan_nama.upper() in row_str and ("JADWAL" in row_str or "TAHUN" in row_str or "BULAN" in row_str):
                month_row = r
                break

    if month_row == -1:
        return nama_teknisi, "", {}

    next_month_row = len(df_jadwal)
    other_months = [m.upper() for m in MONTH_MAP.keys() if m.upper() != bulan_nama.upper()]
    for r in range(month_row + 1, len(df_jadwal)):
        row_str = " ".join([str(val).strip().upper() for val in df_jadwal.iloc[r].values if pd.notna(val)])
        if "BULAN" in row_str and any(m in row_str for m in other_months):
            next_month_row = r
            break

    day_col_map = {}
    for r in range(month_row, min(month_row + 10, next_month_row)):
        row_vals = [str(val).strip() for val in df_jadwal.iloc[r].values]
        if '1' in row_vals and '2' in row_vals:
            for col_idx, val in enumerate(row_vals):
                val_clean = str(val).split('.')[0].strip()
                if val_clean.isdigit():
                    d = int(val_clean)
                    if 1 <= d <= 31 and d not in day_col_map:
                        day_col_map[d] = col_idx
            break

    if not day_col_map:
        for d in range(1, 32):
            day_col_map[d] = d + 1

    tech_row = -1
    for r in range(month_row, next_month_row):
        for c in range(min(5, df_jadwal.shape[1])):
            cell_val = str(df_jadwal.iloc[r, c])
            if pd.notna(cell_val) and cell_val.strip().lower() != 'nan':
                cell_words = get_name_keywords(cell_val)
                if len(target_words) >= 2 and all(w in cell_words for w in target_words):
                    tech_row = r
                    break
                elif len(target_words) == 1 and target_words[0] in cell_words:
                    tech_row = r
                    break
        if tech_row != -1:
            break

    if tech_row == -1:
        return nama_teknisi, "", {}

    nip = ""
    for r_nip in range(tech_row, min(tech_row + 3, next_month_row)):
        for c_nip in range(min(5, df_jadwal.shape[1])):
            val_nip = str(df_jadwal.iloc[r_nip, c_nip])
            digits = re.sub(r'\D', '', val_nip)
            if len(digits) >= 12:
                nip = digits
                break
        if nip:
            break

    jadwal_days = {}
    for d in range(1, 32):
        if d in day_col_map:
            c_idx = day_col_map[d]
            if c_idx < df_jadwal.shape[1]:
                val = str(df_jadwal.iloc[tech_row, c_idx]).strip()
                jadwal_days[d] = val if val.lower() != 'nan' else ""
        else:
            jadwal_days[d] = ""

    return nama_teknisi, nip, jadwal_days

# --- HELPER PARSER ITEM OLA SLA ---
def _extract_ola_item(df_ola, row_st, row_dt):
    status, data = [], []
    for c in range(3, 34):
        v_st = df_ola.iloc[row_st, c] if c < df_ola.shape[1] else None
        v_dt = df_ola.iloc[row_dt, c] if c < df_ola.shape[1] else None
        
        st_str = "" if pd.isna(v_st) or str(v_st).strip().lower() == 'nan' else str(v_st).strip()
        if pd.isna(v_dt) or str(v_dt).strip().lower() in ['nan', '']:
            dt_str = ""
        elif isinstance(v_dt, (int, float)):
            dt_str = f"{float(v_dt):.2f}"
        else:
            dt_str = str(v_dt).strip()
            
        status.append(st_str)
        data.append(dt_str)
        
    return {
        'no': str(df_ola.iloc[row_st, 0]).strip().replace('.0', ''),
        'kode': str(df_ola.iloc[row_st, 1]).strip().replace('.0', ''),
        'nama': str(df_ola.iloc[row_st, 2]).strip(),
        'status_days': status,
        'status_accum': df_ola.iloc[row_st, 34] if 34 < df_ola.shape[1] else 1,
        'data_ratios': data,
        'data_accum': df_ola.iloc[row_dt, 34] if 34 < df_ola.shape[1] else 1
    }

def parse_olasla(df_ola, bulan_nama):
    bulan_target = bulan_nama.upper()
    idx_bulan = [i for i in range(len(df_ola)) if str(df_ola.iloc[i, 2]).strip().upper() == bulan_target]
            
    if not idx_bulan:
        return None
        
    row_m = idx_bulan[0]
    return {
        'bulan': bulan_nama,
        'item1': _extract_ola_item(df_ola, row_m + 3, row_m + 4),
        'item2': _extract_ola_item(df_ola, row_m + 5, row_m + 6)
    }

# --- CLASS GENERATOR PDF ---
class PDFKinerja(FPDF):
    def cetak_kop_surat(self):
        y_awal = self.get_y()
        if os.path.exists(LOGO_FILE):
            self.image(LOGO_FILE, x=14, y=y_awal + 1, w=17)
            
        self.set_y(y_awal)
        self.set_font('helvetica', 'B', 13)
        self.cell(0, 5, 'BADAN METEOROLOGI, KLIMATOLOGI, DAN GEOFISIKA', align='C', ln=1)
        self.set_font('helvetica', 'B', 11)
        self.cell(0, 5, STATION_NAME, align='C', ln=1)
        
        self.set_font('helvetica', '', 9)
        self.cell(0, 4, 'Jl. Adi Sucipto, Waingapu, Sumba Timur', align='C', ln=1)
        self.cell(0, 4, 'Telp. (0387) 61227 | Fax: (0387) 61228 | Kode Pos 87114', align='C', ln=1)
        self.cell(0, 4, 'Email: stamet.sumbatimur@bmkg.go.id | Website: http://ntt.bmkg.go.id', align='C', ln=1)
        
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

def draw_logbook_page(pdf, title, data_rows, total_alat_keseluruhan=DEFAULT_ALAT_COUNT, is_posmet=False):
    pdf.set_font('helvetica', 'B', 10)
    pdf.cell(0, 6, 'LAPORAN HASIL MONITORING KONDISI PERALATAN OPERASIONAL UTAMA METEOROLOGI', align='C', ln=1)
    pdf.cell(0, 6, title, align='C', ln=1)
    pdf.ln(5)
    
    if not data_rows:
        pdf.set_font('helvetica', '', 10)
        pdf.cell(0, 6, "Data tidak ditemukan. Silakan cek format Excel / Bulan yang dipilih.", align='C', ln=1)
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

def draw_jadwal_page(pdf, nama_teknisi, nip, bulan_nama, tahun, triwulan_label, days_dict):
    pdf.set_font('helvetica', 'B', 10)
    pdf.cell(0, 5, 'JADWAL PEMELIHARAAN', align='L', ln=1)
    pdf.cell(0, 5, 'STASIUN METEOROLOGI UMBU MEHANG KUNDA', align='L', ln=1)
    pdf.cell(0, 5, f'TAHUN : {tahun}', align='L', ln=1)
    if triwulan_label:
        pdf.cell(0, 5, f'{triwulan_label.upper()}', align='L', ln=1)
    pdf.cell(0, 5, f'BULAN : {bulan_nama.upper()}', align='L', ln=1)
    pdf.ln(4)
    
    w_name = 60
    w_day = 7
    
    y_hdr = pdf.get_y()
    pdf.set_font('helvetica', 'B', 8)
    pdf.set_fill_color(220, 220, 220)
    
    pdf.rect(10, y_hdr, w_name, 10, style='DF')
    pdf.rect(10 + w_name, y_hdr, w_day * 31, 5, style='DF')
    
    pdf.set_xy(10, y_hdr + 3)
    pdf.cell(w_name, 4, 'NAMA PEGAWAI', align='C')
    pdf.set_xy(10 + w_name, y_hdr + 0.5)
    pdf.cell(w_day * 31, 4, 'TANGGAL', align='C')
    
    for d in range(1, 32):
        x_d = 10 + w_name + (d - 1) * w_day
        pdf.rect(x_d, y_hdr + 5, w_day, 5, style='DF')
        pdf.set_xy(x_d, y_hdr + 5.5)
        pdf.cell(w_day, 4, str(d), align='C')
        
    y_data = y_hdr + 10
    row_h = 6
    total_row_h = row_h * 2
    
    pdf.rect(10, y_data, w_name, row_h)
    pdf.set_xy(10, y_data + 1)
    pdf.set_font('helvetica', 'B', 7)
    pdf.cell(w_name, 4, nama_teknisi, align='L')
    
    pdf.rect(10, y_data + row_h, w_name, row_h)
    pdf.set_xy(10, y_data + row_h + 1)
    pdf.set_font('helvetica', '', 7)
    pdf.cell(w_name, 4, f"NIP. {nip}" if nip else "", align='L')
    
    pdf.set_font('helvetica', 'B', 6.5)
    for d in range(1, 32):
        x_d = 10 + w_name + (d - 1) * w_day
        val = str(days_dict.get(d, '')).strip().upper()
        tokens = [t for t in val.split() if t in COLOR_MAP_JADWAL]
        
        if len(tokens) == 1:
            r, g, b = COLOR_MAP_JADWAL[tokens[0]]
            pdf.set_fill_color(r, g, b)
            pdf.rect(x_d, y_data, w_day, total_row_h, style='DF')
        elif len(tokens) > 1:
            sub_w = w_day / len(tokens)
            for idx, tok in enumerate(tokens):
                r, g, b = COLOR_MAP_JADWAL[tok]
                pdf.set_fill_color(r, g, b)
                pdf.rect(x_d + (idx * sub_w), y_data, sub_w, total_row_h, style='DF')
            pdf.rect(x_d, y_data, w_day, total_row_h)
        else:
            pdf.rect(x_d, y_data, w_day, total_row_h)
            
        pdf.set_xy(x_d, y_data + (total_row_h / 2) - 2)
        pdf.cell(w_day, 4, val, align='C')
        
    pdf.set_y(y_data + total_row_h + 8)
    y_sec = pdf.get_y()
    
    pdf.set_font('helvetica', 'B', 8)
    pdf.set_xy(10, y_sec)
    pdf.cell(40, 4, 'KETERANGAN:', ln=1)
    
    legenda = [
        ("O", "= Pemeliharaan Taman Alat"),
        ("D", "= Pemeliharaan Display Bandara"),
        ("A", "= Pemeliharaan AWOS"),
        ("DL", "= Pengolahan OLA dan SLA"),
        ("G", "= Pembuatan Gas")
    ]
    
    for kode, ket in legenda:
        pdf.set_x(10)
        curr_y = pdf.get_y()
        if kode in COLOR_MAP_JADWAL:
            r, g, b = COLOR_MAP_JADWAL[kode]
            pdf.set_fill_color(r, g, b)
            pdf.rect(10, curr_y, 8, 4, style='DF')
            
        pdf.set_font('helvetica', 'B', 8)
        pdf.cell(8, 4, kode, align='C')
        pdf.set_font('helvetica', '', 8)
        pdf.cell(60, 4, ket, ln=1)
        
    last_day = get_last_day_of_month(bulan_nama, tahun)
    pdf.set_xy(200, y_sec)
    pdf.set_font('helvetica', '', 9)
    pdf.cell(80, 4, f'{STATION_CITY}, {last_day} {bulan_nama} {tahun}', align='C', ln=1)
    pdf.set_x(200)
    pdf.cell(80, 4, 'Kepala Stasiun Meteorologi', align='C', ln=1)
    pdf.set_x(200)
    pdf.cell(80, 4, 'Umbu Mehang Kunda', align='C', ln=1)
    pdf.ln(18)
    pdf.set_x(200)
    pdf.set_font('helvetica', 'BU', 9)
    pdf.cell(80, 4, HEAD_OF_STATION_NAME, align='C', ln=1)
    pdf.set_x(200)
    pdf.set_font('helvetica', '', 9)
    pdf.cell(80, 4, f'NIP. {HEAD_OF_STATION_NIP}', align='C', ln=1)

def draw_olasla_page(pdf, ola_data, nama_teknisi, nip_teknisi, tahun):
    if not ola_data:
        pdf.set_font('helvetica', '', 10)
        pdf.cell(0, 10, "Data OLA SLA tidak ditemukan untuk bulan ini.", align='C', ln=1)
        return
        
    bulan_nama = ola_data['bulan']
    
    pdf.set_font('helvetica', 'B', 10)
    pdf.cell(0, 5, 'MONITORING HARIAN KONDISI AWOS KAT.I DAN AWS STRENGTHENING', align='C', ln=1)
    pdf.cell(0, 5, STATION_NAME, align='C', ln=1)
    pdf.cell(0, 5, f'TAHUN {tahun}', align='C', ln=1)
    pdf.cell(0, 5, f'{bulan_nama.upper()}', align='C', ln=1)
    pdf.ln(4)
    
    w_no = 6
    w_kode = 12
    w_nama = 42
    w_day = 6.2
    w_accum = 24
    
    y_hdr = pdf.get_y()
    pdf.set_font('helvetica', 'B', 7)
    pdf.set_fill_color(220, 220, 220)
    
    w_left = w_no + w_kode + w_nama
    pdf.rect(10, y_hdr, w_left, 10, style='DF')
    pdf.rect(10 + w_left, y_hdr, w_day * 31, 5, style='DF')
    pdf.rect(10 + w_left + w_day * 31, y_hdr, w_accum, 10, style='DF')
    
    pdf.set_xy(10, y_hdr + 3)
    pdf.cell(w_left, 4, 'NAMA DAN LOKASI PERALATAN', align='C')
    pdf.set_xy(10 + w_left, y_hdr + 0.5)
    pdf.cell(w_day * 31, 4, 'KONDISI PERALATAN DAN DATA TERSEDIA', align='C')
    pdf.set_xy(10 + w_left + w_day * 31, y_hdr + 1)
    pdf.multi_cell(w_accum, 3.5, 'AKUMULASI\nON-LINE', align='C')
    
    for d in range(1, 32):
        x_d = 10 + w_left + (d - 1) * w_day
        pdf.rect(x_d, y_hdr + 5, w_day, 5, style='DF')
        pdf.set_xy(x_d, y_hdr + 5.5)
        pdf.cell(w_day, 4, str(d), align='C')
        
    y_data = y_hdr + 10
    row_h = 5.5
    items = [ola_data['item1'], ola_data['item2']]
    
    for item in items:
        pdf.rect(10, y_data, w_no, row_h)
        pdf.set_xy(10, y_data + 1)
        pdf.set_font('helvetica', 'B', 7)
        pdf.cell(w_no, 4, str(item['no']), align='C')
        
        pdf.rect(10 + w_no, y_data, w_kode, row_h)
        pdf.set_xy(10 + w_no, y_data + 1)
        pdf.set_font('helvetica', '', 7)
        pdf.cell(w_kode, 4, str(item['kode']), align='C')
        
        pdf.rect(10 + w_no + w_kode, y_data, w_nama, row_h)
        pdf.set_xy(10 + w_no + w_kode, y_data + 1)
        pdf.set_font('helvetica', 'B', 7)
        pdf.cell(w_nama, 4, str(item['nama']), align='L')
        
        pdf.set_font('helvetica', 'B', 5.5)
        for d in range(1, 32):
            x_d = 10 + w_left + (d - 1) * w_day
            st_val = str(item['status_days'][d-1] if d-1 < len(item['status_days']) else '').strip().upper()
            
            if st_val == "ON":
                pdf.set_fill_color(169, 223, 191)
                pdf.rect(x_d, y_data, w_day, row_h, style='DF')
            elif st_val == "OFF":
                pdf.set_fill_color(241, 148, 138)
                pdf.rect(x_d, y_data, w_day, row_h, style='DF')
            else:
                pdf.rect(x_d, y_data, w_day, row_h)
                
            pdf.set_xy(x_d, y_data + 1)
            pdf.cell(w_day, 4, st_val, align='C')
            
        acc_st = item['status_accum']
        r, g, b = get_percentage_color(acc_st)
        pdf.set_fill_color(r, g, b)
        pdf.rect(10 + w_left + w_day * 31, y_data, w_accum, row_h, style='DF')
        pdf.set_xy(10 + w_left + w_day * 31, y_data + 1)
        pdf.set_font('helvetica', 'B', 7)
        acc_st_str = f"{float(acc_st)*100:.0f}%" if isinstance(acc_st, (int, float)) and acc_st <= 1 else str(acc_st)
        pdf.cell(w_accum, 4, acc_st_str, align='C')
        
        y_data += row_h
        
        pdf.rect(10, y_data, w_no, row_h)
        pdf.rect(10 + w_no, y_data, w_kode, row_h)
        pdf.set_xy(10 + w_no, y_data + 1)
        pdf.set_font('helvetica', 'B', 7)
        pdf.cell(w_kode, 4, 'DATA', align='C')
        
        pdf.rect(10 + w_no + w_kode, y_data, w_nama, row_h)
        
        pdf.set_font('helvetica', '', 5)
        for d in range(1, 32):
            x_d = 10 + w_left + (d - 1) * w_day
            dt_str = item['data_ratios'][d-1] if d-1 < len(item['data_ratios']) else ''
            
            if dt_str != "":
                r_d, g_d, b_d = get_percentage_color(dt_str)
                pdf.set_fill_color(r_d, g_d, b_d)
                pdf.rect(x_d, y_data, w_day, row_h, style='DF')
            else:
                pdf.rect(x_d, y_data, w_day, row_h)
                
            pdf.set_xy(x_d, y_data + 1)
            pdf.cell(w_day, 4, str(dt_str), align='C')
            
        acc_dt = item['data_accum']
        r_da, g_da, b_da = get_percentage_color(acc_dt)
        pdf.set_fill_color(r_da, g_da, b_da)
        pdf.rect(10 + w_left + w_day * 31, y_data, w_accum, row_h, style='DF')
        pdf.set_xy(10 + w_left + w_day * 31, y_data + 1)
        pdf.set_font('helvetica', 'B', 7)
        acc_dt_str = f"{float(acc_dt)*100:.2f}%" if isinstance(acc_dt, (int, float)) and acc_dt <= 1 else str(acc_dt)
        pdf.cell(w_accum, 4, acc_dt_str, align='C')
        
        y_data += row_h
        
    pdf.set_y(y_data + 8)
    y_sec = pdf.get_y()
    
    pdf.set_xy(20, y_sec)
    pdf.set_font('helvetica', '', 8)
    pdf.cell(80, 4, 'Mengetahui,', ln=1)
    pdf.set_x(20)
    pdf.cell(80, 4, 'Kepala Stasiun', ln=1)
    pdf.ln(14)
    pdf.set_x(20)
    pdf.set_font('helvetica', 'BU', 8)
    pdf.cell(80, 4, HEAD_OF_STATION_NAME, ln=1)
    pdf.set_x(20)
    pdf.set_font('helvetica', '', 8)
    pdf.cell(80, 4, f'NIP. {HEAD_OF_STATION_NIP}', ln=1)
    
    last_day = get_last_day_of_month(bulan_nama, tahun)
    pdf.set_xy(200, y_sec)
    pdf.set_font('helvetica', '', 8)
    pdf.cell(80, 4, f'{STATION_CITY}, {last_day} {bulan_nama} {tahun}', align='C', ln=1)
    pdf.set_x(200)
    pdf.cell(80, 4, 'Teknisi', align='C', ln=1)
    pdf.ln(14)
    pdf.set_x(200)
    pdf.set_font('helvetica', 'BU', 8)
    pdf.cell(80, 4, nama_teknisi, align='C', ln=1)
    pdf.set_x(200)
    pdf.set_font('helvetica', '', 8)
    pdf.cell(80, 4, f"NIP. {nip_teknisi}" if nip_teknisi else "", align='C', ln=1)

def generate_pdf_bytes(nama_teknisi, label_periode, triwulan_label, tahun, df_kegiatan, uploaded_excel, poin_korektif, list_bulan_logbook):
    pdf = PDFKinerja()
    
    # --- 1. KOVER BERKOP REKAPITULASI ---
    pdf.add_page(orientation='portrait')
    pdf.cetak_kop_surat()
    pdf.set_font('helvetica', 'B', 11)
    pdf.cell(0, 6, 'MENINGKATNYA LAYANAN OPERASIONAL', align='C', ln=1)
    pdf.cell(0, 6, 'ALOPTAMA METEOROLOGI YANG PRIMA', align='C', ln=1)
    pdf.cell(0, 6, STATION_NAME, align='C', ln=1)
    pdf.cell(0, 6, f'{label_periode.upper()} TAHUN {tahun}', align='C', ln=1)
    pdf.ln(8)
    
    pdf.set_font('helvetica', '', 11)
    pdf.multi_cell(0, 6, f"Persentase Alat Operasional Utama Meteorologi yang Laik Operasi dengan Target 97% di {STATION_NAME} diperoleh menggunakan formula perhitungan sebagai berikut:")
    pdf.ln(5)
    
    pdf.set_font('helvetica', 'I', 11)
    pdf.cell(0, 6, "Laik Operasi Aloptama MET = (Jumlah aloptama meteorologi yg terpelihara / Jumlah Aloptama meteorologi) x 100%", align='C', ln=1)
    pdf.ln(3)
    pdf.set_font('helvetica', 'B', 11)
    pdf.cell(0, 6, f"Laik Operasi Aloptama MET = {DEFAULT_ALAT_COUNT}/{DEFAULT_ALAT_COUNT} x 100%", align='C', ln=1)
    pdf.ln(5)
    
    pdf.set_font('helvetica', '', 11)
    pdf.multi_cell(0, 6, "Sehingga hasil persentase untuk peralatan operasional pada periode terkait menghasilkan nilai akurasi sebesar 100%. Adapun capaian hasil tersebut disusun dan didukung oleh laporan yang menjadi dasar pemantauan, pemeliharaan, serta evaluasi kinerja peralatan operasional utama meteorologi, yaitu sebagai berikut:")
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
        pdf.cell(0, 6, item, ln=1)
        pdf.set_x(10)

    df_log = None
    df_jadwal_sheet = None
    df_ola_sheet = None
    
    if uploaded_excel is not None:
        try:
            df_log = pd.read_excel(uploaded_excel, sheet_name='LOGBOOK', header=None)
        except Exception:
            pass
        try:
            df_jadwal_sheet = pd.read_excel(uploaded_excel, sheet_name=f'JADWAL {tahun}', header=None)
        except Exception:
            try:
                xl = pd.ExcelFile(uploaded_excel)
                match = [s for s in xl.sheet_names if 'JADWAL' in s.upper()]
                if match: df_jadwal_sheet = pd.read_excel(uploaded_excel, sheet_name=match[0], header=None)
            except Exception: pass
        try:
            df_ola_sheet = pd.read_excel(uploaded_excel, sheet_name=f'OLA SLA {tahun}', header=None)
        except Exception:
            try:
                xl = pd.ExcelFile(uploaded_excel)
                match = [s for s in xl.sheet_names if 'OLA' in s.upper()]
                if match: df_ola_sheet = pd.read_excel(uploaded_excel, sheet_name=match[0], header=None)
            except Exception: pass

    # --- 2. LOGBOOK EXCEL ---
    if df_log is not None:
        for bulan_item in list_bulan_logbook:
            stamet_data, posmet_data = parse_logbook(df_log, bulan_item)
            total_aloptama = sum(1 for r in stamet_data + posmet_data if str(r[0]).strip().isdigit())
            if total_aloptama == 0:
                total_aloptama = DEFAULT_ALAT_COUNT
            
            pdf.add_page(orientation='landscape')
            pdf.cetak_kop_surat()
            draw_logbook_page(pdf, f"STASIUN METEOROLOGI UMBU MEHANG KUNDA - BULAN {bulan_item.upper()}", stamet_data, total_aloptama, is_posmet=False)
            
            pdf.add_page(orientation='landscape')
            pdf.set_y(15) 
            draw_logbook_page(pdf, f"POS METEOROLOGI TAMBOLAKA - BULAN {bulan_item.upper()}", posmet_data, total_aloptama, is_posmet=True)

    nip_teknisi_dinamis = ""

    # --- 3. JADWAL PEMELIHARAAN ---
    if df_jadwal_sheet is not None:
        for bulan_item in list_bulan_logbook:
            name_res, nip_res, days_dict = parse_jadwal(df_jadwal_sheet, nama_teknisi, bulan_item)
            if nip_res: 
                nip_teknisi_dinamis = nip_res
            pdf.add_page(orientation='landscape')
            pdf.set_y(15)
            draw_jadwal_page(pdf, name_res, nip_res, bulan_item, tahun, triwulan_label, days_dict)

    # --- 4. LAMPIRAN DOKUMENTASI FOTO (LAYOUT GRID MULTI-FOTO DINAMIS) ---
    for bulan_item in list_bulan_logbook:
        b_code = MONTH_MAP[bulan_item]
        df_kegiatan_bulan = df_kegiatan[df_kegiatan['tanggal'].astype(str).str.slice(5, 7) == b_code] if not df_kegiatan.empty else pd.DataFrame()
        
        pdf.add_page(orientation='portrait')
        pdf.set_y(15)
        pdf.set_font('helvetica', 'B', 11)
        pdf.cell(0, 10, f'LAMPIRAN KEGIATAN TEKNISI REGULER: {nama_teknisi.upper()} - BULAN {bulan_item.upper()}', align='L', ln=1)
        pdf.ln(2)
        
        pdf.set_fill_color(180, 200, 255)
        pdf.set_font('helvetica', 'B', 10)
        pdf.cell(15, 8, 'NO', border=1, align='C', fill=True)
        pdf.cell(175, 8, 'PENJELASAN', border=1, align='C', fill=True, ln=1)
        
        if df_kegiatan_bulan.empty:
            pdf.set_font('helvetica', '', 10)
            y_start = pdf.get_y()
            pdf.rect(10, y_start, 15, 15)
            pdf.rect(25, y_start, 175, 15)
            pdf.set_xy(25, y_start + 4)
            pdf.cell(175, 6, f"Belum ada data dokumentasi untuk bulan {bulan_item}.", align='C')
        else:
            pdf.set_font('helvetica', '', 10)
            y_start = pdf.get_y()
            
            for idx, (_, row) in enumerate(df_kegiatan_bulan.iterrows(), 1):
                tgl_indo = format_tanggal_indo(row['tanggal'])
                teks = f"Tanggal: {tgl_indo}\n{row['penjelasan_kegiatan']}"
                
                raw_photos = parse_foto_paths(row.get('foto_path'))
                valid_photos = [p for p in raw_photos if pd.notna(p) and os.path.exists(p)]
                num_photos = len(valid_photos)
                
                # Hitung jumlah baris foto untuk layout grid (2 kolom)
                img_rows = 1 if num_photos <= 1 else (num_photos + 1) // 2
                row_h = max(65, 15 + (img_rows * 48))
                
                if y_start + row_h > 270:
                    pdf.add_page(orientation='portrait')
                    pdf.set_y(15)
                    pdf.set_font('helvetica', 'B', 10)
                    pdf.cell(15, 8, 'NO', border=1, align='C', fill=True)
                    pdf.cell(175, 8, 'PENJELASAN', border=1, align='C', fill=True, ln=1)
                    pdf.set_font('helvetica', '', 10)
                    y_start = pdf.get_y()
                    
                pdf.rect(10, y_start, 15, row_h)
                pdf.rect(25, y_start, 175, row_h)
                
                pdf.set_xy(10, y_start + (row_h/2) - 3)
                pdf.cell(15, 6, str(idx), align='C')
                
                pdf.set_xy(25, y_start + 3)
                pdf.multi_cell(175, 5, teks, align='C')
                
                img_base_y = y_start + 15
                
                if valid_photos:
                    if num_photos == 1:
                        # 1 Foto: Posisi Tengah
                        fp = valid_photos[0]
                        try:
                            with Image.open(fp) as img:
                                img_w, img_h = img.size
                            max_w, max_h = 130.0, 45.0
                            ratio = min(max_w / img_w, max_h / img_h)
                            fit_w = img_w * ratio
                            fit_h = img_h * ratio
                            fit_x = 25.0 + (175.0 - fit_w) / 2.0
                            fit_y = img_base_y + (max_h - fit_h) / 2.0
                            pdf.image(fp, x=fit_x, y=fit_y, w=fit_w, h=fit_h)
                        except Exception:
                            pdf.set_xy(25, img_base_y + 15)
                            pdf.cell(175, 6, "[Format Gambar Error / Corrupt]", align='C')
                    else:
                        # > 1 Foto: Render Grid 2 Kolom
                        box_w, box_h = 82.0, 44.0
                        for p_idx, fp in enumerate(valid_photos):
                            r_idx = p_idx // 2
                            c_idx = p_idx % 2
                            x_pos = 28.0 + c_idx * (box_w + 5.0)
                            y_pos = img_base_y + r_idx * (box_h + 4.0)
                            
                            try:
                                with Image.open(fp) as img:
                                    img_w, img_h = img.size
                                ratio = min(box_w / img_w, box_h / img_h)
                                fit_w = img_w * ratio
                                fit_h = img_h * ratio
                                fit_x = x_pos + (box_w - fit_w) / 2.0
                                fit_y = y_pos + (box_h - fit_h) / 2.0
                                pdf.image(fp, x=fit_x, y=fit_y, w=fit_w, h=fit_h)
                            except Exception:
                                pdf.set_xy(x_pos, y_pos + 15)
                                pdf.cell(box_w, 6, "[Format Gambar Error]", align='C')
                else:
                    pdf.set_xy(25, img_base_y + 15)
                    pdf.cell(175, 6, "[Tidak Ada Gambar Diunggah]", align='C')
                    
                y_start += row_h + 3

    # --- 5. OLA SLA ---
    if df_ola_sheet is not None:
        for bulan_item in list_bulan_logbook:
            ola_data = parse_olasla(df_ola_sheet, bulan_item)
            pdf.add_page(orientation='landscape')
            pdf.set_y(15)
            draw_olasla_page(pdf, ola_data, nama_teknisi, nip_teknisi_dinamis, tahun)

    out = pdf.output(dest='S')
    if isinstance(out, str):
        return out.encode('latin-1')
    return bytes(out)

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
            
        fotos = st.file_uploader(
            "Upload Foto Dokumentasi (Bisa Pilih Banyak Foto)", 
            type=['jpg', 'jpeg', 'png'], 
            accept_multiple_files=True
        )
        submit = st.form_submit_button("Simpan Data")
        
        if submit:
            if not kegiatan_dipilih:
                st.error("Silakan pilih minimal satu Penjelasan Kegiatan!")
            else:
                kegiatan_str = ", ".join(kegiatan_dipilih) 
                saved_paths = []
                
                if fotos:
                    for idx, foto in enumerate(fotos):
                        timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
                        fname = f"{timestamp_str}_{idx}_{foto.name}"
                        file_path = os.path.join(UPLOAD_DIR, fname)
                        with open(file_path, "wb") as f:
                            f.write(foto.getbuffer())
                        saved_paths.append(file_path)
                    
                foto_db_val = json.dumps(saved_paths) if saved_paths else ""
                
                with get_db_connection() as conn:
                    c = conn.cursor()
                    c.execute(
                        "INSERT INTO e_kinerja (tanggal, nama_teknisi, penjelasan_kegiatan, foto_path) VALUES (?, ?, ?, ?)",
                        (tanggal.strftime("%Y-%m-%d"), nama, kegiatan_str, foto_db_val)
                    )
                    conn.commit()
                st.success(f"Data kegiatan berhasil disimpan! ({len(saved_paths)} foto tersimpan)")

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
            
            with st.spinner('Menyusun Logbook, Jadwal, Lampiran Foto, OLA SLA & Memformat PDF...'):
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
                    base_path = extra_path = final_path = None
                    try:
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
                    finally:
                        # Pembersihan berkas temporer aman
                        for p in [base_path, extra_path, final_path]:
                            if p and os.path.exists(p):
                                try:
                                    os.remove(p)
                                except Exception:
                                    pass
                else:
                    output_data = pdf_bytes

            filename = f"E_Kinerja_{filter_nama.replace(' ', '_')}_{jenis_laporan}_{filter_tahun}.pdf"
            st.success("✅ Dokumen PDF E-Kinerja Lengkap berhasil dibuat!")
            st.download_button(
                label="⬇️ Download Hasil PDF E-Kinerja",
                data=output_data,
                file_name=filename,
                mime="application/pdf"
            )
