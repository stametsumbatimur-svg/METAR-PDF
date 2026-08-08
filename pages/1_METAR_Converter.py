import streamlit as st
import pandas as pd
import re
import io
import os
from datetime import datetime
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- DICTIONARY BULAN BAHASA INDONESIA (KAPITAL) ---
BULAN_INDO = {
    1: "JANUARI", 2: "FEBRUARI", 3: "MARET", 4: "APRIL",
    5: "MEI", 6: "JUNI", 7: "JULI", 8: "AGUSTUS",
    9: "SEPTEMBER", 10: "OKTOBER", 11: "NOVEMBER", 12: "DESEMBER"
}

# ==========================================
# ===== METAR & SPECI PARSER FUNCTIONS =====
# ==========================================
def parse_metar_speci(sandi_str):
    if pd.isna(sandi_str):
        return None
    sandi_str = sandi_str.replace('\n', ' ').replace('\r', '').strip()
    
    report_type = None
    if 'METAR' in sandi_str:
        report_type = 'METAR'
    elif 'SPECI' in sandi_str:
        report_type = 'SPECI'
    else:
        return None
        
    start_idx = sandi_str.find(report_type)
    core_str = sandi_str[start_idx:].replace('=', '').strip()
    tokens = core_str.split()
    
    header_title, loc, time_str, wind, vis, wx, cloud, t_dp, qnh, rmk = report_type, "NIL", "NIL", "NIL", "NIL", "NIL", "NIL", "NIL", "NIL", "NOSIG"
    
    is_cor = False
    cc_type = ""
    remaining_tokens = []
    
    for t in tokens:
        if t in ['METAR', 'SPECI']: continue
        elif t == 'COR':
            is_cor = True
            continue
        elif re.match(r'^CC[A-Z]$', t):
            cc_type = t.upper()
            continue
        elif re.match(r'^[A-Z]{4}$', t) and loc == "NIL":
            loc = t
            continue
        elif re.match(r'^\d{6}Z$', t) and time_str == "NIL":
            time_str = t
            continue
        else:
            remaining_tokens.append(t)
            
    if 'CAVOK' in remaining_tokens:
        vis = 'CAVOK'
        wx = ''      
        cloud = ''   
        for t in remaining_tokens:
            if re.match(r'^\d{5}(G\d{2})?KT$', t) or re.match(r'^VRB\d{2}KT$', t) or t == '00000KT': wind = t
            elif re.match(r'^\d{2}/\d{2}$', t) or re.match(r'^M\d{2}/\d{2}$', t): t_dp = t.replace('/', ' / ')
            elif re.match(r'^Q\d{4}$', t): qnh = t
            elif t in ['NOSIG', 'TEMPO', 'BECMG']: rmk = t
    else:
        cloud_list = []
        for t in remaining_tokens:
            if re.match(r'^\d{5}(G\d{2})?KT$', t) or re.match(r'^VRB\d{2}KT$', t) or t == '00000KT': wind = t
            elif re.match(r'^\d{4}$', t): vis = t
            elif t in ['RA', 'DZ', 'SHRA', 'TSRA', 'TS', 'BR', 'HZ', 'FG', '-RA', '+RA', 'VCTS', '+TSRA', '-TSRA']: wx = t
            elif re.match(r'^(FEW|SCT|BKN|OVC)\d{3}(CB|TCU)?$', t) or t in ['NSC', 'SKC', 'CLR']: cloud_list.append(t)
            elif re.match(r'^\d{2}/\d{2}$', t) or re.match(r'^M\d{2}/\d{2}$', t): t_dp = t.replace('/', ' / ')
            elif re.match(r'^Q\d{4}$', t): qnh = t
            elif t in ['NOSIG', 'TEMPO', 'BECMG']: rmk = t
        if cloud_list:
            cloud = " ".join(cloud_list)
            
    return [header_title, loc, time_str, wind, vis, wx, cloud, t_dp, qnh, rmk, is_cor, cc_type]

def calculate_priority(row):
    score = 0
    if row['is_cor']: score = 1
    if row['cc_type'] and len(row['cc_type']) == 3:
        char_code = ord(row['cc_type'][2]) - ord('A')
        score = max(score, 2 + char_code)
    return score

def check_is_thunderstorm(wx, cloud):
    wx_str = str(wx).upper() if pd.notna(wx) else ""
    cloud_str = str(cloud).upper() if pd.notna(cloud) else ""
    
    # Kriteria TS: Memuat TS / TSRA / VCTS pada WX atau CB pada CLOUD
    has_ts = any(code in wx_str for code in ['TS', 'TSRA', 'VCTS'])
    has_cb = 'CB' in cloud_str
    
    return has_ts or has_cb

# ==========================================
# ===== THUNDERSTORM PDF & EXCEL GENERATOR =
# ==========================================
def generate_pdf_bytes_thunderstorm(df_clean, station_name, kepala_nama, kepala_nip):
    buffer = io.BytesIO()
    # Format Landscape A4 agar 31 kolom muat dengan rapi
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    story = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TSTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=12, leading=16, alignment=1)
    meta_style = ParagraphStyle('TSMeta', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=14, alignment=1)
    text_style = ParagraphStyle('TSText', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=11)
    
    tahun = df_clean['datetime'].dt.year.iloc[0]
    
    story.append(Paragraph("DATA THUNDERSTORM", title_style))
    story.append(Paragraph(f"{station_name.upper()}", meta_style))
    story.append(Paragraph(f"TAHUN: {tahun}", meta_style))
    story.append(Spacer(1, 10))
    
    # Pengolahan Matriks 12 Bulan x 31 Hari
    # Inisialisasi status default kosong
    matrix = {m: {d: "" for d in range(1, 32)} for m in range(1, 13)}
    
    # Cari hari-hari yang memiliki data pengamatan
    df_clean['day'] = df_clean['datetime'].dt.day
    df_clean['month'] = df_clean['datetime'].dt.month
    df_clean['is_ts'] = df_clean.apply(lambda r: check_is_thunderstorm(r['WX'], r['CLOUD']), axis=1)
    
    # Set default 'O' untuk tanggal yang ada pengamatannya
    observed_dates = df_clean[['month', 'day']].drop_duplicates()
    for _, r in observed_dates.iterrows():
        matrix[r['month']][r['day']] = "O"
        
    # Set 'X' untuk tanggal yang terdapat thunderstorm
    ts_dates = df_clean[df_clean['is_ts'] == True][['month', 'day']].drop_duplicates()
    for _, r in ts_dates.iterrows():
        matrix[r['month']][r['day']] = "X"
        
    # Tabel Data PDF
    header_row = ['TANGGAL\nBULAN'] + [str(d) for d in range(1, 32)]
    table_data = [header_row]
    
    for m in range(1, 13):
        row = [BULAN_INDO[m]]
        for d in range(1, 32):
            row.append(matrix[m][d])
        table_data.append(row)
        
    col_widths = [85] + [22] * 31  # Total lebar 767 pt (Sesuai area cetak Landscape A4)
    ts_table = Table(table_data, colWidths=col_widths)
    
    ts_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('ALIGN', (0, 1), (0, -1), 'LEFT'), # Nama Bulan Rata Kiri
        ('FONTNAME', (0, 1), (0, -1), 'Helvetica-Bold'),
    ]))
    
    story.append(ts_table)
    story.append(Spacer(1, 15))
    
    # Blok Keterangan dan Tanda Tangan
    tgl_sekarang = datetime.now()
    tgl_ttd = f"WAINGAPU, {tgl_sekarang.day} {BULAN_INDO[tgl_sekarang.month]} {tgl_sekarang.year}"
    
    ket_text = """<b>KETERANGAN:</b><br/>
    <b>X</b> : Ada satu atau lebih kilat / thunderstorm dalam sandi wwW1W2.<br/>
    <b>O</b> : Tidak ada kilat / thunderstorm dalam sandi wwW1W2.<br/>
    &nbsp;&nbsp;&nbsp;&nbsp;: Tidak ada pengamatan.
    """
    
    ttd_text = f"""{tgl_ttd}<br/>
    <b>KEPALA STASIUN</b><br/>
    <b>BADAN METEOROLOGI KLIMATOLOGI DAN GEOFISIKA</b><br/>
    <b>{station_name.upper()}</b><br/><br/><br/><br/>
    <b><u>{kepala_nama}</u></b>
    """
    
    ket_p = Paragraph(ket_text, text_style)
    ttd_p = Paragraph(ttd_text, ParagraphStyle('TTD', parent=text_style, alignment=1))
    
    footer_table = Table([[ket_p, ttd_p]], colWidths=[400, 367])
    footer_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
    ]))
    
    story.append(KeepTogether(footer_table))

    doc.build(story)
    buffer.seek(0)
    return buffer

def generate_excel_bytes_thunderstorm(df_clean, station_name):
    buffer = io.BytesIO()
    tahun = df_clean['datetime'].dt.year.iloc[0]
    
    matrix = {m: {d: "" for d in range(1, 32)} for m in range(1, 13)}
    df_clean['day'] = df_clean['datetime'].dt.day
    df_clean['month'] = df_clean['datetime'].dt.month
    df_clean['is_ts'] = df_clean.apply(lambda r: check_is_thunderstorm(r['WX'], r['CLOUD']), axis=1)
    
    observed_dates = df_clean[['month', 'day']].drop_duplicates()
    for _, r in observed_dates.iterrows():
        matrix[r['month']][r['day']] = "O"
        
    ts_dates = df_clean[df_clean['is_ts'] == True][['month', 'day']].drop_duplicates()
    for _, r in ts_dates.iterrows():
        matrix[r['month']][r['day']] = "X"
        
    rows_data = []
    for m in range(1, 13):
        r_dict = {'BULAN': BULAN_INDO[m]}
        for d in range(1, 32):
            r_dict[str(d)] = matrix[m][d]
        rows_data.append(r_dict)
        
    df_ts = pd.DataFrame(rows_data)
    
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_ts.to_excel(writer, sheet_name=f'Thunderstorm {tahun}', index=False)
        worksheet = writer.sheets[f'Thunderstorm {tahun}']
        
        header_font = Font(name='Segoe UI', size=10, bold=True, color='FFFFFF')
        header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        align_center = Alignment(horizontal='center', vertical='center')
        thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
        
        for cell in worksheet[1]:
            cell.font, cell.fill, cell.alignment, cell.border = header_font, header_fill, align_center, thin_border
            
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
            for cell in row:
                cell.border, cell.font, cell.alignment = thin_border, Font(name='Segoe UI', size=10), align_center
                
    buffer.seek(0)
    return buffer

# ==========================================
# ======== ANTARMUKA WEB STREAMLIT =========
# ==========================================
st.set_page_config(page_title="BMKG Data Generator", layout="centered")

st.sidebar.title("🎛️ Navigasi Menu")
menu = st.sidebar.radio("Pilih Converter", ["METAR Converter", "SPECI Converter", "Thunderstorm Exporter"])
st.sidebar.markdown("---")
st.sidebar.info("Aplikasi ekstraksi data METAR & SPECI menjadi Laporan Rekapitulasi PDF & Excel.")

LOGO_FILE = "logo_bmkg.png"

# --- HALAMAN THUNDERSTORM EXPORTER ---
if menu == "Thunderstorm Exporter":
    st.title("⚡ Thunderstorm Data Exporter")
    st.write("Ekstraksi data kejadian Kilat / Thunderstorm tahunan dari gabungan file METAR & SPECI.")

    with st.expander("⚙️ Pengaturan Header & Tanda Tangan PDF"):
        st_nama = st.text_input("Nama Stasiun", "STASIUN METEOROLOGI UMBU MEHANG KUNDA")
        kp_nama = st.text_input("Nama Kepala Stasiun", "CARLES ALEXANDER TARI, S.TP")
        kp_nip = st.text_input("NIP Kepala Stasiun (Opsional)", "")

    uploaded_files = st.file_uploader("Upload File CSV (Bisa Upload Multiple METAR/SPECI)", type=["csv"], accept_multiple_files=True)

    if uploaded_files:
        all_rows = []
        for file in uploaded_files:
            try:
                df = pd.read_csv(file)
                if 'sandi' in df.columns and 'data_timestamp' in df.columns:
                    for idx, row in df.iterrows():
                        res = parse_metar_speci(row['sandi'])
                        if res:
                            station = row['station_name'] if 'station_name' in df.columns else st_nama
                            all_rows.append(res + [row['data_timestamp'], station])
            except Exception as e:
                st.error(f"Gagal membaca file {file.name}: {e}")

        if all_rows:
            columns = ['TYPE', 'LOC', 'TIME', 'WIND', 'VIS', 'WX', 'CLOUD', 'T/DP', 'QNH', 'RMK', 'is_cor', 'cc_type', 'raw_timestamp', 'station_name']
            df_clean = pd.DataFrame(all_rows, columns=columns)
            df_clean['raw_timestamp'] = df_clean['raw_timestamp'].str.replace(" +0000 UTC", "", regex=False)
            df_clean['datetime'] = pd.to_datetime(df_clean['raw_timestamp'])
            df_clean = df_clean.sort_values(by='datetime').reset_index(drop=True)

            pdf_data = generate_pdf_bytes_thunderstorm(df_clean, st_nama, kp_nama, kp_nip)
            excel_data = generate_excel_bytes_thunderstorm(df_clean, st_nama)

            tahun = df_clean['datetime'].dt.year.iloc[0]
            st.success(f"Berhasil memproses data Thunderstorm Tahun {tahun}!")

            st.write("---")
            col_pdf, col_xlsx = st.columns(2)
            with col_pdf:
                st.download_button(label="📥 Download PDF Thunderstorm", data=pdf_data, file_name=f"THUNDERSTORM_{tahun}.pdf", mime="application/pdf")
            with col_xlsx:
                st.download_button(label="📊 Download Excel Thunderstorm", data=excel_data, file_name=f"THUNDERSTORM_{tahun}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
