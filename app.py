import streamlit as st
import pandas as pd
import re
import io
import os
import requests
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# --- DICTIONARY BULAN BAHASA INDONESIA (KAPITAL) ---
BULAN_INDO = {
    1: "JANUARI", 2: "FEBRUARI", 3: "MARET", 4: "APRIL",
    5: "MEI", 6: "JUNI", 7: "JULI", 8: "AGUSTUS",
    9: "SEPTEMBER", 10: "OKTOBER", 11: "NOVEMBER", 12: "DESEMBER"
}

# --- FUNGSI AMBIL LOGO (MENGEMBALIKAN RAW BYTES) ---
def get_bmkg_logo_bytes():
    """Mengambil logo dalam bentuk RAW BYTES untuk mencegah stream exhaustion penyebab segfault"""
    # 1. Cek file lokal di GitHub terlebih dahulu
    if os.path.exists("logo_bmkg.png"):
        with open("logo_bmkg.png", "rb") as f:
            return f.read()
        
    # 2. Jika tidak ada, download dari Wikipedia dengan header bypass
    url = "https://upload.wikimedia.org/wikipedia/commons/1/12/Logo_BMKG_%282010%29.png"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    try:
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            return response.content
    except Exception:
        pass
    return None

# --- FUNGSI PARSING SANDI METAR ---
def parse_metar(sandi_str):
    if pd.isna(sandi_str):
        return None
    sandi_str = sandi_str.replace('\n', ' ').replace('\r', '').strip()
    match_main = re.search(r'(METAR\s+\w+\s+\d+Z.*)', sandi_str)
    if not match_main:
        return None
    
    metar_core = match_main.group(1).replace('=', '').strip()
    tokens = metar_core.split()
    
    metar, loc, time_str, wind, vis, wx, cloud, t_dp, qnh, rmk = "METAR", "NIL", "NIL", "NIL", "NIL", "NIL", "NIL", "NIL", "NIL", "NOSIG"
    
    if len(tokens) >= 3:
        metar, loc, time_str = tokens[0], tokens[1], tokens[2]
        remaining = tokens[3:]
        cloud_list = []
        for t in remaining:
            # PERBAIKAN: Menggunakan regex raw string secara konsisten untuk menghilangkan SyntaxWarning
            if re.match(r'^\d{5}(G\d{2})?KT$', t) or re.match(r'^VRB\d{2}KT$', t) or t == '00000KT':
                wind = t
            elif re.match(r'^\d{4}$', t) or t == 'CAVOK':
                vis = t
            elif t in ['RA', 'DZ', 'SHRA', 'TSRA', 'TS', 'BR', 'HZ', 'FG', '-RA', '+RA']:
                wx = t
            elif re.match(r'^(FEW|SCT|BKN|OVC)\d{3}$', t) or t == 'NSC' or t == 'SKC':
                cloud_list.append(t)
            elif re.match(r'^\d{2}/\d{2}$', t) or re.match(r'^M\d{2}/\d{2}$', t):
                t_dp = t
            elif re.match(r'^Q\d{4}$', t):
                qnh = t
            elif t in ['NOSIG', 'TEMPO', 'BECMG']:
                rmk = t
        if cloud_list:
            cloud = " ".join(cloud_list)
            
    return [metar, loc, time_str, wind, vis, wx, cloud, t_dp, qnh, rmk]

# --- FUNGSI GENERATE PDF ---
def generate_pdf_bytes(df_clean, logo_bytes):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=letter,
        rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36
    )
    story = []
    
    styles = getSampleStyleSheet()
    
    header_text_style = ParagraphStyle(
        'HeaderCenterText',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=14,
        alignment=1
    )
    
    rekap_style = ParagraphStyle(
        'RekapStyle', 
        parent=styles['Heading2'], 
        fontName='Helvetica-Bold',
        fontSize=11, 
        leading=14, 
        alignment=1
    )
    
    table_text_style = ParagraphStyle(
        'TableText', 
        parent=styles['Normal'], 
        fontSize=8.5, 
        leading=10, 
        alignment=1
    )
    
    nama_stasiun = df_clean['station_name'].iloc[0].upper() if 'station_name' in df_clean.columns else "STASIUN METEOROLOGI"
    grouped = df_clean.groupby('date_group')
    
    for count, (date, group) in enumerate(grouped):
        if count > 0:
            story.append(PageBreak())
            
        nama_bulan = BULAN_INDO[date.month]
        tanggal_format = f"{date.day:02d} {nama_bulan} {date.year}"
        
        text_block = [
            Paragraph("<b>BADAN METEOROLOGI KLIMATOLOGI DAN GEOFISIKA</b>", header_text_style),
            Paragraph(f"<b>{nama_stasiun}</b>", header_text_style),
        ]
        
        if logo_bytes:
            # PERBAIKAN UTAMA: Membuat instance BytesIO baru dari raw bytes di SETIAP Halaman
            # Ini mencegah kegagalan pembacaan pointer internal objek gambar pada multi-halaman
            fresh_logo_stream = io.BytesIO(logo_bytes)
            logo_img = Image(fresh_logo_stream, width=45, height=45)
            header_table = Table([[logo_img, text_block, ""]], colWidths=[50, 440, 50])
            header_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN', (0,0), (0,0), 'LEFT'),
                ('ALIGN', (1,0), (1,0), 'CENTER'),
                ('BOTTOMPADDING', (0,0), (-1,-1), 0),
                ('TOPPADDING', (0,0), (-1,-1), 0),
            ]))
        else:
            header_table = Table([[text_block]], colWidths=[540])
            header_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ]))
            
        story.append(header_table)
        story.append(Spacer(1, 15))
        
        judul_rekap = f"REKAP DATA METAR: {tanggal_format}".upper()
        story.append(Paragraph(f"<b>{judul_rekap}</b>", rekap_style))
        story.append(Spacer(1, 10))
        
        headers = ['METAR', 'LOC', 'TIME', 'WIND', 'VIS', 'WX', 'CLOUD', 'T/DP', 'QNH', 'RMK']
        table_data = [[Paragraph(f"<b>{h}</b>", table_text_style) for h in headers]]
        
        for _, row in group.iterrows():
            row_data = [Paragraph(str(row[h]), table_text_style) for h in headers]
            table_data.append(row_data)
            
        col_widths = [45, 40, 55, 65, 40, 35, 60, 45, 45, 50]
        metar_table = Table(table_data, colWidths=col_widths)
        metar_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
        ]))
        story.append(metar_table)

    doc.build(story)
    buffer.seek(0)
    return buffer

# --- ANTARMUKA WEB STREAMLIT ---
st.set_page_config(page_title="METAR PDF Generator", layout="centered")

st.title("✈️ METAR to PDF Converter")
st.write("Aplikasi pengubah otomatis extract data CSV METAR menjadi PDF formal per jam (00-23 UTC) dengan layout Kop Surat Resmi.")

# Mengambil raw bytes dari logo
logo_bytes = get_bmkg_logo_bytes()

uploaded_file = st.file_uploader("Upload file CSV hasil extract sistem Anda", type=["csv"])

if uploaded_file is not None:
    try:
        df = pd.read_csv(uploaded_file)
        
        if 'sandi' not in df.columns or 'data_timestamp' not in df.columns:
            st.error("Format CSV tidak sesuai! Pastikan terdapat kolom 'sandi' dan 'data_timestamp'.")
        else:
            with st.spinner("Sedang memproses data..."):
                parsed_rows = []
                for idx, row in df.iterrows():
                    res = parse_metar(row['sandi'])
                    if res:
                        station = row['station_name'] if 'station_name' in df.columns else "STASIUN METEOROLOGI"
                        parsed_rows.append(res + [row['data_timestamp'], station])
                        
                columns = ['METAR', 'LOC', 'TIME', 'WIND', 'VIS', 'WX', 'CLOUD', 'T/DP', 'QNH', 'RMK', 'raw_timestamp', 'station_name']
                df_clean = pd.DataFrame(parsed_rows, columns=columns)
                
                df_clean['raw_timestamp'] = df_clean['raw_timestamp'].str.replace(" +0000 UTC", "", regex=False)
                df_clean['datetime'] = pd.to_datetime(df_clean['raw_timestamp'])
                
                df_clean = df_clean[df_clean['datetime'].dt.minute == 0]
                df_clean['date_group'] = df_clean['datetime'].dt.date
                df_clean = df_clean.sort_values(by='datetime').reset_index(drop=True)
                
                if df_clean.empty:
                    st.warning("Tidak ditemukan data dengan menit :00 (per jam) di dalam file ini.")
                else:
                    st.success(f"Berhasil! Data per jam siap diunduh.")
                    
                    st.subheader("Preview Data (Hanya Jam Genap)")
                    st.dataframe(df_clean[['METAR', 'LOC', 'TIME', 'WIND', 'VIS', 'CLOUD', 'T/DP', 'QNH']].head(10), use_container_width=True)
                    
                    pdf_data = generate_pdf_bytes(df_clean, logo_bytes)
                    
                    first_date = df_clean['date_group'].iloc[0]
                    nama_file_pdf = f"REKAP_METAR_{first_date.day:02d}_{BULAN_INDO[first_date.month]}_{first_date.year}.pdf"
                    
                    st.download_button(
                        label="📥 Download PDF Rekap METAR Resmi",
                        data=pdf_data,
                        file_name=nama_file_pdf,
                        mime="application/pdf"
                    )
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {e}")
