import streamlit as st
import pandas as pd
import re
import io
import os
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
            if re.match(r'^\d{5}(G\d{2})?KT$', t) or re.match(r'^VRB\d{2}KT$', t) or t == '00000KT':
                wind = t
            elif re.match(r'^\d{4}$', t) or t == 'CAVOK':
                vis = t
            elif t in ['RA', 'DZ', 'SHRA', 'TSRA', 'TS', 'BR', 'HZ', 'FG', '-RA', '+RA']:
                wx = t
            elif re.match(r'^(FEW|SCT|BKN|OVC)\d{3}$', t) or re.match(r'^NSC$', t) or re.match(r'^SKC$', t):
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
def generate_pdf_bytes(df_clean, logo_path):
    buffer = io.BytesIO()
    
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=letter,
        rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25
    )
    story = []
    
    styles = getSampleStyleSheet()
    
    header_text_style = ParagraphStyle(
        'HeaderCenterText',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=15,
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
        
        if logo_path and os.path.exists(logo_path):
            logo_img = Image(logo_path, width=48, height=48)
            header_table = Table([[logo_img, text_block, ""]], colWidths=[55, 452, 55])
            header_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN', (0,0), (0,0), 'CENTER'),
                ('ALIGN', (1,0), (1,0), 'CENTER'),
                ('BOTTOMPADDING', (0,0), (-1,-1), 6),
                ('TOPPADDING', (0,0), (-1,-1), 0),
                ('LINEBELOW', (0,0), (-1,-1), 1.5, colors.black),
            ]))
        else:
            header_table = Table([[text_block]], colWidths=[562])
            header_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('BOTTOMPADDING', (0,0), (-1,-1), 6),
                ('LINEBELOW', (0,0), (-1,-1), 1.5, colors.black),
            ]))
            
        story.append(header_table)
        story.append(Spacer(1, 10))
        
        judul_rekap = f"REKAP DATA METAR: {tanggal_format}".upper()
        story.append(Paragraph(f"<b>{judul_rekap}</b>", rekap_style))
        story.append(Spacer(1, 8))
        
        headers = ['METAR', 'LOC', 'TIME', 'WIND', 'VIS', 'WX', 'CLOUD', 'T/DP', 'QNH', 'RMK']
        
        # KUNCI UTAMA: Kita masukkan RAW STRING (teks biasa), bukan objek Paragraph()!
        table_data = [headers]
        
        for _, row in group.iterrows():
            row_data = [str(row[h]) for h in headers]
            table_data.append(row_data)
            
        col_widths = [45, 40, 55, 75, 42, 40, 105, 45, 50, 65]
        
        metar_table = Table(table_data, colWidths=col_widths)
        metar_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2.5),
            ('TOPPADDING', (0, 0), (-1, -1), 2.5),
            # KUNCI KEDUA: Mengatur jenis dan ukuran font langsung lewat TableStyle secara global
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
        ]))
        story.append(metar_table)

    doc.build(story)
    buffer.seek(0)
    return buffer

# --- ANTARMUKA WEB STREAMLIT ---
st.set_page_config(page_title="METAR PDF Generator", layout="centered")

st.title("✈️ METAR to PDF Converter")
st.write("Aplikasi pengubah otomatis extract data CSV METAR menjadi PDF formal per jam (00-23 UTC) dengan layout Kop Surat Resmi.")

LOGO_FILE = "logo_bmkg.png"

if not os.path.exists(LOGO_FILE):
    st.warning(f"⚠️ File gambar '{LOGO_FILE}' tidak terdeteksi di folder utama. Harap pastikan file logo sudah di-upload ke GitHub.")

uploaded_file = st.file_uploader("Upload file CSV hasil extract sistem Anda", type=["csv"])

if uploaded_file is not None:
    try:
        df = pd.read_csv(uploaded_file)
        
        if 'sandi' not in df.columns or 'data_timestamp' not in df.columns:
            st.error("Format CSV tidak sesuai! Pastikan terdapat kolom 'sandi' dan 'data_timestamp'.")
        else:
            with st.spinner("Sedang memproses seluruh data METAR skala besar... Mohon tunggu sebentar."):
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
                    st.success(f"Berhasil memproses {len(df_clean)} baris data per jam!")
                    
                    st.subheader("Preview Data (Hanya Jam Genap)")
                    st.dataframe(df_clean[['METAR', 'LOC', 'TIME', 'WIND', 'VIS', 'CLOUD', 'T/DP', 'QNH']].head(10), width='stretch')
                    
                    pdf_data = generate_pdf_bytes(df_clean, LOGO_FILE)
                    
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
