import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# --- FUNGSI PARSING & GENERATE PDF ---
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
            if re.match(r'^\d{5}(G\d{2})?KT$', t) or t == 'VRB\d{2}KT' or t == '00000KT':
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

def generate_pdf_bytes(df_clean):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=letter,
        rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36
    )
    story = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('HeaderTitle', parent=styles['Heading2'], fontSize=11, leading=13)
    table_text_style = ParagraphStyle('TableText', parent=styles['Normal'], fontSize=8.5, leading=10, alignment=1)
    
    nama_stasiun = df_clean['station_name'].iloc[0].upper() if 'station_name' in df_clean.columns else "STASIUN METEOROLOGI"
    grouped = df_clean.groupby('date_group')
    
    for count, (date, group) in enumerate(grouped):
        if count > 0:
            story.append(PageBreak())
            
        tanggal_format = date.strftime('%d %B %Y')
        
        story.append(Paragraph("<b>BMKG</b>", title_style))
        story.append(Paragraph("<b>BALAI BESAR METEOROLOGI KLIMATOLOGI DAN GEOFISIKA WILAYAH III</b>", title_style))
        story.append(Paragraph(f"<b>{nama_stasiun}</b>", title_style))
        story.append(Spacer(1, 15))
        story.append(Paragraph(f"<b>REKAP DATA METAR: {tanggal_format}</b>", title_style))
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
st.write("Aplikasi pengubah otomatis extract data CSV METAR (per 30 menit) menjadi PDF formal bulanan/harian per jam (00-23 UTC).")

uploaded_file = st.file_uploader("Upload file CSV hasil extract sistem Anda", type=["csv"])

if uploaded_file is not None:
    try:
        df = pd.read_csv(uploaded_file)
        
        if 'sandi' not in df.columns or 'data_timestamp' not in df.columns:
            st.error("Format CSV tidak sesuai! Pastikan terdapat kolom 'sandi' dan 'data_timestamp'.")
        else:
            with st.spinner("Sedang memproses dan menyaring data..."):
                parsed_rows = []
                for idx, row in df.iterrows():
                    res = parse_metar(row['sandi'])
                    if res:
                        station = row['station_name'] if 'station_name' in df.columns else "STASIUN METEOROLOGI"
                        parsed_rows.append(res + [row['data_timestamp'], station])
                        
                columns = ['METAR', 'LOC', 'TIME', 'WIND', 'VIS', 'WX', 'CLOUD', 'T/DP', 'QNH', 'RMK', 'raw_timestamp', 'station_name']
                df_clean = pd.DataFrame(parsed_rows, columns=columns)
                
                # --- PERBAIKAN DI BAGIAN INI ---
                # Bersihkan "+0000 UTC" agar bisa dibaca oleh Pandas
                df_clean['raw_timestamp'] = df_clean['raw_timestamp'].str.replace(" +0000 UTC", "", regex=False)
                df_clean['datetime'] = pd.to_datetime(df_clean['raw_timestamp'])
                
                # Filter menit berkembar :00 saja (per jam)
                df_clean = df_clean[df_clean['datetime'].dt.minute == 0]
                df_clean['date_group'] = df_clean['datetime'].dt.date
                df_clean = df_clean.sort_values(by='datetime').reset_index(drop=True)
                
                if df_clean.empty:
                    st.warning("Tidak ditemukan data dengan menit :00 (per jam) di dalam file ini.")
                else:
                    st.success(f"Berhasil memproses data! Ditemukan {len(df_clean)} baris data per jam.")
                    
                    st.subheader("Preview Data (Hanya Jam Genap)")
                    st.dataframe(df_clean[['METAR', 'LOC', 'TIME', 'WIND', 'VIS', 'CLOUD', 'T/DP', 'QNH']].head(10), use_container_width=True)
                    
                    pdf_data = generate_pdf_bytes(df_clean)
                    
                    st.download_button(
                        label="📥 Download PDF Rekap METAR",
                        data=pdf_data,
                        file_name=f"Rekap_METAR_{df_clean['date_group'].iloc[0]}.pdf",
                        mime="application/pdf"
                    )
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {e}")
