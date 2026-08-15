import streamlit as st
import pandas as pd
import re
import io
import os
import calendar
import zipfile
import openpyxl
from datetime import datetime
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell

# ==========================================
# ===== KONFIGURASI HALAMAN UTAMA ==========
# ==========================================
st.set_page_config(page_title="BMKG Data Generator", layout="centered", page_icon="🌤️")

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
    has_ts = any(code in wx_str for code in ['TS', 'TSRA', 'VCTS'])
    has_cb = 'CB' in cloud_str
    return has_ts or has_cb

def get_cloud_category(cloud_str, vis_str):
    if vis_str == 'CAVOK' or pd.isna(cloud_str) or str(cloud_str).strip() in ['', 'NIL', 'SKC', 'CLR', 'NSC']:
        return 'CERAH'
    
    c_str = str(cloud_str).upper()
    if 'OVC' in c_str:
        return 'BERAWAN_BANYAK'
    elif 'BKN' in c_str:
        return 'BERAWAN'
    elif 'SCT' in c_str or 'FEW' in c_str:
        return 'BERAWAN_SEBAGIAN'
    else:
        return 'CERAH'

def generate_pdf_bytes_metar(df_clean, logo_path):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=20, leftMargin=75, topMargin=25, bottomMargin=25)
    story = []
    
    styles = getSampleStyleSheet()
    header_text_style = ParagraphStyle('HeaderCenterText', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=11, leading=15, alignment=1)
    
    nama_stasiun = df_clean['station_name'].iloc[0].upper() if 'station_name' in df_clean.columns else "STASIUN METEOROLOGI"
    grouped = df_clean.groupby('date_group')
    
    for count, (date, group) in enumerate(grouped):
        if count > 0: story.append(PageBreak())
            
        nama_bulan = BULAN_INDO[date.month]
        tanggal_at = f"{date.day:02d} {nama_bulan} {date.year}"
        judul_rekap = f"REKAP DATA METAR: {tanggal_at}".upper()
        text_block = [
            Paragraph("<b>BALAI BESAR METEOROLOGI KLIMATOLOGI DAN GEOFISIKA WILAYAH III</b>", header_text_style),
            Paragraph(f"<b>{nama_stasiun}</b>", header_text_style),
            Paragraph("<b>JL. ADI SUCIPTO NO. 3</b>", header_text_style),
            Paragraph(f"<b>{judul_rekap}</b>", header_text_style)
        ]
        
        if logo_path and os.path.exists(logo_path):
            logo_img = Image(logo_path, width=48, height=48)
            header_table = Table([[logo_img, text_block, ""]], colWidths=[50, 430, 20])
            header_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('ALIGN', (0,0), (0,0), 'CENTER'), ('ALIGN', (1,0), (1,0), 'CENTER'),
                ('BOTTOMPADDING', (0,0), (-1,-1), 6), ('TOPPADDING', (0,0), (-1,-1), 0), ('LINEBELOW', (0,0), (-1,-1), 1.5, colors.black),
                ('LEFTPADDING', (1,0), (1,0), 0), ('RIGHTPADDING', (1,0), (1,0), 0)
            ]))
        else:
            header_table = Table([[text_block]], colWidths=[500])
            header_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('BOTTOMPADDING', (0,0), (-1,-1), 6), ('LINEBELOW', (0,0), (-1,-1), 1.5, colors.black),
            ]))
            
        story.append(header_table)
        story.append(Spacer(1, 10))
        
        headers = ['METAR', 'LOC', 'TIME', 'WIND', 'VIS', 'WX', 'CLOUD', 'T/DP', 'QNH', 'RMK']
        table_data = [headers]
        
        base_table_style = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey), ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black), ('BOTTOMPADDING', (0, 0), (-1, -1), 2.5), ('TOPPADDING', (0, 0), (-1, -1), 2.5),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), ('FONTSIZE', (0, 0), (-1, 0), 8), ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'), ('FONTSIZE', (0, 1), (-1, -1), 8),
        ]
        
        for idx, row in group.iterrows():
            current_row_idx = len(table_data)
            if row['VIS'] == 'CAVOK':
                row_data = [str(row['TYPE']), str(row['LOC']), str(row['TIME']), str(row['WIND']), 'CAVOK', '', '', str(row['T/DP']), str(row['QNH']), str(row['RMK'])]
                base_table_style.append(('SPAN', (4, current_row_idx), (6, current_row_idx)))
            else:
                row_data = [str(row['TYPE']), str(row['LOC']), str(row['TIME']), str(row['WIND']), str(row['VIS']), str(row['WX']), str(row['CLOUD']), str(row['T/DP']), str(row['QNH']), str(row['RMK'])]
            table_data.append(row_data)
            
        col_widths = [40, 35, 50, 65, 35, 35, 90, 40, 45, 55] 
        metar_table = Table(table_data, colWidths=col_widths)
        metar_table.setStyle(TableStyle(base_table_style))
        story.append(metar_table)

    doc.build(story)
    buffer.seek(0)
    return buffer

def generate_pdf_bytes_speci(df_clean, logo_path):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=20, leftMargin=75, topMargin=25, bottomMargin=25)
    story = []
    
    styles = getSampleStyleSheet()
    header_text_style = ParagraphStyle('HeaderCenterText', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=11, leading=15, alignment=1)
    sub_header_style = ParagraphStyle('SubHeader', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=14, spaceAfter=5, spaceBefore=5)
    
    nama_stasiun = df_clean['station_name'].iloc[0].upper() if 'station_name' in df_clean.columns else "STASIUN METEOROLOGI"
    
    df_clean['year'] = df_clean['datetime'].dt.year
    df_clean['month'] = df_clean['datetime'].dt.month
    grouped_month = df_clean.groupby(['year', 'month'])
    
    for count_month, ((year, month), month_group) in enumerate(grouped_month):
        if count_month > 0: 
            story.append(PageBreak())
            
        nama_bulan = BULAN_INDO[month]
        judul_rekap = f"REKAP DATA SPECI BULAN: {nama_bulan} {year}".upper()
        text_block = [
            Paragraph("<b>BALAI BESAR METEOROLOGI KLIMATOLOGI DAN GEOFISIKA WILAYAH III</b>", header_text_style),
            Paragraph(f"<b>{nama_stasiun}</b>", header_text_style),
            Paragraph("<b>JL. ADI SUCIPTO NO. 3</b>", header_text_style),
            Paragraph(f"<b>{judul_rekap}</b>", header_text_style)
        ]
        
        if logo_path and os.path.exists(logo_path):
            logo_img = Image(logo_path, width=48, height=48)
            header_table = Table([[logo_img, text_block, ""]], colWidths=[50, 430, 20])
            header_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('ALIGN', (0,0), (0,0), 'CENTER'), ('ALIGN', (1,0), (1,0), 'CENTER'),
                ('BOTTOMPADDING', (0,0), (-1,-1), 6), ('TOPPADDING', (0,0), (-1,-1), 0), ('LINEBELOW', (0,0), (-1,-1), 1.5, colors.black),
                ('LEFTPADDING', (1,0), (1,0), 0), ('RIGHTPADDING', (1,0), (1,0), 0)
            ]))
        else:
            header_table = Table([[text_block]], colWidths=[500])
            header_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('BOTTOMPADDING', (0,0), (-1,-1), 6), ('LINEBELOW', (0,0), (-1,-1), 1.5, colors.black),
            ]))
            
        story.append(header_table)
        story.append(Spacer(1, 10))
        
        grouped_date = month_group.groupby('date_group')
        for date, date_group in grouped_date:
            tanggal_format = f"{date.day:02d} {nama_bulan} {year}"
            
            group_story = []
            group_story.append(Paragraph(f"<b>Tanggal: {tanggal_format}</b>", sub_header_style))
            
            headers = ['SPECI', 'LOC', 'TIME', 'WIND', 'VIS', 'WX', 'CLOUD', 'T/DP', 'QNH', 'RMK']
            table_data = [headers]
            
            base_table_style = [
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey), ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black), ('BOTTOMPADDING', (0, 0), (-1, -1), 2.5), ('TOPPADDING', (0, 0), (-1, -1), 2.5),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), ('FONTSIZE', (0, 0), (-1, 0), 8), ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'), ('FONTSIZE', (0, 1), (-1, -1), 8),
            ]
            
            for idx, row in date_group.iterrows():
                current_row_idx = len(table_data)
                if row['VIS'] == 'CAVOK':
                    row_data = [str(row['TYPE']), str(row['LOC']), str(row['TIME']), str(row['WIND']), 'CAVOK', '', '', str(row['T/DP']), str(row['QNH']), str(row['RMK'])]
                    base_table_style.append(('SPAN', (4, current_row_idx), (6, current_row_idx)))
                else:
                    row_data = [str(row['TYPE']), str(row['LOC']), str(row['TIME']), str(row['WIND']), str(row['VIS']), str(row['WX']), str(row['CLOUD']), str(row['T/DP']), str(row['QNH']), str(row['RMK'])]
                table_data.append(row_data)
                
            col_widths = [40, 35, 50, 65, 35, 35, 90, 40, 45, 55] 
            speci_table = Table(table_data, colWidths=col_widths)
            speci_table.setStyle(TableStyle(base_table_style))
            
            group_story.append(speci_table)
            group_story.append(Spacer(1, 15))
            
            story.append(KeepTogether(group_story))

    doc.build(story)
    buffer.seek(0)
    return buffer

def generate_excel_bytes_metar_speci(df_clean, report_type="METAR"):
    buffer = io.BytesIO()
    headers = ['TYPE', 'LOC', 'TIME', 'WIND', 'VIS', 'WX', 'CLOUD', 'T/DP', 'QNH', 'RMK', 'datetime']
    df_export = df_clean[headers].copy()
    df_export.rename(columns={'TYPE': report_type}, inplace=True)
    df_export['datetime'] = df_export['datetime'].dt.strftime('%Y-%m-%d %H:%M:%S')
    
    sheet_title = f"Rekap {report_type}"
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_export.to_excel(writer, sheet_name=sheet_title, index=False)
        worksheet = writer.sheets[sheet_title]
        header_font = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
        header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        align_center, align_left = Alignment(horizontal='center', vertical='center'), Alignment(horizontal='left', vertical='center')
        thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
        
        for cell in worksheet[1]:
            cell.font, cell.fill, cell.alignment = header_font, header_fill, align_center
            
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
            for cell in row:
                cell.border, cell.font = thin_border, Font(name='Segoe UI', size=10)
                col_header = worksheet.cell(row=1, column=cell.column).value
                cell.alignment = align_center if col_header in [report_type, 'LOC', 'TIME', 'WIND', 'VIS', 'WX', 'T/DP', 'QNH'] else align_left
        
        for col in worksheet.columns:
            max_len = max([len(str(cell.value)) for cell in col if cell.value] + [0])
            worksheet.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 3, 11)
            
    buffer.seek(0)
    return buffer

def parse_wxrev(sandi_str):
    if pd.isna(sandi_str):
        return None
    sandi_str = sandi_str.replace('\n', ' ').replace('\r', '').replace('=', '').strip()
    
    start_idx = sandi_str.find('WXREV')
    if start_idx == -1:
        return None
        
    wxrev_core = sandi_str[start_idx:].strip()
    tokens = wxrev_core.split()
    
    if len(tokens) < 7: 
        return None
        
    mmyygp = tokens[1] if len(tokens) > 1 else ""
    tgl = mmyygp[2:4] if len(mmyygp) >= 4 else ""
    iiiii = tokens[2] if len(tokens) > 2 else ""
    att = tokens[3] if len(tokens) > 3 else ""
    app = tokens[4] if len(tokens) > 4 else ""
    auu = tokens[5] if len(tokens) > 5 else ""
    arr = tokens[6] if len(tokens) > 6 else ""
    rdrd1 = tokens[7] if len(tokens) > 7 else ""
    rdrd2 = tokens[8] if len(tokens) > 8 else ""
    
    return [tgl, mmyygp, iiiii, att, app, auu, arr, rdrd1, rdrd2]

def generate_pdf_bytes_wxrev(df_clean, logo_path):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=20, leftMargin=75, topMargin=20, bottomMargin=20)
    story = []
    
    styles = getSampleStyleSheet()
    header_title_style = ParagraphStyle('HeaderTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=13, alignment=1)
    header_sub_style = ParagraphStyle('HeaderSub', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=11, alignment=1)
    center_title_style = ParagraphStyle('CenterTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=11, leading=15, alignment=1)
    
    station_name = df_clean['station_name'].iloc[0].upper() if 'station_name' in df_clean.columns else "STASIUN METEOROLOGI"
    
    text_block = [
        Paragraph("<b>BADAN METEOROLOGI KLIMATOLOGI DAN GEOFISIKA</b>", header_title_style),
        Paragraph(f"<b>{station_name}</b>", header_title_style),
        Paragraph("Alamat : JL.ADI SUCIPTO NO.3 | Telp. (0387)61227 | Email : stamet.waingapu@gmail.com", header_sub_style),
    ]
    
    total_width = 500  
    if logo_path and os.path.exists(logo_path):
        logo_img = Image(logo_path, width=48, height=48)
        header_table = Table([[logo_img, text_block, ""]], colWidths=[50, 430, 20])
        header_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('ALIGN', (0,0), (0,0), 'CENTER'), ('ALIGN', (1,0), (1,0), 'CENTER'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4), ('TOPPADDING', (0,0), (-1,-1), 0), ('LINEBELOW', (0,0), (-1,-1), 1.2, colors.black),
            ('LEFTPADDING', (1,0), (1,0), 0), ('RIGHTPADDING', (1,0), (1,0), 0)
        ]))
    else:
        header_table = Table([[text_block]], colWidths=[total_width])
        header_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4), ('LINEBELOW', (0,0), (-1,-1), 1.2, colors.black),
        ]))
        
    story.append(header_table)
    story.append(Spacer(1, 8))
    
    dt_first = df_clean['datetime'].iloc[0] if not df_clean.empty else datetime.now()
    bulan_str = BULAN_INDO.get(dt_first.month, "").upper()
    tahun_str = dt_first.year
    
    story.append(Paragraph("<b>KUMPULAN BERITA WXREV</b>", center_title_style))
    story.append(Paragraph(f"<b>{bulan_str} {tahun_str}</b>", center_title_style))
    story.append(Spacer(1, 8))
    
    headers = ['TGL', 'MMYYGp', 'IIiii', 'atTxTxTnTn', 'apPxPxPnPn', 'auUxUxUnUn', 'arRRRR', 'rDrDdfmfm', 'rDrDdfmfm']
    table_data = [headers]
    
    for _, row in df_clean.iterrows():
        table_data.append([
            str(row['TGL']), str(row['MMYYGp']), str(row['IIiii']),
            str(row['atTxTxTnTn']), str(row['apPxPxPnPn']), str(row['auUxUxUnUn']),
            str(row['arRRRR']), str(row['rDrDdfmfm_1']), str(row['rDrDdfmfm_2'])
        ])
        
    base_table_style = [
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 8), ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('TOPPADDING', (0, 0), (-1, -1), 2), ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
    ]
    
    col_widths = [30, 50, 45, 66, 66, 66, 55, 61, 61]
    wx_table = Table(table_data, colWidths=col_widths)
    wx_table.setStyle(TableStyle(base_table_style))
    story.append(wx_table)
    
    doc.build(story)
    buffer.seek(0)
    return buffer

def generate_excel_bytes_wxrev(df_clean):
    buffer = io.BytesIO()
    headers_excel = ['TGL', 'MMYYGp', 'IIiii', 'atTxTxTnTn', 'apPxPxPnPn', 'auUxUxUnUn', 'arRRRR', 'rDrDdfmfm_1', 'rDrDdfmfm_2']
    df_export = df_clean[headers_excel].copy()
    
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_export.to_excel(writer, sheet_name='WXREV', index=False, header=['TGL', 'MMYYGp', 'IIiii', 'atTxTxTnTn', 'apPxPxPnPn', 'auUxUxUnUn', 'arRRRR', 'rDrDdfmfm', 'rDrDdfmfm'])
        worksheet = writer.sheets['WXREV']
        
        header_font = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
        header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        align_center = Alignment(horizontal='center', vertical='center')
        thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
        
        for cell in worksheet[1]:
            cell.font, cell.fill, cell.alignment, cell.border = header_font, header_fill, align_center, thin_border
            
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
            for cell in row:
                cell.border, cell.font, cell.alignment = thin_border, Font(name='Segoe UI', size=10), align_center
        
        for col in worksheet.columns:
            max_len = max([len(str(cell.value)) for cell in col if cell.value] + [0])
            worksheet.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 3, 11)
            
    buffer.seek(0)
    return buffer

def generate_pdf_bytes_thunderstorm(df_clean, station_name, kepala_nama, kepala_nip):
    buffer = io.BytesIO()
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
        
    header_row = ['TANGGAL\nBULAN'] + [str(d) for d in range(1, 32)]
    table_data = [header_row]
    
    for m in range(1, 13):
        row = [BULAN_INDO[m]]
        for d in range(1, 32):
            row.append(matrix[m][d])
        table_data.append(row)
        
    col_widths = [85] + [22] * 31
    ts_table = Table(table_data, colWidths=col_widths)
    
    ts_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey), ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3), ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('ALIGN', (0, 1), (0, -1), 'LEFT'), ('FONTNAME', (0, 1), (0, -1), 'Helvetica-Bold'),
    ]))
    
    story.append(ts_table)
    story.append(Spacer(1, 15))
    
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
        ('VALIGN', (0, 0), (-1, -1), 'TOP'), ('LEFTPADDING', (0, 0), (-1, -1), 0), ('RIGHTPADDING', (0, 0), (-1, -1), 0),
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

def generate_pdf_bytes_perawanan(df_clean, station_name, kepala_nama, kepala_nip, form_code="Klim / PWN/WPU-2001"):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=25, bottomMargin=25)
    story = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('PWNTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=13, alignment=1)
    text_style = ParagraphStyle('PWNText', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=11)
    
    df_clean['year'] = df_clean['datetime'].dt.year
    df_clean['month'] = df_clean['datetime'].dt.month
    df_clean['day'] = df_clean['datetime'].dt.day
    df_clean['cat'] = df_clean.apply(lambda r: get_cloud_category(r['CLOUD'], r['VIS']), axis=1)
    
    grouped_month = df_clean.groupby(['year', 'month'])
    
    for count_month, ((year, month), month_group) in enumerate(grouped_month):
        if count_month > 0:
            story.append(PageBreak())
            
        bulan_str = BULAN_INDO[month]
        num_days = calendar.monthrange(year, month)[1]
        
        story.append(Paragraph("DAFTAR IKHTISAR FREKUENSI PERAWANAN", title_style))
        story.append(Paragraph(f"{station_name.upper()}", title_style))
        story.append(Paragraph(f"BULAN {bulan_str} {year}", title_style))
        story.append(Spacer(1, 10))
        
        header_row = ["TANGGAL", "CERAH\n(N=0)", "BERAWAN SEBAGIAN\n(N=1-3)", "BERAWAN\n(N=4-6)", "BERAWAN BANYAK\n(N=7-8)", "JUMLAH"]
        table_data = [header_row]
        
        tot_cerah = tot_sebagian = tot_berawan = tot_banyak = tot_semua = 0
        
        for d in range(1, num_days + 1):
            df_day = month_group[month_group['day'] == d]
            counts = df_day['cat'].value_counts()
            
            c_cerah = counts.get('CERAH', 0)
            c_sebagian = counts.get('BERAWAN_SEBAGIAN', 0)
            c_berawan = counts.get('BERAWAN', 0)
            c_banyak = counts.get('BERAWAN_BANYAK', 0)
            c_total = len(df_day)
            
            tot_cerah += c_cerah
            tot_sebagian += c_sebagian
            tot_berawan += c_berawan
            tot_banyak += c_banyak
            tot_semua += c_total
            
            table_data.append([
                f"{d:02d}", str(c_cerah) if c_cerah > 0 else "-", str(c_sebagian) if c_sebagian > 0 else "-",
                str(c_berawan) if c_berawan > 0 else "-", str(c_banyak) if c_banyak > 0 else "-", str(c_total) if c_total > 0 else "-"
            ])
            
        table_data.append(["JUMLAH", str(tot_cerah), str(tot_sebagian), str(tot_berawan), str(tot_banyak), str(tot_semua)])
        
        pct_c = f"{round((tot_cerah/tot_semua)*100)}%" if tot_semua > 0 else "0%"
        pct_s = f"{round((tot_sebagian/tot_semua)*100)}%" if tot_semua > 0 else "0%"
        pct_bw = f"{round((tot_berawan/tot_semua)*100)}%" if tot_semua > 0 else "0%"
        pct_by = f"{round((tot_banyak/tot_semua)*100)}%" if tot_semua > 0 else "0%"
        
        table_data.append(["% - TAGE", pct_c, pct_s, pct_bw, pct_by, "100%" if tot_semua > 0 else "0%"])
        
        col_widths = [60, 90, 120, 100, 110, 60]
        pwn_table = Table(table_data, colWidths=col_widths)
        pwn_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey), ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), ('FONTSIZE', (0, 0), (-1, -1), 7.5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2), ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('FONTNAME', (0, -2), (-1, -1), 'Helvetica-Bold'), ('BACKGROUND', (0, -2), (-1, -1), colors.whitesmoke),
        ]))
        
        story.append(pwn_table)
        story.append(Spacer(1, 15))
        
        tgl_sekarang = datetime.now()
        tgl_ttd = f"WAINGAPU, {tgl_sekarang.day:02d} {BULAN_INDO[tgl_sekarang.month]} {tgl_sekarang.year}"
        
        form_text = f"FORM: {form_code}"
        nip_str = f"<br/>NIP. {kepala_nip}" if kepala_nip else ""
        ttd_text = f"""{tgl_ttd}<br/>
        <b>KEPALA STASIUN METEOROLOGI</b><br/>
        <b>{station_name.upper()}</b><br/><br/><br/><br/>
        <b><u>{kepala_nama}</u></b>{nip_str}
        """
        
        form_p = Paragraph(form_text, text_style)
        ttd_p = Paragraph(ttd_text, ParagraphStyle('TTDPWN', parent=text_style, alignment=1))
        
        footer_table = Table([[form_p, ttd_p]], colWidths=[200, 340])
        footer_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'), ('LEFTPADDING', (0, 0), (-1, -1), 0), ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ]))
        
        story.append(KeepTogether(footer_table))

    doc.build(story)
    buffer.seek(0)
    return buffer

def generate_excel_bytes_perawanan(df_clean, station_name):
    buffer = io.BytesIO()
    df_clean['year'] = df_clean['datetime'].dt.year
    df_clean['month'] = df_clean['datetime'].dt.month
    df_clean['day'] = df_clean['datetime'].dt.day
    df_clean['cat'] = df_clean.apply(lambda r: get_cloud_category(r['CLOUD'], r['VIS']), axis=1)
    
    grouped_month = df_clean.groupby(['year', 'month'])
    
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for (year, month), month_group in grouped_month:
            bulan_str = BULAN_INDO[month]
            num_days = calendar.monthrange(year, month)[1]
            rows_data = []
            tot_cerah = tot_sebagian = tot_berawan = tot_banyak = tot_semua = 0
            
            for d in range(1, num_days + 1):
                df_day = month_group[month_group['day'] == d]
                counts = df_day['cat'].value_counts()
                
                c_cerah = counts.get('CERAH', 0)
                c_sebagian = counts.get('BERAWAN_SEBAGIAN', 0)
                c_berawan = counts.get('BERAWAN', 0)
                c_banyak = counts.get('BERAWAN_BANYAK', 0)
                c_total = len(df_day)
                
                tot_cerah += c_cerah; tot_sebagian += c_sebagian
                tot_berawan += c_berawan; tot_banyak += c_banyak; tot_semua += c_total
                
                rows_data.append({
                    'TANGGAL': f"{d:02d}",
                    'CERAH (N=0)': c_cerah if c_cerah > 0 else "-",
                    'BERAWAN SEBAGIAN (N=1-3)': c_sebagian if c_sebagian > 0 else "-",
                    'BERAWAN (N=4-6)': c_berawan if c_berawan > 0 else "-",
                    'BERAWAN BANYAK (N=7-8)': c_banyak if c_banyak > 0 else "-",
                    'JUMLAH': c_total if c_total > 0 else "-"
                })
                
            rows_data.append({
                'TANGGAL': "JUMLAH", 'CERAH (N=0)': tot_cerah, 'BERAWAN SEBAGIAN (N=1-3)': tot_sebagian,
                'BERAWAN (N=4-6)': tot_berawan, 'BERAWAN BANYAK (N=7-8)': tot_banyak, 'JUMLAH': tot_semua
            })
            
            pct_c = f"{round((tot_cerah/tot_semua)*100)}%" if tot_semua > 0 else "0%"
            pct_s = f"{round((tot_sebagian/tot_semua)*100)}%" if tot_semua > 0 else "0%"
            pct_bw = f"{round((tot_berawan/tot_semua)*100)}%" if tot_semua > 0 else "0%"
            pct_by = f"{round((tot_banyak/tot_semua)*100)}%" if tot_semua > 0 else "0%"
            
            rows_data.append({
                'TANGGAL': "% - TAGE", 'CERAH (N=0)': pct_c, 'BERAWAN SEBAGIAN (N=1-3)': pct_s,
                'BERAWAN (N=4-6)': pct_bw, 'BERAWAN BANYAK (N=7-8)': pct_by, 'JUMLAH': "100%" if tot_semua > 0 else "0%"
            })
            
            df_pwn = pd.DataFrame(rows_data)
            sheet_name = f"PWN {bulan_str[:3]} {year}"[:31]
            
            df_pwn.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            
            header_font = Font(name='Segoe UI', size=10, bold=True, color='FFFFFF')
            header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
            align_center = Alignment(horizontal='center', vertical='center')
            thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
            
            for cell in worksheet[1]:
                cell.font, cell.fill, cell.alignment, cell.border = header_font, header_fill, align_center, thin_border
                
            for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
                for cell in row:
                    cell.border, cell.font, cell.alignment = thin_border, Font(name='Segoe UI', size=10), align_center
                    
            for col in worksheet.columns:
                max_len = max([len(str(cell.value)) for cell in col if cell.value] + [0])
                worksheet.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 3, 12)
            
    buffer.seek(0)
    return buffer

# ==========================================
# ==== PIBAL / RADIOSONDE FUNCTIONS ========
# ==========================================
PENGAMAT_MAP = {
    "Mitra Agritami": "Mitra", "M. Fauzan Fachrurozi": "Fauzan", "Zulqha Ariandi Al Zikri": "Zulqha",    
    "Hakim": "Hakim", "Andreas Yoga": "Andreas", "Imam Zacky Anwar": "Zacky",               
    "Mohammad Hasyim Hanif": "Hanif", "Ni Luh Ayu Agnes": "Ayu", "Yenny Thenu": "Yenny", "Adi": "Adi"
}

def resolve_alasan_and_pengamat(raw_alasan, raw_pengamat, is_norep=False):
    alasan_res = str(raw_alasan).strip() if pd.notna(raw_alasan) and str(raw_alasan).strip().lower() not in ['nan', 'none'] else ""
    raw_p = str(raw_pengamat).strip() if pd.notna(raw_pengamat) and str(raw_pengamat).strip().lower() not in ['nan', 'none'] else ""
    pengamat_res = ""
    if raw_p:
        name_upper = raw_p.upper()
        matched = False
        for full_name, short_name in PENGAMAT_MAP.items():
            if full_name.upper() in name_upper:
                pengamat_res = short_name
                matched = True
                break
        if not matched:
            pengamat_res = raw_p.title()
    if is_norep and not alasan_res:
        alasan_res = "NOREP"
    return alasan_res, pengamat_res

def detect_columns(columns):
    col_alasan = col_pengamat = None
    for c in columns:
        clow = c.lower().strip()
        if any(kw in clow for kw in ['pengamat', 'observer', 'petugas', 'nama', 'user']):
            if not any(kw in clow for kw in ['alasan', 'keterangan', 'penghentian', 'stasiun']):
                col_pengamat = c
                break
    for c in columns:
        clow = c.lower().strip()
        if any(kw in clow for kw in ['alasan', 'reason', 'keterangan', 'ket', 'penghentian', 'henti']):
            if c != col_pengamat:
                col_alasan = c
                break
    return col_alasan, col_pengamat

def extract_valid_value(series):
    valid = series.dropna().astype(str).str.strip()
    valid = valid[~valid.str.lower().isin(['', 'nan', 'none', 'null', '-', 'nat'])]
    if len(valid) > 0:
        return valid.iloc[0]
    return None

def parse_sandi_pibal_complete(sandi_text, surface_dir, surface_speed):
    wind_data = {}
    logs = [] 
    if pd.isna(sandi_text):
        return wind_data, ["Teks sandi kosong (NaN)."]
    text = str(sandi_text).replace('=', ' ')
    if 'NIL' in text:
        return wind_data, []
    wind_data['PERMUKAAN'] = (surface_dir, surface_speed)

    ppaa_data = {}
    if 'PPAA' in text:
        try:
            ppaa_part = text.split('PPAA', 1)[1]
            for end_mark in ['PPBB', 'PPCC', 'PPDD']:
                if end_mark in ppaa_part: ppaa_part = ppaa_part.split(end_mark)[0]
            tokens_a = ppaa_part.split()
            idx = 0
            while idx < len(tokens_a):
                try:
                    token = tokens_a[idx]
                    if token.startswith('55') and len(token) == 5:
                        if idx + 1 < len(tokens_a) and tokens_a[idx+1].startswith('97'):
                            idx += 2; continue
                        indicators = token[2:]
                        level_map = []
                        for ind in indicators:
                            if ind in ['2', '3'] and '5000 feet' not in [l[0] for l in level_map]: level_map.append(('5000 feet', ind))
                            elif ind == '8': level_map.append(('10000 feet', ind))
                            elif ind == '5': level_map.append(('18000 feet', ind))
                            elif ind == '4': level_map.append(('25000 feet', ind))
                            elif ind == '3': level_map.append(('30000 feet', ind))
                            elif ind == '1': level_map.append(('10000 feet' if '5000 feet' in [l[0] for l in level_map] else '18000 feet', ind))
                            elif ind == '6': level_map.append(('45000 feet', ind))
                            elif ind == '7': level_map.append(('50000 feet', ind))
                        for i, (lvl_name, ind) in enumerate(level_map):
                            try:
                                if idx + 1 + i < len(tokens_a):
                                    w_str = tokens_a[idx + 1 + i]
                                    if len(w_str) == 5 and w_str != '/////' and w_str != '77999':
                                        ppaa_data[lvl_name] = (int(w_str[:3]), int(w_str[3:]))
                                    elif w_str != '/////' and w_str != '77999': logs.append(f"Format angin PPAA tidak valid ({w_str}) di level {lvl_name}.")
                            except Exception as e: logs.append(f"Gagal ekstrak angin PPAA level {lvl_name}: {str(e)}")
                        idx += len(level_map)
                except Exception as e: logs.append(f"Gagal memproses token PPAA '{tokens_a[idx]}': {str(e)}")
                idx += 1
        except Exception as e: logs.append(f"Gagal membaca blok PPAA secara keseluruhan: {str(e)}")

    ppbb_data = {}
    if 'PPBB' in text:
        try:
            ppbb_part = text.split('PPBB', 1)[1]
            for end_mark in ['PPAA', 'PPCC', 'PPDD']:
                if end_mark in ppbb_part: ppbb_part = ppbb_part.split(end_mark)[0]
            tokens = ppbb_part.split()
            idx = 0
            while idx < len(tokens):
                try:
                    token = tokens[idx]
                    is_station_id = (idx < 2 and len(token) == 5 and token.startswith('97'))
                    if len(token) == 5 and token.startswith('9') and not is_station_id:
                        t_n = int(token[1]) if token[1].isdigit() else 0
                        levels = [int(ch) if ch.isdigit() else (0 if ch == '/' else None) for ch in token[2:5]]
                        valid_levels = [u for u in levels if u is not None]
                        for i, u in enumerate(levels):
                            try:
                                if u is not None and (idx + 1 + i) < len(tokens):
                                    wind_str = tokens[idx + 1 + i]
                                    if len(wind_str) == 5 and wind_str != '/////':
                                        lvl = (t_n * 10000) + (u * 1000)
                                        if lvl > 0: ppbb_data[f'{lvl} feet'] = (int(wind_str[:3]), int(wind_str[3:]))
                            except Exception as e: logs.append(f"Gagal ekstrak angin PPBB di sekitar token {token}: {str(e)}")
                        idx += len(valid_levels)
                except Exception as e: pass
                idx += 1
        except Exception as e: logs.append(f"Gagal membaca blok PPBB secara keseluruhan: {str(e)}")

    final_data = dict(ppaa_data)
    final_data.update(ppbb_data)
    final_data['PERMUKAAN'] = (surface_dir, surface_speed)
    return final_data, logs

def get_max_height_from_parsed(parsed_data):
    max_feet = 0
    for key in parsed_data.keys():
        if 'feet' in key:
            try:
                val = int(key.split(' ')[0])
                if val > max_feet: max_feet = val
            except ValueError: pass
    return max_feet if max_feet > 0 else 0

def safe_set_cell(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    if not isinstance(cell, MergedCell): cell.value = value  

def process_template_uw1(wb, df_records, nama_bulan, sample_year, sample_month_num):
    ws_uw1 = wb['UW I'] if 'UW I' in wb.sheetnames else None
    ws_uw2 = wb['UW II'] if 'UW II' in wb.sheetnames else None
        
    col_mapping_uw1 = {
        '1000 feet': (12, 13, 14, 15), '2000 feet': (16, 17, 18, 19), 
        '3000 feet': (20, 21, 22, 23), '5000 feet': (24, 25, 26, 27), 
        '7000 feet': (28, 29, 30, 31), '10000 feet': (32, 33, 34, 35), 
        '12000 feet': (36, 37, 38, 39), '15000 feet': (40, 41, 42, 43)
    }
    col_mapping_uw2 = {
        '18000 feet': (12, 13, 14, 15), '20000 feet': (16, 17, 18, 19), 
        '25000 feet': (20, 21, 22, 23), '30000 feet': (24, 25, 26, 27), 
        '35000 feet': (28, 29, 30, 31), '40000 feet': (32, 33, 34, 35),
        '45000 feet': (36, 37, 38, 39), '50000 feet': (40, 41, 42, 43)
    }

    def get_row_uw(d, h):
        if h == 0: return 9 + (d - 1) * 3
        elif h == 6: return 9 + (d - 1) * 3 + 1
        elif h == 12: return 9 + (d - 1) * 3 + 2
        return None

    def split_digits(d_val, f_val):
        d_str = str(int(d_val)).zfill(3)[:2]
        f_str = str(int(f_val)).zfill(2)[-2:]
        return str(d_str[0]), str(d_str[1]), str(f_str[0]), str(f_str[1])
        
    year_2d = str(sample_year % 100).zfill(2)
    month_str = str(sample_month_num)

    for ws in [ws_uw1, ws_uw2]:
        if ws is None: continue
        safe_set_cell(ws, 4, 6, f": {nama_bulan}")
        safe_set_cell(ws, 5, 6, f": {sample_year}")
        row_limit = 9 if ws.title == 'UW I' else 8 
        for c in range(5, 10): safe_set_cell(ws, row_limit, c, None)
        safe_set_cell(ws, 9, 1, "9"); safe_set_cell(ws, 9, 2, "7"); safe_set_cell(ws, 9, 3, "3")
        safe_set_cell(ws, 9, 4, "4"); safe_set_cell(ws, 9, 5, "0"); safe_set_cell(ws, 9, 6, year_2d[0])
        safe_set_cell(ws, 9, 7, year_2d[1]); safe_set_cell(ws, 9, 8, month_str)
        safe_set_cell(ws, 9, 9, "0"); safe_set_cell(ws, 9, 10, "1")
            
    for _, row in df_records.iterrows():
        day = row['day']; hour = row['hour_z']
        target_row = get_row_uw(day, hour)
        if target_row is None: continue
        wind_dict = row['parsed_wind']
        hour_str = str(hour).zfill(2)
        
        if ws_uw1:
            safe_set_cell(ws_uw1, target_row, 11, hour_str) 
            for layer, cols in col_mapping_uw1.items():
                if layer in wind_dict:
                    d1, d2, f1, f2 = split_digits(*wind_dict[layer])
                    for i, val in enumerate([d1, d2, f1, f2]): safe_set_cell(ws_uw1, target_row, cols[i], val)
        if ws_uw2:
            safe_set_cell(ws_uw2, target_row, 11, hour_str) 
            for layer, cols in col_mapping_uw2.items():
                if layer in wind_dict:
                    d1, d2, f1, f2 = split_digits(*wind_dict[layer])
                    for i, val in enumerate([d1, d2, f1, f2]): safe_set_cell(ws_uw2, target_row, cols[i], val)
    return wb

def process_template_uw2(wb, df_records, nama_bulan, sample_year):
    for sheet_name in ['UPPER WIND', 'KETERANGAN']:
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            safe_set_cell(ws, 10, 6, f": {nama_bulan} {sample_year}")
            for r in range(16, 47):
                for c in range(3, 12): safe_set_cell(ws, r, c, None)
            for _, row in df_records.iterrows():
                day = row['day']; hour = row['hour_z']
                target_row = 15 + day
                col_pibal = 3 if hour == 0 else (6 if hour == 6 else (9 if hour == 12 else None))
                if col_pibal is not None:
                    max_height = row['max_height']
                    is_norep = (max_height == 0)
                    safe_set_cell(ws, target_row, col_pibal, "NOREP" if is_norep else f"{max_height:,}".replace(',', '.'))
                    alasan_final, pengamat_final = resolve_alasan_and_pengamat(row.get('alasan_data'), row.get('pengamat_data'), is_norep=is_norep)
                    safe_set_cell(ws, target_row, col_pibal + 1, alasan_final if alasan_final else "")
                    safe_set_cell(ws, target_row, col_pibal + 2, pengamat_final if pengamat_final else "-")
    return wb

def process_template_komponen(wb, df_records, nama_bulan, sample_year):
    START_ROW = 9
    sheet_mapping = {
        0:  {'P-5': '00 P-5', '7-20': '00 7-20', '25-40': '00 25-40', '45-50': '00 45-50'},
        6:  {'P-5': '06 P-5', '7-20': '06 7-20', '25-40': '06 25-40', '45-50': '06 45-50'},
        12: {'P-5': '12 P-5', '7-20': '12 7-20', '25-40': '12 25-40', '45-50': '12 45-50'},
    }
    col_mapping = {
        'PERMUKAAN': (2, 3), '1000 feet': (6, 7), '3000 feet': (10, 11), '5000 feet': (14, 15),
        '7000 feet': (2, 3), '10000 feet': (6, 7), '15000 feet': (10, 11), '20000 feet': (14, 15),
        '25000 feet': (2, 3), '30000 feet': (6, 7), '35000 feet': (10, 11), '40000 feet': (14, 15),
        '45000 feet': (2, 3), '50000 feet': (6, 7),
    }

    for ws_name in wb.sheetnames:
        ws = wb[ws_name]
        safe_set_cell(ws, 3, 3, f": {nama_bulan} {sample_year}")
        for r in range(START_ROW, START_ROW + 31):
            for c in [2, 3, 6, 7, 10, 11, 14, 15]:
                if c <= ws.max_column: safe_set_cell(ws, r, c, None)

    layer_groups = [
        ('P-5', ['PERMUKAAN', '1000 feet', '3000 feet', '5000 feet']),
        ('7-20', ['7000 feet', '10000 feet', '15000 feet', '20000 feet']),
        ('25-40', ['25000 feet', '30000 feet', '35000 feet', '40000 feet']),
        ('45-50', ['45000 feet', '50000 feet'])
    ]

    for _, row in df_records.iterrows():
        hour, day = row['hour_z'], row['day']
        if hour not in sheet_mapping: continue
        target_row = START_ROW + (day - 1)
        wind_dict = row['parsed_wind']
        for group_key, group_layers in layer_groups:
            sheet_title = sheet_mapping[hour].get(group_key)
            if sheet_title and sheet_title in wb.sheetnames:
                ws_target = wb[sheet_title]
                for layer in group_layers:
                    if layer in wind_dict:
                        d, f = wind_dict[layer]
                        safe_set_cell(ws_target, target_row, col_mapping[layer][0], int(d))
                        safe_set_cell(ws_target, target_row, col_mapping[layer][1], int(f))
    return wb

# ==========================================
# ======== ANTARMUKA WEB STREAMLIT =========
# ==========================================
# --- SIDEBAR MENU ---
st.sidebar.title("🎛️ Navigasi Menu")
menu = st.sidebar.radio(
    "Pilih Converter", 
    [
        "METAR Converter", 
        "SPECI Converter", 
        "WXREV Converter", 
        "Thunderstorm Exporter", 
        "Form Perawanan Exporter",
        "Form A/B"
    ]
)
st.sidebar.markdown("---")
st.sidebar.info("Aplikasi ekstraksi data Sandi Cuaca (METAR, SPECI, WXREV, Thunderstorm, Perawanan, & Pibal) menjadi Laporan PDF & Excel otomatis.")

LOGO_FILE = "logo_bmkg.png"

if not os.path.exists(LOGO_FILE):
    st.sidebar.warning(f"⚠️ File gambar '{LOGO_FILE}' tidak terdeteksi di folder utama.")

# --- HALAMAN METAR ---
if menu == "METAR Converter":
    st.title("✈️ METAR to PDF & Excel Converter")
    with st.expander("ℹ️ Klik di sini untuk melihat Petunjuk Penggunaan"):
        st.markdown("""
        **Syarat File CSV:**
        - File hasil extract dari "https://bmkgsatu.bmkg.go.id/extractgts" data METAR dengan status SENT.
        """)
    uploaded_file = st.file_uploader("Upload file CSV METAR", type=["csv"])
    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file)
            if 'sandi' not in df.columns or 'data_timestamp' not in df.columns:
                st.error("Format CSV tidak sesuai! Pastikan terdapat kolom 'sandi' dan 'data_timestamp'.")
            else:
                with st.spinner("Sedang memvalidasi data METAR..."):
                    parsed_rows = []
                    for idx, row in df.iterrows():
                        res = parse_metar_speci(row['sandi'])
                        if res and res[0] == 'METAR':
                            station = row['station_name'] if 'station_name' in df.columns else "STASIUN METEOROLOGI"
                            msg_id = row['id'] if 'id' in df.columns else idx
                            parsed_rows.append(res + [row['data_timestamp'], station, msg_id])
                            
                    columns = ['TYPE', 'LOC', 'TIME', 'WIND', 'VIS', 'WX', 'CLOUD', 'T/DP', 'QNH', 'RMK', 'is_cor', 'cc_type', 'raw_timestamp', 'station_name', 'msg_id']
                    df_clean = pd.DataFrame(parsed_rows, columns=columns)
                    
                    if df_clean.empty: st.warning("Tidak ditemukan data METAR di dalam file ini.")
                    else:
                        df_clean['raw_timestamp'] = df_clean['raw_timestamp'].str.replace(" +0000 UTC", "", regex=False)
                        df_clean['datetime'] = pd.to_datetime(df_clean['raw_timestamp'])
                        df_clean = df_clean[df_clean['datetime'].dt.minute == 0]
                        df_clean['priority_score'] = df_clean.apply(calculate_priority, axis=1)
                        df_clean = df_clean.sort_values(by=['datetime', 'priority_score', 'msg_id'])
                        df_clean = df_clean.drop_duplicates(subset=['datetime'], keep='last')
                        df_clean['date_group'] = df_clean['datetime'].dt.date
                        df_clean = df_clean.sort_values(by='datetime').reset_index(drop=True)
                        
                        if df_clean.empty: st.warning("Tidak ditemukan data METAR dengan menit :00 (per jam) di dalam file ini.")
                        else:
                            st.success(f"Berhasil memproses data METAR!")
                            pdf_data = generate_pdf_bytes_metar(df_clean, LOGO_FILE)
                            excel_data = generate_excel_bytes_metar_speci(df_clean, report_type="METAR")
                            first_date = df_clean['date_group'].iloc[0]
                            nama_file_base = f"REKAP_METAR_{first_date.day:02d}_{BULAN_INDO[first_date.month]}_{first_date.year}"
                            
                            st.write("---")
                            st.subheader("Unduh Laporan METAR")
                            col_pdf, col_xlsx = st.columns(2)
                            with col_pdf: st.download_button(label="📥 Download PDF", data=pdf_data, file_name=f"{nama_file_base}.pdf", mime="application/pdf")
                            with col_xlsx: st.download_button(label="📊 Download Excel", data=excel_data, file_name=f"{nama_file_base}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e: st.error(f"Terjadi kesalahan: {e}")

# --- HALAMAN SPECI ---
elif menu == "SPECI Converter":
    st.title("🛩️ SPECI to PDF & Excel Converter")
    with st.expander("ℹ️ Klik di sini untuk melihat Petunjuk Penggunaan"):
        st.markdown("""
        **Syarat File CSV:**
        - File hasil extract dari "https://bmkgsatu.bmkg.go.id/extractgts" data SPECI dengan status SENT.
        """)
    uploaded_file = st.file_uploader("Upload file CSV SPECI", type=["csv"])
    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file)
            if 'sandi' not in df.columns or 'data_timestamp' not in df.columns:
                st.error("Format CSV tidak sesuai! Pastikan terdapat kolom 'sandi' dan 'data_timestamp'.")
            else:
                with st.spinner("Sedang memvalidasi data SPECI..."):
                    parsed_rows = []
                    for idx, row in df.iterrows():
                        res = parse_metar_speci(row['sandi'])
                        if res and res[0] == 'SPECI':
                            station = row['station_name'] if 'station_name' in df.columns else "STASIUN METEOROLOGI"
                            msg_id = row['id'] if 'id' in df.columns else idx
                            parsed_rows.append(res + [row['data_timestamp'], station, msg_id])
                            
                    columns = ['TYPE', 'LOC', 'TIME', 'WIND', 'VIS', 'WX', 'CLOUD', 'T/DP', 'QNH', 'RMK', 'is_cor', 'cc_type', 'raw_timestamp', 'station_name', 'msg_id']
                    df_clean = pd.DataFrame(parsed_rows, columns=columns)
                    
                    if df_clean.empty: st.warning("Tidak ditemukan data SPECI di dalam file ini.")
                    else:
                        df_clean['raw_timestamp'] = df_clean['raw_timestamp'].str.replace(" +0000 UTC", "", regex=False)
                        df_clean['datetime'] = pd.to_datetime(df_clean['raw_timestamp'])
                        df_clean['priority_score'] = df_clean.apply(calculate_priority, axis=1)
                        df_clean = df_clean.sort_values(by=['datetime', 'priority_score', 'msg_id'])
                        df_clean = df_clean.drop_duplicates(subset=['datetime'], keep='last')
                        df_clean['date_group'] = df_clean['datetime'].dt.date
                        df_clean = df_clean.sort_values(by='datetime').reset_index(drop=True)
                        
                        st.success(f"Berhasil memproses {len(df_clean)} data SPECI!")
                        pdf_data = generate_pdf_bytes_speci(df_clean, LOGO_FILE)
                        excel_data = generate_excel_bytes_metar_speci(df_clean, report_type="SPECI")
                        first_date = df_clean['datetime'].iloc[0]
                        nama_file_base = f"REKAP_SPECI_BULAN_{BULAN_INDO[first_date.month]}_{first_date.year}"
                        
                        st.write("---")
                        st.subheader("Unduh Laporan SPECI")
                        col_pdf, col_xlsx = st.columns(2)
                        with col_pdf: st.download_button(label="📥 Download PDF", data=pdf_data, file_name=f"{nama_file_base}.pdf", mime="application/pdf")
                        with col_xlsx: st.download_button(label="📊 Download Excel", data=excel_data, file_name=f"{nama_file_base}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e: st.error(f"Terjadi kesalahan: {e}")

# --- HALAMAN WXREV ---
elif menu == "WXREV Converter":
    st.title("🌧️ WXREV to PDF & Excel Converter")
    with st.expander("ℹ️ Klik di sini untuk melihat Petunjuk Penggunaan"):
        st.markdown("""
        **Syarat File CSV:**
        - File hasil extract dari "https://bmkgsatu.bmkg.go.id/extractgts" data WXREV.
        """)
    uploaded_file = st.file_uploader("Upload file CSV WXREV", type=["csv"])
    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file)
            if 'sandi' not in df.columns or 'data_timestamp' not in df.columns:
                st.error("Format CSV tidak sesuai! Pastikan terdapat kolom 'sandi' dan 'data_timestamp'.")
            else:
                with st.spinner("Sedang memvalidasi data WXREV..."):
                    parsed_rows = []
                    for idx, row in df.iterrows():
                        res = parse_wxrev(row['sandi'])
                        if res:
                            station = row['station_name'] if 'station_name' in df.columns else "Stasiun Meteorologi"
                            parsed_rows.append(res + [row['data_timestamp'], station])
                            
                    if not parsed_rows: st.warning("Tidak ditemukan data sandi WXREV di dalam file ini.")
                    else:
                        columns = ['TGL', 'MMYYGp', 'IIiii', 'atTxTxTnTn', 'apPxPxPnPn', 'auUxUxUnUn', 'arRRRR', 'rDrDdfmfm_1', 'rDrDdfmfm_2', 'raw_timestamp', 'station_name']
                        df_wxrev = pd.DataFrame(parsed_rows, columns=columns)
                        df_wxrev['raw_timestamp'] = df_wxrev['raw_timestamp'].str.replace(" +0000 UTC", "", regex=False)
                        df_wxrev['datetime'] = pd.to_datetime(df_wxrev['raw_timestamp'])
                        df_wxrev = df_wxrev.sort_values(by='datetime')
                        df_wxrev = df_wxrev.drop_duplicates(subset=['TGL'], keep='last')
                        df_wxrev = df_wxrev.sort_values(by='TGL').reset_index(drop=True)
                        
                        st.success(f"Berhasil memproses {len(df_wxrev)} data WXREV harian!")
                        pdf_data = generate_pdf_bytes_wxrev(df_wxrev, LOGO_FILE)
                        excel_data = generate_excel_bytes_wxrev(df_wxrev)
                        dt_first = df_wxrev['datetime'].iloc[0]
                        nama_file_base = f"REKAP_WXREV_{BULAN_INDO.get(dt_first.month, '').upper()}_{dt_first.year}"
                        
                        st.write("---")
                        st.subheader("Unduh Laporan WXREV")
                        col_pdf, col_xlsx = st.columns(2)
                        with col_pdf: st.download_button(label="📥 Download PDF", data=pdf_data, file_name=f"{nama_file_base}.pdf", mime="application/pdf")
                        with col_xlsx: st.download_button(label="📊 Download Excel", data=excel_data, file_name=f"{nama_file_base}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e: st.error(f"Terjadi kesalahan: {e}")

# --- HALAMAN THUNDERSTORM EXPORTER ---
elif menu == "Thunderstorm Exporter":
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
            except Exception as e: st.error(f"Gagal membaca file {file.name}: {e}")
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
            with col_pdf: st.download_button(label="📥 Download PDF Thunderstorm", data=pdf_data, file_name=f"THUNDERSTORM_{tahun}.pdf", mime="application/pdf")
            with col_xlsx: st.download_button(label="📊 Download Excel Thunderstorm", data=excel_data, file_name=f"THUNDERSTORM_{tahun}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- HALAMAN FORM PERAWANAN EXPORTER ---
elif menu == "Form Perawanan Exporter":
    st.title("☁️ Form Perawanan Exporter")
    st.write("Ekstraksi data Ikhtisar Frekuensi Perawanan bulanan dari file METAR per jam.")
    with st.expander("⚙️ Pengaturan Header, Form & Tanda Tangan PDF"):
        st_nama = st.text_input("Nama Stasiun", "STASIUN METEOROLOGI UMBU MEHANG KUNDA", key="pwn_st")
        form_code = st.text_input("Kode Form (Pojok Kiri Bawah)", "Klim / PWN/WPU-2001")
        kp_nama = st.text_input("Nama Kepala Stasiun", "CARLES ALEXANDER TARI, S.TP", key="pwn_kp")
        kp_nip = st.text_input("NIP Kepala Stasiun", "197712082001121001", key="pwn_nip")
    uploaded_files = st.file_uploader("Upload File CSV METAR", type=["csv"], accept_multiple_files=True, key="pwn_uploader")
    if uploaded_files:
        all_rows = []
        for file in uploaded_files:
            try:
                df = pd.read_csv(file)
                if 'sandi' in df.columns and 'data_timestamp' in df.columns:
                    for idx, row in df.iterrows():
                        res = parse_metar_speci(row['sandi'])
                        if res and res[0] == 'METAR':
                            station = row['station_name'] if 'station_name' in df.columns else st_nama
                            msg_id = row['id'] if 'id' in df.columns else idx
                            all_rows.append(res + [row['data_timestamp'], station, msg_id])
            except Exception as e: st.error(f"Gagal membaca file {file.name}: {e}")
        if all_rows:
            columns = ['TYPE', 'LOC', 'TIME', 'WIND', 'VIS', 'WX', 'CLOUD', 'T/DP', 'QNH', 'RMK', 'is_cor', 'cc_type', 'raw_timestamp', 'station_name', 'msg_id']
            df_clean = pd.DataFrame(all_rows, columns=columns)
            df_clean['raw_timestamp'] = df_clean['raw_timestamp'].str.replace(" +0000 UTC", "", regex=False)
            df_clean['datetime'] = pd.to_datetime(df_clean['raw_timestamp'])
            df_clean = df_clean[df_clean['datetime'].dt.minute == 0]
            df_clean['priority_score'] = df_clean.apply(calculate_priority, axis=1)
            df_clean = df_clean.sort_values(by=['datetime', 'priority_score', 'msg_id'])
            df_clean = df_clean.drop_duplicates(subset=['datetime'], keep='last')
            df_clean = df_clean.sort_values(by='datetime').reset_index(drop=True)
            if df_clean.empty: st.warning("Tidak ditemukan data METAR dengan menit :00 di dalam file ini.")
            else:
                pdf_data = generate_pdf_bytes_perawanan(df_clean, st_nama, kp_nama, kp_nip, form_code)
                excel_data = generate_excel_bytes_perawanan(df_clean, st_nama)
                months_count = len(df_clean.groupby([df_clean['datetime'].dt.year, df_clean['datetime'].dt.month]))
                dt_min = df_clean['datetime'].min(); dt_max = df_clean['datetime'].max()
                st.success(f"Berhasil memproses Data Perawanan untuk {months_count} bulan ({dt_min.strftime('%b %Y')} - {dt_max.strftime('%b %Y')})!")
                st.write("---")
                col_pdf, col_xlsx = st.columns(2)
                with col_pdf: st.download_button(label="📥 Download PDF Perawanan", data=pdf_data, file_name=f"PERAWANAN_{dt_min.strftime('%Y%m')}_{dt_max.strftime('%Y%m')}.pdf", mime="application/pdf")
                with col_xlsx: st.download_button(label="📊 Download Excel Perawanan", data=excel_data, file_name=f"PERAWANAN_{dt_min.strftime('%Y%m')}_{dt_max.strftime('%Y%m')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- HALAMAN PIBAL / RADIOSONDE ---
elif menu == "Form A/B":
    st.title("🎈 Otomatisasi Rekap Udara Atas (Pibal / Radiosonde)")
    st.subheader("Stasiun Meteorologi BMKG")
    st.markdown("---")

    DEFAULT_TEMPLATE_UW1 = "template uw 1.xlsx"
    DEFAULT_TEMPLATE_UW2 = "template uw 2.xlsx"
    DEFAULT_TEMPLATE_KOMPONEN = "template komponen angin.xlsx"

    uploaded_pibal_files = st.file_uploader(
        "Upload File CSV BMKG (Bisa Upload 2 File Sekaligus: Raw Pibal & Data Pibal)", 
        type=['csv'], 
        accept_multiple_files=True
    )

    if uploaded_pibal_files:
        use_uw1 = DEFAULT_TEMPLATE_UW1 if os.path.exists(DEFAULT_TEMPLATE_UW1) else None
        use_uw2 = DEFAULT_TEMPLATE_UW2 if os.path.exists(DEFAULT_TEMPLATE_UW2) else None
        use_komp = DEFAULT_TEMPLATE_KOMPONEN if os.path.exists(DEFAULT_TEMPLATE_KOMPONEN) else None

        if not use_uw1 and not use_uw2 and not use_komp:
            st.error("❌ File Template Excel lokal tidak ditemukan di direktori aplikasi! Pastikan file template UW1, UW2, dan Komponen Angin ada di folder utama bersama skrip.")
        else:
            try:
                df_raw = None; df_data = None
                for file in uploaded_pibal_files:
                    df_temp = pd.read_csv(file)
                    if 'data_timestamp' not in df_temp.columns:
                        file.seek(0)
                        df_temp = pd.read_csv(file, skiprows=1)
                    cols = [c.lower() for c in df_temp.columns]
                    if 'lapisan' in cols or 'wd' in cols: df_data = df_temp
                    elif 'pembacaan' in cols or 'observer_name' in cols or 'sandi_pibal' in cols: df_raw = df_temp
                    else:
                        if df_raw is None: df_raw = df_temp

                if df_raw is None and df_data is not None: df_raw = df_data

                col_alasan, col_pengamat = detect_columns(df_raw.columns)
                timestamps = df_raw['data_timestamp'].unique()
                records = []

                for ts in timestamps:
                    sub_raw = df_raw[df_raw['data_timestamp'] == ts]
                    raw_alasan = extract_valid_value(sub_raw[col_alasan]) if col_alasan and col_alasan in sub_raw.columns else None
                    raw_pengamat = extract_valid_value(sub_raw[col_pengamat]) if col_pengamat and col_pengamat in sub_raw.columns else None
                    s_dir = sub_raw['wind_dir_surface'].iloc[0] if 'wind_dir_surface' in sub_raw.columns else 0
                    s_spd = sub_raw['wind_speed_surface'].iloc[0] if 'wind_speed_surface' in sub_raw.columns else 0
                    sandi = extract_valid_value(sub_raw['sandi_pibal']) if 'sandi_pibal' in sub_raw.columns else ""

                    wind_dict = {}
                    parse_logs = []
                    if df_data is not None and 'lapisan' in df_data.columns:
                        sub_data = df_data[df_data['data_timestamp'] == ts]
                        wind_dict['PERMUKAAN'] = (s_dir, s_spd)
                        for _, d_row in sub_data.iterrows():
                            if pd.notna(d_row.get('lapisan')) and pd.notna(d_row.get('wd')) and pd.notna(d_row.get('ws')):
                                wind_dict[f"{int(d_row['lapisan'])} feet"] = (int(d_row['wd']), int(d_row['ws']))
                    
                    if not wind_dict or len(wind_dict) <= 1:
                        wind_dict, parse_logs = parse_sandi_pibal_complete(sandi, s_dir, s_spd)

                    max_h = get_max_height_from_parsed(wind_dict)
                    dt = pd.to_datetime(ts)
                    records.append({
                        'data_timestamp': ts, 'day': dt.day, 'month': dt.month, 'year': dt.year, 'hour_z': dt.hour,
                        'alasan_data': raw_alasan, 'pengamat_data': raw_pengamat, 'parsed_wind': wind_dict,
                        'max_height': max_h, 'parse_logs': parse_logs
                    })

                df_records = pd.DataFrame(records).sort_values(['month', 'day', 'hour_z']).reset_index(drop=True)
                error_records = [r for r in records if len(r.get('parse_logs', [])) > 0]
                if error_records:
                    with st.expander("⚠️ Peringatan: Terdapat Isu Parsing Sandi Pibal (Klik untuk detail)"):
                        st.warning(f"Ditemukan {len(error_records)} pengamatan dengan sandi malformed/korup. Sistem mengabaikan blok yang rusak dan mengambil data yang valid.")
                        for err in error_records:
                            st.markdown(f"**Waktu Pengamatan (UTC):** `{err['data_timestamp']}`")
                            for log_msg in err['parse_logs']: st.markdown(f" - {log_msg}")

                sample_month_num = df_records['month'].iloc[0]
                sample_year = df_records['year'].iloc[0]
                nama_bulan = BULAN_INDO[sample_month_num]

                st.success(f"✅ Data berhasil digabungkan: **{len(df_records)} pengamatan** periode **{nama_bulan} {sample_year}**.")
                if df_data is not None: st.info("💡 **Mode Dual CSV Aktif**: Data angin per lapisan diambil matang dari `Data Pibal`. Metadata diambil dari `Raw Pibal`.")
                else: st.info("ℹ️ **Mode Single CSV**: Data ditarik via Parser Sandi Pibal WMO.")

                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    if use_uw1:
                        wb_uw1 = openpyxl.load_workbook(use_uw1)
                        wb_uw1 = process_template_uw1(wb_uw1, df_records, nama_bulan, sample_year, sample_month_num)
                        excel_buffer_uw1 = io.BytesIO(); wb_uw1.save(excel_buffer_uw1); excel_buffer_uw1.seek(0)
                        zip_file.writestr(f"Laporan_UW1_{nama_bulan}_{sample_year}.xlsx", excel_buffer_uw1.getvalue())
                        st.info("📋 Template UW 1 (UW I & UW II) berhasil diproses.")
                    if use_uw2:
                        wb_uw2 = openpyxl.load_workbook(use_uw2)
                        wb_uw2 = process_template_uw2(wb_uw2, df_records, nama_bulan, sample_year)
                        excel_buffer_uw2 = io.BytesIO(); wb_uw2.save(excel_buffer_uw2); excel_buffer_uw2.seek(0)
                        zip_file.writestr(f"Laporan_UW2_{nama_bulan}_{sample_year}.xlsx", excel_buffer_uw2.getvalue())
                        st.info("📋 Template UW 2 (Form Ae.31) berhasil diproses.")
                    if use_komp:
                        wb_komp = openpyxl.load_workbook(use_komp)
                        wb_komp = process_template_komponen(wb_komp, df_records, nama_bulan, sample_year)
                        excel_buffer_komp = io.BytesIO(); wb_komp.save(excel_buffer_komp); excel_buffer_komp.seek(0)
                        zip_file.writestr(f"Rekap_Komponen_Angin_{nama_bulan}_{sample_year}.xlsx", excel_buffer_komp.getvalue())
                        st.info("📋 Template Rekap Komponen Angin (Lengkap s.d 50.000 ft) berhasil diproses.")

                st.subheader("📥 Unduh Hasil Rekapitulasi")
                zip_buffer.seek(0)
                st.download_button(
                    label="⬇️ Download Semua File (Bentuk ZIP)",
                    data=zip_buffer,
                    file_name=f"Rekap_Udara_Atas_{nama_bulan}_{sample_year}.zip",
                    mime="application/zip",
                )
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan saat memproses data: {e}")
