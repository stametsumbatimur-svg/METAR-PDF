import streamlit as st
import pandas as pd
import re
import io
import os
import calendar
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

# ==========================================
# ======= METAR PDF & EXCEL GENERATOR ======
# ==========================================
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
        tanggal_format = f"{date.day:02d} {nama_bulan} {date.year}"
        judul_rekap = f"REKAP DATA METAR: {tanggal_format}".upper()
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

# ==========================================
# ======= SPECI PDF & EXCEL GENERATOR ======
# ==========================================
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
        workbook = writer.book
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

# ==========================================
# ======= WXREV CONVERTER FUNCTIONS ========
# ==========================================
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
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ALIGN', (0,0), (0,0), 'CENTER'),
            ('ALIGN', (1,0), (1,0), 'CENTER'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING', (0,0), (-1,-1), 0),
            ('LINEBELOW', (0,0), (-1,-1), 1.2, colors.black),
            ('LEFTPADDING', (1,0), (1,0), 0), ('RIGHTPADDING', (1,0), (1,0), 0)
        ]))
    else:
        header_table = Table([[text_block]], colWidths=[total_width])
        header_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('LINEBELOW', (0,0), (-1,-1), 1.2, colors.black),
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
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
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
        workbook = writer.book
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

# ==========================================
# ===== THUNDERSTORM PDF & EXCEL GENERATOR =
# ==========================================
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
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),
        ('FONTNAME', (0, 1), (0, -1), 'Helvetica-Bold'),
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
# ===== PERAWANAN PDF & EXCEL GENERATOR ====
# ==========================================
def generate_pdf_bytes_perawanan(df_clean, station_name, kepala_nama, kepala_nip, form_code="Klim / PWN/WPU-2001"):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=25, bottomMargin=25)
    story = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('PWNTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=13, alignment=1)
    text_style = ParagraphStyle('PWNText', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=11)
    
    dt_first = df_clean['datetime'].iloc[0]
    bulan_str = BULAN_INDO[dt_first.month]
    tahun = dt_first.year
    num_days = calendar.monthrange(tahun, dt_first.month)[1]
    
    story.append(Paragraph("DAFTAR IKHTISAR FREKUENSI PERAWANAN", title_style))
    story.append(Paragraph(f"{station_name.upper()}", title_style))
    story.append(Paragraph(f"BULAN {bulan_str} {tahun}", title_style))
    story.append(Spacer(1, 10))
    
    df_clean['day'] = df_clean['datetime'].dt.day
    df_clean['cat'] = df_clean.apply(lambda r: get_cloud_category(r['CLOUD'], r['VIS']), axis=1)
    
    header_row = [
        "TANGGAL", "CERAH\n(N=0)", "BERAWAN SEBAGIAN\n(N=1-3)",
        "BERAWAN\n(N=4-6)", "BERAWAN BANYAK\n(N=7-8)", "JUMLAH"
    ]
    table_data = [header_row]
    
    tot_cerah = tot_sebagian = tot_berawan = tot_banyak = tot_semua = 0
    
    for d in range(1, num_days + 1):
        df_day = df_clean[df_clean['day'] == d]
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
            f"{d:02d}",
            str(c_cerah) if c_cerah > 0 else "",
            str(c_sebagian) if c_sebagian > 0 else "",
            str(c_berawan) if c_berawan > 0 else "",
            str(c_banyak) if c_banyak > 0 else "",
            str(c_total) if c_total > 0 else ""
        ])
        
    table_data.append([
        "JUMLAH",
        str(tot_cerah),
        str(tot_sebagian),
        str(tot_berawan),
        str(tot_banyak),
        str(tot_semua)
    ])
    
    pct_c = f"{round((tot_cerah/tot_semua)*100)}%" if tot_semua > 0 else "0%"
    pct_s = f"{round((tot_sebagian/tot_semua)*100)}%" if tot_semua > 0 else "0%"
    pct_bw = f"{round((tot_berawan/tot_semua)*100)}%" if tot_semua > 0 else "0%"
    pct_by = f"{round((tot_banyak/tot_semua)*100)}%" if tot_semua > 0 else "0%"
    
    table_data.append([
        "% - TAGE",
        pct_c,
        pct_s,
        pct_bw,
        pct_by,
        "100%" if tot_semua > 0 else "0%"
    ])
    
    col_widths = [60, 90, 120, 100, 110, 60]
    pwn_table = Table(table_data, colWidths=col_widths)
    
    pwn_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 7.5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('FONTNAME', (0, -2), (-1, -1), 'Helvetica-Bold'),
        ('BACKGROUND', (0, -2), (-1, -1), colors.whitesmoke),
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
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
    ]))
    
    story.append(KeepTogether(footer_table))

    doc.build(story)
    buffer.seek(0)
    return buffer

def generate_excel_bytes_perawanan(df_clean, station_name):
    buffer = io.BytesIO()
    dt_first = df_clean['datetime'].iloc[0]
    bulan_str = BULAN_INDO[dt_first.month]
    tahun = dt_first.year
    num_days = calendar.monthrange(tahun, dt_first.month)[1]
    
    df_clean['day'] = df_clean['datetime'].dt.day
    df_clean['cat'] = df_clean.apply(lambda r: get_cloud_category(r['CLOUD'], r['VIS']), axis=1)
    
    rows_data = []
    tot_cerah = tot_sebagian = tot_berawan = tot_banyak = tot_semua = 0
    
    for d in range(1, num_days + 1):
        df_day = df_clean[df_clean['day'] == d]
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
        
        rows_data.append({
            'TANGGAL': f"{d:02d}",
            'CERAH (N=0)': c_cerah if c_cerah > 0 else "",
            'BERAWAN SEBAGIAN (N=1-3)': c_sebagian if c_sebagian > 0 else "",
            'BERAWAN (N=4-6)': c_berawan if c_berawan > 0 else "",
            'BERAWAN BANYAK (N=7-8)': c_banyak if c_banyak > 0 else "",
            'JUMLAH': c_total if c_total > 0 else ""
        })
        
    rows_data.append({
        'TANGGAL': "JUMLAH",
        'CERAH (N=0)': tot_cerah,
        'BERAWAN SEBAGIAN (N=1-3)': tot_sebagian,
        'BERAWAN (N=4-6)': tot_berawan,
        'BERAWAN BANYAK (N=7-8)': tot_banyak,
        'JUMLAH': tot_semua
    })
    
    pct_c = f"{round((tot_cerah/tot_semua)*100)}%" if tot_semua > 0 else "0%"
    pct_s = f"{round((tot_sebagian/tot_semua)*100)}%" if tot_semua > 0 else "0%"
    pct_bw = f"{round((tot_berawan/tot_semua)*100)}%" if tot_semua > 0 else "0%"
    pct_by = f"{round((tot_banyak/tot_semua)*100)}%" if tot_semua > 0 else "0%"
    
    rows_data.append({
        'TANGGAL': "% - TAGE",
        'CERAH (N=0)': pct_c,
        'BERAWAN SEBAGIAN (N=1-3)': pct_s,
        'BERAWAN (N=4-6)': pct_bw,
        'BERAWAN BANYAK (N=7-8)': pct_by,
        'JUMLAH': "100%" if tot_semua > 0 else "0%"
    })
    
    df_pwn = pd.DataFrame(rows_data)
    
    sheet_name = f'Perawanan {bulan_str} {tahun}'[:31]
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
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
# ======== ANTARMUKA WEB STREAMLIT =========
# ==========================================
st.set_page_config(page_title="BMKG Data Generator", layout="centered")

# --- SIDEBAR MENU ---
st.sidebar.title("🎛️ Navigasi Menu")
menu = st.sidebar.radio("Pilih Converter", ["METAR Converter", "SPECI Converter", "WXREV Converter", "Thunderstorm Exporter", "Form Perawanan Exporter"])
st.sidebar.markdown("---")
st.sidebar.info("Aplikasi ekstraksi data Sandi Cuaca (METAR, SPECI, WXREV, Thunderstorm & Perawanan) menjadi Laporan PDF & Excel otomatis.")

LOGO_FILE = "logo_bmkg.png"

if not os.path.exists(LOGO_FILE):
    st.warning(f"⚠️ File gambar '{LOGO_FILE}' tidak terdeteksi di folder utama. Harap pastikan file logo sudah di-upload.")

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
                    
                    if df_clean.empty:
                        st.warning("Tidak ditemukan data METAR di dalam file ini.")
                    else:
                        df_clean['raw_timestamp'] = df_clean['raw_timestamp'].str.replace(" +0000 UTC", "", regex=False)
                        df_clean['datetime'] = pd.to_datetime(df_clean['raw_timestamp'])
                        
                        df_clean = df_clean[df_clean['datetime'].dt.minute == 0]
                        df_clean['priority_score'] = df_clean.apply(calculate_priority, axis=1)
                        df_clean = df_clean.sort_values(by=['datetime', 'priority_score', 'msg_id'])
                        df_clean = df_clean.drop_duplicates(subset=['datetime'], keep='last')
                        
                        df_clean['date_group'] = df_clean['datetime'].dt.date
                        df_clean = df_clean.sort_values(by='datetime').reset_index(drop=True)
                        
                        if df_clean.empty:
                            st.warning("Tidak ditemukan data METAR dengan menit :00 (per jam) di dalam file ini.")
                        else:
                            st.success(f"Berhasil memproses data METAR!")
                            
                            pdf_data = generate_pdf_bytes_metar(df_clean, LOGO_FILE)
                            excel_data = generate_excel_bytes_metar_speci(df_clean, report_type="METAR")
                            
                            first_date = df_clean['date_group'].iloc[0]
                            nama_file_base = f"REKAP_METAR_{first_date.day:02d}_{BULAN_INDO[first_date.month]}_{first_date.year}"
                            
                            st.write("---")
                            st.subheader("Unduh Laporan METAR")
                            
                            col_pdf, col_xlsx = st.columns(2)
                            with col_pdf:
                                st.download_button(label="📥 Download PDF", data=pdf_data, file_name=f"{nama_file_base}.pdf", mime="application/pdf")
                            with col_xlsx:
                                st.download_button(label="📊 Download Excel", data=excel_data, file_name=f"{nama_file_base}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                                
        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")

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
                    
                    if df_clean.empty:
                        st.warning("Tidak ditemukan data SPECI di dalam file ini.")
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
                        with col_pdf:
                            st.download_button(label="📥 Download PDF", data=pdf_data, file_name=f"{nama_file_base}.pdf", mime="application/pdf")
                        with col_xlsx:
                            st.download_button(label="📊 Download Excel", data=excel_data, file_name=f"{nama_file_base}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                            
        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")

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
                            
                    if not parsed_rows:
                        st.warning("Tidak ditemukan data sandi WXREV di dalam file ini.")
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
                        with col_pdf:
                            st.download_button(label="📥 Download PDF", data=pdf_data, file_name=f"{nama_file_base}.pdf", mime="application/pdf")
                        with col_xlsx:
                            st.download_button(label="📊 Download Excel", data=excel_data, file_name=f"{nama_file_base}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                            
        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")

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

# --- HALAMAN FORM PERAWANAN EXPORTER ---
elif menu == "Form Perawanan Exporter":
    st.title("☁️ Form Perawanan Exporter")
    st.write("Ekstraksi data Ikhtisar Frekuensi Perawanan bulanan dari file METAR per jam.")

    with st.expander("⚙️ Pengaturan Header, Form & Tanda Tangan PDF"):
        st_nama = st.text_input("Nama Stasiun", "STASIUN METEOROLOGI UMBU MEHANG KUNDA", key="pwn_st")
        form_code = st.text_input("Kode Form (Pojok Kiri Bawah)", "Klim / PWN/WPU-2001")
        kp_nama = st.text_input("Nama Kepala Stasiun", "CARLES ALEXANDER TARI, S.TP", key="pwn_kp")
        kp_nip = st.text_input("NIP Kepala Stasiun", "197712082001121001", key="pwn_nip")

    uploaded_files = st.file_uploader("Upload File CSV METAR (Bulanan)", type=["csv"], accept_multiple_files=True, key="pwn_uploader")

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
            except Exception as e:
                st.error(f"Gagal membaca file {file.name}: {e}")

        if all_rows:
            columns = ['TYPE', 'LOC', 'TIME', 'WIND', 'VIS', 'WX', 'CLOUD', 'T/DP', 'QNH', 'RMK', 'is_cor', 'cc_type', 'raw_timestamp', 'station_name', 'msg_id']
            df_clean = pd.DataFrame(all_rows, columns=columns)
            df_clean['raw_timestamp'] = df_clean['raw_timestamp'].str.replace(" +0000 UTC", "", regex=False)
            df_clean['datetime'] = pd.to_datetime(df_clean['raw_timestamp'])
            
            # Perawanan hanya dihitung dari observasi METAR per jam (:00)
            df_clean = df_clean[df_clean['datetime'].dt.minute == 0]
            df_clean['priority_score'] = df_clean.apply(calculate_priority, axis=1)
            df_clean = df_clean.sort_values(by=['datetime', 'priority_score', 'msg_id'])
            df_clean = df_clean.drop_duplicates(subset=['datetime'], keep='last')
            df_clean = df_clean.sort_values(by='datetime').reset_index(drop=True)

            if df_clean.empty:
                st.warning("Tidak ditemukan data METAR dengan menit :00 di dalam file ini.")
            else:
                pdf_data = generate_pdf_bytes_perawanan(df_clean, st_nama, kp_nama, kp_nip, form_code)
                excel_data = generate_excel_bytes_perawanan(df_clean, st_nama)

                dt_first = df_clean['datetime'].iloc[0]
                bln_name = BULAN_INDO[dt_first.month]
                thn = dt_first.year
                
                st.success(f"Berhasil memproses Data Perawanan Bulan {bln_name} {thn} ({len(df_clean)} pengamatan)!")

                st.write("---")
                col_pdf, col_xlsx = st.columns(2)
                with col_pdf:
                    st.download_button(label="📥 Download PDF Perawanan", data=pdf_data, file_name=f"PERAWANAN_{bln_name}_{thn}.pdf", mime="application/pdf")
                with col_xlsx:
                    st.download_button(label="📊 Download Excel Perawanan", data=excel_data, file_name=f"PERAWANAN_{bln_name}_{thn}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
