import streamlit as st
import pandas as pd
import numpy as np
import io
import calendar
import re

st.set_page_config(page_title="Data Extract Job", layout="wide")
st.title("📊 EXCEL PARAMETER RATA-RATA KLIMAT, ME45 & LAPBUL TEKNISI")

with st.expander("ℹ️ Klik di sini untuk melihat Petunjuk Penggunaan"):
    st.markdown("""
    **Syarat File CSV:**
    - File hasil extract dari "https://bmkgsatu.bmkg.go.id/exportdata".
    - Harus mengandung kolom `encoded_synop` untuk ekstraksi grup sandi tekanan.
    """)

# --- FUNGSI UTAMA KALKULASI PARAMS MATRIKS (VEKTORISASI VEKTOR NUMPY) ---
def hitung_tekanan_uap_excel(suhu, rh):
    es = 6.112 * np.exp((17.67 * suhu) / (suhu + 243.5))
    return np.round((rh / 100.0) * es * 10, 2) 

def hitung_dewpoint(suhu, rh):
    a, b = 17.27, 237.7
    alpha = ((a * suhu) / (b + suhu)) + np.log(rh / 100.0)
    return np.round((b * alpha) / (a - alpha), 2)

# Ekstraksi Sandi 3 Jam (appp)
def ext_appp(v, syn):
    if pd.isna(v): return np.nan
    if pd.notna(syn):
        exp = f"{int(round(abs(v) * 10)):03d}"
        s_str = str(syn).split('333')[0]
        tokens = re.findall(r'\b5\d{4}\b', s_str)
        for tk in tokens:
            if tk.endswith(exp): return int(tk[1:])
        for a in range(10):
            if f"5{a}{exp}" in s_str: return int(f"{a}{exp}")
    return int(f"{'3' if v >= 0 else '8'}{int(round(abs(v) * 10)):03d}")

# Ekstraksi Sandi 24 Jam (P24P24P24)
def ext_p24(v, syn):
    if pd.isna(v): return np.nan
    if pd.notna(syn) and '333' in str(syn):
        exp = f"{int(round(abs(v) * 10)):03d}"
        s_str = str(syn).split('333')[1]
        tokens = re.findall(r'\b5[89]\d{3}\b', s_str)
        for tk in tokens:
            if tk.endswith(exp): return int(tk[2:]) if tk.startswith('58') else 500 + int(tk[2:])
        for prfx in ['58', '59']:
            if f"{prfx}{exp}" in s_str: return int(exp) if prfx == '58' else 500 + int(exp)
    return int(round(abs(v) * 10)) if v >= 0 else 500 + int(round(abs(v) * 10))

# --- PARAMETER MAPPING UNTUK MATRIKS ---
parameter_mapping = {
    'pressure_qff_mb_derived': 'QFF RATA-2 HARIAN',
    'pressure_qfe_mb_derived': 'QFE RATA-2 HARIAN',
    'temp_drybulb_c_tttttt': 'SUHU UDARA RATA-2 HARIAN',
    'Dewpoint': 'SUHU TITIK EMBUN (TD) RATA-2 HARIAN',
    'relative_humidity_pc': 'KELEMBABAN UDARA RATA-2 HARIAN',
    'wind_speed_ff': 'KECEPATAN ANGIN RATA-2 HARIAN',
    'Tekanan_Uap_x10': 'TEKANAN UAP AIR RATA-2 HARIAN',
    'pressure_3h_diff_mb_ppp': 'PERUBAHAN TEKANAN 3 JAM (5APP) RATA-2 HARIAN',
    'pressure_24h_diff_mb_p24': 'PERUBAHAN TEKANAN 24 JAM (58/59P24) RATA-2 HARIAN'
}

uploaded_file = st.file_uploader("Unggah file CSV...", type=["csv"])

if uploaded_file is not None:
    try:
        df_raw = pd.read_csv(uploaded_file)
        df_raw.replace([9999, 99999, '9999', '/', '//', '///', '#REF!', '#VALUE!', 'STNR', '#N/A'], np.nan, inplace=True)
        
        # Waktu Parsing
        df_raw['data_timestamp'] = pd.to_datetime(df_raw['data_timestamp'])
        df_raw['Tahun'] = df_raw['data_timestamp'].dt.year
        df_raw['Bulan_Angka'] = df_raw['data_timestamp'].dt.month
        df_raw['Tanggal'] = df_raw['data_timestamp'].dt.day
        df_raw['Jam'] = df_raw['data_timestamp'].dt.hour
        
        # Vektorisasi Perhitungan Fisika
        if 'temp_drybulb_c_tttttt' in df_raw.columns and 'relative_humidity_pc' in df_raw.columns:
            df_raw['Tekanan_Uap_x10'] = hitung_tekanan_uap_excel(df_raw['temp_drybulb_c_tttttt'], df_raw['relative_humidity_pc'])
            df_raw['Dewpoint'] = hitung_dewpoint(df_raw['temp_drybulb_c_tttttt'], df_raw['relative_humidity_pc'])
        
        df_raw['Bulan_Tahun'] = df_raw['Tahun'].astype(str) + "-" + df_raw['Bulan_Angka'].astype(str).str.zfill(2)
        
        st.markdown("---")
        bulan_dipilih = st.selectbox("Pilih Bulan untuk di-Generate:", sorted(df_raw['Bulan_Tahun'].unique()))
        df_bulan_ini = df_raw[df_raw['Bulan_Tahun'] == bulan_dipilih].copy()
        
        tahun_val = int(bulan_dipilih.split('-')[0])
        bulan_val = int(bulan_dipilih.split('-')[1])
        nama_bulan = calendar.month_name[bulan_val].upper()
        jml_hari = calendar.monthrange(tahun_val, bulan_val)[1]
        semua_tanggal = pd.Index(range(1, jml_hari + 1), name='NO.')

        # =====================================================================
        # 1. GENERATOR EXCEL MATRIKS KLIMAT RATA-RATA
        # =====================================================================
        buffer_matriks = io.BytesIO()
        with pd.ExcelWriter(buffer_matriks, engine='xlsxwriter') as writer:
            wb = writer.book
            ws = wb.add_worksheet('MATRIKS')
            
            fmt_teks = wb.add_format({'bold': True, 'align': 'left'})
            fmt_judul = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 12})
            fmt_header = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9D9D9'})
            fmt_blank = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
            fmt_blank_rata2 = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#DCE6F1'})
            fmt_qff_biasa = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0000'}) 
            fmt_p24_biasa = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '000'}) 
            fmt_int_biasa = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0'})     
            fmt_float_biasa = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0.0'}) 
            fmt_int_rata2 = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#DCE6F1', 'num_format': '0'})
            fmt_float_rata2 = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#DCE6F1', 'num_format': '0.0'})
            fmt_summary_judul = wb.add_format({'bold': True, 'align': 'right', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFF2CC'})
            fmt_summary_kosong = wb.add_format({'border': 1, 'bg_color': '#FFF2CC'})
            fmt_summary_final_int = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FCD5B4', 'num_format': '0'})
            fmt_summary_final_float = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FCD5B4', 'num_format': '0.0'})
            fmt_summary_final_blank = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FCD5B4'})
            
            ws.set_column('A:A', 6)
            ws.set_column('B:Y', 5)
            ws.set_column('Z:Z', 13)
            ws.set_column('AA:AC', 10) 
            
            start_row = 0
            
            def get_cell_matriks(val, param, col_type):
                if pd.isna(val): return "", (fmt_blank_rata2 if col_type == 'RATA2' else fmt_blank)
                if 'KECEPATAN ANGIN' in param:
                    v = int(round(float(val)))
                    if v == 0: return "", (fmt_blank_rata2 if col_type == 'RATA2' else fmt_blank)
                    return v, (fmt_int_rata2 if col_type == 'RATA2' else fmt_int_biasa)
                if 'QFF' in param or 'QFE' in param or 'SUHU' in param or 'PERUBAHAN TEKANAN' in param:
                    if col_type == '0-23':
                        if 'SUHU' in param: return int(round(float(val) * 10)), fmt_int_biasa
                        if '3 JAM' in param: return int(val), fmt_qff_biasa
                        if '24 JAM' in param: return int(val), fmt_p24_biasa
                        return (int(round(float(val) * 10)) % 10000), fmt_qff_biasa
                    else:
                        if '3 JAM' in param and col_type == 'SPEC': return int(val), fmt_qff_biasa
                        if '24 JAM' in param and col_type == 'SPEC': return int(val), fmt_p24_biasa
                        return round(float(val), 1), (fmt_float_rata2 if col_type == 'RATA2' else fmt_float_rata2 if col_type == 'RATA2' else fmt_float_biasa)
                if 'KELEMBABAN' in param or 'UAP AIR' in param:
                    if col_type == 'RATA2': return round(float(val), 1), fmt_float_rata2
                    return int(round(float(val))), fmt_int_biasa
                return round(float(val), 1), (fmt_float_rata2 if col_type == 'RATA2' else fmt_float_biasa)

            for kolom_csv, judul_param in parameter_mapping.items():
                if kolom_csv in df_bulan_ini.columns:
                    df_float_check = df_bulan_ini.copy()
                    df_float_check[kolom_csv] = pd.to_numeric(df_float_check[kolom_csv], errors='coerce')
                    pivot_float = df_float_check.pivot_table(index='Tanggal', columns='Jam', values=kolom_csv, aggfunc='first')
                    for h in range(24):
                        if h not in pivot_float.columns: pivot_float[h] = np.nan
                    pivot_float = pivot_float[list(range(24))].reindex(semua_tanggal)
                    rata_harian = pivot_float.mean(axis=1)
                    if 'UAP AIR' in judul_param: rata_harian = rata_harian / 10

                    if kolom_csv == 'pressure_3h_diff_mb_ppp' and 'encoded_synop' in df_bulan_ini.columns:
                        df_bulan_ini['sandi_5app'] = [ext_appp(v, syn) for v, syn in zip(df_bulan_ini['pressure_3h_diff_mb_ppp'], df_bulan_ini['encoded_synop'])]
                        pivot = df_bulan_ini.pivot_table(index='Tanggal', columns='Jam', values='sandi_5app', aggfunc='first')
                    elif kolom_csv == 'pressure_24h_diff_mb_p24' and 'encoded_synop' in df_bulan_ini.columns:
                        df_bulan_ini['sandi_p24'] = [ext_p24(v, syn) for v, syn in zip(df_bulan_ini['pressure_24h_diff_mb_p24'], df_bulan_ini['encoded_synop'])]
                        pivot = df_bulan_ini.pivot_table(index='Tanggal', columns='Jam', values='sandi_p24', aggfunc='first')
                    else:
                        df_bulan_ini.loc[:, kolom_csv] = pd.to_numeric(df_bulan_ini[kolom_csv], errors='coerce')
                        pivot = df_bulan_ini.pivot_table(index='Tanggal', columns='Jam', values=kolom_csv, aggfunc='first')
                    
                    for h in range(24):
                        if h not in pivot.columns: pivot[h] = np.nan
                    pivot = pivot[list(range(24))].reindex(semua_tanggal)

                    extra_headers = [] if 'UAP AIR' in judul_param else (['MAX HARIAN'] if 'KECEPATAN ANGIN' in judul_param else ['23 00', '05 00', '10 00'])
                    if 'UAP AIR' in judul_param: summary_labels = []
                    elif 'KECEPATAN ANGIN' in judul_param:
                        daily_max_angin = pivot.max(axis=1)
                        summary_labels = [("MAXIMUM BULAN INI", pivot.max().max())]
                    else:
                        summary_labels = [("MAXIMUM BULAN INI", pivot_float.max().max()), ("MINIMUM BULAN INI", pivot_float.min().min()), ("TOTAL RATA-RATA", rata_harian.mean())]
                    
                    ws.write(start_row, 1, "BULAN", fmt_teks)
                    ws.write(start_row, 3, nama_bulan, fmt_teks)
                    ws.write(start_row, 4, str(tahun_val), fmt_teks)
                    ws.write(start_row + 1, 12, judul_param, fmt_judul)
                    
                    ws.write(start_row + 2, 0, "NO.", fmt_header)
                    for i in range(24): ws.write(start_row + 2, i + 1, str(i), fmt_header)
                    ws.write(start_row + 2, 25, "R A T A   2", fmt_header)
                    for c_i, ext_hdr in enumerate(extra_headers): ws.write(start_row + 2, 26 + c_i, ext_hdr, fmt_header)
                    
                    row_idx = start_row + 3
                    for tgl in semua_tanggal:
                        ws.write(row_idx, 0, tgl, fmt_header)
                        for h in range(24):
                            val, fmt = get_cell_matriks(pivot.loc[tgl, h], judul_param, '0-23')
                            ws.write(row_idx, h + 1, val, fmt)
                        val_rata, fmt_rata = get_cell_matriks(rata_harian.loc[tgl], judul_param, 'RATA2')
                        ws.write(row_idx, 25, val_rata, fmt_rata)
                        
                        if 'UAP AIR' in judul_param: pass
                        elif 'KECEPATAN ANGIN' in judul_param:
                            v_max = daily_max_angin.loc[tgl]
                            ws.write(row_idx, 26, "" if pd.isna(v_max) or float(v_max) == 0 else int(round(float(v_max))), fmt_int_biasa)
                        else:
                            for c_idx, h_spec in zip([26, 27, 28], [23, 5, 10]):
                                val_s, fmt_s = get_cell_matriks(pivot.loc[tgl, h_spec], judul_param, 'SPEC')
                                ws.write(row_idx, c_idx, val_s, fmt_s)
                        row_idx += 1
                        
                    for label, final_val in summary_labels:
                        ws.merge_range(row_idx, 0, row_idx, 24, label, fmt_summary_judul)
                        if pd.isna(final_val): ws.write(row_idx, 25, "", fmt_summary_final_blank)
                        else:
                            if 'KECEPATAN ANGIN' in judul_param:
                                v_ang = int(round(float(final_val)))
                                ws.write(row_idx, 25, "" if v_ang == 0 else v_ang, fmt_summary_final_int)
                            else: ws.write(row_idx, 25, round(float(final_val), 1), fmt_summary_final_float)
                        for c_i in range(len(extra_headers)): ws.write(row_idx, 26 + c_i, "", fmt_summary_kosong)
                        row_idx += 1
                    start_row = row_idx + 3

        # =====================================================================
        # 2. GENERATOR FORMULIR ME45 DENGAN ATURAN SANDI BARU + AUTO-FILTER
        # =====================================================================
        buffer_me45 = io.BytesIO()
        
        me45_rows = []
        station_id = int(df_raw['station_name'].iloc[0]) if ('station_name' in df_raw.columns and str(df_raw['station_name'].iloc[0]).isdigit()) else 97340
        station_label = "Stasiun Meteorologi Umbu Mehang Kunda"
        
        for d in range(1, jml_hari + 1):
            for h in range(24):
                me45_rows.append({'NoSta': station_id, 'Station': station_label, 'YY': tahun_val, 'MM': bulan_val, 'DD': d, 'HH': h})
        df_me45_final = pd.DataFrame(me45_rows)
        
        df_merge = pd.merge(df_me45_final, df_bulan_ini, left_on=['DD', 'HH'], right_on=['Tanggal', 'Jam'], how='left')
        
        df_out = pd.DataFrame()
        df_out['NoSta'] = df_merge['NoSta']
        df_out['Station'] = df_merge['Station']
        df_out['YY'] = df_merge['YY']
        df_out['MM'] = df_merge['MM']
        df_out['DD'] = df_merge['DD']
        df_out['HH'] = df_merge['HH']
        
        df_out['TdTdTd'] = (df_merge['temp_dewpoint_c_tdtdtd'] * 10).round().fillna(np.nan)
        df_out['N'] = df_merge['cloud_cover_oktas_m'].fillna(np.nan)
        df_out['dd'] = (df_merge['wind_dir_deg_dd'] / 10).round().fillna(np.nan)
        df_out['ff'] = df_merge['wind_speed_ff'].round().fillna(np.nan)
        df_out['VV'] = (df_merge['visibility_vv'] + 50).round().fillna(np.nan)
        df_out['ww'] = df_merge['present_weather_ww'].fillna(np.nan)
        df_out['W1'] = df_merge['past_weather_w1'].fillna(np.nan)
        df_out['W2'] = df_merge['past_weather_w2'].fillna(np.nan)
        df_out['QFF'] = (df_merge['pressure_qff_mb_derived'] * 10).round().fillna(np.nan) % 10000
        df_out['TtTtTt'] = (df_merge['temp_drybulb_c_tttttt'] * 10).round().fillna(np.nan)
        df_out['Nh'] = df_merge['cloud_low_cover_oktas'].fillna(np.nan)
        df_out['CL'] = df_merge['cloud_low_type_cl'].fillna(np.nan)
        df_out['h'] = df_merge['cloud_low_base_1'].fillna(np.nan)  
        df_out['CM'] = df_merge['cloud_med_type_cm'].fillna(np.nan)
        df_out['CH'] = df_merge['cloud_high_type_ch'].fillna(np.nan)
        
        for blank_c in ['Ns', 'C', 'hshs', 'Ns.1', 'C.1', 'hshs.1', 0, 'C.2', 'hshs.2', 'C.3', 'D', 'e']:
            df_out[blank_c] = np.nan
            
        df_out['UU'] = df_merge['relative_humidity_pc'].round().fillna(np.nan)
        df_out['QFE'] = (df_merge['pressure_qfe_mb_derived'] * 10).round().fillna(np.nan) % 10000
        df_out['TwTwTw'] = (df_merge['temp_wetbulb_c'] * 10).round().fillna(np.nan)
        df_out['RRR'] = df_merge['rainfall_6h_rrr'].fillna(np.nan)
        df_out['tR'] = df_merge['rainfall_indicator_ir'].fillna(np.nan)
        df_out['TxTxTx'] = (df_merge['temp_max_c_txtxtx'] * 10).round().fillna(np.nan)
        df_out['TnTnTn'] = (df_merge['temp_min_c_tntntn'] * 10).round().fillna(np.nan)
        df_out['EEE'] = (df_merge['evaporation_24hours_mm_eee'] * 10).round().fillna(np.nan)
        df_out['F24F24F24F24'] = np.nan
        df_out['SSS'] = (df_merge['sunshine_h_sss'] * 10).round().fillna(np.nan)
        df_out['E'] = df_merge['land_cond'].fillna(np.nan)
        
        for blank_c2 in ['DL', 'DM', 'DH']:
            df_out[blank_c2] = np.nan
            
        df_out['appp'] = [ext_appp(v, syn) for v, syn in zip(df_merge['pressure_3h_diff_mb_ppp'], df_merge['encoded_synop'])]
        df_out['P24P24P24'] = [ext_p24(v, syn) for v, syn in zip(df_merge['pressure_24h_diff_mb_p24'], df_merge['encoded_synop'])]
        df_out['iW'] = df_merge['wind_indicator_iw'].fillna(np.nan)
        df_out['iX'] = df_merge['weather_indicator_ix'].fillna(np.nan)
        df_out['iR'] = df_merge['rainfall_indicator_ir'].fillna(np.nan)
        df_out['iE'] = df_merge['evaporation_eq_indicator_ie'].fillna(np.nan)

        with pd.ExcelWriter(buffer_me45, engine='xlsxwriter') as writer_me45:
            df_out.to_excel(writer_me45, sheet_name='Sheet1', index=False)
            
            wb_m = writer_me45.book
            ws_m = writer_me45.sheets['Sheet1']
            
            ws_m.hide_gridlines(2)
            
            fmt_header_me45 = wb_m.add_format({
                'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#1F4E78',
                'border': 1, 'border_color': '#D9D9D9', 'align': 'center', 'valign': 'vcenter'
            })
            fmt_data_me45 = wb_m.add_format({
                'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#E0E0E0'
            })
            
            fmt_2dig = wb_m.add_format({'num_format': '00', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#E0E0E0'})
            fmt_3dig = wb_m.add_format({'num_format': '000', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#E0E0E0'})
            fmt_4dig = wb_m.add_format({'num_format': '0000', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#E0E0E0'})
            
            for col_num, value in enumerate(df_out.columns.values):
                ws_m.write(0, col_num, value, fmt_header_me45)
                
            ws_m.set_row(0, 26)
            ws_m.set_column(0, len(df_out.columns) - 1, 9, fmt_data_me45)
            
            ws_m.set_column(df_out.columns.get_loc('dd'), df_out.columns.get_loc('dd'), 9, fmt_2dig)
            ws_m.set_column(df_out.columns.get_loc('ff'), df_out.columns.get_loc('ff'), 9, fmt_2dig)
            ws_m.set_column(df_out.columns.get_loc('VV'), df_out.columns.get_loc('VV'), 9, fmt_2dig)
            ws_m.set_column(df_out.columns.get_loc('EEE'), df_out.columns.get_loc('EEE'), 9, fmt_3dig)
            ws_m.set_column(df_out.columns.get_loc('SSS'), df_out.columns.get_loc('SSS'), 9, fmt_3dig)
            
            ws_m.set_column(df_out.columns.get_loc('QFF'), df_out.columns.get_loc('QFF'), 9, fmt_4dig)
            ws_m.set_column(df_out.columns.get_loc('QFE'), df_out.columns.get_loc('QFE'), 9, fmt_4dig)
            ws_m.set_column(df_out.columns.get_loc('appp'), df_out.columns.get_loc('appp'), 10, fmt_4dig)
            ws_m.set_column(df_out.columns.get_loc('P24P24P24'), df_out.columns.get_loc('P24P24P24'), 11, fmt_3dig)
            
            ws_m.autofilter(0, 0, len(df_out), len(df_out.columns) - 1)

        # =====================================================================
        # 3. GENERATOR LAPBUL TEKNISI (MONITORING PERALATAN OPERASIONAL)
        # =====================================================================
        buffer_lapbul = io.BytesIO()
        triwulan_map = {1: "I", 2: "I", 3: "I", 4: "II", 5: "II", 6: "II", 7: "III", 8: "III", 9: "III", 10: "IV", 11: "IV", 12: "IV"}
        triwulan_str = f"TRIWULAN {triwulan_map[bulan_val]}"

        # Master Data Peralatan Stamet & Posmet
        items_stamet = [
            ("1", "AWS Synoptic Strengthening", "Taman Alat Stamet", "Degreane Horizone", "", "", "", "", "", "2015"),
            ("", "Sensor Suhu dan Kelembapan", "", "", "B", "B", "B", "B", "September 2025", ""),
            ("", "Sensor Tekanan", "", "", "B", "B", "B", "B", "September 2025", ""),
            ("", "Sensor Anemometer", "", "", "B", "B", "B", "B", "September 2025", ""),
            ("", "Sensor Tipping Bucket", "", "", "B", "B", "B", "B", "September 2025", ""),
            ("", "Sensor Radiasi Matahari", "", "", "B", "B", "B", "B", "September 2025", ""),
            ("", "Sensor Soil Temperature", "", "", "B", "B", "B", "B", "September 2025", ""),
            ("", "Pyranometer", "", "", "B", "B", "B", "B", "", ""),
            ("2", "AWOS kategori 1", "Bandara Umbu Mehang Kunda", "All Weather", "", "", "", "", "", "2018"),
            ("", "Sensor Suhu dan Kelembapan", "", "All Weather", "B", "B", "B", "B", "September 2025", ""),
            ("", "Sensor Tekanan", "", "All Weather", "B", "B", "B", "B", "September 2025", ""),
            ("", "Sensor Tipping Bucket", "", "All Weather", "B", "B", "B", "B", "September 2025", ""),
            ("", "Sensor Radiasi Matahari", "", "All Weather", "B", "B", "B", "B", "September 2025", ""),
            ("", "Sensor Windsonic", "", "All Weather", "B", "B", "B", "B", "Desember 2025", ""),
            ("3", "Anemometer", "Taman Alat Stamet", "R.M. Young", "B", "B", "B", "B", "September 2025", "2008"),
            ("4", "Barometer Digital", "Taman Alat Stamet", "Vaisala", "B", "B", "B", "B", "September 2025", "2012"),
            ("5", "Thermometer Bola Basah Bola Kering", "Taman Alat Stamet", "Thermo Schneider", "B", "B", "B", "B", "September 2025", "2012"),
            ("6", "Thermometer Bola Basah Bola Kering", "Taman Alat Stamet", "Franz Ketterer", "B", "B", "B", "B", "September 2025", "2004"),
            ("7", "Thermometer Max", "Taman Alat Stamet", "Thermo Schneider", "B", "B", "B", "B", "September 2025", "2012"),
            ("8", "Thermometer Min", "Taman Alat Stamet", "Thermo Schneider", "B", "B", "B", "B", "September 2025", "2012"),
            ("9", "Thermohigrograph", "Taman Alat Stamet", "R. Fuess", "B", "B", "B", "B", "September 2025", "2010"),
            ("10", "Penakar Hujan Obs", "Taman Alat Stamet", "Lokal", "B", "B", "B", "B", "September 2025", "2012"),
            ("11", "Campbel Stoke", "Taman Alat Stamet", "R. Fuess", "B", "B", "B", "B", "", "1975"),
            ("12", "Rainfall Recorder", "Taman Alat Stamet", "Hellmann", "B", "B", "B", "B", "September 2025", "2021"),
            ("13", "Panci Penguapan", "Taman Alat Stamet", "Lokal", "B", "B", "B", "B", "", "1982"),
            ("14", "Theodolite", "Taman Alat Stamet", "Warren/Knight", "RR", "RR", "RR", "RR", "", "2009")
        ]

        items_posmet = [
            ("15", "AWS Rekayasa", "Taman Alat Posmet", "Campbell Scientific", "", "", "", "", "April 2026", "2009"),
            ("", "Sensor Tekanan", "", "", "B", "B", "B", "B", "April 2026", ""),
            ("", "Sensor Suhu dan Kelembapan", "", "", "B", "B", "B", "B", "April 2026", ""),
            ("", "Sensor Windmarine", "", "", "B", "B", "B", "B", "April 2026", ""),
            ("", "Sensor Tipping Bucket", "", "", "B", "B", "B", "B", "April 2026", ""),
            ("", "Pyranometer", "", "", "B", "B", "B", "B", "April 2026", ""),
            ("16", "AWOS kategori 2", "Bandara Lede Kalumbang", "All Weather", "", "", "", "", "", "2016"),
            ("", "Sensor Suhu dan Kelembapan", "", "All Weather", "B", "B", "B", "B", "April 2026", ""),
            ("", "Sensor Tekanan", "", "All Weather", "B", "B", "B", "B", "April 2026", ""),
            ("", "Sensor Tipping Bucket", "", "All Weather", "B", "B", "B", "B", "April 2026", ""),
            ("", "Sensor Radiasi Matahari", "", "All Weather", "B", "B", "B", "B", "April 2026", ""),
            ("", "Sensor Windsonic", "", "All Weather", "B", "B", "B", "B", "April 2026", ""),
            ("", "Sensor Present Weather", "", "All Weather", "B", "B", "B", "B", "April 2026", ""),
            ("", "Sensor Perawanan (Ceilometer)", "", "All Weather", "B", "B", "B", "B", "April 2026", ""),
            ("", "Sensor Jarak Pandang dan RVR", "", "All Weather", "B", "B", "B", "B", "April 2026", ""),
            ("17", "Anemometer", "Taman Alat Posmet", "R.M. Young", "B", "B", "B", "B", "April 2026", "2014"),
            ("18", "Barometer Digital", "Taman Alat Posmet", "Vaisala", "B", "B", "B", "B", "April 2026", "2013"),
            ("19", "Thermometer Bola Basah Bola Kering", "Taman Alat Posmet", "Thermo Schneider", "B", "B", "B", "B", "April 2026", "2012"),
            ("20", "Thermometer Max", "Taman Alat Posmet", "Thermo Schneider", "B", "B", "B", "B", "April 2026", "2012"),
            ("21", "Thermometer Min", "Taman Alat Posmet", "Thermo Schneider", "B", "B", "B", "B", "April 2026", "2012"),
            ("22", "Penakar Hujan Obs", "Taman Alat Posmet", "Lokal", "B", "B", "B", "B", "", "2014"),
            ("23", "Campbel Stoke", "Taman Alat Posmet", "", "B", "B", "B", "B", "", "2014"),
            ("24", "Panci Penguapan", "Taman Alat Posmet", "Lokal", "B", "B", "B", "B", "", "2012")
        ]

        with pd.ExcelWriter(buffer_lapbul, engine='xlsxwriter') as writer_lapbul:
            wb_l = writer_lapbul.book
            ws_l = wb_l.add_worksheet('LOGBOOK')
            
            fmt_title = wb_l.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 11})
            fmt_th_main = wb_l.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9D9D9'})
            fmt_cell_center = wb_l.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
            fmt_cell_left = wb_l.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1})
            fmt_total = wb_l.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFF2CC'})

            ws_l.set_column('A:A', 5)
            ws_l.set_column('B:B', 38)
            ws_l.set_column('C:C', 26)
            ws_l.set_column('D:D', 22)
            ws_l.set_column('E:H', 10)
            ws_l.set_column('I:J', 16)

            def render_block(start_r, nama_lokasi, items):
                ws_l.merge_range(start_r, 0, start_r, 9, "LAPORAN HASIL MONITORING KONDISI PERALATAN OPERASIONAL UTAMA METEOROLOGI", fmt_title)
                ws_l.merge_range(start_r + 1, 0, start_r + 1, 9, nama_lokasi, fmt_title)
                ws_l.merge_range(start_r + 2, 0, start_r + 2, 9, f"TAHUN {tahun_val}", fmt_title)
                ws_l.merge_range(start_r + 3, 0, start_r + 3, 9, triwulan_str, fmt_title)
                ws_l.merge_range(start_r + 4, 0, start_r + 4, 9, f"BULAN {nama_bulan}", fmt_title)

                ws_l.merge_range(start_r + 5, 0, start_r + 6, 0, "No", fmt_th_main)
                ws_l.merge_range(start_r + 5, 1, start_r + 6, 1, "Nama Alat", fmt_th_main)
                ws_l.merge_range(start_r + 5, 2, start_r + 6, 2, "Lokasi", fmt_th_main)
                ws_l.merge_range(start_r + 5, 3, start_r + 6, 3, "Merk/Type", fmt_th_main)
                ws_l.merge_range(start_r + 5, 4, start_r + 5, 7, "KONDISI", fmt_th_main)
                
                pekan = ["PEKAN I", "PEKAN II", "PEKAN III", "PEKAN IV"]
                for i, p in enumerate(pekan):
                    ws_l.write(start_r + 6, 4 + i, p, fmt_th_main)

                ws_l.merge_range(start_r + 5, 8, start_r + 6, 8, "Tahun Kalibrasi", fmt_th_main)
                ws_l.merge_range(start_r + 5, 9, start_r + 6, 9, "Tahun Pengadaan", fmt_th_main)

                curr_r = start_r + 7
                for it in items:
                    ws_l.write(curr_r, 0, it[0], fmt_cell_center)
                    ws_l.write(curr_r, 1, it[1], fmt_cell_left)
                    ws_l.write(curr_r, 2, it[2], fmt_cell_left)
                    ws_l.write(curr_r, 3, it[3], fmt_cell_left)
                    ws_l.write(curr_r, 4, it[4], fmt_cell_center)
                    ws_l.write(curr_r, 5, it[5], fmt_cell_center)
                    ws_l.write(curr_r, 6, it[6], fmt_cell_center)
                    ws_l.write(curr_r, 7, it[7], fmt_cell_center)
                    ws_l.write(curr_r, 8, it[8], fmt_cell_center)
                    ws_l.write(curr_r, 9, it[9], fmt_cell_center)
                    curr_r += 1
                return curr_r

            r_next = render_block(0, "STASIUN METEOROLOGI UMBU MEHANG KUNDA", items_stamet)
            r_next = render_block(r_next + 2, "POS METEOROLOGI TAMBOLAKA", items_posmet)

            ws_l.merge_range(r_next, 0, r_next, 4, "TOTAL", fmt_total)
            ws_l.write(r_next, 5, "24", fmt_total)
            for c_tot in range(6, 10):
                ws_l.write(r_next, c_tot, "", fmt_total)

        st.success("🎉 Berhasil! Seluruh berkas otomatisasi telah diperbarui.")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.download_button(
                label=f"📥 1. Unduh Matriks Klimat ({bulan_dipilih})",
                data=buffer_matriks.getvalue(),
                file_name=f"Matriks_Klimat_{bulan_dipilih}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col2:
            st.download_button(
                label=f"📥 2. Unduh Berkas ME45 ({bulan_dipilih})",
                data=buffer_me45.getvalue(),
                file_name=f"ME45_Standard_{bulan_dipilih}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col3:
            st.download_button(
                label=f"📥 3. Unduh Lapbul Teknisi ({bulan_dipilih})",
                data=buffer_lapbul.getvalue(),
                file_name=f"Lapbul_Teknisi_{bulan_dipilih}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses berkas: {e}")
