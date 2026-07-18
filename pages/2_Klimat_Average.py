import streamlit as st
import pandas as pd
import numpy as np
import io
import calendar
import re

st.set_page_config(page_title="Data Extract Job", layout="wide")
st.title("📊 EXCEL PARAMETER RATA-RATA KLIMAT & ME45")

with st.expander("ℹ️ Klik di sini untuk melihat Petunjuk Penggunaan"):
    st.markdown("""
    **Syarat File CSV:**
    - File hasil extract dari "https://bmkgsatu.bmkg.go.id/exportdata".
    - Harus mengandung kolom `encoded_synop` untuk ekstraksi grup sandi tekanan.
    """)

# --- FUNGSI UTAMA KALKULASI PARAMS ---
def hitung_tekanan_uap_excel(suhu, rh):
    if pd.isna(suhu) or pd.isna(rh): return np.nan
    es = 6.112 * np.exp((17.67 * suhu) / (suhu + 243.5))
    return round((rh / 100.0) * es * 10, 2) 

def hitung_dewpoint(suhu, rh):
    if pd.isna(suhu) or pd.isna(rh): return np.nan
    a, b = 17.27, 237.7
    alpha = ((a * suhu) / (b + suhu)) + np.log(rh / 100.0)
    return round((b * alpha) / (a - alpha), 2)

# Ekstraksi Sandi 3 Jam (appp) lengkap tanpa awalan 5
def ext_appp(row):
    v = row['pressure_3h_diff_mb_ppp']
    if pd.isna(v): return np.nan
    syn = row['encoded_synop']
    if pd.notna(syn):
        exp = f"{int(round(abs(v) * 10)):03d}"
        tokens = re.findall(r'\b5\d{4}\b', syn.split('333')[0])
        for tk in tokens:
            if tk.endswith(exp): return int(tk[1:])
        for a in range(10):
            if f"5{a}{exp}" in syn.split('333')[0]: return int(f"{a}{exp}")
    return int(f"{'3' if v >= 0 else '8'}{int(round(abs(v) * 10)):03d}")

# Ekstraksi Sandi 24 Jam (P24P24P24)
def ext_p24(row):
    v = row['pressure_24h_diff_mb_p24']
    if pd.isna(v): return np.nan
    syn = row['encoded_synop']
    if pd.notna(syn) and '333' in syn:
        exp = f"{int(round(abs(v) * 10)):03d}"
        tokens = re.findall(r'\b5[89]\d{3}\b', syn.split('333')[1])
        for tk in tokens:
            if tk.endswith(exp): return int(tk[2:]) if tk.startswith('58') else 500 + int(tk[2:])
        for prfx in ['58', '59']:
            if f"{prfx}{exp}" in syn.split('333')[1]: return int(exp) if prfx == '58' else 500 + int(exp)
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
        
        if 'temp_drybulb_c_tttttt' in df_raw.columns and 'relative_humidity_pc' in df_raw.columns:
            df_raw['Tekanan_Uap_x10'] = df_raw.apply(lambda r: hitung_tekanan_uap_excel(r['temp_drybulb_c_tttttt'], r['relative_humidity_pc']), axis=1)
            df_raw['Dewpoint'] = df_raw.apply(lambda r: hitung_dewpoint(r['temp_drybulb_c_tttttt'], r['relative_humidity_pc']), axis=1)
        
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
            
            # Formats
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
                        df_bulan_ini['sandi_5app'] = df_bulan_ini.apply(ext_appp, axis=1)
                        pivot = df_bulan_ini.pivot_table(index='Tanggal', columns='Jam', values='sandi_5app', aggfunc='first')
                    elif kolom_csv == 'pressure_24h_diff_mb_p24' and 'encoded_synop' in df_bulan_ini.columns:
                        df_bulan_ini['sandi_p24'] = df_bulan_ini.apply(ext_p24, axis=1)
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
        # 2. GENERATOR FORMULIR ME45 BEAUTIFIED + AUTO-FILTER
        # =====================================================================
        buffer_me45 = io.BytesIO()
        
        # Buat Kerangka Dasar Tanggal x Jam (1-31 Bulan Ini x 24 Jam Berurutan)
        me45_rows = []
        station_id = int(df_raw['station_name'].iloc[0]) if ('station_name' in df_raw.columns and str(df_raw['station_name'].iloc[0]).isdigit()) else 97340
        station_label = "STASIUN METEOROLOGI"
        
        for d in range(1, jml_hari + 1):
            for h in range(24):
                me45_rows.append({'NoSta': station_id, 'Station': station_label, 'YY': tahun_val, 'MM': bulan_val, 'DD': d, 'HH': h})
        df_me45_final = pd.DataFrame(me45_rows)
        
        # Gabungkan data mentah dengan kerangka dasar waktu
        df_merge = pd.merge(df_me45_final, df_bulan_ini, left_on=['DD', 'HH'], right_on=['Tanggal', 'Jam'], how='left')
        
        # --- PROSES VECTORIZED MAPPING MAKSIMAL KE KOLOM ME45 ---
        df_out = pd.DataFrame()
        df_out['NoSta'] = df_merge['NoSta']
        df_out['Station'] = df_merge['Station']
        df_out['YY'] = df_merge['YY']
        df_out['MM'] = df_merge['MM']
        df_out['DD'] = df_merge['DD']
        df_out['HH'] = df_merge['HH']
        
        # Parameter Int x10 & Direct Mapping
        df_out['TdTdTd'] = (df_merge['temp_dewpoint_c_tdtdtd'] * 10).round().fillna(np.nan)
        df_out['N'] = df_merge['cloud_cover_oktas_m'].fillna(np.nan)
        df_out['dd'] = df_merge['wind_dir_deg_dd'].fillna(np.nan)
        df_out['ff'] = df_merge['wind_speed_ff'].fillna(np.nan)
        df_out['VV'] = df_merge['visibility_vv'].fillna(np.nan)
        df_out['ww'] = df_merge['present_weather_ww'].fillna(np.nan)
        df_out['W1'] = df_merge['past_weather_w1'].fillna(np.nan)
        df_out['W2'] = df_merge['past_weather_w2'].fillna(np.nan)
        df_out['QFF'] = (df_merge['pressure_qff_mb_derived'] * 10).round().fillna(np.nan) % 10000
        df_out['TtTtTt'] = (df_merge['temp_drybulb_c_tttttt'] * 10).round().fillna(np.nan)
        df_out['Nh'] = df_merge['cloud_low_cover_oktas'].fillna(np.nan)
        df_out['CL'] = df_merge['cloud_low_type_cl'].fillna(np.nan)
        df_out['h'] = df_merge['cloud_low_base_1'].fillna(np.nan)  # Dasar awan jika ada
        df_out['CM'] = df_merge['cloud_med_type_cm'].fillna(np.nan)
        df_out['CH'] = df_merge['cloud_high_type_ch'].fillna(np.nan)
        
        # Kolom kosong untuk menyesuaikan struktur template standard 
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
        
        # Kolom kosong pelengkap sebelum grup sandi tekanan
        for blank_c2 in ['DL', 'DM', 'DH']:
            df_out[blank_c2] = np.nan
            
        # Pemuatan Sandi Tekanan Khusus yang di-extract cerdas dari text SYNOP
        df_out['appp'] = df_merge.apply(ext_appp, axis=1)
        df_out['P24P24P24'] = df_merge.apply(ext_p24, axis=1)
        
        # Indikator Tambahan Akhir
        df_out['iW'] = df_merge['wind_indicator_iw'].fillna(np.nan)
        df_out['iX'] = df_merge['weather_indicator_ix'].fillna(np.nan)
        df_out['iR'] = df_merge['rainfall_indicator_ir'].fillna(np.nan)
        df_out['iE'] = df_merge['evaporation_eq_indicator_ie'].fillna(np.nan)

        # SIMPAN DAN HIAS DESIGN SHEET EXCEL ME45
        with pd.ExcelWriter(buffer_me45, engine='xlsxwriter') as writer_me45:
            df_out.to_excel(writer_me45, sheet_name='Sheet1', index=False)
            
            wb_m = writer_me45.book
            ws_m = writer_me45.sheets['Sheet1']
            
            # Pengaturan Gridline Excel agar Tetap Terlihat Secara Default
            ws_m.hide_gridlines(2)
            
            # --- PALET DECORATION FORMATS ---
            # Header: Deep Navy Background, Bold White Text
            fmt_header_me45 = wb_m.add_format({
                'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#1F4E78',
                'border': 1, 'border_color': '#D9D9D9', 'align': 'center', 'valign': 'vcenter'
            })
            
            # Isi Data Standar: Rata Tengah dengan Garis Batas Abu-abu Tipis
            fmt_data_me45 = wb_m.add_format({
                'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#E0E0E0'
            })
            
            # Format Angka Ber-Leading Zero Proteksi (3 & 4 Digit)
            fmt_3d = wb_m.add_format({'num_format': '000', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#E0E0E0'})
            fmt_4d = wb_m.add_format({'num_format': '0000', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#E0E0E0'})
            
            # Tulis ulang Header dengan Format Baru yang Elegan
            for col_num, value in enumerate(df_out.columns.values):
                ws_m.write(0, col_num, value, fmt_header_me45)
                
            # Set Lebar Kolom Otomatis & Pasang Format Standar Seluruh Baris
            ws_m.set_row(0, 26)  # Mengatur row header sedikit lebih tinggi
            ws_m.set_column(0, len(df_out.columns) - 1, 9, fmt_data_me45)
            
            # Tempelkan Format leading zero di kolom spesifik sandi tekanan 3 jam & 24 jam
            idx_qff = df_out.columns.get_loc('QFF')
            idx_qfe = df_out.columns.get_loc('QFE')
            idx_appp = df_out.columns.get_loc('appp')
            idx_p24 = df_out.columns.get_loc('P24P24P24')
            
            ws_m.set_column(idx_qff, idx_qff, 9, fmt_4d)
            ws_m.set_column(idx_qfe, idx_qfe, 9, fmt_4d)
            ws_m.set_column(idx_appp, idx_appp, 10, fmt_4d)
            ws_m.set_column(idx_p24, idx_p24, 11, fmt_3d)
            
            # --- AKTIFKAN FITUR UTAMA FILTER OTOMATIS EXCEL ---
            ws_m.autofilter(0, 0, len(df_out), len(df_out.columns) - 1)

        # =====================================================================
        # ANTARMUKA TOMBOL UNDUH TERPISAH
        # =====================================================================
        st.success("🎉 Sukses! Koding diefisiensikan & Berkas ME45 Siap Diunduh.")
        
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label=f"📥 1. Unduh Matriks Klimat Rata-Rata ({bulan_dipilih})",
                data=buffer_matriks.getvalue(),
                file_name=f"Matriks_Klimat_{bulan_dipilih}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col2:
            st.download_button(
                label=f"📥 2. Unduh Berkas ME45 Beautified + Filter ({bulan_dipilih})",
                data=buffer_me45.getvalue(),
                file_name=f"ME45_Standard_{bulan_dipilih}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses berkas: {e}")
