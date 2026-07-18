import streamlit as st
import pandas as pd
import numpy as np
import io
import calendar
import re

st.set_page_config(page_title="Data Extract Job", layout="wide")

st.title("📊 EXCEL PARAMETER RATA-RATA KLIMAT & ME45")

# --- KODE PETUNJUK PENGGUNAAN ---
with st.expander("ℹ️ Klik di sini untuk melihat Petunjuk Penggunaan"):
    st.markdown("""
    **Syarat File CSV:**
    - File hasil extract dari "https://bmkgsatu.bmkg.go.id/exportdata".
    - Mengandung kolom `encoded_synop` dan parameter dasar klimatologi.
    """)
# --------------------------------

# --- RUMUS TEKANAN UAP AIR ---
def hitung_tekanan_uap_excel(suhu, rh):
    if pd.isna(suhu) or pd.isna(rh): return np.nan
    es = 6.112 * np.exp((17.67 * suhu) / (suhu + 243.5))
    e_actual = (rh / 100.0) * es
    return round(e_actual * 10, 2) 

# --- RUMUS DEWPOINT (TD) ---
def hitung_dewpoint(suhu, rh):
    if pd.isna(suhu) or pd.isna(rh): return np.nan
    a = 17.27
    b = 237.7
    alpha = ((a * suhu) / (b + suhu)) + np.log(rh / 100.0)
    td = (b * alpha) / (a - alpha)
    return round(td, 2)

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
        
        # Cleansing nilai sandi hilang
        NA_VALUES = [9999, 99999, '9999', '/', '//', '///', '#REF!', '#VALUE!', 'STNR', '#N/A']
        df_raw.replace(NA_VALUES, np.nan, inplace=True)
        
        # Parsing Waktu
        df_raw['data_timestamp'] = pd.to_datetime(df_raw['data_timestamp'])
        df_raw['Tahun'] = df_raw['data_timestamp'].dt.year
        df_raw['Bulan_Angka'] = df_raw['data_timestamp'].dt.month
        df_raw['Tanggal'] = df_raw['data_timestamp'].dt.day
        df_raw['Jam'] = df_raw['data_timestamp'].dt.hour
        
        # Hitung parameter turunan
        if 'temp_drybulb_c_tttttt' in df_raw.columns and 'relative_humidity_pc' in df_raw.columns:
            df_raw['Tekanan_Uap_x10'] = df_raw.apply(
                lambda row: hitung_tekanan_uap_excel(row['temp_drybulb_c_tttttt'], row['relative_humidity_pc']), axis=1)
            df_raw['Dewpoint'] = df_raw.apply(
                lambda row: hitung_dewpoint(row['temp_drybulb_c_tttttt'], row['relative_humidity_pc']), axis=1)
        
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
        # PEMBUATAN 1: PROSES DOWNLOAD MATRIKS UTAMA
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
            
            def get_cell(val, param, col_type):
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
                        return round(float(val), 1), (fmt_float_rata2 if col_type == 'RATA2' else fmt_float_biasa)
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
                        def to_sandi_3h_4d_int(row):
                            val = row['pressure_3h_diff_mb_ppp']
                            if pd.isna(val): return np.nan
                            synop = row['encoded_synop']
                            if pd.notna(synop):
                                expected_ppp = f"{int(round(abs(val) * 10)):03d}"
                                sec1 = synop.split('333')[0]
                                tokens = re.findall(r'\b5\d{4}\b', sec1)
                                for tk in tokens:
                                    if tk.endswith(expected_ppp): return int(tk[1:])
                                for a in range(10):
                                    if f"5{a}{expected_ppp}" in sec1: return int(f"{a}{expected_ppp}")
                            return int(f"{'3' if val >= 0 else '8'}{int(round(abs(val) * 10)):03d}")
                        
                        df_bulan_ini['sandi_5app'] = df_bulan_ini.apply(to_sandi_3h_4d_int, axis=1)
                        pivot = df_bulan_ini.pivot_table(index='Tanggal', columns='Jam', values='sandi_5app', aggfunc='first')
                    elif kolom_csv == 'pressure_24h_diff_mb_p24' and 'encoded_synop' in df_bulan_ini.columns:
                        def to_sandi_24h_3d_int(row):
                            val = row['pressure_24h_diff_mb_p24']
                            if pd.isna(val): return np.nan
                            synop = row['encoded_synop']
                            if pd.notna(synop):
                                expected_p24 = f"{int(round(abs(val) * 10)):03d}"
                                if '333' in synop:
                                    sec3 = synop.split('333')[1]
                                    tokens = re.findall(r'\b5[89]\d{3}\b', sec3)
                                    for tk in tokens:
                                        if tk.endswith(expected_p24):
                                            return int(tk[2:]) if tk.startswith('58') else 500 + int(tk[2:])
                                    for prfx in ['58', '59']:
                                        if f"{prfx}{expected_p24}" in sec3:
                                            return int(expected_p24) if prfx == '58' else 500 + int(expected_p24)
                            return int(expected_p24) if val >= 0 else 500 + int(expected_p24)
                        
                        df_bulan_ini['sandi_p24'] = df_bulan_ini.apply(to_sandi_24h_3d_int, axis=1)
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
                            val, fmt = get_cell(pivot.loc[tgl, h], judul_param, '0-23')
                            ws.write(row_idx, h + 1, val, fmt)
                        val_rata, fmt_rata = get_cell(rata_harian.loc[tgl], judul_param, 'RATA2')
                        ws.write(row_idx, 25, val_rata, fmt_rata)
                        
                        if 'UAP AIR' in judul_param: pass
                        elif 'KECEPATAN ANGIN' in judul_param:
                            v_max = daily_max_angin.loc[tgl]
                            if pd.isna(v_max) or float(v_max) == 0: ws.write(row_idx, 26, "", fmt_blank)
                            else: ws.write(row_idx, 26, int(round(float(v_max))), fmt_int_biasa)
                        else:
                            for c_idx, h_spec in zip([26, 27, 28], [23, 5, 10]):
                                val_s, fmt_s = get_cell(pivot.loc[tgl, h_spec], judul_param, 'SPEC')
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
        # PEMBUATAN 2: GENERATOR FORMULIR ME45 TERPISAH
        # =====================================================================
        buffer_me45 = io.BytesIO()
        
        # Buat kombinasi baris lengkap Tanggal x Jam (1 s/d 31 bulan berjalan x 24 jam)
        me45_rows = []
        station_id = df_bulan_ini['station_name'].iloc[0] if 'station_name' in df_bulan_ini.columns else "97340"
        
        for d in range(1, jml_hari + 1):
            for h in range(24):
                me45_rows.append({
                    'NoSta': 97340,
                    'Station': station_id,
                    'YY': tahun_val,
                    'MM': bulan_val,
                    'DD': d,
                    'HH': h
                })
        df_me45_base = pd.DataFrame(me45_rows)
        
        # Merge data riil GTS ke dalam kerangka waktu dasar ME45
        df_merge = pd.merge(df_me45_base, df_bulan_ini, left_on=['DD', 'HH'], right_on=['Tanggal', 'Jam'], how='left')
        
        # Konversi kalkulasi skala BMKG (Kali 10)
        df_merge['TtTtTt'] = df_merge['temp_drybulb_c_tttttt'].apply(lambda x: int(round(x * 10)) if pd.notna(x) else np.nan)
        df_merge['TdTdTd'] = df_merge['temp_dewpoint_c_tdtdtd'].apply(lambda x: int(round(x * 10)) if pd.notna(x) else np.nan)
        df_merge['UU'] = df_merge['relative_humidity_pc'].apply(lambda x: int(round(x)) if pd.notna(x) else np.nan)
        
        # Skala Tekanan QFE & QFF (Modulus 10000 atau konversi standard)
        df_merge['QFE'] = df_merge['pressure_qfe_mb_derived'].apply(lambda x: int(round(x * 10)) % 10000 if pd.notna(x) else np.nan)
        df_merge['QFF'] = df_merge['pressure_qff_mb_derived'].apply(lambda x: int(round(x * 10)) % 10000 if pd.notna(x) else np.nan)
        
        # Arah dan Kecepatan Angin
        df_merge['dd'] = df_merge['wind_dir_deg_dd'].apply(lambda x: int(round(x)) if pd.notna(x) else np.nan)
        df_merge['ff'] = df_merge['wind_speed_ff'].apply(lambda x: int(round(x)) if pd.notna(x) else np.nan)
        
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
            return int(f"{'3' if v >= 0 else '8'}{int(round(abs(v) * 10)):03d}")
        df_merge['appp'] = df_merge.apply(ext_appp, axis=1)
        
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
            return int(round(abs(v) * 10)) if v >= 0 else 500 + int(round(abs(v) * 10))
        df_merge['P24P24P24'] = df_merge.apply(ext_p24, axis=1)

        # Definisikan semua kolom standard ME45 sesuai file acuan template
        me45_columns = [
            'NoSta', 'Station', 'YY', 'MM', 'DD', 'HH', 'TdTdTd', 'N', 'dd', 'ff', 'VV', 'ww', 'W1', 'W2', 
            'QFF', 'TtTtTt', 'Nh', 'CL', 'h', 'CM', 'CH', 'Ns', 'C', 'hshs', 'Ns.1', 'C.1', 'hshs.1', 0, 
            'C.2', 'hshs.2', 'C.3', 'D', 'e', 'UU', 'QFE', 'TwTwTw', 'RRR', 'tR', 'TxTxTx', 'TnTnTn', 
            'EEE', 'F24F24F24F24', 'SSS', 'E', 'DL', 'DM', 'DH', 'appp', 'P24P24P24', 'iW', 'iX', 'iR', 'iE'
        ]
        
        # Pastikan kolom opsional lain yang tidak terisi tetap dibuat kosong agar struktur pas
        for col in me45_columns:
            if col not in df_merge.columns:
                df_merge[col] = np.nan
                
        df_me45_final = df_merge[me45_columns]
        
        with pd.ExcelWriter(buffer_me45, engine='xlsxwriter') as writer_me45:
            df_me45_final.to_excel(writer_me45, sheet_name='Sheet1', index=False)
            
            # Berikan Number Formatting otomatis agar kode sandi 000, 007, 7002 tampil sempurna
            wb_m = writer_me45.book
            ws_m = writer_me45.sheets['Sheet1']
            fmt_3dig = wb_m.add_format({'num_format': '000', 'align': 'center'})
            fmt_4dig = wb_m.add_format({'num_format': '0000', 'align': 'center'})
            
            # Format kolom appp (Kolom AV / indeks ke-47) dan P24P24P24 (Kolom AW / indeks ke-48)
            ws_m.set_column(47, 47, 10, fmt_4dig)
            ws_m.set_column(48, 48, 12, fmt_3dig)

        # =====================================================================
        # INTERFACE TOMBOL DOWNLOAD TERPISAH
        # =====================================================================
        st.success("✅ Seluruh data bulan ini berhasil diproses!")
        
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label=f"📥 1. Unduh Matriks Rata-Rata Klimat ({bulan_dipilih})",
                data=buffer_matriks.getvalue(),
                file_name=f"Matriks_{bulan_dipilih}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
        with col2:
            st.download_button(
                label=f"📥 2. Unduh Formulir ME45 Standard ({bulan_dipilih})",
                data=buffer_me45.getvalue(),
                file_name=f"ME45_{bulan_dipilih}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
