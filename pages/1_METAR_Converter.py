import streamlit as st
import pandas as pd
import numpy as np
import io
import calendar

st.set_page_config(page_title="Data Extract Job", layout="wide")

st.title("📊 EXCEL PARAMETER RATA-RATA KLIMAT")

# --- RUMUS TEKANAN UAP AIR ---
def hitung_tekanan_uap_excel(suhu, rh):
    if pd.isna(suhu) or pd.isna(rh): return np.nan
    es = 6.112 * np.exp((17.67 * suhu) / (suhu + 243.5))
    e_actual = (rh / 100.0) * es
    return round(e_actual * 10, 2) 

# --- RUMUS DEWPOINT (TD) ---
def hitung_dewpoint(suhu, rh):
    if pd.isna(suhu) or pd.isna(rh): return np.nan
    # Menggunakan konstanta Magnus-Tetens
    a = 17.27
    b = 237.7
    alpha = ((a * suhu) / (b + suhu)) + np.log(rh / 100.0)
    td = (b * alpha) / (a - alpha)
    return round(td, 2)

# --- PARAMETER MAPPING ---
parameter_mapping = {
    'pressure_qff_mb_derived': 'QFF RATA-2 HARIAN',
    'pressure_qfe_mb_derived': 'QFE RATA-2 HARIAN',
    'temp_drybulb_c_tttttt': 'SUHU UDARA RATA-2 HARIAN',
    'Dewpoint': 'SUHU TITIK EMBUN (TD) RATA-2 HARIAN',  # <-- PARAMETER DEWPOINT DITAMBAHKAN
    'relative_humidity_pc': 'KELEMBABAN UDARA RATA-2 HARIAN',
    'wind_speed_ff': 'KECEPATAN ANGIN RATA-2 HARIAN',
    'Tekanan_Uap_x10': 'TEKANAN UAP AIR RATA-2 HARIAN'
}

uploaded_file = st.file_uploader("Unggah file...", type=["csv"])

if uploaded_file is not None:
    try:
        df_raw = pd.read_csv(uploaded_file)
        
        # Cleansing
        NA_VALUES = [9999, 99999, '9999', '/', '//', '///', '#REF!', '#VALUE!', 'STNR', '#N/A']
        df_raw.replace(NA_VALUES, np.nan, inplace=True)
        
        # Parsing Waktu
        df_raw['data_timestamp'] = pd.to_datetime(df_raw['data_timestamp'])
        df_raw['Tahun'] = df_raw['data_timestamp'].dt.year
        df_raw['Bulan_Angka'] = df_raw['data_timestamp'].dt.month
        df_raw['Tanggal'] = df_raw['data_timestamp'].dt.day
        df_raw['Jam'] = df_raw['data_timestamp'].dt.hour
        
        if 'temp_drybulb_c_tttttt' in df_raw.columns and 'relative_humidity_pc' in df_raw.columns:
            # Hitung Tekanan Uap
            df_raw['Tekanan_Uap_x10'] = df_raw.apply(
                lambda row: hitung_tekanan_uap_excel(row['temp_drybulb_c_tttttt'], row['relative_humidity_pc']), axis=1)
            
            # Hitung Dewpoint
            df_raw['Dewpoint'] = df_raw.apply(
                lambda row: hitung_dewpoint(row['temp_drybulb_c_tttttt'], row['relative_humidity_pc']), axis=1)
        
        df_raw['Bulan_Tahun'] = df_raw['Tahun'].astype(str) + "-" + df_raw['Bulan_Angka'].astype(str).str.zfill(2)
        
        st.markdown("---")
        bulan_dipilih = st.selectbox("Pilih Bulan untuk di-Generate:", sorted(df_raw['Bulan_Tahun'].unique()))
        df_bulan_ini = df_raw[df_raw['Bulan_Tahun'] == bulan_dipilih]
        
        tahun_val = int(bulan_dipilih.split('-')[0])
        bulan_val = int(bulan_dipilih.split('-')[1])
        nama_bulan = calendar.month_name[bulan_val].upper()
        jml_hari = calendar.monthrange(tahun_val, bulan_val)[1]
        semua_tanggal = pd.Index(range(1, jml_hari + 1), name='NO.')

        # =====================================================================
        # PROSES EXCEL DENGAN NUMBER FORMATTING
        # =====================================================================
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            wb = writer.book
            ws = wb.add_worksheet('MATRIKS')
            
            # Format Cell
            fmt_teks = wb.add_format({'bold': True, 'align': 'left'})
            fmt_judul = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 12})
            fmt_header = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9D9D9'})
            fmt_blank = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
            fmt_blank_rata2 = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#DCE6F1'})
            fmt_qff_biasa = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0000'}) 
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
                if pd.isna(val): 
                    return "", (fmt_blank_rata2 if col_type == 'RATA2' else fmt_blank)
                
                # KECEPATAN ANGIN
                if 'KECEPATAN ANGIN' in param:
                    v = int(round(float(val)))
                    if v == 0:
                        return "", (fmt_blank_rata2 if col_type == 'RATA2' else fmt_blank)
                    return v, (fmt_int_rata2 if col_type == 'RATA2' else fmt_int_biasa)
                
                # QFF / QFE / SUHU UDARA / DEWPOINT (0-23)
                if 'QFF' in param or 'QFE' in param or 'SUHU' in param:
                    if col_type == '0-23':
                        if 'SUHU' in param: # Ini juga mencakup "SUHU TITIK EMBUN (TD)"
                            return int(round(float(val) * 10)), fmt_int_biasa
                        
                        return (int(round(float(val) * 10)) % 10000), fmt_qff_biasa
                    else:
                        return round(float(val), 1), (fmt_float_rata2 if col_type == 'RATA2' else fmt_float_biasa)
                
                # RH / UAP AIR
                if 'KELEMBABAN' in param or 'UAP AIR' in param:
                    if col_type == 'RATA2': 
                        return round(float(val), 1), fmt_float_rata2
                    return int(round(float(val))), fmt_int_biasa
                
                return round(float(val), 1), (fmt_float_rata2 if col_type == 'RATA2' else fmt_float_biasa)

            # LOOOPING PARAMETER
            for kolom_csv, judul_param in parameter_mapping.items():
                if kolom_csv in df_bulan_ini.columns:
                    
                    df_bulan_ini.loc[:, kolom_csv] = pd.to_numeric(df_bulan_ini[kolom_csv], errors='coerce')
                    pivot = df_bulan_ini.pivot_table(index='Tanggal', columns='Jam', values=kolom_csv, aggfunc='first')
                    
                    for h in range(24):
                        if h not in pivot.columns: pivot[h] = np.nan
                    pivot = pivot[list(range(24))]
                    pivot = pivot.reindex(semua_tanggal)
                    
                    rata_harian = pivot.mean(axis=1)
                    if 'UAP AIR' in judul_param: rata_harian = rata_harian / 10

                    if 'UAP AIR' in judul_param:
                        extra_headers = []
                        summary_labels = []
                    elif 'KECEPATAN ANGIN' in judul_param:
                        extra_headers = ['MAX HARIAN']
                        daily_max_angin = pivot.max(axis=1)
                        summary_labels = [("MAXIMUM BULAN INI", pivot.max().max())]
                    else:
                        extra_headers = ['23 00', '05 00', '10 00']
                        summary_labels = [
                            ("MAXIMUM BULAN INI", pivot.max().max()), 
                            ("MINIMUM BULAN INI", pivot.min().min()), 
                            ("TOTAL RATA-RATA", rata_harian.mean())
                        ]
                    
                    # Tulis Header
                    ws.write(start_row, 1, "BULAN", fmt_teks)
                    ws.write(start_row, 3, nama_bulan, fmt_teks)
                    ws.write(start_row, 4, str(tahun_val), fmt_teks)
                    ws.write(start_row + 1, 12, judul_param, fmt_judul)
                    
                    ws.write(start_row + 2, 0, "NO.", fmt_header)
                    for i in range(24): ws.write(start_row + 2, i + 1, str(i), fmt_header)
                    ws.write(start_row + 2, 25, "R A T A   2", fmt_header)
                    
                    for c_i, ext_hdr in enumerate(extra_headers):
                        ws.write(start_row + 2, 26 + c_i, ext_hdr, fmt_header)
                    
                    row_idx = start_row + 3
                    
                    # TULIS DATA TANGGAL 1-31
                    for tgl in semua_tanggal:
                        ws.write(row_idx, 0, tgl, fmt_header)
                        
                        # Data 0-23
                        for h in range(24):
                            val, fmt = get_cell(pivot.loc[tgl, h], judul_param, '0-23')
                            ws.write(row_idx, h + 1, val, fmt)
                        
                        # Data RATA 2
                        val_rata, fmt_rata = get_cell(rata_harian.loc[tgl], judul_param, 'RATA2')
                        ws.write(row_idx, 25, val_rata, fmt_rata)
                        
                        # Data Tambahan Kanan
                        if 'UAP AIR' in judul_param:
                            pass
                        elif 'KECEPATAN ANGIN' in judul_param:
                            v_max = daily_max_angin.loc[tgl]
                            if pd.isna(v_max) or float(v_max) == 0:
                                ws.write(row_idx, 26, "", fmt_blank)
                            else:
                                ws.write(row_idx, 26, int(round(float(v_max))), fmt_int_biasa) # NUMBER
                        else:
                            for c_idx, h_spec in zip([26, 27, 28], [23, 5, 10]):
                                val_s, fmt_s = get_cell(pivot.loc[tgl, h_spec], judul_param, 'SPEC')
                                ws.write(row_idx, c_idx, val_s, fmt_s)
                        
                        row_idx += 1
                        
                    # TULIS SUMMARY BAWAH
                    for label, final_val in summary_labels:
                        ws.merge_range(row_idx, 0, row_idx, 24, label, fmt_summary_judul)
                        
                        if pd.isna(final_val):
                            ws.write(row_idx, 25, "", fmt_summary_final_blank)
                        else:
                            if 'KECEPATAN ANGIN' in judul_param:
                                v_ang = int(round(float(final_val)))
                                if v_ang == 0:
                                    ws.write(row_idx, 25, "", fmt_summary_final_blank)
                                else:
                                    ws.write(row_idx, 25, v_ang, fmt_summary_final_int) # NUMBER
                            else:
                                ws.write(row_idx, 25, round(float(final_val), 1), fmt_summary_final_float) # NUMBER
                        
                        for c_i in range(len(extra_headers)):
                            ws.write(row_idx, 26 + c_i, "", fmt_summary_kosong)
                        
                        row_idx += 1
                        
                    start_row = row_idx + 3 

        st.success("✅ Sukses!")
        st.download_button(
            label=f"📥 Unduh... ({bulan_dipilih})",
            data=buffer.getvalue(),
            file_name=f"Matriks_{bulan_dipilih}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
