import os
import sys
import struct
import gc
import io
import tempfile
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# --- KONFIGURASI HALAMAN STREAMLIT ---
st.set_page_config(
    page_title="Pengolah Data ALOPTAMA Pro",
    page_icon="🌤️",
    layout="wide"
)

# --- ENGINE OLAH DATA & HELPER ---
class EnginePerapihData:
    @staticmethod
    def smart_parse_datetime(date_series, time_series):
        combined = date_series.astype(str).str.strip() + ' ' + time_series.astype(str).str.strip()
        return pd.to_datetime(combined, format='mixed', dayfirst=True, errors='coerce')

    @classmethod
    def normalize_dataframe(cls, df):
        if df is None or df.empty:
            return df
        if 'Date' not in df.columns or 'Time' not in df.columns:
            return df

        valid_mask = (
            df['Date'].notna() & 
            df['Date'].astype(str).str.strip().ne('') & 
            (df['Date'].astype(str).str.strip().str.lower() != 'nan')
        )
        df_clean = df[valid_mask].copy()

        df_clean['datetime_temp'] = cls.smart_parse_datetime(df_clean['Date'], df_clean['Time'])
        df_clean.dropna(subset=['datetime_temp'], inplace=True)

        # Deduplikasi & Urutkan Berdasarkan Waktu
        df_clean.drop_duplicates(subset=['datetime_temp'], inplace=True)
        df_clean.set_index('datetime_temp', inplace=True)
        df_clean.sort_index(inplace=True)

        # Resample Kontinu 1-Menit (Time Series Regularity)
        if not df_clean.empty:
            df_clean = df_clean.resample('1min').asfreq()
        
        df_clean.reset_index(inplace=True)
        df_clean['Date'] = df_clean['datetime_temp'].dt.strftime('%Y-%m-%d')
        df_clean['Time'] = df_clean['datetime_temp'].dt.strftime('%H:%M:00')

        # Type Casting Numerik
        cols_to_exclude = ['Date', 'Time', 'datetime_temp', 'Date_Time_Raw', 'S']
        for col in df_clean.columns:
            if col not in cols_to_exclude:
                df_clean[col] = df_clean[col].astype(str).replace(r'^/+$', np.nan, regex=True)
                df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce')

        # Quality Control (QC) Range Check Fisik
        for col in ['T', 'Air Tmp (C) 33']:
            if col in df_clean.columns:
                df_clean.loc[(df_clean[col] < -50) | (df_clean[col] > 60), col] = np.nan
        for col in ['RH', 'RH (%) 33']:
            if col in df_clean.columns:
                df_clean.loc[(df_clean[col] < 0) | (df_clean[col] > 100), col] = np.nan
        for col in ['DD', 'Mag WD (deg) 33']:
            if col in df_clean.columns:
                df_clean.loc[(df_clean[col] < 0) | (df_clean[col] > 360), col] = np.nan
        for col in ['FF', 'WS (kt) 33']:
            if col in df_clean.columns:
                df_clean.loc[(df_clean[col] < 0) | (df_clean[col] > 150), col] = np.nan
        for p_col in ['STAP', 'MSLP', 'QFE (hPa) 33', 'QNH (hPa) 33']:
            if p_col in df_clean.columns:
                df_clean.loc[(df_clean[p_col] < 600) | (df_clean[p_col] > 1150), p_col] = np.nan

        # Kalkulasi Titik Embun (DP) Otomatis
        t_col = 'T' if 'T' in df_clean.columns else ('Air Tmp (C) 33' if 'Air Tmp (C) 33' in df_clean.columns else None)
        rh_col = 'RH' if 'RH' in df_clean.columns else ('RH (%) 33' if 'RH (%) 33' in df_clean.columns else None)
        dp_col = 'DP' if 'DP' in df_clean.columns else ('Dew Pt (C) 33' if 'Dew Pt (C) 33' in df_clean.columns else 'DP')

        if t_col and rh_col:
            valid_trh = df_clean[t_col].notna() & df_clean[rh_col].notna()
            a, b = 17.27, 237.7
            alpha = ((a * df_clean[t_col]) / (b + df_clean[t_col])) + np.log(df_clean[rh_col] / 100.0)
            dp_calc = (b * alpha) / (a - alpha)
            dp_calc = dp_calc.round(1)
            
            if dp_col not in df_clean.columns:
                df_clean[dp_col] = np.nan
                
            missing_dp = df_clean[dp_col].isna()
            df_clean.loc[missing_dp & valid_trh, dp_col] = dp_calc[missing_dp & valid_trh]

        df_clean.drop(columns=['datetime_temp'], inplace=True, errors='ignore')
        return df_clean

    @staticmethod
    def parse_text_lines(lines, kolom_aws_resmi):
        if not lines: return []
        data_rows = []
        start_idx = 0
        first_line = lines[0].strip()
        tokens_first = first_line.split()
        if tokens_first:
            t0 = tokens_first[0].lower()
            if 'date' in t0 or 'tgl' in t0 or 'tanggal' in t0 or 'waktu' in t0 or not any(c.isdigit() for c in t0):
                start_idx = 1

        for baris in lines[start_idx:]:
            tokens = baris.strip().split()
            if not tokens: continue
            
            if len(tokens) > 1 and ':' in tokens[1]:
                date_val, time_val = tokens[0], tokens[1]
                remaining_tokens = tokens[2:]
            else:
                date_val, time_val = tokens[0], '00:00:00'
                remaining_tokens = tokens[1:]
            
            while len(remaining_tokens) < 22:
                remaining_tokens.append('')
            data_rows.append([date_val, time_val] + remaining_tokens[:22])
            
        return data_rows

    @staticmethod
    def buat_rangkuman_per_sensor(df, sensor_labels):
        TOTAL_HARUSNYA_PER_HARI = 1440
        list_sensor = list(sensor_labels.keys())
        
        valid_df = df[df['Date'].notna() & (df['Date'] != '')].copy()
        if valid_df.empty:
            cols = ['Date', 'Total Menit Log'] + [f'{label} (%)' for label in sensor_labels.values()] + ['PERSENTASE TOTAL KESELURUHAN (%)']
            return pd.DataFrame(columns=cols)

        grouped = valid_df.groupby('Date', sort=False)
        summary_data = []
        
        for date, group in grouped:
            row_summary = {'Date': date, 'Total Menit Log': int(len(group))}
            total_valid_all_sensors = 0
            total_possible_all_sensors = len(list_sensor) * TOTAL_HARUSNYA_PER_HARI
            
            for sensor in list_sensor:
                label = sensor_labels[sensor]
                valid_count = int(group[sensor].notna().sum()) if sensor in group.columns else 0
                percentage = min((valid_count / TOTAL_HARUSNYA_PER_HARI) * 100, 100.0)
                row_summary[f'{label} (%)'] = float(round(percentage, 2))
                total_valid_all_sensors += valid_count
            
            overall_percentage = min((total_valid_all_sensors / total_possible_all_sensors) * 100, 100.0)
            row_summary['PERSENTASE TOTAL KESELURUHAN (%)'] = float(round(overall_percentage, 2))
            summary_data.append(row_summary)
            
        summary_df = pd.DataFrame(summary_data)
        summary_df['date_temp'] = pd.to_datetime(summary_df['Date'], errors='coerce')
        summary_df.sort_values('date_temp', inplace=True)
        summary_df.drop(columns=['date_temp'], inplace=True, errors='ignore')
        
        if not summary_df.empty:
            avg_row = {
                'Date': 'RATA-RATA & TOTAL BULANAN', 
                'Total Menit Log': int(summary_df['Total Menit Log'].sum())
            }
            for col in summary_df.columns:
                if col not in ['Date', 'Total Menit Log']:
                    avg_row[col] = float(round(summary_df[col].mean(), 2))
            summary_df = pd.concat([summary_df, pd.DataFrame([avg_row])], ignore_index=True)
            
        return summary_df

    @staticmethod
    def simpan_ke_excel_bytes(df_data, df_summary, data_sheet_name):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_data.to_excel(writer, index=False, sheet_name=data_sheet_name)
            df_summary.to_excel(writer, index=False, sheet_name="Rangkuman_Ketersediaan")
            
            font_header = Font(name='Arial', size=11, bold=True, color='FFFFFF')
            fill_header = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
            align_header = Alignment(horizontal='center', vertical='center', wrap_text=True)
            font_data = Font(name='Arial', size=10)
            fill_zebra = PatternFill(start_color='F2F5F9', end_color='F2F5F9', fill_type='solid')
            align_center = Alignment(horizontal='center', vertical='center')
            align_left = Alignment(horizontal='left', vertical='center')
            font_total = Font(name='Arial', size=11, bold=True, color='000000')
            fill_total = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
            side_thin = Side(style='thin', color='D9D9D9')
            border_thin = Border(left=side_thin, right=side_thin, top=side_thin, bottom=side_thin)
            border_total = Border(left=side_thin, right=side_thin, top=side_thin, bottom=Side(style='double', color='000000'))

            ws_data = writer.sheets[data_sheet_name]
            ws_data.freeze_panes = 'A2'
            ws_data.auto_filter.ref = ws_data.dimensions
            
            for cell in ws_data[1]:
                cell.font = font_header
                cell.fill = fill_header
                cell.alignment = align_header
                cell.border = border_thin

            ws_sum = writer.sheets["Rangkuman_Ketersediaan"]
            ws_sum.freeze_panes = 'A2'
            if not df_summary.empty:
                ws_sum.auto_filter.ref = ws_sum.dimensions
                
            max_row_sum = ws_sum.max_row
            max_col_sum = ws_sum.max_column
            
            for r_idx, row in enumerate(ws_sum.iter_rows(min_row=1, max_row=max_row_sum, max_col=max_col_sum), start=1):
                is_total_row = (r_idx == max_row_sum) and (max_row_sum > 1)
                for c_idx, cell in enumerate(row, start=1):
                    if r_idx == 1:
                        cell.font, cell.fill, cell.alignment, cell.border = font_header, fill_header, align_header, border_thin
                    elif is_total_row:
                        cell.font, cell.fill, cell.border = font_total, fill_total, border_total
                        cell.alignment = align_left if c_idx == 1 else align_center
                        if c_idx > 2: cell.number_format = '0.00"%"'
                    else:
                        cell.font, cell.border = font_data, border_thin
                        if r_idx % 2 == 0: cell.fill = fill_zebra
                        cell.alignment = align_left if c_idx == 1 else align_center
                        if c_idx > 2: cell.number_format = '0.00"%"'

        return output.getvalue()

    @classmethod
    def baca_database_fdb(cls, fdb_file, kolom_aws_resmi):
        # Menyimpan berkas sementara untuk koneksi driver Firebird
        with tempfile.NamedTemporaryFile(delete=False, suffix=".fdb") as tmp:
            tmp.write(fdb_file.getbuffer())
            tmp_path = tmp.name

        conn = None
        try:
            try:
                import fdb
                conn = fdb.connect(dsn=tmp_path, user='sysdba', password='masterkey')
            except Exception:
                from firebird.driver import connect
                conn = connect(tmp_path, user='sysdba', password='masterkey')
                
            cursor = conn.cursor()
            cursor.execute("SELECT RDB$RELATION_NAME FROM RDB$RELATIONS WHERE RDB$SYSTEM_FLAG = 0 AND RDB$VIEW_BLR IS NULL")
            all_tables = [row[0].strip().upper() for row in cursor.fetchall()]
            
            priority_order = ['T1MN', 'TBRUT', 'T1H']
            sorted_tables = [t for t in priority_order if t in all_tables] + [t for t in all_tables if t not in priority_order]
            
            extracted_dfs = []
            for target_table in sorted_tables:
                try:
                    cursor.execute(f'SELECT * FROM "{target_table}"')
                    df_tbl = pd.DataFrame(cursor.fetchall(), columns=[f[0] for f in cursor.description])
                    if not df_tbl.empty:
                        # Parsing sederhana
                        df_mapped = pd.DataFrame(index=df_tbl.index)
                        dt_col = next((c for c in df_tbl.columns if c.upper() in ['DATACQ', 'DATETIME', 'DATE_TIME', 'DATE']), None)
                        if dt_col:
                            dt_parsed = pd.to_datetime(df_tbl[dt_col], errors='coerce')
                            df_mapped['Date'] = dt_parsed.dt.strftime('%Y-%m-%d')
                            df_mapped['Time'] = dt_parsed.dt.strftime('%H:%M:00')
                            for c in kolom_aws_resmi:
                                if c not in ['Date', 'Time']:
                                    df_mapped[c] = df_tbl[c] if c in df_tbl.columns else ''
                            extracted_dfs.append(df_mapped[kolom_aws_resmi])
                            if target_table in ['T1MN', 'TBRUT']: break
                except Exception:
                    continue
            
            if extracted_dfs:
                return pd.concat(extracted_dfs, ignore_index=True)
            return pd.DataFrame(columns=kolom_aws_resmi)
        finally:
            if conn: conn.close()
            if os.path.exists(tmp_path): os.remove(tmp_path)


# --- INTERFASE PENGGUNA (STREAMLIT UI) ---
st.title("🌤️ Pengolah Data ALOPTAMA Meteorologi")
st.caption("Aplikasi Rekapitulasikan & Normalisasi Data AWOS & AWS (Power BI Ready) - By Luqmanul Hakim, S.Tr")

tab_awos, tab_fdb, tab_aws, tab_info = st.tabs([
    "1. 📊 Panel AWOS", 
    "2. ⚡ Ekstrak FDB ke CSV", 
    "3. 🛰️ Panel AWS", 
    "4. ℹ️ Informasi Sistem"
])

# --- TAB 1: AWOS ---
with tab_awos:
    st.subheader("Pengolahan Data AWOS (.CSV / .XLSX)")
    awos_labels = {
        'Air Tmp (C) 33': 'Suhu Udara', 'Dew Pt (C) 33': 'Titik Embun',
        'RH (%) 33': 'Kelembaban', 'QFE (hPa) 33': 'Tekanan QFE',
        'QNH (hPa) 33': 'Tekanan QNH', 'WS (kt) 33': 'Kecepatan Angin',
        'Mag WD (deg) 33': 'Arah Angin', 'Solar Rad (W/m^2) 33': 'Radiasi Matahari',
        'Precip 1Hr (mm) 33': 'Curah Hujan'
    }
    
    excel_lama_awos = st.file_uploader(" (Opsional) Gabungkan dengan File Excel AWOS Lama", type=['xlsx'], key="awos_old")
    files_awos = st.file_uploader("Upload File Mentah AWOS Tambahan / Baru", type=['csv', 'xlsx', 'xls'], accept_multiple_files=True, key="awos_new")
    
    if st.button("Proses Data AWOS 🚀", key="btn_awos"):
        if not files_awos and not excel_lama_awos:
            st.error("Silakan upload minimal satu file AWOS!")
        else:
            with st.spinner("Memproses, Menyelaraskan Waktu, & Melakukan Quality Control AWOS..."):
                list_df = []
                if excel_lama_awos:
                    list_df.append(pd.read_excel(excel_lama_awos, sheet_name="Data_AWOS_Clean"))
                
                for fp in files_awos:
                    if fp.name.endswith(('.xlsx', '.xls')):
                        df = pd.read_excel(fp)
                        mapping_eg = {
                            'Date and Time': 'Date_Time_Raw', 'airTemperatureValidated.RWYA': 'Air Tmp (C) 33',
                            'dewPointValidated.RWYA': 'Dew Pt (C) 33', 'relativeHumidityValidated.RWYA': 'RH (%) 33',
                            'QFE.RWYA': 'QFE (hPa) 33', 'QNH.RWYA': 'QNH (hPa) 33', 'windSpeedSelected.RWYA': 'WS (kt) 33',
                            'instant.windDirection.RWYA': 'Mag WD (deg) 33', 'solarRadiation.RWYA': 'Solar Rad (W/m^2) 33',
                            'oneHour.precipSelected.RWYA': 'Precip 1Hr (mm) 33'
                        }
                        df.rename(columns=mapping_eg, inplace=True)
                        if 'Date_Time_Raw' in df.columns:
                            split_dt = df['Date_Time_Raw'].astype(str).str.split(r'\s+', n=1, expand=True)
                            df.insert(0, 'Time', split_dt[1].fillna('00:00:00') if split_dt.shape[1] > 1 else '00:00:00')
                            df.insert(0, 'Date', split_dt[0])
                            df.drop(columns=['Date_Time_Raw'], inplace=True)
                        list_df.append(df)
                    elif fp.name.endswith('.csv'):
                        df = pd.read_csv(fp)
                        orig_col = df.columns[0]
                        split_dt = df[orig_col].astype(str).str.split(r'\s+', n=1, expand=True)
                        df.insert(0, 'Time', split_dt[1].fillna('00:00:00') if split_dt.shape[1] > 1 else '00:00:00')
                        df.insert(0, 'Date', split_dt[0])
                        df.drop(columns=[orig_col], inplace=True)
                        list_df.append(df)
                
                df_all = pd.concat(list_df, ignore_index=True)
                df_clean = EnginePerapihData.normalize_dataframe(df_all)
                df_summary = EnginePerapihData.buat_rangkuman_per_sensor(df_clean, awos_labels)
                excel_bytes = EnginePerapihData.simpan_ke_excel_bytes(df_clean, df_summary, "Data_AWOS_Clean")
                
                st.success("✅ Data AWOS Berhasil Diproses!")
                st.download_button(
                    label="⬇️ Download Excel AWOS Clean",
                    data=excel_bytes,
                    file_name="AWOS_Gabungan_Clean.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.dataframe(df_clean.head(100), use_container_width=True)

# --- TAB 2: FDB TO CSV ---
with tab_fdb:
    st.subheader("Ekstrak Database Firebird AWS (.FDB) Kilat ke CSV")
    fdb_file = st.file_uploader("Upload File Database Firebird (.FDB)", type=['fdb'], key="fdb_single")
    
    if st.button("⚡ Ekstrak Kilat ke CSV", key="btn_fdb"):
        if not fdb_file:
            st.error("Upload file .FDB terlebih dahulu!")
        else:
            with st.spinner("Mengekstrak tabel database FDB..."):
                kolom_aws = ['Date', 'Time', 'S', 'DD', 'FF', 'DM10', 'FM10', 'DD2', 'FF2', 'DVN', 'DVX', 'FVN', 'FVX', 'RR', 'RH', 'TSV', 'DP', 'T', 'GLOR', 'GLORP', 'INSD', 'STAP', 'MSLP', 'GNDT']
                df_fdb = EnginePerapihData.baca_database_fdb(fdb_file, kolom_aws)
                df_clean = EnginePerapihData.normalize_dataframe(df_fdb)
                
                csv_bytes = df_clean.to_csv(index=False).encode('utf-8')
                st.success("⚡ Ekstraksi FDB Selesai!")
                st.download_button(
                    label="⬇️ Download Hasil CSV",
                    data=csv_bytes,
                    file_name="AWS_Ekstrak_FDB.csv",
                    mime="text/csv"
                )
                st.dataframe(df_clean.head(100), use_container_width=True)

# --- TAB 3: AWS MULTI-FORMAT ---
with tab_aws:
    st.subheader("Pengolahan Data AWS Multi-Format (.XLSX / .CSV / .TXT / .FDB)")
    aws_labels = {
        'DD': 'Arah Angin', 'FF': 'Kecepatan Angin', 'RR': 'Curah Hujan',
        'RH': 'Kelembaban', 'DP': 'Titik Embun', 'T': 'Suhu Udara',
        'STAP': 'Tekanan Stasiun', 'GNDT': 'Suhu Tanah'
    }
    kolom_aws_resmi = ['Date', 'Time', 'S', 'DD', 'FF', 'DM10', 'FM10', 'DD2', 'FF2', 'DVN', 'DVX', 'FVN', 'FVX', 'RR', 'RH', 'TSV', 'DP', 'T', 'GLOR', 'GLORP', 'INSD', 'STAP', 'MSLP', 'GNDT']
    
    excel_lama_aws = st.file_uploader(" (Opsional) Gabungkan dengan File Excel AWS Lama", type=['xlsx'], key="aws_old")
    files_aws = st.file_uploader("Upload File Mentah AWS (Format Text/Excel/CSV/FDB)", type=['xlsx', 'xls', 'csv', 'txt', 'fdb'], accept_multiple_files=True, key="aws_new")
    
    if st.button("Proses Data AWS 🚀", key="btn_aws"):
        if not files_aws and not excel_lama_aws:
            st.error("Silakan upload minimal satu file AWS!")
        else:
            with st.spinner("Memproses & Merekonstruksi Data AWS..."):
                list_df = []
                if excel_lama_aws:
                    list_df.append(pd.read_excel(excel_lama_aws, sheet_name="Data_AWS_Clean"))
                
                for fp in files_aws:
                    ext = fp.name.lower()
                    if ext.endswith('.fdb'):
                        df_fdb = EnginePerapihData.baca_database_fdb(fp, kolom_aws_resmi)
                        if not df_fdb.empty: list_df.append(df_fdb)
                    elif ext.endswith('.csv'):
                        list_df.append(pd.read_csv(fp))
                    elif ext.endswith(('.xlsx', '.xls')):
                        df_ex = pd.read_excel(fp)
                        if len(df_ex.columns) == 1:
                            lines = df_ex.iloc[:, 0].dropna().astype(str).tolist()
                            data_rows = EnginePerapihData.parse_text_lines(lines, kolom_aws_resmi)
                            if data_rows: list_df.append(pd.DataFrame(data_rows, columns=kolom_aws_resmi))
                        else: list_df.append(df_ex)
                    else:
                        lines = fp.getvalue().decode('utf-8', errors='ignore').splitlines()
                        data_rows = EnginePerapihData.parse_text_lines(lines, kolom_aws_resmi)
                        if data_rows: list_df.append(pd.DataFrame(data_rows, columns=kolom_aws_resmi))
                
                df_all = pd.concat(list_df, ignore_index=True)
                df_clean = EnginePerapihData.normalize_dataframe(df_all)
                df_summary = EnginePerapihData.buat_rangkuman_per_sensor(df_clean, aws_labels)
                excel_bytes = EnginePerapihData.simpan_ke_excel_bytes(df_clean, df_summary, "Data_AWS_Clean")
                
                st.success("✅ Data AWS Berhasil Diproses!")
                st.download_button(
                    label="⬇️ Download Excel AWS Clean",
                    data=excel_bytes,
                    file_name="AWS_Gabungan_Clean.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.dataframe(df_clean.head(100), use_container_width=True)

# --- TAB 4: INFO SISTEM ---
with tab_info:
    st.markdown("""
    ### 📌 Fitur Utama Sistem Refactored (Streamlit Version)
    
    1. **Format Tanggal Standar ISO (`YYYY-MM-DD`)**: Mencegah timbulnya `DataFormat.Error` atau konflik *locale* saat diimpor ke **Power BI Desktop**[cite: 1, 2].
    2. **Kontinuitas Waktu Kontinu (1-Min Resampling)**: Memunculkan baris waktu yang hilang (*missing time gaps*) secara teratur agar grafik *time-series* tidak meloncat.
    3. **Quality Control (QC) Fisik**: Otomatis menyaring dan membuang outlier/anomali nilai ekstrem fisik meteorologi.
    4. **In-Memory Download**: Seluruh proses ekspor dikemas dalam memori (`io.BytesIO`) tanpa mengotori folder server web[cite: 2].
    """)
    
    if st.button("♻️ Bersihkan Memori Server (RAM)"):
        freed = gc.collect()
        st.info(f"RAM berhasil dibersihkan! ({freed} objek dibuang).")
