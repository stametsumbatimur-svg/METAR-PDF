import streamlit as st
import math
import random
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.cm as cm
from datetime import datetime

# --- KONFIGURASI HALAMAN STREAMLIT ---
st.set_page_config(
    page_title="Pibal Generator Stamet Waingapu",
    layout="wide",
    initial_sidebar_state="expanded"
)

elevation_waingapu = 32.8
MAX_WIND_SPEED_KT = 25.0  # Batas maksimum ketat kecepatan angin (ws <= 25 kt)

# --- LOAD DAN GABUNGKAN SECOND DATASET HISTORIS ---
@st.cache_data
def load_historical_pibal():
    fn_raw = 'Raw Pibal 2024-05-15 to 2026-08-16.csv'
    fn_data = 'Data Pibal 2024-05-15 to 2026-08-16.csv'
    
    def find_file(filename):
        possible_paths = [
            filename,
            os.path.join(os.path.dirname(__file__), '..', filename),
            os.path.join(os.path.dirname(__file__), filename),
        ]
        for path in possible_paths:
            if os.path.exists(path):
                return path
        return None

    path_raw = find_file(fn_raw)
    path_data = find_file(fn_data)
    
    if not path_raw:
        st.warning(f"⚠️ File raw dataset (`{fn_raw}`) tidak ditemukan.")
        return None, None

    try:
        # 1. BACA & SANITASI RAW PIBAL (Azimuth & Elevasi Per Menit)
        with open(path_raw, 'r') as f:
            first_line = f.readline()
        skip_raw = 1 if 'Raw Pibal' in first_line else 0
        df_raw = pd.read_csv(path_raw, skiprows=skip_raw)
        
        df_raw = df_raw.dropna(subset=['azimuth', 'elevasi']).copy()
        df_raw = df_raw[(df_raw['azimuth'] != 9999) & (df_raw['elevasi'] != 9999)]
        
        # Konversi Tipe Data
        df_raw['pembacaan'] = pd.to_numeric(df_raw['pembacaan'], errors='coerce')
        df_raw['azimuth'] = pd.to_numeric(df_raw['azimuth'], errors='coerce')
        df_raw['elevasi'] = pd.to_numeric(df_raw['elevasi'], errors='coerce')
        df_raw['wind_dir_surface'] = pd.to_numeric(df_raw['wind_dir_surface'], errors='coerce')
        df_raw['wind_speed_surface'] = pd.to_numeric(df_raw['wind_speed_surface'], errors='coerce')
        
        # FILTER KETAT: ws <= 25 knot
        df_raw = df_raw[(df_raw['wind_speed_surface'] >= 0) & (df_raw['wind_speed_surface'] <= MAX_WIND_SPEED_KT)]
        df_raw['datetime'] = pd.to_datetime(df_raw['data_timestamp'])
        df_raw['month'] = df_raw['datetime'].dt.month

        # 2. BACA & SANITASI DATA PIBAL (Vektor Angin Per Lapisan Standard & Sandi Pibal)
        df_data = None
        if path_data:
            with open(path_data, 'r') as f:
                first_line_d = f.readline()
            skip_data = 1 if 'Data Pibal' in first_line_d else 0
            df_data = pd.read_csv(path_data, skiprows=skip_data)
            
            df_data['lapisan'] = pd.to_numeric(df_data['lapisan'], errors='coerce')
            df_data['wd'] = pd.to_numeric(df_data['wd'], errors='coerce')
            df_data['ws'] = pd.to_numeric(df_data['ws'], errors='coerce')
            
            # FILTER KETAT ANGIN LAPISAN ATAS: ws <= 25 knot (membuang outlier > 60 kt & 542 kt)
            df_data = df_data[(df_data['ws'].isna()) | (df_data['ws'] <= MAX_WIND_SPEED_KT)]

        return df_raw, df_data

    except Exception as e:
        st.warning(f"Gagal membaca/mengolah file dataset historis. Detail: {e}")
        return None, None

df_historical, df_pibal_layers = load_historical_pibal()

# --- INSTANSIASI STATE MEMORI SIMULASI ---
if 'generated_records' not in st.session_state:
    st.session_state.generated_records = []
if 'hodo_points' not in st.session_state:
    st.session_state.hodo_points = []
if 'last_idx' not in st.session_state:
    st.session_state.last_idx = 0
if 'matched_info' not in st.session_state:
    st.session_state.matched_info = ""
if 'active_row' not in st.session_state:
    st.session_state.active_row = 1
if 'sandi_text' not in st.session_state:
    st.session_state.sandi_text = ""

# --- FUNGSI CALLBACK NAVIGASI MOBILE ---
def prev_row():
    if st.session_state.active_row > 1:
        st.session_state.active_row -= 1

def next_row():
    if st.session_state.active_row < len(st.session_state.generated_records):
        st.session_state.active_row += 1

# --- HEADER APLIKASI ---
st.markdown(
    f"""
    <div style='background-color:#0d3b66; padding:15px; border-radius:8px; text-align:center; color:white; margin-bottom:20px;'>
        <h2 style='margin:0; color:white;'>APLIKASI SIMULATOR PIBAL HISTORIS</h2>
        <p style='margin:5px 0 0 0; font-style:italic; font-size:14px;'>Stasiun Meteorologi Umbu Mehang Kunda Waingapu (97340) | Elevasi: {elevation_waingapu} ft</p>
    </div>
    """, 
    unsafe_allow_html=True
)

# --- DETEKSI OTOMATIS MUSIM ---
current_month = datetime.now().month
if current_month in [5, 6, 7, 8, 9]:
    status_deteksi = f"*Deteksi Otomatis: Musim Timur (Bulan {current_month})"
elif current_month in [11, 12, 1, 2, 3]:
    status_deteksi = f"*Deteksi Otomatis: Musim Barat (Bulan {current_month})"
else:
    status_deteksi = f"*Deteksi Otomatis: Pancaroba (Bulan {current_month})"

# --- FUNGSI SEARCH MATCHING HISTORIS ---
def find_best_historical_match(input_ddd, input_ff, input_month, df):
    if df is None or df.empty:
        return None
        
    obs_meta = df.groupby('data_timestamp').first().reset_index()[
        ['data_timestamp', 'wind_dir_surface', 'wind_speed_surface', 'month', 'datetime']
    ]
    
    rad_in = math.radians(input_ddd)
    rad_hist = np.radians(obs_meta['wind_dir_surface'])
    
    angle_diff = np.degrees(np.arccos(np.clip(
        np.cos(rad_in) * np.cos(rad_hist) + np.sin(rad_in) * np.sin(rad_hist), -1.0, 1.0
    )))
    
    speed_diff = np.abs(obs_meta['wind_speed_surface'] - input_ff)
    month_diff = np.minimum(np.abs(obs_meta['month'] - input_month), 12 - np.abs(obs_meta['month'] - input_month))
    
    score = angle_diff + 4.0 * speed_diff + 2.0 * month_diff
    best_idx = score.idxmin()
    best_obs = obs_meta.loc[best_idx]
    
    return best_obs['data_timestamp'], best_obs['datetime'], best_obs['wind_dir_surface'], best_obs['wind_speed_surface']

# --- FUNGSI CORE GENERATOR DENGAN SANITASI KETAT ANGIN <= 25 KT ---
def run_generation_core(target_readings, surf_ddd, surf_ff, month_idx, fresh=False):
    if fresh or not st.session_state.generated_records:
        st.session_state.generated_records = []
        st.session_state.hodo_points = []
        st.session_state.last_idx = 0
        st.session_state.active_row = 1
        st.session_state.sandi_text = ""

    if target_readings <= st.session_state.last_idx:
        st.warning(f"Data sudah ter-generate sebanyak {st.session_state.last_idx} baris. Naikkan target untuk melanjutkan.")
        return

    match_ts, match_dt, hist_ddd, hist_ff = None, None, surf_ddd, surf_ff
    hist_rows = []
    
    if df_historical is not None:
        match_res = find_best_historical_match(surf_ddd, surf_ff, month_idx, df_historical)
        if match_res:
            match_ts, match_dt, hist_ddd, hist_ff = match_res
            hist_rows = df_historical[df_historical['data_timestamp'] == match_ts].sort_values('pembacaan').to_dict('records')
            
            # Cek ketersediaan sandi pibal di file Data Pibal
            if df_pibal_layers is not None:
                sandi_match = df_pibal_layers[(df_pibal_layers['data_timestamp'] == match_ts) & (df_pibal_layers['sandi_pibal'].notna())]
                if not sandi_match.empty:
                    st.session_state.sandi_text = sandi_match.iloc[0]['sandi_pibal']

            dt_str = match_dt.strftime('%d %B %Y %H:%M UTC')
            st.session_state.matched_info = f"📌 **Pola Berdasarkan Data Historis Riil:** {dt_str} (Angin Permukaan Historis: {hist_ddd:.0f}° / {hist_ff:.0f} kt)"
        else:
            st.session_state.matched_info = "📌 **Pola Simulasi Matematika (Fallback)**"
    else:
        st.session_state.matched_info = "📌 **Pola Simulasi Matematika (Dataset Tidak Ditemukan)**"

    start_loop = st.session_state.last_idx + 1
    hist_dict = {r['pembacaan']: r for r in hist_rows}

    rate_ft_min = 600.0
    max_step_change_kt = 3.0  # Batas perubahan kecepatan antar layer agar grafik halus
    
    curr_x, curr_y = 0.0, 0.0
    
    # Komponen angin permukaan (dibatasi maks 25 kt)
    surf_ff_clean = min(surf_ff, MAX_WIND_SPEED_KT)
    u_surf = -surf_ff_clean * math.sin(math.radians(surf_ddd))
    v_surf = -surf_ff_clean * math.cos(math.radians(surf_ddd))
    
    prev_u, prev_v = u_surf, v_surf

    for idx in range(1, target_readings + 1):
        target_level = math.ceil((idx - 1) / 2) * 1000 if idx > 1 else 0
        level_target_str = "Diabaikan (Rilis)" if idx == 1 else f"Level {target_level} ft"
        height_above_stn = 100.0 if idx == 1 else (idx - 1) * 500.0
        
        # 1. Ambil nilai azimut & elevasi mentah
        if idx <= len(st.session_state.generated_records):
            existing_rec = st.session_state.generated_records[idx - 1]
            raw_az = float(existing_rec["AZIMUT"])
            raw_el = float(existing_rec["ELEVASI"])
        elif idx in hist_dict:
            raw_az = hist_dict[idx]['azimuth']
            raw_el = hist_dict[idx]['elevasi']
        else:
            if len(hist_rows) > 0:
                last_r = hist_rows[-1]
                delta_idx = idx - last_r['pembacaan']
                raw_az = (last_r['azimuth'] + delta_idx * random.uniform(-0.5, 0.5)) % 360
                raw_el = max(1.5, last_r['elevasi'] - delta_idx * random.uniform(0.1, 0.3))
            else:
                raw_az = (surf_ddd + random.uniform(-5, 5)) % 360
                raw_el = max(5.0, 45.0 - idx * 1.2)

        if idx == 1:
            clean_az, clean_el = raw_az, raw_el
            curr_x, curr_y = 0.0, 0.0
            u_comp, v_comp = u_surf, v_surf
            prev_u, prev_v = u_surf, v_surf
        else:
            prev_h = 100.0 if idx == 2 else (idx - 2) * 500.0
            dt = ((height_above_stn - prev_h) / rate_ft_min) * 60.0
            
            safe_el = max(0.5, min(89.0, raw_el))
            d_raw = height_above_stn / math.tan(math.radians(safe_el))
            x_raw_target = d_raw * math.sin(math.radians(raw_az))
            y_raw_target = d_raw * math.cos(math.radians(raw_az))
            
            dx_raw = x_raw_target - curr_x
            dy_raw = y_raw_target - curr_y
            u_raw = (dx_raw / dt) / 1.68781
            v_raw = (dy_raw / dt) / 1.68781
            
            # Low-Pass Filter (Smoothing)
            alpha = 0.25
            u_smooth = alpha * u_raw + (1 - alpha) * prev_u
            v_smooth = alpha * v_raw + (1 - alpha) * prev_v
            
            # Batasi Step-Change Vector
            du = u_smooth - prev_u
            dv = v_smooth - prev_v
            delta_spd = math.hypot(du, dv)
            if delta_spd > max_step_change_kt:
                scale = max_step_change_kt / delta_spd
                u_smooth = prev_u + du * scale
                v_smooth = prev_v + dv * scale
            
            # BATASI MAKSIMUM KECEPATAN ANGIN PER LAYER <= 25 KT
            total_spd = math.hypot(u_smooth, v_smooth)
            if total_spd > MAX_WIND_SPEED_KT:
                scale = MAX_WIND_SPEED_KT / total_spd
                u_smooth *= scale
                v_smooth *= scale
            
            u_comp, v_comp = u_smooth, v_smooth
            prev_u, prev_v = u_smooth, v_smooth
            
            # Recalculate Azimuth & Elevasi
            dx_clean = u_comp * 1.68781 * dt
            dy_clean = v_comp * 1.68781 * dt
            curr_x += dx_clean
            curr_y += dy_clean
            
            d_clean = math.hypot(curr_x, curr_y)
            if d_clean > 0:
                clean_az = math.degrees(math.atan2(curr_x, curr_y)) % 360
                clean_el = math.degrees(math.atan2(height_above_stn, d_clean))
            else:
                clean_az, clean_el = raw_az, raw_el

        # Simpan Hasil
        if idx >= start_loop:
            height_display = "Awal" if idx == 1 else f"{int(height_above_stn)} ft"
            st.session_state.generated_records.append({
                "Pembacaan Ke-": idx,
                "Tinggi Balon (ft)": height_display,
                "Level Target (BMKG)": level_target_str,
                "AZIMUT": round(clean_az, 1),
                "ELEVASI": round(clean_el, 1)
            })

        if idx >= start_loop or fresh:
            st.session_state.hodo_points.append((u_comp, v_comp, idx))

    st.session_state.last_idx = target_readings
    st.session_state.active_row = target_readings

# --- LAYOUT KANAN & KIRI ---
col_left, col_right = st.columns([7, 5], gap="large")

# === KOLOM KIRI: INPUT & TABEL DATA ===
with col_left:
    st.subheader("⚙️ Parameter Kontrol Pengamatan")
    
    c1, c2 = st.columns(2)
    with c1:
        target_readings = st.number_input("Target Jumlah Pembacaan:", min_value=1, value=25, step=1)
        surf_ddd = st.number_input("Angin Permukaan ddd (°):", min_value=0.0, max_value=360.0, value=190.0, step=5.0)
    with c2:
        surf_ff = st.number_input("Kec Angin Perm ff (kt, Max 25):", min_value=0.0, max_value=25.0, value=5.0, step=1.0)
        selected_month = st.selectbox("Bulan Pengamatan:", list(range(1, 13)), index=current_month-1, 
                                      format_func=lambda x: datetime(2024, x, 1).strftime('%B'))

    b1, b2 = st.columns(2)
    with b1:
        if st.button("⚡ Generate dari Historis", type="primary", use_container_width=True):
            run_generation_core(target_readings, surf_ddd, surf_ff, selected_month, fresh=True)
    with b2:
        if st.button("⏩ Lanjutkan ke Target", use_container_width=True):
            if st.session_state.last_idx == 0:
                st.error("Belum ada data awal. Silakan klik 'Generate dari Historis' terlebih dahulu.")
            else:
                run_generation_core(target_readings, surf_ddd, surf_ff, selected_month, fresh=False)

    st.markdown("---")
    
    if st.session_state.matched_info:
        st.info(st.session_state.matched_info)
        
    st.subheader("📊 Tabel Hasil Pembacaan")
    
    if st.session_state.generated_records:
        df_result = pd.DataFrame(st.session_state.generated_records)
        st.dataframe(df_result, use_container_width=True, hide_index=True)
        
        # Fitur Tambahan: Download CSV
        csv_data = df_result.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download Hasil Pembacaan (CSV)",
            data=csv_data,
            file_name=f"pibal_simulation_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime='text/csv'
        )
    else:
        st.info("Belum ada data yang dibuat. Atur parameter lalu pilih 'Generate dari Historis'.")

# === KOLOM KANAN: HODOGRAPH & ASISTEN FORM ===
with col_right:
    st.subheader("🎯 Verifikasi Kelurusan Angin (Hodograph)")
    
    fig, ax = plt.subplots(figsize=(6, 6), facecolor='#f8f9fa')
    ax.set_facecolor('#ffffff')
    ax.set_aspect('equal')
    
    # Ring Hodograph disesuaikan dengan skala max 25-30 kt
    for knots in [5, 10, 15, 20, 25]:
        circle = plt.Circle((0, 0), knots, color='#cbd5e1', fill=False, linestyle='-', linewidth=1)
        ax.add_patch(circle)
        ax.text(0, knots, f"{knots} kt", color='#64748b', fontsize=8, ha='center', va='center',
                bbox=dict(facecolor='white', edgecolor='none', pad=2, alpha=0.8))
        
    ax.axhline(0, color='#94a3b8', linestyle='--', linewidth=0.8)
    ax.axvline(0, color='#94a3b8', linestyle='--', linewidth=0.8)
    
    c_props = dict(boxstyle='round,pad=0.3', facecolor='#0d3b66', edgecolor='none', alpha=0.9)
    ax.text(0, 28, "U", weight='bold', ha='center', va='center', color='white', bbox=c_props)
    ax.text(0, -28, "S", weight='bold', ha='center', va='center', color='white', bbox=c_props)
    ax.text(28, 0, "T", weight='bold', ha='center', va='center', color='white', bbox=c_props)
    ax.text(-28, 0, "B", weight='bold', ha='center', va='center', color='white', bbox=c_props)
    
    if st.session_state.hodo_points:
        u_pts = [p[0] for p in st.session_state.hodo_points]
        v_pts = [p[1] for p in st.session_state.hodo_points]
        
        ax.plot(u_pts, v_pts, color='#94a3b8', linewidth=1.5, zorder=1)
        colors = [cm.plasma(i/len(u_pts)) for i in range(len(u_pts))]
        ax.scatter(u_pts, v_pts, color=colors, edgecolor='white', s=55, zorder=2)
        
        ax.plot(u_pts[0], v_pts[0], marker='s', color='#10b981', markersize=9, markeredgecolor='white', zorder=3, label='Mulai')
        ax.plot(u_pts[-1], v_pts[-1], marker='X', color='#ef4444', markersize=10, markeredgecolor='white', zorder=3, label='Akhir')
        ax.legend(loc='lower right', fontsize=8, framealpha=0.9)
        
    ax.set_xlim(-30, 30)
    ax.set_ylim(-30, 30)
    ax.axis('off')
    
    st.pyplot(fig)
    plt.close(fig)
    
    st.caption(status_deteksi)
    
    if st.session_state.sandi_text:
        with st.expander("📝 Referensi Sandi Pibal Historis (PPAA/PPBB)"):
            st.code(st.session_state.sandi_text, language='text')

    st.markdown("---")
    
    # --- PANEL ASISTEN KETIK MANUAL ---
    st.subheader("🔍 Panel Bantuan Ketik Manual")
    
    if st.session_state.generated_records:
        total_rec = len(st.session_state.generated_records)
        
        if st.session_state.active_row > total_rec:
            st.session_state.active_row = total_rec
        if st.session_state.active_row < 1:
            st.session_state.active_row = 1
        
        st.markdown("<p style='text-align:center; font-weight:bold; margin-bottom:5px;'>Navigasi Baris Form (Khusus HP):</p>", unsafe_allow_html=True)
        nav1, nav2, nav3 = st.columns([1, 2, 1])
        with nav1:
            st.button("⬅️ Mundur", on_click=prev_row, use_container_width=True)
        with nav2:
            st.slider("Pilih Baris", min_value=1, max_value=total_rec, key='active_row', label_visibility="collapsed")
        with nav3:
            st.button("Maju ➡️", on_click=next_row, use_container_width=True)
        
        active_rec = st.session_state.generated_records[st.session_state.active_row - 1]
        
        azimuth_fmt = f"{active_rec['AZIMUT']:.1f}".replace('.', ',')
        elevation_fmt = f"{active_rec['ELEVASI']:.1f}".replace('.', ',')
        
        st.markdown(
            f"""
            <div style="background-color: #fffdf0; padding: 20px; border-radius: 8px; border: 2px solid #d62828; text-align: center; margin-top:10px;">
                <div style="text-align: left; margin-bottom: 10px;">
                    <span style="font-size: 16px; font-weight: bold; color: #333;">Pembacaan Ke: {active_rec['Pembacaan Ke-']}</span><br>
                    <span style="font-size: 14px; font-style: italic; color: #e67e22; font-weight: bold;">Target Form: {active_rec['Level Target (BMKG)']}</span>
                </div>
                <div style="display: flex; justify-content: space-around; margin-top: 15px;">
                    <div>
                        <div style="color: gray; font-size: 12px; font-weight: bold; letter-spacing: 1px;">AZIMUT</div>
                        <div style="color: #005b96; font-size: 45px; font-weight: bold; line-height: 1;">{azimuth_fmt}</div>
                    </div>
                    <div>
                        <div style="color: gray; font-size: 12px; font-weight: bold; letter-spacing: 1px;">ELEVASI</div>
                        <div style="color: #d62828; font-size: 45px; font-weight: bold; line-height: 1;">{elevation_fmt}</div>
                    </div>
                </div>
            </div>
            """, 
            unsafe_allow_html=True
        )
    else:
        st.info("Silakan generate data terlebih dahulu untuk memunculkan panel bantuan ketik manual.")
