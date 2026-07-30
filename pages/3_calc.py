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

# --- LOAD DATASET HISTORIS UMBU MEHANG KUNDA ---
@st.cache_data
def load_historical_pibal():
    # Perbaikan path agar file dapat ditemukan baik di root maupun folder pages
    filename = 'Raw Pibal 2024-01-01 to 2026-07-21.csv'
    if not os.path.exists(filename):
        filename = os.path.join(os.path.dirname(__file__), '..', 'Raw Pibal 2024-01-01 to 2026-07-21.csv')
        
    try:
        with open(filename, 'r') as f:
            first_line = f.readline()
        skip = 1 if 'Raw Pibal' in first_line else 0
        df = pd.read_csv(filename, skiprows=skip)
        
        # Pembersihan Data Mentah
        df = df.dropna(subset=['azimuth', 'elevasi']).copy()
        df = df[(df['azimuth'] != 9999) & (df['elevasi'] != 9999)]
        df['pembacaan'] = df['pembacaan'].astype(int)
        df['azimuth'] = df['azimuth'].astype(float)
        df['elevasi'] = df['elevasi'].astype(float)
        df['wind_dir_surface'] = df['wind_dir_surface'].astype(float)
        df['wind_speed_surface'] = df['wind_speed_surface'].astype(float)
        
        df['datetime'] = pd.to_datetime(df['data_timestamp'])
        df['month'] = df['datetime'].dt.month
        return df
    except Exception as e:
        st.warning(f"File dataset historis tidak ditemukan atau gagal dibaca. Detail: {e}")
        return None

df_historical = load_historical_pibal()

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

# --- FUNGSI CORE GENERATOR DATA DENGAN SANITASI FISIKA AZIMUT/ELEVASI ---
def run_generation_core(target_readings, surf_ddd, surf_ff, month_idx, fresh=False):
    if fresh or not st.session_state.generated_records:
        st.session_state.generated_records = []
        st.session_state.hodo_points = []
        st.session_state.last_idx = 0
        st.session_state.active_row = 1

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
            dt_str = match_dt.strftime('%d %B %Y %H:%M UTC')
            st.session_state.matched_info = f"📌 **Pola Berdasarkan Data Historis Riil:** {dt_str} (Angin Permukaan Historis: {hist_ddd:.0f}° / {hist_ff:.0f} kt)"
        else:
            st.session_state.matched_info = "📌 **Pola Simulasi Matematika (Fallback)**"
    else:
        st.session_state.matched_info = "📌 **Pola Simulasi Matematika (Dataset Tidak Ditemukan)**"

    start_loop = st.session_state.last_idx + 1
    hist_dict = {r['pembacaan']: r for r in hist_rows}

    rate_ft_min = 600.0
    max_allowed_speed_kt = 35.0
    
    prev_x, prev_y = 0.0, 0.0

    for idx in range(1, target_readings + 1):
        target_level = math.ceil((idx - 1) / 2) * 1000 if idx > 1 else 0
        level_target_str = "Diabaikan (Rilis)" if idx == 1 else f"Level {target_level} ft"
        height_above_stn = 100.0 if idx == 1 else (idx - 1) * 500.0
        
        # PERBAIKAN BUG 1: Gunakan data yang sudah tersimpan jika ada untuk menghindari pergeseran acak
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

        # Sanitasi fisika Azimut/Elevasi
        if idx == 1:
            dt = 10.0
            clean_az = raw_az
            clean_el = raw_el
            x, y = 0.0, 0.0
            u_comp = -surf_ff * math.sin(math.radians(surf_ddd))
            v_comp = -surf_ff * math.cos(math.radians(surf_ddd))
        else:
            prev_h = 100.0 if idx == 2 else (idx - 2) * 500.0
            dt = ((height_above_stn - prev_h) / rate_ft_min) * 60.0
            
            safe_el = max(0.5, min(89.0, raw_el))
            d = height_above_stn / math.tan(math.radians(safe_el))
            
            x = d * math.sin(math.radians(raw_az))
            y = d * math.cos(math.radians(raw_az))
            
            dx = x - prev_x
            dy = y - prev_y
            dist_ft = math.hypot(dx, dy)
            speed_kt = (dist_ft / dt) / 1.68781
            
            if speed_kt > max_allowed_speed_kt:
                adj_dist_ft = max_allowed_speed_kt * 1.68781 * dt
                move_dir = math.degrees(math.atan2(dx, dy)) % 360
                
                x = prev_x + adj_dist_ft * math.sin(math.radians(move_dir))
                y = prev_y + adj_dist_ft * math.cos(math.radians(move_dir))
                
                d = math.hypot(x, y)
                if d > 0:
                    clean_az = math.degrees(math.atan2(x, y)) % 360
                    clean_el = math.degrees(math.atan2(height_above_stn, d))
                else:
                    clean_az, clean_el = raw_az, raw_el
            else:
                clean_az, clean_el = raw_az, raw_el

            move_dir = math.degrees(math.atan2(x - prev_x, y - prev_y)) % 360
            wind_dir = (move_dir + 180) % 360
            calc_speed_kt = min(speed_kt, max_allowed_speed_kt)
            
            u_comp = -calc_speed_kt * math.sin(math.radians(wind_dir))
            v_comp = -calc_speed_kt * math.cos(math.radians(wind_dir))

        prev_x, prev_y = x, y

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

# --- LAYOUT DENGAN DUA KOLOM UTAMA ---
col_left, col_right = st.columns([7, 5], gap="large")

# === KOLOM KIRI: INPUT & TABEL DATA ===
with col_left:
    st.subheader("⚙️ Parameter Kontrol Pengamatan")
    
    c1, c2 = st.columns(2)
    with c1:
        target_readings = st.number_input("Target Jumlah Pembacaan:", min_value=1, value=25, step=1)
        surf_ddd = st.number_input("Angin Permukaan ddd (°):", min_value=0.0, max_value=360.0, value=190.0, step=5.0)
    with c2:
        surf_ff = st.number_input("Kec Angin Perm ff (kt):", min_value=0.0, value=5.0, step=1.0)
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
        df = pd.DataFrame(st.session_state.generated_records)
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("Belum ada data yang dibuat. Atur parameter lalu pilih 'Generate dari Historis'.")

# === KOLOM KANAN: HODOGRAPH & ASISTEN FORM ===
with col_right:
    st.subheader("🎯 Verifikasi Kelurusan Angin (Hodograph)")
    
    fig, ax = plt.subplots(figsize=(6, 6), facecolor='#f8f9fa')
    ax.set_facecolor('#ffffff')
    ax.set_aspect('equal')
    
    for knots in [10, 20, 30, 40]:
        circle = plt.Circle((0, 0), knots, color='#cbd5e1', fill=False, linestyle='-', linewidth=1)
        ax.add_patch(circle)
        ax.text(0, knots, f"{knots} kt", color='#64748b', fontsize=8, ha='center', va='center',
                bbox=dict(facecolor='white', edgecolor='none', pad=2, alpha=0.8))
        
    ax.axhline(0, color='#94a3b8', linestyle='--', linewidth=0.8)
    ax.axvline(0, color='#94a3b8', linestyle='--', linewidth=0.8)
    
    c_props = dict(boxstyle='round,pad=0.3', facecolor='#0d3b66', edgecolor='none', alpha=0.9)
    ax.text(0, 45, "U", weight='bold', ha='center', va='center', color='white', bbox=c_props)
    ax.text(0, -45, "S", weight='bold', ha='center', va='center', color='white', bbox=c_props)
    ax.text(45, 0, "T", weight='bold', ha='center', va='center', color='white', bbox=c_props)
    ax.text(-45, 0, "B", weight='bold', ha='center', va='center', color='white', bbox=c_props)
    
    if st.session_state.hodo_points:
        u_pts = [p[0] for p in st.session_state.hodo_points]
        v_pts = [p[1] for p in st.session_state.hodo_points]
        
        ax.plot(u_pts, v_pts, color='#94a3b8', linewidth=1.5, zorder=1)
        colors = [cm.plasma(i/len(u_pts)) for i in range(len(u_pts))]
        ax.scatter(u_pts, v_pts, color=colors, edgecolor='white', s=55, zorder=2)
        
        ax.plot(u_pts[0], v_pts[0], marker='s', color='#10b981', markersize=9, markeredgecolor='white', zorder=3, label='Mulai (Bawah)')
        ax.plot(u_pts[-1], v_pts[-1], marker='X', color='#ef4444', markersize=10, markeredgecolor='white', zorder=3, label='Akhir (Atas)')
        ax.legend(loc='lower right', fontsize=8, framealpha=0.9)
        
    ax.set_xlim(-50, 50)
    ax.set_ylim(-50, 50)
    ax.axis('off')
    
    st.pyplot(fig)
    plt.close(fig)
    
    st.caption(status_deteksi)
    st.markdown("---")
    
    # --- PANEL ASISTEN KETIK MANUAL ---
    st.subheader("🔍 Panel Bantuan Ketik Manual")
    
    if st.session_state.generated_records:
        total_rec = len(st.session_state.generated_records)
        
        # PERBAIKAN BUG 2: Validasi nilai active_row agar tidak melebihi total_rec
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
