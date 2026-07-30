import streamlit as st
import math
import random
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.cm as cm
from datetime import datetime, timedelta

# --- KONFIGURASI GLOBAL & KONSTANTA ---
ELEVATION_WAINGAPU = 32.8  # Elevasi Stasiun Umbu Mehang Kunda (ft)
MAX_SPEED_KT = 19.0        # Batas maksimum kecepatan angin realistis Waingapu (kt)

st.set_page_config(
    page_title="Pibal Generator Stamet Waingapu",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- FUNGSI DETEKSI SESI PIBAL & DETEKSI IKLIM WAINGAPU ---
def get_waingapu_schedule_and_climate():
    # Mengambil waktu lokal WITA (UTC+8)
    now_wita = datetime.utcnow() + timedelta(hours=8)
    hour = now_wita.hour
    month = now_wita.month

    # 1. Tentukan Sesi Pibal Rutin Waingapu
    if 0 <= hour < 10:
        session_name = "Pagi"
        release_time = "06:00 WITA"
        transmit_time = "00:00 UTC (08:00 WITA)"
        default_speed = 7.0
    elif 10 <= hour < 16:
        session_name = "Siang"
        release_time = "13:00 WITA"
        transmit_time = "06:00 UTC (14:00 WITA)"
        default_speed = 14.0
    else:
        session_name = "Malam"
        release_time = "19:00 WITA"
        transmit_time = "12:00 UTC (20:00 WITA)"
        default_speed = 8.0

    # 2. Tentukan Musim & Arah Dominan Waingapu
    if month in [5, 6, 7, 8, 9, 10]:
        season_name = "Musim Timur (Kemarau)"
        default_dir = 140.0  # Dominan Angin Tenggara (SE)
        season_desc = "Dominan Angin Pasat Tenggara (SE) khas Australia."
    elif month in [12, 1, 2, 3]:
        season_name = "Musim Barat (Hujan)"
        default_dir = 290.0  # Dominan Angin Barat Laut / Barat Daya
        season_desc = "Dominan Angin Barat Laut / Barat Daya bertiup dari Asia/Samudra Hindia."
    else:
        season_name = "Pancaroba"
        default_dir = 180.0  # Arah bervariasi
        season_desc = "Arah angin cenderung bervariasi dengan kecepatan fluktuatif."

    return {
        "now_wita": now_wita,
        "month": month,
        "session": session_name,
        "release": release_time,
        "transmit": transmit_time,
        "season": season_name,
        "season_desc": season_desc,
        "default_dir": default_dir,
        "default_speed": min(default_speed, MAX_SPEED_KT)
    }

climate_info = get_waingapu_schedule_and_climate()

# --- LOAD DATASET HISTORIS UMBU MEHANG KUNDA ---
@st.cache_data
def load_historical_pibal():
    filename = 'Raw Pibal 2024-01-01 to 2026-07-21.csv'
    try:
        df = pd.read_csv(filename)
        df = df.dropna(subset=['azimuth', 'elevasi']).copy()
        df = df[(df['azimuth'] != 9999) & (df['elevasi'] != 9999)]
        df['pembacaan'] = df['pembacaan'].astype(int)
        df['azimuth'] = df['azimuth'].astype(float)
        df['elevasi'] = df['elevasi'].astype(float)
        df['wind_dir_surface'] = df['wind_dir_surface'].astype(float)
        df['wind_speed_surface'] = df['wind_speed_surface'].astype(float)
        
        df['datetime'] = pd.to_datetime(df['data_timestamp'])
        df['month'] = df['datetime'].dt.month

        obs_meta = df.groupby('data_timestamp').first().reset_index()[
            ['data_timestamp', 'wind_dir_surface', 'wind_speed_surface', 'month', 'datetime']
        ]
        return df, obs_meta
    except Exception:
        return None, None

df_historical, obs_metadata = load_historical_pibal()

# --- INSTANSIASI STATE MEMORI ---
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

# --- CALLBACK NAVIGASI MOBILE ---
def prev_row():
    if st.session_state.active_row > 1:
        st.session_state.active_row -= 1

def next_row():
    if st.session_state.active_row < len(st.session_state.generated_records):
        st.session_state.active_row += 1

# --- HEADER APLIKASI ---
st.markdown(
    f"""
    <div style='background-color:#0d3b66; padding:15px; border-radius:8px; text-align:center; color:white; margin-bottom:15px;'>
        <h2 style='margin:0; color:white;'>APLIKASI SIMULATOR PIBAL HISTORIS WAINGAPU</h2>
        <p style='margin:5px 0 0 0; font-style:italic; font-size:13px;'>
            Stamet Umbu Mehang Kunda (97340) | Elevasi: {ELEVATION_WAINGAPU} ft | Waktu Sistem: {climate_info['now_wita'].strftime('%H:%M WITA')}
        </p>
    </div>
    """, 
    unsafe_allow_html=True
)

# --- CARD STATUS OTOMATIS SESI & IKLIM ---
st.markdown(
    f"""
    <div style='background-color:#eef2f6; padding:12px; border-radius:6px; border-left:5px solid #0d3b66; margin-bottom:20px; font-size:13px;'>
        <b>📌 Deteksi Otomatis Sesi Pibal:</b> Sesi <b>{climate_info['session']}</b> (Rilis ~{climate_info['release']} | Laporan {climate_info['transmit']})<br>
        <b>🌤️ Karakater Iklim Lokal:</b> {climate_info['season']} — <i>{climate_info['season_desc']}</i>
    </div>
    """,
    unsafe_allow_html=True
)

# --- FUNGSI MATCHING HISTORIS ---
def find_best_historical_match(input_ddd, input_ff, input_month, obs_meta):
    if obs_meta is None or obs_meta.empty:
        return None
        
    rad_in = math.radians(input_ddd)
    rad_hist = np.radians(obs_meta['wind_dir_surface'])
    
    angle_diff = np.degrees(np.arccos(np.clip(
        np.cos(rad_in) * np.cos(rad_hist) + np.sin(rad_in) * np.sin(rad_hist), -1.0, 1.0
    )))
    speed_diff = np.abs(obs_meta['wind_speed_surface'] - input_ff)
    month_diff = np.minimum(np.abs(obs_meta['month'] - input_month), 12 - np.abs(obs_meta['month'] - input_month))
    
    score = angle_diff + 5.0 * speed_diff + 3.0 * month_diff
    best_idx = score.idxmin()
    best_obs = obs_meta.loc[best_idx]
    
    return best_obs['data_timestamp'], best_obs['datetime'], best_obs['wind_dir_surface'], best_obs['wind_speed_surface']

# --- GENERATOR INTEGRASI MAJU (SMOOTH & PHYSICAL) ---
# --- GENERATOR INTEGRASI MAJU (REALISTIS: ADA TREN & TURBULENSI HALUS) ---
def run_generation_core(target_readings, surf_ddd, surf_ff, month_idx, fresh=False):
    if fresh or not st.session_state.generated_records:
        st.session_state.generated_records = []
        st.session_state.hodo_points = []
        st.session_state.last_idx = 0
        st.session_state.active_row = 1

    if target_readings <= st.session_state.last_idx:
        st.warning(f"Data sudah ter-generate sebanyak {st.session_state.last_idx} baris.")
        return

    # Cari referensi data historis
    hist_rows = []
    if df_historical is not None and obs_metadata is not None:
        match_res = find_best_historical_match(surf_ddd, surf_ff, month_idx, obs_metadata)
        if match_res:
            match_ts, match_dt, h_ddd, h_ff = match_res
            hist_rows = df_historical[df_historical['data_timestamp'] == match_ts].sort_values('pembacaan').to_dict('records')
            dt_str = match_dt.strftime('%d %B %Y %H:%M UTC')
            st.session_state.matched_info = f"📌 **Pola Historis Acuan:** {dt_str} (Permukaan: {h_ddd:.0f}° / {min(h_ff, MAX_SPEED_KT):.0f} kt)"
        else:
            st.session_state.matched_info = "📌 **Pola Simulasi Profil Atmosfer Waingapu (Fallback)**"
    else:
        st.session_state.matched_info = "📌 **Pola Simulasi Profil Atmosfer Waingapu (Dataset Offline)**"

    hist_dict = {r['pembacaan']: r for r in hist_rows}
    rate_ft_min = 600.0  # Laju naik balon standar BMKG (600 ft/menit)
    
    x_curr, y_curr = 0.0, 0.0
    
    # Vektor Awal
    u_surf = -surf_ff * math.sin(math.radians(surf_ddd))
    v_surf = -surf_ff * math.cos(math.radians(surf_ddd))
    u_prev, v_prev = u_surf, v_surf
    
    # Alpha dinaikkan ke 0.65 agar gerigi alami atmosfer tidak terhapus total
    alpha = 0.65 

    for idx in range(1, target_readings + 1):
        target_level = math.ceil((idx - 1) / 2) * 1000 if idx > 1 else 0
        level_target_str = "Diabaikan (Rilis)" if idx == 1 else f"Level {target_level} ft"
        height_above_stn = 100.0 if idx == 1 else (idx - 1) * 500.0
        
        # 1. Penentuan Vektor Mentah
        if idx in hist_dict:
            az_m = hist_dict[idx]['azimuth']
            el_m = max(2.0, min(88.0, hist_dict[idx]['elevasi']))
            d_m = height_above_stn / math.tan(math.radians(el_m))
            x_m = d_m * math.sin(math.radians(az_m))
            y_m = d_m * math.cos(math.radians(az_m))
            
            dt_step = 10.0 if idx == 1 else ((height_above_stn - ((idx - 2) * 500.0 if idx > 2 else 100.0)) / rate_ft_min) * 60.0
            u_raw = ((x_m - x_curr) / dt_step) / 1.68781
            v_raw = ((y_m - y_curr) / dt_step) / 1.68781
        else:
            # Profil sintetis non-linier khas lapisan batas Waingapu
            shear_angle = math.radians(surf_ddd + (math.log(idx + 1) * 8.0))
            speed_target = min(surf_ff + (math.sqrt(idx) * 1.5), MAX_SPEED_KT)
            u_raw = -speed_target * math.sin(shear_angle)
            v_raw = -speed_target * math.cos(shear_angle)

        # 2. Injeksi Micro-Turbulence Atmosfer (Derau Halus 0.3 - 0.5 kt)
        if idx > 1:
            u_raw += np.random.normal(0, 0.4)
            v_raw += np.random.normal(0, 0.4)

        # 3. Smoothing Ringan (Menjaga Tren Tanpa Menghilangkan Karakter Alami)
        u_smooth = alpha * u_raw + (1 - alpha) * u_prev
        v_smooth = alpha * v_raw + (1 - alpha) * v_prev
        
        # 4. Batas Maksimum Kecepatan Strictly 19 Knot
        current_speed = math.hypot(u_smooth, v_smooth)
        if current_speed > MAX_SPEED_KT:
            scale = MAX_SPEED_KT / current_speed
            u_smooth *= scale
            v_smooth *= scale

        u_prev, v_prev = u_smooth, v_smooth

        # 5. Integrasi Maju Posisi Balon (x, y)
        if idx == 1:
            dt = 10.0
            x_curr = u_smooth * 1.68781 * dt
            y_curr = v_smooth * 1.68781 * dt
        else:
            prev_h = 100.0 if idx == 2 else (idx - 2) * 500.0
            dt = ((height_above_stn - prev_h) / rate_ft_min) * 60.0
            x_curr += u_smooth * 1.68781 * dt
            y_curr += v_smooth * 1.68781 * dt

        # 6. Hitung Ulang Azimut & Elevasi Berbasis Geometri Posisi Riil
        d_horiz = math.hypot(x_curr, y_curr)
        if d_horiz > 0:
            clean_az = math.degrees(math.atan2(x_curr, y_curr)) % 360
            clean_el = math.degrees(math.atan2(height_above_stn, d_horiz))
        else:
            clean_az, clean_el = surf_ddd, 85.0

        # Simpan Hasil
        if idx > st.session_state.last_idx or fresh:
            height_display = "Awal" if idx == 1 else f"{int(height_above_stn)} ft"
            st.session_state.generated_records.append({
                "Pembacaan Ke-": idx,
                "Tinggi Balon (ft)": height_display,
                "Level Target (BMKG)": level_target_str,
                "AZIMUT": round(clean_az, 1),
                "ELEVASI": round(clean_el, 1)
            })
            st.session_state.hodo_points.append((u_smooth, v_smooth, idx))

    st.session_state.last_idx = target_readings
    st.session_state.active_row = target_readings
# --- LAYOUT DUA KOLOM ---
col_left, col_right = st.columns([7, 5], gap="large")

# === KOLOM KIRI: INPUT & TABEL ===
with col_left:
    st.subheader("⚙️ Parameter Kontrol Pengamatan")
    
    c1, c2 = st.columns(2)
    with c1:
        target_readings = st.number_input("Target Jumlah Pembacaan:", min_value=1, value=25, step=1)
        surf_ddd = st.number_input(
            "Angin Permukaan ddd (°):", 
            min_value=0.0, max_value=360.0, 
            value=float(climate_info['default_dir']), step=5.0
        )
    with c2:
        surf_ff = st.number_input(
            "Kec Angin Perm ff (kt) [Max 19]:", 
            min_value=0.0, max_value=MAX_SPEED_KT, 
            value=float(climate_info['default_speed']), step=1.0
        )
        selected_month = st.selectbox(
            "Bulan Pengamatan:", list(range(1, 13)), 
            index=climate_info['month'] - 1, 
            format_func=lambda x: datetime(2024, x, 1).strftime('%B')
        )

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
        df_out = pd.DataFrame(st.session_state.generated_records)
        st.dataframe(df_out, use_container_width=True, hide_index=True)
    else:
        st.info("Belum ada data. Silakan klik 'Generate dari Historis'.")

# === KOLOM KANAN: HODOGRAPH & PANEL HP ===
with col_right:
    st.subheader("🎯 Verifikasi Kelurusan Angin (Hodograph)")
    
    fig, ax = plt.subplots(figsize=(6, 6), facecolor='#f8f9fa')
    ax.set_facecolor('#ffffff')
    ax.set_aspect('equal')
    
    # Ring Kecepatan (Skala 5, 10, 15, 20 kt)
    for knots in [5, 10, 15, 20]:
        circle = plt.Circle((0, 0), knots, color='#cbd5e1', fill=False, linestyle='-', linewidth=1)
        ax.add_patch(circle)
        ax.text(0, knots, f"{knots} kt", color='#64748b', fontsize=8, ha='center', va='center',
                bbox=dict(facecolor='white', edgecolor='none', pad=1.5, alpha=0.85))
        
    ax.axhline(0, color='#94a3b8', linestyle='--', linewidth=0.8)
    ax.axvline(0, color='#94a3b8', linestyle='--', linewidth=0.8)
    
    c_props = dict(boxstyle='round,pad=0.3', facecolor='#0d3b66', edgecolor='none', alpha=0.9)
    ax.text(0, 22, "U", weight='bold', ha='center', va='center', color='white', bbox=c_props)
    ax.text(0, -22, "S", weight='bold', ha='center', va='center', color='white', bbox=c_props)
    ax.text(22, 0, "T", weight='bold', ha='center', va='center', color='white', bbox=c_props)
    ax.text(-22, 0, "B", weight='bold', ha='center', va='center', color='white', bbox=c_props)
    
    if st.session_state.hodo_points:
        u_pts = [p[0] for p in st.session_state.hodo_points]
        v_pts = [p[1] for p in st.session_state.hodo_points]
        
        # Garis penghubung halus
        ax.plot(u_pts, v_pts, color='#64748b', linewidth=2.0, zorder=1, alpha=0.8)
        colors = [cm.plasma(i/len(u_pts)) for i in range(len(u_pts))]
        ax.scatter(u_pts, v_pts, color=colors, edgecolor='white', s=60, zorder=2)
        
        ax.plot(u_pts[0], v_pts[0], marker='s', color='#10b981', markersize=9, markeredgecolor='white', zorder=3, label='Mulai (Bawah)')
        ax.plot(u_pts[-1], v_pts[-1], marker='X', color='#ef4444', markersize=10, markeredgecolor='white', zorder=3, label='Akhir (Atas)')
        ax.legend(loc='lower right', fontsize=8, framealpha=0.9)
        
    ax.set_xlim(-25, 25)
    ax.set_ylim(-25, 25)
    ax.axis('off')
    
    st.pyplot(fig)
    plt.close(fig)
    
    st.markdown("---")
    
    # --- PANEL ASISTEN KETIK MANUAL ---
    st.subheader("🔍 Panel Bantuan Ketik Manual")
    
    if st.session_state.generated_records:
        total_rec = len(st.session_state.generated_records)
        
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
            <div style="background-color: #fffdf0; padding: 18px; border-radius: 8px; border: 2px solid #d62828; text-align: center; margin-top:10px;">
                <div style="text-align: left; margin-bottom: 8px;">
                    <span style="font-size: 15px; font-weight: bold; color: #333;">Pembacaan Ke: {active_rec['Pembacaan Ke-']}</span><br>
                    <span style="font-size: 13px; font-style: italic; color: #e67e22; font-weight: bold;">Target Form: {active_rec['Level Target (BMKG)']}</span>
                </div>
                <div style="display: flex; justify-content: space-around; margin-top: 10px;">
                    <div>
                        <div style="color: gray; font-size: 11px; font-weight: bold; letter-spacing: 1px;">AZIMUT</div>
                        <div style="color: #005b96; font-size: 42px; font-weight: bold; line-height: 1;">{azimuth_fmt}</div>
                    </div>
                    <div>
                        <div style="color: gray; font-size: 11px; font-weight: bold; letter-spacing: 1px;">ELEVASI</div>
                        <div style="color: #d62828; font-size: 42px; font-weight: bold; line-height: 1;">{elevation_fmt}</div>
                    </div>
                </div>
            </div>
            """, 
            unsafe_allow_html=True
        )
    else:
        st.info("Silakan generate data terlebih dahulu untuk memunculkan panel bantuan ketik manual.")
