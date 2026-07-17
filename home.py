import streamlit as st

st.set_page_config(page_title="hello Kamila", page_icon="⛅", layout="centered")

st.title("⛅ Kamila ")

# Kotak pesan berwarna biru
st.info("👋 Selamat datang! Aplikasi ini dibuat untuk mempermudah pengolahan data laporan bulanan.")

st.markdown("""
Silakan pilih menu di **Sidebar sebelah kiri** untuk mulai menggunakan aplikasi:

1. **METAR Converter** ✈️ 
   *Gunakan ini untuk mengekstrak dan merapikan sandi METAR ke format PDF atau Excel.*
2. **Klimat Average** 📊
   *Gunakan ini untuk menghitung rata-rata harian (suhu, RH, QFF, QFE, dll) dan mengekspor matriks Excel.*

---
*Catatan: Pastikan format file Anda adalah CSV (.csv).*
""")
