import streamlit as st

st.set_page_config(page_title="hello Kamila", page_icon="⛅", layout="centered")

st.title("⛅ Kamila ")

# Kotak pesan berwarna biru
st.info("👋 Selamat datang! Aplikasi ini dibuat untuk mempermudah pengolahan data laporan bulanan.")

st.markdown("""
Silakan pilih menu di **Sidebar sebelah kiri** untuk mulai menggunakan aplikasi:

1. **METAR Converter** ✈️ 
   
2. **Klimat Average** 📊
   
3. **Histori Pilot** 🎈
   

---
*Catatan: Pastikan format file Anda adalah CSV (.csv).*
""")
