import streamlit as st

st.set_page_config(
    page_title="Portal Data Cuaca",
    page_icon="⛅",
    layout="centered"
)

st.title("⛅ Selamat Datang di Portal Data Cuaca")
st.write("Silakan pilih alat yang ingin Anda gunakan melalui menu di sebelah kiri (sidebar):")
st.markdown("""
- **METAR Converter**: Untuk memparsing data sandi METAR ke PDF & Excel.
- **Klimat Average**: Untuk menghitung rata-rata parameter klimatologi dan mengekspor matriks Excel.
""")
