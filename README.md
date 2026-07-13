# Dashboard Keuangan IDX 📊

Aplikasi dashboard interaktif untuk visualisasi data laporan keuangan perusahaan yang terdaftar di **Bursa Efek Indonesia (BEI/IDX)**.

## 🚀 Live Demo
[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://projekkp.streamlit.app)

## 📋 Fitur
- **Overview** — KPI cards, Top 15 perusahaan, distribusi subsektor, heatmap ketersediaan data
- **Analisis Perusahaan** — Tren keuangan per perusahaan, waterfall breakdown laba
- **Tren Periode** — Perbandingan hingga 8 perusahaan, heatmap growth rate QoQ
- **Analisis Sektor** — Agregasi per subsektor, scatter plot, box plot distribusi

## 📊 Data
- **969 perusahaan** terdaftar di IDX
- **9 periode**: 2024 TW1 s/d 2026 TW1
- **5 variabel**: Pendapatan, Beban Penjualan, Laba Kotor, Laba Usaha, Penambahan Aset Tetap
- Sumber: Laporan Keuangan XBRL BEI

## 🛠️ Teknologi
- Python 3.11
- Streamlit
- Plotly
- Pandas

## 💻 Cara Menjalankan Lokal
```bash
pip install -r requirements.txt
streamlit run dashboard.py
```
