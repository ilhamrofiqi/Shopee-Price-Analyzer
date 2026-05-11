# Shopee Price Analyzer 🛒🔍

![Python Version](https://img.shields.io/badge/python-3.9+-blue.svg)
![Selenium](https://img.shields.io/badge/selenium-webdriver-green.svg)
![License](https://img.shields.io/badge/license-MIT-blue.svg)

**Shopee Price Analyzer** adalah *tool* otomatisasi berbasis Python yang dirancang untuk melakukan riset pasar, memantau harga kompetitor, dan memberikan rekomendasi optimasi harga di platform Shopee. 

Skrip ini melakukan *scraping* secara visual dari DOM untuk menghindari pemblokiran API, kemudian membersihkan data dari harga palsu atau tidak wajar menggunakan metode statistik **IQR (Interquartile Range)**, dan menghasilkan laporan intelijen pasar yang komprehensif.

## ✨ Fitur Utama

- 🛡️ **Anti-Block Scraping:** Menggunakan `undetected-chromedriver` untuk meniru perilaku manusia, membaca elemen DOM secara visual (mengabaikan restriksi API tradisional).
- 📈 **Filter Statistik IQR:** Secara otomatis mendeteksi dan membuang *outlier* (seperti harga aksesoris receh atau produk *bundling* mahal) agar rata-rata harga pasar akurat.
- 📦 **Single & Bulk Input Mode:** Bisa mengecek satu per satu, atau langsung mengecek ratusan SKU sekaligus menggunakan file `.txt`.
- 🏪 **Smart Own-Shop Detection:** Dapat mendeteksi toko milikmu sendiri di antara deretan kompetitor (berdasarkan validasi kecocokan selisih harga dan domisili toko, misal: Surabaya).
- 📊 **Excel Report Generation:** Mengekspor laporan lengkap ke format `.xlsx` dengan 3 sheet rapi (Summary Analisa, Top 15 Terlaris, dan Raw Data).
- 🎨 **Rich Terminal UI:** Tampilan di terminal/CMD sangat cantik, terstruktur, dan berwarna berkat library `rich`.

## 🛠️ Prasyarat (Prerequisites)

Pastikan kamu sudah menginstal **Python 3.9** atau lebih baru. Kemudian, instal semua *library* yang dibutuhkan dengan menjalankan perintah berikut di terminal:

```bash
pip install rich undetected-chromedriver openpyxl
