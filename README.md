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

Catatan: Pastikan browser Google Chrome versi terbaru sudah terinstal di komputer/laptop kamu.

🚀 Cara Penggunaan
Clone repositori ini:

Bash
git clone [https://github.com/ilhamrofiqi/shopee_price_analyzer.git](https://github.com/ilhamrofiqi/shopee_price_analyzer.git)
cd shopee_price_analyzer
Jalankan skrip utama:

Bash
python shopee_price_analyzer.py
Pilih Mode Input:
Saat program berjalan, kamu akan diminta memilih mode:

Mode 1 (Single): Masukkan nama produk/keyword dan harga jualmu secara manual di terminal.

Mode 2 (Bulk): Buat file teks (contoh: produk.txt) di folder yang sama. Isi file tersebut menggunakan format Nama Keyword | Harga Jual Kamu. Contoh isi file produk.txt:

Plaintext
Power Supply MSI MAG A650BN 650W | 751000
Power Supply FSP HV PRO 550W 85+ | 683000
Memory RAM Kingston Fury Beast 16GB | 850000
Tunggu proses berjalan. Skrip akan membuka Chrome (otomatis scroll untuk lazy-loading), menarik data, melakukan perhitungan statistik, dan menampilkannya di terminal.

Setelah selesai, tekan y jika kamu ingin menyimpan hasil analisa ke format Excel.

📊 Contoh Hasil Analisa
Tool ini akan memberikan matriks dan rekomendasi detail seperti:

Statistik Kompetitor: Nilai Minimum, Q1, Median, Average (Bersih tanpa Outlier), Q3, dan Maksimum.

Posisi Harga: Status apakah harga kamu "Sangat Murah", "Rata-rata Pasar", atau "Mahal".

Rekomendasi Shopee Ads: Penilaian apakah produk aman dan memiliki potensi Conversion Rate (CVR) yang baik jika diiklankan.

Top 15 Terlaris: Menampilkan produk kompetitor yang paling banyak terjual untuk dianalisa lebih lanjut (foto, deskripsi, voucher yang mereka pakai).

⚠️ Disclaimer
Project ini dibuat khusus untuk tujuan riset internal, pembelajaran analitik data, dan optimasi e-commerce. Web scraping harus dilakukan secara bertanggung jawab dan mematuhi Ketentuan Layanan (Terms of Service) dari platform terkait. Jangan membebani server secara masif (lakukan delay yang wajar). Pengembang tidak bertanggung jawab atas penyalahgunaan tool ini.
