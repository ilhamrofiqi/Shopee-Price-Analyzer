# Shopee Price Analyzer 🛒🔍

![Python Version](https://img.shields.io/badge/python-3.9+-blue.svg)
![Selenium](https://img.shields.io/badge/selenium-webdriver-green.svg)
![License](https://img.shields.io/badge/license-MIT-blue.svg)

**Shopee Price Analyzer** adalah *tool* Python otomatis untuk riset harga kompetitor di Shopee secara massal. Dilengkapi metode *scraping* visual DOM untuk menghindari *blocking*, filter statistik IQR untuk membuang harga palsu/tidak wajar, dan *export* Excel guna keperluan optimasi strategi harga e-commerce.

## ✨ Fitur Utama

- 🛡️ **Anti-Block Scraping:** Menggunakan `undetected-chromedriver` untuk meniru perilaku manusia dan membaca elemen DOM secara visual (mengabaikan restriksi API tradisional).
- 📈 **Filter Statistik IQR:** Secara otomatis mendeteksi dan membuang *outlier* (seperti harga aksesoris receh atau produk *bundling* super mahal) agar perhitungan rata-rata pasar lebih akurat.
- 📦 **Single & Bulk Input Mode:** Mengecek satu produk secara manual, atau menganalisa ratusan SKU sekaligus dengan mengimpor file `.txt`.
- 🏪 **Smart Own-Shop Detection:** Otomatis mendeteksi toko milikmu sendiri di antara daftar kompetitor (memvalidasi selisih harga 0 dan lokasi/domisili toko).
- 📊 **Excel Report Generation:** Mengekspor laporan riset pasar komprehensif ke format `.xlsx` yang terdiri dari 3 *sheet* (Summary Analisa, Top 15 Terlaris, dan Raw Data).
- 🎨 **Rich Terminal UI:** Menampilkan laporan statistik, CVR (*Conversion Rate*), dan rekomendasi Shopee Ads di terminal secara visual dan berwarna berkat library `rich`.

## 🛠️ Prasyarat (Prerequisites)

Pastikan kamu sudah menginstal **Python 3.9** atau lebih baru, dan memiliki *browser* Google Chrome versi terbaru. 

Instal semua dependensi (*library*) yang dibutuhkan dengan menjalankan perintah berikut di terminal:

```bash
pip install rich undetected-chromedriver openpyxl
```

## 🚀 Cara Penggunaan

1. Clone repositori ini:

```Bash
git clone [https://github.com/ilhamrofiqi/Shopee-Price-Analyzer.git](https://github.com/ilhamrofiqi/Shopee-Price-Analyzer.git)
cd Shopee-Price-Analyzer
```

2. Jalankan skrip utama:

```Bash
python shopee_price_analyzer.py
```

3. Pilih Mode Input saat program berjalan:

- Mode 1 (Single): Masukkan nama produk/keyword dan target harga jualmu secara manual di terminal.

- Mode 2 (Bulk): Buat file teks bernama produk.txt di folder yang sama. Isi file tersebut dengan format Keyword | Harga Jual Kamu. Contoh:

```Plaintext
Power Supply MSI MAG A650BN 650W | 751000
Power Supply FSP HV PRO 550W 85+ | 683000
Memory RAM Kingston Fury Beast 16GB | 850000
```

4. Tunggu proses analisis berjalan. Skrip akan otomatis membuka Chrome, menarik data, dan menampilkan hasil statistik langsung di layar.

5. Di akhir proses, ketik y jika kamu ingin menyimpan keseluruhan analisis ke dalam format laporan Excel (.xlsx).

## 📊 Output Analisis
Tool ini akan menyajikan metrik bisnis yang langsung dapat ditindaklanjuti (actionable insights):

- Statistik Kompetitor (Bersih dari Outlier): Menghitung nilai Min, Q1, Median, Average, Q3, dan Max.

- Analisa Posisi Harga: Menilai apakah hargamu "Sangat Murah", "Rata-rata Pasar", atau "Mahal".

- Rekomendasi Shopee Ads: Memberikan analisis kelayakan iklan berdasarkan Expected Conversion Rate (CVR) dari persaingan harga.

- Top 15 Kompetitor (Terlaris): Mengurutkan kompetitor penguasa pasar di keyword tersebut beserta selisih harga produk mereka dengan produkmu.

##⚠️ Disclaimer
Project ini dibuat khusus untuk tujuan riset internal, pembelajaran data analytics, dan optimasi strategi e-commerce. Web scraping harus dilakukan secara bertanggung jawab dengan mematuhi Ketentuan Layanan (Terms of Service) dari platform terkait. Pengembang tidak bertanggung jawab atas penyalahgunaan tool ini.
