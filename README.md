# Shopee Price Analyzer 🛒🔍

![Python Version](https://img.shields.io/badge/python-3.9+-blue.svg)
![Selenium](https://img.shields.io/badge/selenium-webdriver-green.svg)
![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Status](https://img.shields.io/badge/status-active-brightgreen.svg)

**Shopee Price Analyzer** adalah *tool* Python otomatis untuk riset harga kompetitor di Shopee secara massal. Dilengkapi metode *scraping* visual DOM untuk menghindari *blocking*, filter statistik IQR untuk membuang harga palsu/tidak wajar, dan *export* Excel guna keperluan optimasi strategi harga e-commerce.

---

## 💡 Latar Belakang

Proses riset harga kompetitor di Shopee yang dilakukan secara manual — copy-paste satu per satu untuk ratusan SKU atau keyword — sangat memakan waktu dan rentan *human error*. Tool ini dibuat untuk mengotomasi seluruh proses tersebut, sehingga riset yang biasanya memakan waktu berjam-jam bisa selesai dalam hitungan menit dengan data yang lebih bersih dan terstruktur.

---

## ✨ Fitur Utama

| Fitur | Deskripsi |
|---|---|
| 🛡️ **Anti-Block Scraping** | Menggunakan `undetected-chromedriver` untuk meniru perilaku manusia, menghindari deteksi bot |
| 📈 **Filter Statistik IQR** | Otomatis membuang harga *outlier* agar rata-rata pasar yang dihasilkan lebih akurat |
| 📦 **Single & Bulk Mode** | Analisa satu produk secara manual, atau ratusan SKU sekaligus via file `.txt` |
| 🏪 **Smart Own-Shop Detection** | Otomatis mendeteksi dan mengecualikan toko milik sendiri dari daftar kompetitor |
| 📊 **Excel Report Generation** | Ekspor laporan `.xlsx` lengkap dengan 3 *sheet*: Summary, Top 15 Terlaris, Raw Data |
| 🎨 **Rich Terminal UI** | Tampilan statistik, CVR, dan rekomendasi Shopee Ads yang berwarna di terminal |

---

## 🛠️ Prasyarat (Prerequisites)

Pastikan kamu sudah menginstal:
- **Python 3.9** atau lebih baru
- **Google Chrome** versi terbaru

Instal semua dependensi yang dibutuhkan:

```bash
pip install rich undetected-chromedriver openpyxl
```

---

## 🚀 Cara Penggunaan

**1. Clone repositori ini:**

```bash
git clone https://github.com/ilhamrofiqi/Shopee-Price-Analyzer.git
cd Shopee-Price-Analyzer
```

**2. Jalankan skrip utama:**

```bash
python shopee_price_analyzer.py
```

**3. Pilih Mode Input saat program berjalan:**

- **Mode 1 — Single:** Masukkan nama produk/keyword dan target harga jualmu secara manual di terminal.
- **Mode 2 — Bulk:** Buat file `produk.txt` di folder yang sama dengan format `Keyword | Harga Jual`. Contoh:

```plaintext
Power Supply MSI MAG A650BN 650W | 751000
Power Supply FSP HV PRO 550W 85+ | 683000
Memory RAM Kingston Fury Beast 16GB | 850000
```

**4.** Tunggu proses analisis. Skrip akan otomatis membuka Chrome, menarik data, dan menampilkan hasil statistik langsung di layar.

**5.** Di akhir proses, ketik `y` untuk menyimpan hasil analisis ke laporan Excel (`.xlsx`).

---

## 📊 Output Analisis

Tool ini menghasilkan *actionable insights* yang siap digunakan untuk pengambilan keputusan harga:

- **Statistik Kompetitor (Bersih dari Outlier)** — nilai Min, Q1, Median, Average, Q3, dan Max
- **Analisa Posisi Harga** — menilai apakah hargamu *Sangat Murah*, *Rata-rata Pasar*, atau *Mahal*
- **Rekomendasi Shopee Ads** — kelayakan iklan berdasarkan *Expected Conversion Rate* (CVR)
- **Top 15 Kompetitor Terlaris** — peringkat kompetitor beserta selisih harga terhadap produkmu

---

## 🧰 Tech Stack

| Library | Fungsi |
|---|---|
| `undetected-chromedriver` | Anti-bot scraping via visual DOM |
| `selenium` | Otomasi browser |
| `openpyxl` | Generate laporan Excel |
| `rich` | Terminal UI yang berwarna dan terstruktur |

---

## 📄 Lisensi

Didistribusikan di bawah **MIT License**. Lihat file `LICENSE` untuk detail lebih lanjut.

## ⚠️ Disclaimer

Project ini dibuat untuk tujuan riset internal, pembelajaran *data analytics*, dan optimasi strategi e-commerce. *Web scraping* harus dilakukan secara bertanggung jawab dengan mematuhi Ketentuan Layanan (*Terms of Service*) dari platform terkait. Pengembang tidak bertanggung jawab atas penyalahgunaan *tool* ini.
