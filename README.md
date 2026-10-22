# DataKemkes Automation

**Otomatisasi pengisian form skrining EPDS (Edinburgh Postnatal Depression Scale) pada portal SIM-Keswa Kemenkes RI**

[![Python](https://img.shields.io/badge/Python-3.8+-blue?logo=python&logoColor=white)](https://python.org)
[![Selenium](https://img.shields.io/badge/Selenium-4.x-green?logo=selenium&logoColor=white)](https://selenium.dev)
[![Pandas](https://img.shields.io/badge/Pandas-2.x-darkblue?logo=pandas&logoColor=white)](https://pandas.pydata.org)
[![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white)](https://microsoft.com/windows)
[![License](https://img.shields.io/badge/License-MIT-yellow)](LICENSE)

---

## Deskripsi

Project ini adalah tools otomatisasi berbasis **Python + Selenium** untuk mengisi form skrining kesehatan jiwa ibu hamil (EPDS) di portal [SIM-Keswa Kemenkes RI](https://simkeswa.kemkes.go.id).

Data diambil dari file Excel hasil ekspor **SIMPUS**, difilter berdasarkan kode ICD Z34 (kehamilan normal), lalu diproses dan diinputkan secara otomatis ke sistem — termasuk pengisian kuesioner EPDS 10 soal, data domisili, dan informasi demografis pasien.

> **Disclaimer**: Tools ini dibuat untuk keperluan operasional internal Puskesmas. Penggunaan harus sesuai dengan kebijakan dan kewenangan institusi terkait.

---

## Fitur

| Fitur | Deskripsi |
|-------|-----------|
| Auto-fill Form | Mengisi form skrining EPDS secara otomatis dari data Excel |
| Filter Z34.* | Otomatis memfilter pasien ibu hamil berdasarkan kode ICD Z34 |
| Template EPDS | 3 template jawaban normal yang dipilih acak per pasien |
| Parse Tanggal | Mendukung berbagai format tanggal dari Excel |
| Deteksi Trimester | Auto-deteksi usia kehamilan dari catatan dokter untuk penentuan kunjungan (k1–k6) |
| Logging Detail | Log lengkap setiap proses ke console dan file `.txt` |
| Resume | Dapat dilanjutkan dari baris tertentu jika proses terputus |
| Auto-Delete | Hapus data dari sistem secara batch dengan konfirmasi otomatis |
| Data Cleaning | Preprocessing file Excel sebelum diproses |

---

## Struktur Project

```
DataKemkesAutomation/
│
├── scripts/
│   ├── epds_autofill_v2.py     ← Script utama (gunakan ini)
│   ├── epds_autofill_v1.py     ← Script v1 (legacy, referensi)
│   ├── epds_auto_delete.py     ← Hapus data batch dari SIM-Keswa
│   ├── clean_the_file.py       ← Preprocessing / bersihkan data Excel
│   └── test_koneksi.py         ← Uji koneksi dan baca struktur file
│
├── data/                        ← File Excel data pasien (gitignored)
│   └── .gitkeep
│
├── logs/                        ← Output log dari script (gitignored)
│   └── .gitkeep
│
├── docs/                        ← Gambar & aset dokumentasi
│   ├── contoh_input/
│   │   └── image.png
│   ├── demo.png
│   └── screenshot.png
│
├── examples/                    ← Data dummy untuk demo & pengujian
│   ├── generate_sample.py       ← Script pembuat file contoh
│   └── sample_data.xlsx         ← File Excel contoh (data FIKTIF)
│
├── .env.example                 ← Template konfigurasi kredensial
├── requirements.txt
├── .gitignore
└── README.md
```

---

## Requirements

- Python 3.8 atau lebih baru
- Google Chrome (versi terbaru)
- ChromeDriver (versi harus cocok dengan Chrome yang terinstal)

---

## Instalasi

### 1. Clone repository

```bash
git clone https://github.com/ArdhenisAflah/DataKemkesAutomation.git
cd DataKemkesAutomation
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Setup ChromeDriver

Download ChromeDriver yang sesuai dengan versi Google Chrome kamu:

- Cek versi Chrome: buka `chrome://version/` di browser
- Download di: https://chromedriver.chromium.org/downloads
- Letakkan `chromedriver.exe` di root folder project

### 4. Siapkan file `.env`

```bash
copy .env.example .env
```

Lalu isi kredensial SIM-Keswa kamu di file `.env`.

---

## Coba dengan Data Dummy

Tersedia file contoh di folder `examples/` yang aman untuk pengujian — semua data adalah **fiktif**:

```bash
# Generate file contoh Excel
python examples/generate_sample.py
# → menghasilkan: examples/sample_data.xlsx
```

File ini memiliki struktur **identik** dengan ekspor SIMPUS asli (termasuk 5 baris metadata di atas, header kolom di baris ke-6, dan format NIK perempuan dengan DD+40). Kamu bisa langsung mengarahkan script ke file ini untuk uji coba:

```python
# Di scripts/epds_autofill_v2.py, ubah baris ini:
FILE_DATA_PATH = os.path.join(_ROOT, 'examples', 'sample_data.xlsx')
```

---

## Penggunaan

### Auto-fill Form EPDS (v2 — Direkomendasikan)

**1. Taruh file Excel kamu di folder `data/`**, lalu edit baris ini di `scripts/epds_autofill_v2.py`:

```python
FILE_DATA_PATH = os.path.join(_ROOT, 'data', 'nama_file_kamu.xlsx')
START_INDEX    = 0   # 0 = mulai dari awal | N = resume dari baris ke-N+1
```

> Path log dan root project dihitung otomatis — tidak perlu diubah.

**2. Format Excel yang diharapkan:**

| Kolom Wajib | Keterangan |
|-------------|------------|
| `Nama Pasien` | Nama lengkap pasien |
| `NIK` | 16 digit NIK |
| `Tgl. Kunjungan` | Tanggal kunjungan (berbagai format didukung) |
| `Tgl. Lahir` | Tanggal lahir pasien |
| `Kode ICD` | Harus diawali `Z34` agar diproses |

> Header Excel harus berada di **baris ke-6** (baris index 5). Ini sesuai format ekspor SIMPUS.

**3. Jalankan:**

```bash
python scripts/epds_autofill_v2.py
```

Browser Chrome akan terbuka otomatis dan mulai mengisi form satu per satu.

---

### Bersihkan Data Excel

```bash
python scripts/clean_the_file.py
```

Membaca file Excel mentah dan menghasilkan `file_bersih.xlsx` yang sudah difilter.

---

### Auto-Delete Data (Hati-hati!)

```bash
python scripts/epds_auto_delete.py
```

Script akan login ke SIM-Keswa menggunakan kredensial dari `.env`, lalu menghapus data secara batch. Ubah `MAX_DELETIONS` di script untuk membatasi jumlah data yang dihapus.

---

## Contoh Output Log

```
=================================================================
EPDS Autofill v2 — Mulai: 2025-06-16 09:00:00
File data : D:\...\baru.xlsx
=================================================================
Membaca file Excel...
Baris Z34*  : 250 rows
Baris valid : 248 rows

=================================================================
[Baris 1/248]  NAMA: SITI RAHAYU  |  Kode ICD: Z34.0

  Nama              : SITI RAHAYU
  NIK               : 3674XXXXXXXXXXXX
  Tanggal Skrining  : 2025-05-10
  Tanggal Lahir     : 1995-03-22
  UK (minggu)       : 32  →  kunjungan: k5
  Template EPDS     : Normal_5  →  [0,0,2,2,3,2,2,3,3,3]  (estimasi skor: 5)
  ✓ NIK              : 3674XXXXXXXXXXXX
  ✓ Tanggal Skrining : 2025-05-10
  ...
  ✅ Form terisi lengkap. Menyimpan...
  ✅ Tombol Simpan diklik.

=================================================================
SELESAI  |  Sukses: 248  |  Skip: 0  |  Error: 0
```

---

## Template EPDS

Script menggunakan 3 template jawaban yang dipilih secara acak per pasien. Semua template menghasilkan skor di bawah threshold depresi (< 10).

| Template | Estimasi Skor | Jawaban per Soal |
|----------|:------------:|-----------------|
| `Normal_3` | 3 | `[0,0,2,2,3,3,3,3,3,3]` |
| `Normal_5` | 5 | `[0,0,2,2,3,2,2,3,3,3]` |
| `Normal_6` | 6 | `[0,0,2,2,2,2,2,3,3,3]` |

---

## Catatan Penting

- Pastikan **tidak menutup browser** saat script berjalan
- Jika proses terputus, set `START_INDEX` ke nomor baris terakhir yang berhasil lalu jalankan ulang
- File Excel data pasien **tidak diupload ke GitHub** — bersifat rahasia/sensitif
- Versi ChromeDriver **harus sesuai** dengan versi Google Chrome yang terinstall
- Script ini dikonfigurasi untuk Puskesmas percontohan — field alamat, institusi, dan wilayah dapat disesuaikan pada file konfigurasi / script

---

## Tech Stack

- [Python](https://python.org) — Bahasa pemrograman utama
- [Selenium](https://selenium.dev) — Otomatisasi browser
- [Pandas](https://pandas.pydata.org) — Baca dan manipulasi data Excel
- [OpenPyXL](https://openpyxl.readthedocs.io) — Dukungan baca file `.xlsx`

---

## Lisensi

[MIT](LICENSE) © ArdhenisAflah
