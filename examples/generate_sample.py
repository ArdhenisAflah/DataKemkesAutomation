"""
generate_sample.py
------------------
Buat file sample_data.xlsx dengan format identik ekspor SIMPUS Kemkes.
File ini aman di-commit ke GitHub — semua data adalah DUMMY (fiktif).

Cara pakai:
    python examples/generate_sample.py

Output:
    examples/sample_data.xlsx
"""

import os
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'sample_data.xlsx')

# ---------------------------------------------------------------
# DATA DUMMY — Semua data di bawah ini FIKTIF / tidak nyata
# NIK mengikuti format resmi: BBBBBBDDMMYYXXXX
#   BB = kode wilayah (3674 = Kota Tangsel, 05 = Kec. Serpong)
#   DD = tanggal lahir + 40 untuk perempuan
#   MM = bulan lahir
#   YY = 2 digit tahun lahir
#   XXXX = nomor urut
# ---------------------------------------------------------------
PASIEN_DUMMY = [
    # (No, Nama Pasien, NIK, Tgl. Kunjungan, Tgl. Lahir, L/P,
    #  No. Telp, Kode ICD, Diagnosa, Pekerjaan, Pendidikan,
    #  Catatan Dokter, Anamnesa, Wilayah Kerja, Kelurahan, Kota)
    (
        1, "SITI RAHAYU", "3674054502950001",
        "2025-04-07", "1995-02-05", "P",
        "081234567890", "Z34.0", "Pengawasan kehamilan normal, tidak ditentukan",
        "Tidak Bekerja", "SMA",
        "UK 12 mg. TD 110/70. Kondisi ibu dan janin baik.",
        "Hamil 12 minggu, tidak ada keluhan berarti.",
        "SERPONG", "SERPONG", "TANGERANG SELATAN"
    ),
    (
        2, "DEWI PUSPITASARI", "3674055208930002",
        "2025-04-10", "1993-08-12", "P",
        "082198765432", "Z34.8", "Pengawasan kehamilan normal lainnya",
        "Bekerja", "D3/S1",
        "UK 28 mg. Letak kepala. DJJ 142x/mnt.",
        "Hamil 28 minggu, sering pusing ringan.",
        "SERPONG", "RAWABUNTU", "TANGERANG SELATAN"
    ),
    (
        3, "NURUL HIDAYAH", "3674056011960003",
        "2025-04-14", "1996-11-20", "P",
        "085612345678", "Z34.9", "Pengawasan kehamilan normal, tidak ditentukan",
        "Tidak Bekerja", "SMA",
        "Hamil 20 minggu. ANC rutin ke-4. BB naik 3 kg.",
        "Kehamilan 20 minggu, mual sudah berkurang.",
        "SERPONG", "CILENGGANG", "TANGERANG SELATAN"
    ),
    (
        4, "FITRI HANDAYANI", "3674054304980004",
        "2025-04-21", "1998-04-03", "P",
        "087765432198", "Z34.0", "Pengawasan kehamilan normal trimester pertama",
        "Bekerja", "SMA",
        "UK 8 mgg. HPHT 15/2/2025. Mual ringan.",
        "Hamil 8 minggu, dikeluhkan mual pagi hari.",
        "SERPONG", "BUARAN", "TANGERANG SELATAN"
    ),
    (
        5, "RINA WATI", "3674056807940005",
        "2025-04-28", "1994-07-28", "P",
        "081398765432", "Z34.8", "Pengawasan kehamilan normal",
        "Tidak Bekerja", "SMP",
        "UK 35-36 mg. Letak kepala sudah masuk PAP. Persiapan persalinan.",
        "Hamil tua, sudah terasa kencang-kencang tidak teratur.",
        "SERPONG", "LENGKONG GUDANG", "TANGERANG SELATAN"
    ),
    (
        6, "ANI SURYANI", "3674055501970006",
        "2025-05-05", "1997-01-15", "P",
        "082211223344", "Z34.0", "Pengawasan kehamilan normal, trimester pertama",
        "Tidak Bekerja", "SMA",
        "hamil 15 mg, mual mulai berkurang, BB stabil.",
        "G2P1A0, hamil 15 minggu, ANC ke-3.",
        "SERPONG", "RAWA MEKARJAYA", "TANGERANG SELATAN"
    ),
    (
        7, "DIAN PERTIWI", "3674054903990007",
        "2025-05-12", "1999-03-09", "P",
        "089944556677", "Z34.9", "Pengawasan kehamilan normal",
        "Mahasiswa", "D3/S1",
        "UK 24 mg. BB naik 2 kg dari kunjungan lalu. DJJ normal.",
        "Hamil 24 minggu, tidak ada keluhan.",
        "SERPONG", "CIATER", "TANGERANG SELATAN"
    ),
    (
        8, "SARI INDAH LESTARI", "3674056206920008",
        "2025-05-19", "1992-06-22", "P",
        "081234509876", "Z34.8", "Pengawasan kehamilan normal lainnya",
        "Bekerja", "D3/S1",
        "h 38 mgg. Sudah ada pembukaan 1 cm. Rencana rujuk RS jika HIS kuat.",
        "Hamil 38 minggu, sudah ada tanda-tanda persalinan awal.",
        "SERPONG", "LENGKONG WETAN", "TANGERANG SELATAN"
    ),
    (
        9, "MEGA KURNIAWATI", "3674055710950009",
        "2025-05-26", "1995-10-17", "P",
        "087800112233", "Z34.0", "Pengawasan kehamilan normal",
        "Tidak Bekerja", "SMA",
        "hamil 10 minggu. Trimester 1. USG: janin sesuai usia kehamilan.",
        "G1P0A0. Hamil pertama 10 minggu. Khawatir karena mual hebat.",
        "SERPONG", "SERPONG", "TANGERANG SELATAN"
    ),
    (
        10, "YUNI ASTUTI", "3674057012960010",
        "2025-06-02", "1996-12-30", "P",
        "-", "Z34.9", "Pengawasan kehamilan normal",
        "Tidak Bekerja", "SMP",
        "UK 32 mg. Cek rutin. Tekanan darah normal.",
        "Hamil 32 minggu. Keluhan punggung pegal.",
        "SERPONG", "SERPONG", "TANGERANG SELATAN"
    ),
]

# ---------------------------------------------------------------
# BUAT WORKBOOK
# ---------------------------------------------------------------
wb = Workbook()
ws = wb.active
ws.title = "Data Kunjungan"

# --- Warna & style ---
BIRU_TUA   = "1F4E79"
BIRU_MUDA  = "D6E4F0"
KUNING     = "FFF2CC"
ABU        = "F2F2F2"

def cell_style(cell, bold=False, bg=None, align="left", wrap=False, font_color="000000", size=10):
    cell.font = Font(bold=bold, color=font_color, size=size)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if bg:
        cell.fill = PatternFill(fill_type="solid", fgColor=bg)

thin = Side(style="thin", color="BFBFBF")
border_thin = Border(left=thin, right=thin, top=thin, bottom=thin)

# ---------------------------------------------------------------
# BARIS 1–5: Header metadata SIMPUS (sesuai format ekspor asli)
# Script membaca data mulai header=5 (baris ke-6), jadi 5 baris
# pertama ini adalah metadata yang dilewati oleh pandas.
# ---------------------------------------------------------------
ws.merge_cells("A1:P1")
ws["A1"] = "LAPORAN DATA KUNJUNGAN PASIEN — SKRINING KESEHATAN JIWA (EPDS)"
cell_style(ws["A1"], bold=True, bg=BIRU_TUA, align="center", font_color="FFFFFF", size=13)
ws.row_dimensions[1].height = 22

ws.merge_cells("A2:P2")
ws["A2"] = "PUSKESMAS SERPONG II — DINAS KESEHATAN KOTA TANGERANG SELATAN"
cell_style(ws["A2"], bold=True, bg=BIRU_TUA, align="center", font_color="FFFFFF", size=11)
ws.row_dimensions[2].height = 18

ws.merge_cells("A3:P3")
ws["A3"] = "PERIODE: JANUARI — JUNI 2025"
cell_style(ws["A3"], bold=False, bg=BIRU_MUDA, align="center", size=10)
ws.row_dimensions[3].height = 16

ws.merge_cells("A4:P4")
ws["A4"] = f"Dicetak: {date.today().strftime('%d %B %Y')}    |    Sumber: SIMPUS Web Kemkes RI    |    ⚠ File ini berisi DATA DUMMY — aman untuk demo dan pengujian"
cell_style(ws["A4"], bold=False, bg=KUNING, align="center", size=9)
ws.row_dimensions[4].height = 16

ws.merge_cells("A5:P5")
ws["A5"] = ""
ws.row_dimensions[5].height = 6

# ---------------------------------------------------------------
# BARIS 6: Header kolom (ini yang terbaca oleh pandas dengan header=5)
# ---------------------------------------------------------------
HEADERS = [
    "No", "Nama Pasien", "NIK", "Tgl. Kunjungan", "Tgl. Lahir",
    "L/P", "No. Telp", "Kode ICD", "Diagnosa",
    "Pekerjaan", "Pendidikan", "Catatan Dokter", "Anamnesa",
    "Wilayah Kerja", "Kelurahan", "Kota"
]
for col_idx, header in enumerate(HEADERS, 1):
    cell = ws.cell(row=6, column=col_idx, value=header)
    cell_style(cell, bold=True, bg=BIRU_TUA, align="center", font_color="FFFFFF")
    cell.border = border_thin
ws.row_dimensions[6].height = 20

# ---------------------------------------------------------------
# BARIS 7+: Data pasien dummy
# ---------------------------------------------------------------
for row_idx, pasien in enumerate(PASIEN_DUMMY, 7):
    bg_color = ABU if row_idx % 2 == 0 else "FFFFFF"
    for col_idx, value in enumerate(pasien, 1):
        cell = ws.cell(row=row_idx, column=col_idx, value=value)
        wrap = col_idx in (12, 13)   # Catatan Dokter & Anamnesa boleh wrap
        align = "center" if col_idx in (1, 6, 8) else "left"
        cell_style(cell, bg=bg_color, align=align, wrap=wrap)
        cell.border = border_thin
    ws.row_dimensions[row_idx].height = 32 if row_idx % 2 != 0 else 32

# ---------------------------------------------------------------
# LEBAR KOLOM (disesuaikan konten)
# ---------------------------------------------------------------
COLUMN_WIDTHS = {
    1: 5,    # No
    2: 22,   # Nama Pasien
    3: 18,   # NIK
    4: 14,   # Tgl. Kunjungan
    5: 14,   # Tgl. Lahir
    6: 5,    # L/P
    7: 15,   # No. Telp
    8: 10,   # Kode ICD
    9: 35,   # Diagnosa
    10: 16,  # Pekerjaan
    11: 12,  # Pendidikan
    12: 40,  # Catatan Dokter
    13: 40,  # Anamnesa
    14: 14,  # Wilayah Kerja
    15: 18,  # Kelurahan
    16: 20,  # Kota
}
for col_idx, width in COLUMN_WIDTHS.items():
    ws.column_dimensions[get_column_letter(col_idx)].width = width

# Freeze panes: beku di baris 7 kolom C (scroll data tapi header tetap)
ws.freeze_panes = "C7"

# ---------------------------------------------------------------
# SIMPAN
# ---------------------------------------------------------------
wb.save(OUTPUT_PATH)
print(f"[OK] Sample data berhasil dibuat: {OUTPUT_PATH}")
print(f"     Total baris data  : {len(PASIEN_DUMMY)}")
print(f"     Header ada di baris ke-6 (header=5 untuk pandas)")
print(f"     Siap dipakai oleh: scripts/epds_autofill_v2.py")
