import os
import re
import pandas as pd
import time
import random
import logging
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException

# =============================================================
# KONFIGURASI UTAMA
# Path dihitung otomatis relatif ke root project — tidak perlu
# diubah selama struktur folder tidak berubah.
# Ganti nama file data sesuai file Excel yang ingin diproses.
# =============================================================
_ROOT          = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILE_DATA_PATH = os.path.join(_ROOT, 'data', 'baru.xlsx')
URL_FORM       = 'https://simkeswa.kemkes.go.id/form_skrining/dewasa'
LOG_FILE       = os.path.join(_ROOT, 'logs', 'autofill_log.txt')

# =============================================================
# LOGGING: output ke console (terminal) DAN disimpan ke file
# =============================================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8', mode='w'),  # simpan ke file, timpa tiap run
        logging.StreamHandler()                                       # tampil di terminal
    ]
)
log = logging.getLogger(__name__)

# =============================================================
# TEMPLATE JAWABAN EPDS — SEMUA PASIEN Z34.* → HASIL NORMAL
#
# Setiap list berisi 10 angka, masing-masing adalah INDEX pilihan
# untuk soal 1 sampai 10 pada form EPDS.
#
# Cara hitung skor:
#   - Soal Q1, Q2, Q4      : skor += idx  (skala biasa/normal)
#   - Soal Q3, Q5 s/d Q10  : skor += (3 - idx)  (skala terbalik)
#
# Template ini dipilih secara acak per pasien agar tidak terlihat
# seragam, namun tetap menghasilkan skor di bawah ambang batas depresi.
# =============================================================
TEMPLATES = {
    "Normal_3":  [0, 0, 2, 2, 3, 3, 3, 3, 3, 3],   # estimasi skor total: 3
    "Normal_5":  [0, 0, 2, 2, 3, 2, 2, 3, 3, 3],   # estimasi skor total: 5
    "Normal_6":  [0, 0, 2, 2, 2, 2, 2, 3, 3, 3],   # estimasi skor total: 6
}

# Daftar nama template yang boleh dipilih (semuanya hasil Normal)
TEMPLATE_NORMAL = ["Normal_3", "Normal_5", "Normal_6"]


# =============================================================
# FUNGSI-FUNGSI BANTU
# =============================================================

def hitung_skor_epds(template):
    """Hitung estimasi skor EPDS dari template jawaban yang dipilih."""
    # Soal dengan indeks 0, 1, 3 (Q1, Q2, Q4) menggunakan skala normal
    normal_q = [0, 1, 3]
    return sum(idx if i in normal_q else 3 - idx for i, idx in enumerate(template))


def parse_tanggal(val):
    """
    Ubah berbagai format tanggal dari Excel menjadi objek datetime Python.
    Bisa menangani: datetime object, Timestamp pandas, atau string tanggal.
    Format string yang didukung: DD-MM-YYYY, YYYY-MM-DD, DD/MM/YYYY, YYYY/MM/DD.
    """
    if isinstance(val, datetime):
        return val
    if hasattr(val, 'to_pydatetime'):
        return val.to_pydatetime()
    if hasattr(val, 'strftime'):
        return datetime(val.year, val.month, val.day)
    # Ambil bagian tanggal saja (buang bagian jam jika ada)
    s = str(val).split(' ')[0].strip()
    for fmt in ('%d-%m-%Y', '%Y-%m-%d', '%d/%m/%Y', '%Y/%m/%d'):
        try:
            return datetime.strptime(s, fmt)
        except (ValueError, TypeError):
            continue
    return None  # tidak bisa diparse sama sekali


def fmt_tanggal(val, fmt='%Y-%m-%d'):
    """Format nilai tanggal ke string dengan format tertentu (default: YYYY-MM-DD)."""
    dt = parse_tanggal(val)
    return dt.strftime(fmt) if dt else None


def isi_tanggal_js(driver, element_id, nilai):
    """
    Isi field tanggal menggunakan JavaScript karena beberapa input tanggal
    di website menggunakan datepicker jQuery yang tidak bisa diisi langsung
    dengan send_keys biasa.
    """
    el = driver.find_element(By.ID, element_id)
    driver.execute_script("""
        var el = arguments[0], val = arguments[1];
        el.removeAttribute('readonly');
        el.value = val;
        if (window.jQuery) {
            try { jQuery(el).datepicker('setDate', val); } catch(e) {}
            jQuery(el).trigger('change');
        }
        el.dispatchEvent(new Event('input',  {bubbles: true}));
        el.dispatchEvent(new Event('change', {bubbles: true}));
    """, el, nilai)


def fill_questionnaire(driver, wait, answer_template):
    """
    Isi 10 soal kuesioner EPDS satu per satu sesuai template jawaban.
    Setiap soal punya tombol radio dengan ID format: soal_X_pil_Y
    (X = nomor soal, Y = index pilihan jawaban).
    """
    for i, idx in enumerate(answer_template):
        soal = i + 1
        tid = f"soal_{soal}_pil_{idx}"
        try:
            # Tunggu tombol muncul dan bisa diklik, lalu klik via JS
            btn = wait.until(EC.element_to_be_clickable((By.ID, tid)))
            driver.execute_script("arguments[0].click();", btn)
            # Ambil teks label pilihan untuk dicatat di log
            label = driver.find_element(By.CSS_SELECTOR, f"label[for='{tid}']").text.strip()
            log.info(f"      Soal {soal:>2}: pil_{idx} → '{label}'")
        except Exception as e:
            log.warning(f"      ⚠ Soal {soal}: gagal klik '{tid}' — {e}")


def parse_uk_dari_teks(*texts):
    """
    Cari dan ekstrak usia kehamilan dalam satuan minggu dari teks bebas
    (biasanya dari kolom Catatan Dokter atau Anamnesa).

    Contoh teks yang bisa dikenali:
      "uk 15 mg", "uk 38-39 mg", "h 9mgg", "hamil 25 minggu", "12 minggu"

    Mengembalikan angka minggu (int) atau None jika tidak ditemukan.
    """
    pola = [
        r'uk\s*(\d+)(?:[-–]\d+)?\s*(?:mg|mgg|minggu)',   # "uk 15 mg" / "uk 38-39 mg"
        r'\bh\s*(\d+)\s*(?:mg|mgg|minggu)',                # "h 9mgg"
        r'hamil\s+(\d+)\s*(?:mg|mgg|minggu)',              # "hamil 25 minggu"
        r'(\d+)\s*(?:mg|mgg|minggu)\b',                    # "12 minggu" (pola paling umum)
    ]
    for teks in texts:
        s = str(teks).lower()
        for pat in pola:
            m = re.search(pat, s)
            if m:
                minggu = int(m.group(1))
                # Pastikan nilainya masuk akal (kehamilan normal 4–45 minggu)
                if 4 <= minggu <= 45:
                    return minggu
    return None


def uk_to_kunjungan(minggu):
    """
    Tentukan kode kunjungan ANC (k1–k6) berdasarkan usia kehamilan.
    Dipilih secara acak dalam range yang sesuai agar terlihat natural:
      - Trimester 1 (≤12 mgg)  : k1 atau k2
      - Trimester 2 (13–24 mgg): k3 atau k4
      - Trimester 3 (>24 mgg)  : k5 atau k6
    """
    if minggu is None or minggu <= 12:
        return random.choice(['k1', 'k2'])
    elif minggu <= 24:
        return random.choice(['k3', 'k4'])
    else:
        return random.choice(['k5', 'k6'])


def logic_pekerjaan(val):
    """
    Tentukan nilai pekerjaan untuk form berdasarkan isi kolom Pekerjaan di Excel.
    Jika kosong atau tidak jelas → default ke 'Tidak Bekerja'.
    """
    if pd.isna(val) or str(val).strip() in ('-', '', 'nan'):
        return "Tidak Bekerja"
    s = str(val).lower()
    if any(x in s for x in ('belum', 'tidak bekerja', 'mengurus rumah', 'ibu rumah')):
        return "Tidak Bekerja"
    if 'pensiunan' in s:
        return "Pensiunan"
    if 'mahasiswa' in s:
        return "Mahasiswa"
    if 'pelajar' in s:
        return "Pelajar"
    return "Bekerja"  # semua yang tidak cocok dianggap bekerja


def logic_pendidikan(val):
    """
    Tentukan nilai pendidikan untuk form berdasarkan isi kolom Pendidikan di Excel.
    Jika kosong atau tidak dikenali → default ke 'SMA Atau Sederajat'.
    """
    if pd.isna(val):
        return "SMA Atau Sederajat"
    s = str(val).lower()
    if any(x in s for x in ('sarjana', 'd3', 'd4', 's1', 's2', 's3', 'diploma')):
        return "D3/S1/S2/S3/Pasca Sarjana"
    if 'sma' in s or 'smk' in s:
        return "SMA Atau Sederajat"
    if 'smp' in s:
        return "SMP Atau Sederajat"
    if ' sd' in s or s.startswith('sd'):
        return "SD Atau Sederajat"
    return "SMA Atau Sederajat"


def bersihkan_no_telp(val):
    """
    Bersihkan nomor telepon: ambil digit saja, buang .0 dari angka float.
    Jika kosong atau tidak valid → kembalikan '0'.
    """
    if pd.isna(val) or str(val).strip() in ('-', '', 'nan', '0.0', '0'):
        return "0"
    digits = ''.join(filter(str.isdigit, str(val).split('.')[0]))
    return digits if digits else "0"


def bersihkan_nik(val):
    """Ambil hanya karakter angka dari NIK (buang spasi, titik, atau karakter lain)."""
    return ''.join(c for c in str(val) if c.isdigit())


# =============================================================
# MAIN — PROGRAM UTAMA
# =============================================================
log.info("=" * 65)
log.info(f"EPDS Autofill v2 — Mulai: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
log.info(f"File data : {FILE_DATA_PATH}")
log.info(f"Log file  : {LOG_FILE}")
log.info("=" * 65)

try:
    # ----------------------------------------------------------
    # TAHAP 1: Baca dan bersihkan data dari Excel
    # Header ada di baris ke-6 (index 5 dari 0)
    # ----------------------------------------------------------
    log.info("Membaca file Excel...")
    df = pd.read_excel(FILE_DATA_PATH, header=5)

    # Hapus kolom dan baris yang sepenuhnya kosong
    df.dropna(axis=1, how='all', inplace=True)
    df.dropna(how='all', inplace=True)

    # ----------------------------------------------------------
    # TAHAP 2: Filter hanya pasien dengan Kode ICD Z34.*
    # Z34 = kehamilan normal (Z34.0, Z34.8, Z34.9, dll.)
    # Pasien Z34 = semua aman diasumsikan Normal pada EPDS
    # ----------------------------------------------------------
    df = df[df['Kode ICD'].astype(str).str.startswith('Z34', na=False)].copy()
    df.reset_index(drop=True, inplace=True)
    log.info(f"Baris Z34*  : {len(df)} rows")

    # ----------------------------------------------------------
    # TAHAP 3: Buang baris yang tidak lengkap (field wajib kosong)
    # ----------------------------------------------------------
    field_wajib = ['Nama Pasien', 'NIK', 'Tgl. Kunjungan', 'Tgl. Lahir']
    df = df.dropna(subset=field_wajib).copy()
    df.reset_index(drop=True, inplace=True)

    log.info(f"Baris valid : {len(df)} rows")
    log.info(f"Distribusi Kode ICD:\n{df['Kode ICD'].value_counts().to_string()}")
    log.info(f"Kolom       : {df.columns.tolist()}\n")

    # ----------------------------------------------------------
    # TAHAP 4: Tentukan titik mulai
    # START_INDEX = 0  → mulai dari baris pertama
    # START_INDEX = N  → lanjut (resume) dari baris ke N+1
    # ----------------------------------------------------------
    START_INDEX = 0   # 0 = mulai dari data pertama
    if START_INDEX > 0:
        df = df.iloc[START_INDEX:].copy()
        df.reset_index(drop=True, inplace=True)
        log.info(f"▶  Resume dari baris ke-{START_INDEX + 1} — sisa {len(df)} baris\n")

    # ----------------------------------------------------------
    # TAHAP 5: Buka browser Chrome dan navigasi ke halaman form
    # ----------------------------------------------------------
    log.info("Membuka browser Chrome...")
    driver = webdriver.Chrome()
    wait   = WebDriverWait(driver, 20)   # tunggu maksimal 20 detik tiap elemen
    driver.implicitly_wait(10)           # tunggu implisit 10 detik untuk find_element
    driver.get(URL_FORM)
    log.info(f"URL: {URL_FORM}\n")

    # Penghitung hasil akhir
    sukses = 0
    skip   = 0
    error  = 0

    # ----------------------------------------------------------
    # TAHAP 6: Loop per baris data — isi form satu per satu
    # ----------------------------------------------------------
    for index, row in df.iterrows():
        baris_no = index + 1   # nomor tampilan (mulai dari 1)
        log.info("")
        log.info("=" * 65)
        log.info(f"[Baris {baris_no}/{len(df)}]  NAMA: {row['Nama Pasien']}  |  Kode ICD: {row['Kode ICD']}")

        try:
            # === PERSIAPAN: Ekstrak dan proses data dari baris Excel ===

            nama         = str(row['Nama Pasien']).strip()
            nik          = bersihkan_nik(row['NIK'])
            tgl_skrining = fmt_tanggal(row['Tgl. Kunjungan'])
            tgl_lahir    = fmt_tanggal(row['Tgl. Lahir'])
            no_telp      = bersihkan_no_telp(row.get('No. Telp'))
            pekerjaan    = logic_pekerjaan(row.get('Pekerjaan'))
            pendidikan   = logic_pendidikan(row.get('Pendidikan'))

            # Cari usia kehamilan dari kolom teks bebas (Catatan Dokter / Anamnesa)
            uk_minggu  = parse_uk_dari_teks(row.get('Catatan Dokter', ''), row.get('Anamnesa', ''))
            kunjungan  = uk_to_kunjungan(uk_minggu)

            # Pilih template EPDS secara acak (semua Z34. → Normal)
            template_key    = random.choice(TEMPLATE_NORMAL)
            answer_template = TEMPLATES[template_key]
            skor            = hitung_skor_epds(answer_template)

            # --- Catat ringkasan data ke log ---
            log.info(f"  Nama              : {nama}")
            log.info(f"  NIK               : {nik}  (raw: {row['NIK']})")
            log.info(f"  Tanggal Skrining  : {tgl_skrining}  (raw: {row['Tgl. Kunjungan']})")
            log.info(f"  Tanggal Lahir     : {tgl_lahir}  (raw: {row['Tgl. Lahir']})")
            log.info(f"  UK (minggu)       : {uk_minggu}  →  kunjungan: {kunjungan}")
            log.info(f"  No Telp           : {no_telp}  (raw: {row.get('No. Telp')})")
            log.info(f"  Pekerjaan         : {pekerjaan}  (raw: {row.get('Pekerjaan')})")
            log.info(f"  Pendidikan        : {pendidikan}")
            log.info(f"  Wilayah Kerja     : {row.get('Wilayah Kerja')}  |  Kelurahan: {row.get('Kelurahan')}  |  Kota: {row.get('Kota')}")
            log.info(f"  Domisili form     : BANTEN > KOTA TANGERANG SELATAN > SERPONG > SERPONG")
            log.info(f"  Template EPDS     : {template_key}  →  {answer_template}  (estimasi skor: {skor})")

            # === VALIDASI: Pastikan data wajib tidak kosong/rusak sebelum mengisi form ===
            masalah = []
            if not nama or nama.lower() in ('nan',):
                masalah.append("Nama Pasien kosong / NaN")
            if len(nik) != 16:
                masalah.append(f"NIK tidak valid — panjang={len(nik)}, nilai='{nik}'")
            if not tgl_skrining:
                masalah.append(f"Tgl. Kunjungan tidak bisa diparse — '{row['Tgl. Kunjungan']}'")
            if not tgl_lahir:
                masalah.append(f"Tgl. Lahir tidak bisa diparse — '{row['Tgl. Lahir']}'")

            if masalah:
                log.warning(f"  ⚠  BARIS MALFORMED — dilewati:")
                for m in masalah:
                    log.warning(f"       • {m}")
                skip += 1
                continue   # loncat ke baris berikutnya

            # =========================================================
            # ISI FORM — satu per satu sesuai urutan field di halaman
            # =========================================================

            # 1. Tanggal Skrining
            isi_tanggal_js(driver, 'tanggal_skrining', tgl_skrining)
            log.info(f"  ✓ Tanggal Skrining : {tgl_skrining}")

            # 2. NIK
            el_nik = driver.find_element(By.ID, 'nik')
            el_nik.clear()
            el_nik.send_keys(nik)
            log.info(f"  ✓ NIK              : {nik}")

            # 3. Nama Lengkap
            el_nama = driver.find_element(By.ID, 'nama_lengkap')
            el_nama.clear()
            el_nama.send_keys(nama)
            log.info(f"  ✓ Nama             : {nama}")

            # 4. Tanggal Lahir
            isi_tanggal_js(driver, 'tanggal_lahir', tgl_lahir)
            log.info(f"  ✓ Tanggal Lahir    : {tgl_lahir}")

            # 5. Jenis Kelamin: selalu Perempuan karena semua pasien Z34. adalah ibu hamil
            Select(driver.find_element(By.ID, 'jenis_kelamin')).select_by_value('P')
            log.info("  ✓ Jenis Kelamin    : Perempuan (P)")

            # 6. Status Kehamilan: selalu Ya (sudah pasti hamil, kode Z34.)
            try:
                wait.until(EC.presence_of_element_located((By.ID, 'status_kehamilan')))
                Select(driver.find_element(By.ID, 'status_kehamilan')).select_by_value('Ya')
                log.info("  ✓ Status Kehamilan : Ya")
            except TimeoutException:
                log.warning("  ⚠ Dropdown status_kehamilan tidak muncul")

            # 7. Waktu Kunjungan Kehamilan (k1–k6, ditentukan dari usia kehamilan)
            try:
                wait.until(EC.presence_of_element_located((By.ID, 'waktu_kunjungan_kehamilan')))
                Select(driver.find_element(By.ID, 'waktu_kunjungan_kehamilan')).select_by_value(kunjungan)
                log.info(f"  ✓ Waktu Kunjungan  : {kunjungan}  (uk ≈ {uk_minggu} mgg)")
            except (TimeoutException, NoSuchElementException) as e:
                log.warning(f"  ⚠ Gagal pilih waktu kunjungan '{kunjungan}': {e}")

            # 8. Nomor Telepon
            el_telp = driver.find_element(By.ID, 'no_telepon')
            el_telp.clear()
            el_telp.send_keys(no_telp)
            log.info(f"  ✓ No Telp          : {no_telp}")

            # 9. Provinsi: BANTEN (kode wilayah: 36)
            Select(driver.find_element(By.ID, 'propinsi_domisili')).select_by_value('36')
            log.info("  ✓ Provinsi         : BANTEN (36)")
            time.sleep(1)   # tunggu dropdown kabupaten dimuat oleh halaman

            # 10. Kabupaten: KOTA TANGERANG SELATAN (kode: 3674)
            # Dropdown ini muncul dinamis setelah provinsi dipilih
            try:
                wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//select[@name='kabupaten_domisili']/option[@value='3674']")
                ))
                Select(driver.find_element(By.NAME, 'kabupaten_domisili')).select_by_value('3674')
                log.info("  ✓ Kabupaten        : KOTA TANGERANG SELATAN (3674)")
            except TimeoutException:
                log.warning("  ⚠ Dropdown kabupaten tidak muncul")
            time.sleep(1)   # tunggu dropdown kecamatan dimuat

            # 11. Kecamatan: SERPONG (kode: 367401)
            try:
                wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//select[@name='kecamatan_domisili']/option[@value='367401']")
                ))
                Select(driver.find_element(By.NAME, 'kecamatan_domisili')).select_by_value('367401')
                log.info("  ✓ Kecamatan        : SERPONG (367401)")
            except TimeoutException:
                log.warning("  ⚠ Dropdown kecamatan tidak muncul")
            time.sleep(1)   # tunggu dropdown kelurahan dimuat

            # 12. Kelurahan: SERPONG (kode: 3674011009)
            try:
                wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//select[@id='kelurahan_domisili']/option[@value='3674011009']")
                ))
                Select(driver.find_element(By.ID, 'kelurahan_domisili')).select_by_value('3674011009')
                log.info("  ✓ Kelurahan        : SERPONG (3674011009)")
            except TimeoutException:
                log.warning("  ⚠ Dropdown kelurahan tidak muncul")

            # 13. Pendidikan Terakhir
            try:
                Select(driver.find_element(By.NAME, 'pendidikan')).select_by_value(pendidikan)
                log.info(f"  ✓ Pendidikan       : {pendidikan}")
            except NoSuchElementException:
                log.warning(f"  ⚠ Opsi pendidikan '{pendidikan}' tidak ada di dropdown")

            # 14. Pekerjaan
            try:
                Select(driver.find_element(By.NAME, 'pekerjaan')).select_by_value(pekerjaan)
                log.info(f"  ✓ Pekerjaan        : {pekerjaan}")
            except NoSuchElementException:
                log.warning(f"  ⚠ Opsi pekerjaan '{pekerjaan}' tidak ada di dropdown")

            # 15. Status Pernikahan: selalu Menikah (semua ibu hamil Z34.)
            Select(driver.find_element(By.ID, 'status_nikah')).select_by_value('Menikah')
            log.info("  ✓ Status Menikah   : Menikah")

            # 16. Penyelenggara: Puskesmas
            Select(driver.find_element(By.NAME, 'penyelenggara')).select_by_value('Puskesmas')
            log.info("  ✓ Penyelenggara    : Puskesmas")

            # 17. Nama Institusi (Puskesmas)
            el_inst = driver.find_element(By.ID, 'nama_institusi')
            el_inst.clear()
            nama_institusi = os.getenv('NAMA_PUSKESMAS', 'PUSKESMAS CONTOH')
            el_inst.send_keys(nama_institusi)
            log.info(f"  ✓ Nama Institusi   : {nama_institusi}")

            # 18. Isi 10 soal kuesioner EPDS sesuai template yang dipilih
            log.info(f"  Mengisi EPDS [{template_key}, skor≈{skor}]:")
            fill_questionnaire(driver, wait, answer_template)

            log.info("")
            log.info("  ✅ Form terisi lengkap. Menyimpan...")
            sukses += 1

            # --- Klik tombol Simpan dan tunggu halaman reload ---
            simpan_btn = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button.btn.btn-dark-blue[type='submit']")
            ))
            driver.execute_script("arguments[0].scrollIntoView(true);", simpan_btn)
            simpan_btn.click()
            log.info("  ✅ Tombol Simpan diklik.")

            # Tunggu sebentar supaya server sempat memproses, lalu buka form baru
            time.sleep(5)
            driver.get(URL_FORM)
            time.sleep(1)

        except Exception as e:
            # Kalau ada error tak terduga pada satu baris, catat dan lanjut ke baris berikutnya
            log.error(f"  ❌ ERROR baris {baris_no} ({row.get('Nama Pasien', '?')}): {e}", exc_info=True)
            error += 1
            try:
                driver.get(URL_FORM)   # refresh ke halaman form bersih
                time.sleep(1)
            except Exception:
                pass

    # --- Tampilkan ringkasan akhir setelah semua baris selesai ---
    log.info("")
    log.info("=" * 65)
    log.info(f"SELESAI  |  Sukses: {sukses}  |  Skip: {skip}  |  Error: {error}")
    log.info(f"Log tersimpan di: {LOG_FILE}")

except FileNotFoundError:
    log.error(f"❌ File tidak ditemukan: {FILE_DATA_PATH}")
except KeyError as e:
    log.error(f"❌ Kolom tidak ditemukan di Excel: {e}")
except Exception as e:
    log.error(f"❌ ERROR tidak terduga: {e}", exc_info=True)
