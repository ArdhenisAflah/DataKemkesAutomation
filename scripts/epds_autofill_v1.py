# import pandas as pd
# import time
# from selenium import webdriver
# import random # ### <<< Tambahkan import ini untuk fungsi acak
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import Select
# from selenium.common.exceptions import NoSuchElementException
# from datetime import datetime
# # ### <<< INI ADALAH PERBAIKANNYA >>> ###
# from selenium.common.exceptions import NoSuchElementException, TimeoutException
# # Tambahkan import untuk WebDriverWait & ExpectedConditions
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC

# from typing import Dict, List, Tuple, Optional
# import random
# from selenium.webdriver.common.by import By
# from selenium.common.exceptions import NoSuchElementException


# # KONFIGURASI
# FILE_DATA_PATH = 'file_bersih.xlsx'
# URL_FORM = 'https://simkeswa.kemkes.go.id/form_skrining/dewasa'

# # --- FUNGSI BANTU ---
# def ubah_format_tanggal(tanggal_obj):
#     if isinstance(tanggal_obj, datetime):
#         return tanggal_obj.strftime('%Y-%m-%d')
#     try:
#         return datetime.strptime(str(tanggal_obj), '%d-%m-%Y').strftime('%Y-%m-%d')
#     except (ValueError, TypeError):
#         return None
    
# def fill_questionnaire(driver, answer_template):
#     """Fungsi untuk mengisi 10 pertanyaan kuesioner berdasarkan templat."""
#     wait = WebDriverWait(driver, 10)
#     for i, answer_index in enumerate(answer_template):
#         nomor_soal = i + 1
#         target_id = f"soal_{nomor_soal}_pil_{answer_index}"
        
#         try:
#             radio_button = wait.until(EC.element_to_be_clickable((By.ID, target_id)))
#             # Gunakan JavaScript click untuk menghindari masalah elemen terhalang
#             driver.execute_script("arguments[0].click();", radio_button)
            
#             label_text = driver.find_element(By.CSS_SELECTOR, f"label[for='{target_id}']").text.strip()
#             print(f"    Soal {nomor_soal}: Memilih opsi ke-{answer_index + 1} ('{label_text}')")
#         except Exception as e:
#             print(f"    ❌ Gagal pada Soal {nomor_soal}: Tidak dapat mengklik elemen ID '{target_id}'. Error: {e}")

#  # --- DATABASE TEMPLAT JAWABAN EPDS ---
#  # Kunci dictionary (e.g., 'Normal') harus cocok dengan nilai di kolom 'KATEGORI' file CSV.
# # --- DATABASE TEMPLAT JAWABAN EPDS ---
# TEMPLATES = {
#     "Normal_3": [0, 0, 2, 2, 3, 3, 3, 3, 3, 3], #score 3
#     "Normal_5": [0, 0, 2, 2, 3, 2, 2, 3, 3, 3], #score 5
#     "Normal_6": [0, 0, 2, 2, 2, 2, 2, 3, 3, 3], #score 6
# }

# # Fungsi untuk generate jawaban random jika kategori bukan string template (misal kode Python di kolom)
# def generate_random_template(kategori):
#     if "Normal" in kategori:  # Jika mengandung kode untuk Normal
#         return [random.choices([0, 1, 2, 3], weights=[0.7, 0.2, 0.1, 0])[0] for _ in range(10)]
#     elif "Depresi Ringan" in kategori:
#         return [random.choices([0, 1, 2, 3], weights=[0.2, 0.4, 0.3, 0.1])[0] for _ in range(10)]
#     # Tambahkan untuk Sedang/Berat jika diperlukan
#     else:
#         return TEMPLATES["Normal"]  # Default

# try:
#     print(f"Membaca file: '{FILE_DATA_PATH}'...")
#     # Kita pertahankan cara membaca Anda yang sudah 95% benar
#     # df = pd.read_excel(FILE_DATA_PATH, header=4) //for xls only
#     df = pd.read_excel(FILE_DATA_PATH, header=0)

#     # --- CLEANING + FIX URUTAN ---
#     df.dropna(axis=1, how='all', inplace=True)
#     df.dropna(subset=['NIK'], inplace=True)
#     df.reset_index(drop=True, inplace=True)   # <<< FIX urutan index
#     print("✅ Data berhasil dibersihkan dan index sudah direset.")
    
#     # --- LANGKAH PRINT HEADER TABLE (SANGAT PENTING) ---
#     print("\n✅ File XLS berhasil dibaca. Nama kolom yang terdeteksi adalah:")
#     print("=============================================================")
#     print(df.columns.tolist())
#     print("=============================================================")

#     print("\n✅ File XLS berhasil dibaca. Row yang terdeteksi adalah:")
#     print("=============================================================")
#     df_to_process = df;
#     total_rows = len(df_to_process)
#     print(f"\n🚀 Akan memulai proses automasi untuk {total_rows} baris data...\n")
#     print("=============================================================")



#     print("\n✅ File XLS berhasil dibaca. List Rows terdeteksi adalah:")
#     print("=============================================================")
#     df_to_process = df.iloc[33:]
#     total_rows = len(df_to_process)
    
    
#     # --- Inisialisasi Selenium sekali saja ---
#     print("Mempersiapkan browser...")
#     driver = webdriver.Chrome()
#     wait = WebDriverWait(driver, 20)
#     driver.implicitly_wait(10)
#     driver.get(URL_FORM)

#     for index, row in df_to_process.iterrows():
        
#         kunjungan_kehamilan = row["BULAN"][:2].lower()
        
        
#         nama_lengkap = row['Nama']
#         nik = str(row['NIK']).strip() # .strip() untuk menghapus spasi tak terlihat
#         print(f"--- Mengisi data untuk: {nama_lengkap} (NIK: {nik}) ---")

        


#         # --- Mengisi Field Form (Kode ini sama seperti sebelumnya) ---
#         tanggal_periksa = ubah_format_tanggal(row['tanggal kunjungan'])
#         if tanggal_periksa:
#             driver.find_element(By.ID, 'tanggal_skrining').send_keys(tanggal_periksa)
        
#         driver.find_element(By.ID, 'nama_lengkap').send_keys(nama_lengkap)
#         driver.find_element(By.ID, 'nik').send_keys(nik)
        
#         # ### <<< INI BAGIAN YANG DIPERBAIKI KELAMIN >>> ###
#         # Menggunakan Select untuk menangani dropdown jenis kelamin
#         jenis_kelamin = 'P'
#         if pd.notna(jenis_kelamin):
#             # 1. Cari elemen <select> berdasarkan ID-nya
#             select_element = driver.find_element(By.ID, 'jenis_kelamin')
#             # 2. Buat objek Select dari elemen tersebut
#             select_object = Select(select_element)
#             # 3. Pilih opsi berdasarkan 'value' ('L' atau 'P'), yang cocok dengan data Excel Anda
#             select_object.select_by_value(jenis_kelamin.upper())
                    
            
#         # ### <<< INI BAGIAN BARUNYA (LOGIKA FORM DINAMIS) >>> ###
#         # Jika gendernya Perempuan, tunggu dan isi form tambahan
#         if jenis_kelamin.upper() == 'P':
#             print("   - Gender Perempuan terdeteksi, mengisi form kehamilan...")
#             try:
#                 # Tunggu hingga dropdown kehamilan muncul
#                 wait.until(EC.presence_of_element_located((By.ID, 'status_kehamilan')))
#                 Select(driver.find_element(By.ID, 'status_kehamilan')).select_by_value('Ya')

#                 # Tunggu hingga dropdown nifas muncul
#                 wait.until(EC.presence_of_element_located((By.ID, 'waktu_kunjungan_kehamilan')))
#                 Select(driver.find_element(By.ID, 'waktu_kunjungan_kehamilan')).select_by_value(kunjungan_kehamilan)
                
#                 print("   - Form kehamilan dan kunjungan hamil berhasil diisi.")
#             except TimeoutException:
#                 print("   - PERINGATAN: Form kehamilan/kunjungan tidak muncul tepat waktu.")
#         # ### <<< AKHIR DARI BAGIAN BARU >>> ###
        
        
#           # ### <<< INI BAGIAN YANG DIPERBAIKI >>> ###
#         tanggal_lahir_str = ubah_format_tanggal(row['tanggal_lahir'])
#         if tanggal_lahir_str:
#             # Menggunakan ID yang benar: 'tanggal_lahir'
#             driver.find_element(By.ID, 'tanggal_lahir').send_keys(tanggal_lahir_str)
        
#         # ### <<< BLOK LOGIKA PEKERJAAN (DENGAN PERBAIKAN TIPE DATA) >>> ###
#         print("   - Menerapkan logika untuk pekerjaan...")
#         pekerjaan_dari_excel = row['pekerjaan']
#         pekerjaan_dipilih = None
        
#         # 1. Terapkan Aturan Logika (tidak ada perubahan di sini)
#         if pd.isna(pekerjaan_dari_excel) or str(pekerjaan_dari_excel).strip() == '-':
#             pekerjaan_dipilih = "Tidak Bekerja"
#         elif 'pensiunan' in str(pekerjaan_dari_excel).lower():
#             pekerjaan_dipilih = "Pensiunan"
#         elif 'pelajar' in str(pekerjaan_dari_excel).lower() or 'mahasiswa' in str(pekerjaan_dari_excel).lower():
#             pekerjaan_dipilih = "Mahasiswa"
#         else:
#             pekerjaan_dipilih = "Bekerja"

#         # 3. Eksekusi Pilihan di Dropdown
#         if pekerjaan_dipilih:
#             try:
#                 print(f"   - Hasil logika : Memilih '{pekerjaan_dipilih}'")
#                 # Menggunakan By.NAME karena lebih simpel dari HTML yang Anda berikan
#                 Select(driver.find_element(By.NAME, 'pekerjaan')).select_by_value(pekerjaan_dipilih)
#             except NoSuchElementException:
#                 print(f"   - Peringatan: Opsi '{pekerjaan_dipilih}' tidak ditemukan di dropdown.")

#         # ### <<< BLOK LOGIKA PENDIDIKAN (SESUAI ATURAN BARU) >>> ###
#         print("   - Menerapkan logika untuk pendidikan...")
#         pendidikan_dari_excel = row['pendidikan']
#         pendidikan_dipilih = None # Variabel untuk menyimpan hasil pilihan
        
#         # Aturan 1 & 2: Jika pekerjaan adalah pelajar/mahasiswa
#         if 'sma' in str(pendidikan_dari_excel).lower() or 'smk' in str(pendidikan_dari_excel).lower():
#             pendidikan_dipilih = "SMA Atau Sederajat" 
                
#         elif 'smp' in str(pendidikan_dari_excel).lower():
#             pendidikan_dipilih = "SMP Atau Sederajat"
        
#         elif 'sd' in str(pendidikan_dari_excel).lower():
#             pendidikan_dipilih = "SD Atau Sederajat"
        
#         # Aturan 3: Jika pekerjaan kosong, "-", atau TNI
#         elif pd.isna(row['pekerjaan']) or pekerjaan_dari_excel == '-' in pekerjaan_dari_excel:
#             pendidikan_dipilih = "SMA Atau Sederajat"
            
#         # Aturan 4: Jika pekerjaan adalah BUMN atau PNS
#         elif 'sarjana' in str(pendidikan_dari_excel).lower():
#             pendidikan_dipilih = "D3/S1/S2/S3/Pasca Sarjana"  
#         # Aturan 5: Catch-all untuk kondisi lainnya
#         else:
#             pendidikan_dipilih = "SMA Atau Sederajat"
            
#         # Eksekusi Pilihan di Dropdown Pendidikan
#         if pendidikan_dipilih:
#             try:
#                 print(f"   - Hasil logika: Memilih pendidikan '{pendidikan_dipilih}'")
#                 Select(driver.find_element(By.NAME, 'pendidikan')).select_by_value(pendidikan_dipilih)
#             except NoSuchElementException:
#                 print(f"   - Peringatan: Opsi '{pendidikan_dipilih}' tidak ditemukan di dropdown.")
                
                
#         # ### <<< INI BAGIAN YANG DITAMBAHKAN INSTITUSI >>> ###
#         print("   - Memilih jenis institusi 'Puskesmas'...")
#         Select(driver.find_element(By.NAME, 'penyelenggara')).select_by_value('Puskesmas')
#         # ### <<< INI BAGIAN YANG DITAMBAHKAN INSTITUSI NAMA >>> ###
#         driver.find_element(By.ID, 'nama_institusi').send_keys('Puskesmas SERPONG II')
        
        
#         # ### <<< BLOK LOGIKA STATUS MENIKAH (SESUAI ATURAN BARU) >>> ###
#         print("   - Menerapkan logika untuk status menikah...")
#         status_nikah_dipilih = "Menikah"

#         if status_nikah_dipilih:
#             Select(driver.find_element(By.ID, 'status_nikah')).select_by_value(status_nikah_dipilih)
            
#         # ### <<< AWAL BLOK LOGIKA ALAMAT DINAMIS >>> ###
#         print("   - Menerapkan logika alamat dinamis...")

#         # Mengambil data relevan dari baris Excel saat ini
#         # Pastikan nama kolom 'Kab' dan 'Kelurahan' sudah benar
#         kabupaten_dari_excel = str("TANGERANG SELATAN").strip().upper()
#         kelurahan_dari_excel = str("SERPONG").strip().upper()

#         # Daftar kelurahan spesifik untuk aturan Tangerang Selatan
#         kelurahan_serpong = [
#             "CIATER", "RAWABUNTU", "RAWA MEKARJAYA", "LENGKONG GUDANG", 
#             "LENGKONG WETAN", "BUARAN", "LENGKONG GUDANG TIMUR", 
#             "CILENGGANG", "SERPONG"
#         ]

#         # Pemetaan nama provinsi ke 'value' di dropdown untuk kemudahan
#         provinsi_map = {
#             "BANTEN": "36", 
#             "JAWA BARAT": "32", 
#             "DKI JAKARTA": "31"
#         }

#         # --- Skenario 1: Kabupaten adalah TANGERANG SELATAN ---
#         if "TANGERANG SELATAN" in kabupaten_dari_excel:
#             print("   - Skenario 1: Tangerang Selatan terdeteksi.")
#             try:
#                 # 1. Pilih Provinsi BANTEN
#                 Select(driver.find_element(By.ID, 'propinsi_domisili')).select_by_value(provinsi_map["BANTEN"])
#                 time.sleep(1) # Jeda krusial untuk AJAX memuat kabupaten

#                 # 2. Pilih Kabupaten KOTA TANGERANG SELATAN
#                 wait.until(EC.presence_of_element_located((By.XPATH, "//select[@id='kabupaten_domisili']/option[contains(text(), 'TANGERANG SELATAN')]")))
#                 Select(driver.find_element(By.ID, 'kabupaten_domisili')).select_by_visible_text("KOTA TANGERANG SELATAN")
#                 time.sleep(1) # Jeda krusial untuk AJAX memuat kecamatan

#                 print(f"   - Kelurahan '{kelurahan_dari_excel}' cocok dengan daftar Serpong.")
#                 wait.until(EC.presence_of_element_located((By.XPATH, "//select[@id='kecamatan_domisili']/option[contains(text(), 'SERPONG')]")))
#                 Select(driver.find_element(By.ID, 'kecamatan_domisili')).select_by_visible_text("SERPONG")
#                 time.sleep(1) # Jeda krusial untuk AJAX memuat kelurahan
                
#                 print(f" ISI  - Kelurahan '{kelurahan_dari_excel}' cocok dengan daftar Serpong.")
#                 wait.until(EC.presence_of_element_located((By.XPATH, "//select[@id='kelurahan_domisili']/option[contains(text(), 'SERPONG')]")))
#                 Select(driver.find_element(By.ID, 'kelurahan_domisili')).select_by_visible_text("SERPONG")
#                 time.sleep(1) # Jeda krusial untuk AJAX memuat kelurahan

#             except TimeoutException:
#                 print("      - Gagal memuat dropdown dinamis untuk Tangsel.")

#         # --- Skenario 3: Kondisi Lain (RANDOM PENUH) ---
#         else:
#             print("   - Skenario 3: Tidak ada aturan cocok, menjalankan mode acak penuh.")
#             try:
#                 # 1. Pilih Provinsi Acak
#                 provinsi_acak_nama = random.choice(["BANTEN", "JAWA BARAT", "DKI JAKARTA"])
#                 provinsi_acak_value = provinsi_map[provinsi_acak_nama]
#                 Select(driver.find_element(By.ID, 'propinsi_domisili')).select_by_value(provinsi_acak_value)
#                 print(f"   - Memilih provinsi acak: {provinsi_acak_nama}")
#                 time.sleep(1)

#                 # 2. Pilih Kabupaten Acak
#                 wait.until(EC.numberOfElementsToBeMoreThan((By.XPATH, "//select[@id='kabupaten_domisili']/option"), 1))
#                 kabupaten_select = Select(driver.find_element(By.ID, 'kabupaten_domisili'))
#                 opsi_kabupaten = kabupaten_select.options[1:]
#                 kabupaten_acak = random.choice(opsi_kabupaten)
#                 kabupaten_select.select_by_visible_text(kabupaten_acak.text)
#                 print(f"   - Memilih kabupaten acak: {kabupaten_acak.text}")
#                 time.sleep(1)

#                 # 3. Pilih Kecamatan Acak
#                 wait.until(EC.numberOfElementsToBeMoreThan((By.XPATH, "//select[@id='kecamatan_domisili']/option"), 1))
#                 kecamatan_select = Select(driver.find_element(By.ID, 'kecamatan_domisili'))
#                 opsi_kecamatan = kecamatan_select.options[1:]
#                 kecamatan_acak = random.choice(opsi_kecamatan)
#                 kecamatan_select.select_by_visible_text(kecamatan_acak.text)
#                 print(f"   - Memilih kecamatan acak: {kecamatan_acak.text}")
#                 time.sleep(1)

#                 # 4. Pilih Kelurahan Acak
#                 wait.until(EC.numberOfElementsToBeMoreThan((By.XPATH, "//select[@id='kelurahan_domisili']/option"), 1))
#                 kelurahan_select = Select(driver.find_element(By.ID, 'kelurahan_domisili'))
#                 opsi_kelurahan = kelurahan_select.options[1:]
#                 kelurahan_acak = random.choice(opsi_kelurahan)
#                 kelurahan_select.select_by_visible_text(kelurahan_acak.text)
#                 print(f"   - Memilih kelurahan acak: {kelurahan_acak.text}")

#             except TimeoutException:
#                 print("      - Gagal memuat dropdown dinamis dalam mode acak penuh.")

#         # ### <<< AKHIR BLOK LOGIKA ALAMAT DINAMIS >>> ###
        
#         # ### <<< BLOK LOGIKA NOMOR TELEPON (SESUAI ATURAN BARU) >>> ###
#         print("   - Menerapkan logika untuk nomor telepon...")
#         no_telp_dari_excel = str(row['no_telp']).strip()
#         no_telp_dikirim = "0" # Nilai default

#         # Aturan 1: Jika tidak kosong dan bukan '-', ambil nilainya
#         if pd.notna(row['no_telp']) and no_telp_dari_excel != '-':
#             # Membersihkan nomor telepon dari karakter non-numerik jika perlu
#             no_telp_dikirim = ''.join(filter(str.isdigit, no_telp_dari_excel))
#         # Aturan 2 sudah ditangani oleh nilai default di atas

#         print(f"   - Hasil logika: Mengisi No. Telp dengan '{no_telp_dikirim}'")
#         driver.find_element(By.ID, 'no_telepon').send_keys(no_telp_dikirim)
        
        
#         # --- ISI KUESIONER DENGAN TEMPLAT RANDOM (TOTAL SKOR < 9) ---
#         answer_template = TEMPLATES[random.choice(list(TEMPLATES.keys()))]
#         print(f"    Using template: {answer_template} (Total score: {sum(answer_template)})")
#         fill_questionnaire(driver, answer_template);
        
#         time.sleep(2)
#         # === Klik tombol Simpan ===
#         print("📝 Menyimpan data...")
#         try:
#             simpan_btn = wait.until(
#                 EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-dark-blue[type='submit']"))
#             )
#             driver.execute_script("arguments[0].scrollIntoView(true);", simpan_btn)
#             simpan_btn.click()
#             print("✅ Tombol 'Simpan' berhasil diklik.")
#             print(f"⏳ Menunggu {2:.1f} detik sebelum lanjut ke baris berikutnya...")
#             time.sleep(5)
#             # setelah selesai isi baris ini, kembali ke halaman awal form
#             driver.get(URL_FORM)
#         except Exception as e:
#             print(f"❌ ERROR saat klik tombol Simpan: {e}")


# except FileNotFoundError:
#     print(f"❌ ERROR: File '{FILE_DATA_PATH}' tidak ditemukan.")
#     # exit()
# except KeyError as e:
#     # Error ini sekarang seharusnya tidak terjadi untuk 'NKK / NIK'
#     # Tapi bisa terjadi jika nama 'Unnamed' di blok rename salah.
#     print(f"❌ ERROR 'KeyError': Kolom {e} tidak ditemukan.")
#     print("   PASTIKAN nama kolom 'Unnamed' yang Anda masukkan di blok .rename() sudah SAMA PERSIS dengan yang tercetak di atas.")
#     # exit()
# except Exception as e:
#     print(f"❌ ERROR: Gagal membaca file XLS. Kesalahan: {e}")
#     # exit()

# # --- INISIALISASI SELENIUM & LANJUTKAN PROSES ---


import pandas as pd
import time
import random
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException

FILE_DATA_PATH = 'file_bersih.xlsx'
URL_FORM = 'https://simkeswa.kemkes.go.id/form_skrining/dewasa'

TEMPLATES = {
    "Normal_3": [0, 0, 2, 2, 3, 3, 3, 3, 3, 3],
    "Normal_5": [0, 0, 2, 2, 3, 2, 2, 3, 3, 3],
    "Normal_6": [0, 0, 2, 2, 2, 2, 2, 3, 3, 3],
}


def parse_tanggal(tanggal_obj):
    """Konversi berbagai format input ke objek datetime. Return None jika gagal."""
    if isinstance(tanggal_obj, datetime):
        return tanggal_obj
    if hasattr(tanggal_obj, 'to_pydatetime'):
        return tanggal_obj.to_pydatetime()
    if hasattr(tanggal_obj, 'strftime'):
        return datetime(tanggal_obj.year, tanggal_obj.month, tanggal_obj.day)
    s = str(tanggal_obj).split(' ')[0].strip()
    for fmt in ('%d-%m-%Y', '%Y-%m-%d', '%d/%m/%Y', '%Y/%m/%d'):
        try:
            return datetime.strptime(s, fmt)
        except (ValueError, TypeError):
            continue
    return None


def ubah_format_tanggal(tanggal_obj, fmt='%Y-%m-%d'):
    """Parse tanggal lalu format ke string sesuai fmt yang diminta."""
    dt = parse_tanggal(tanggal_obj)
    return dt.strftime(fmt) if dt else None


def isi_tanggal_js(driver, element_id, nilai_tanggal):
    el = driver.find_element(By.ID, element_id)
    driver.execute_script("""
        var el = arguments[0];
        var val = arguments[1];
        el.removeAttribute('readonly');
        el.value = val;
        if (window.jQuery) {
            try { jQuery(el).datepicker('setDate', val); } catch(e) {}
            jQuery(el).trigger('change');
        }
        el.dispatchEvent(new Event('input', {bubbles: true}));
        el.dispatchEvent(new Event('change', {bubbles: true}));
    """, el, nilai_tanggal)


def fill_questionnaire(driver, answer_template):
    wait = WebDriverWait(driver, 10)
    for i, answer_index in enumerate(answer_template):
        nomor_soal = i + 1
        target_id = f"soal_{nomor_soal}_pil_{answer_index}"
        try:
            radio_button = wait.until(EC.element_to_be_clickable((By.ID, target_id)))
            driver.execute_script("arguments[0].click();", radio_button)
            label_text = driver.find_element(By.CSS_SELECTOR, f"label[for='{target_id}']").text.strip()
            print(f"    Soal {nomor_soal}: Memilih opsi ke-{answer_index + 1} ('{label_text}')")
        except Exception as e:
            print(f"    ❌ Gagal pada Soal {nomor_soal}: ID '{target_id}'. Error: {e}")


try:
    print(f"Membaca file: '{FILE_DATA_PATH}'...")
    df = pd.read_excel(FILE_DATA_PATH, header=0)

    df.dropna(axis=1, how='all', inplace=True)
    df.dropna(subset=['NIK'], inplace=True)
    df.reset_index(drop=True, inplace=True)

    print("✅ Data berhasil dibersihkan.")
    print("Kolom terdeteksi:", df.columns.tolist())

    df_to_process = df.iloc[33:]
    total_rows = len(df_to_process)
    print(f"\n🚀 Akan memproses {total_rows} baris data...\n")

    print("Mempersiapkan browser...")
    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 20)
    driver.implicitly_wait(10)
    driver.get(URL_FORM)

    provinsi_map = {
        "BANTEN": "36",
        "JAWA BARAT": "32",
        "DKI JAKARTA": "31"
    }

    for index, row in df_to_process.iterrows():
        kunjungan_kehamilan = row["BULAN"][:2].lower()
        nama_lengkap = row['Nama']
        nik = str(row['NIK']).strip()
        print(f"\n--- Mengisi data untuk: {nama_lengkap} (NIK: {nik}) ---")

        tanggal_periksa = ubah_format_tanggal(row['tanggal kunjungan'], '%Y-%m-%d')
        if tanggal_periksa:
            isi_tanggal_js(driver, 'tanggal_skrining', tanggal_periksa)

        driver.find_element(By.ID, 'nama_lengkap').send_keys(nama_lengkap)
        driver.find_element(By.ID, 'nik').send_keys(nik)

        # Isi tanggal lahir SEBELUM trigger dropdown gender (hindari gangguan AJAX)
        tanggal_lahir_str = ubah_format_tanggal(row['tanggal_lahir'], '%Y-%m-%d')
        if tanggal_lahir_str:
            print(f"   - Mengisi tanggal lahir: {tanggal_lahir_str}")
            isi_tanggal_js(driver, 'tanggal_lahir', tanggal_lahir_str)
        else:
            print(f"   - tanggal_lahir tidak dapat diparse: {row['tanggal_lahir']}")

        jenis_kelamin = 'P'
        select_object = Select(driver.find_element(By.ID, 'jenis_kelamin'))
        select_object.select_by_value(jenis_kelamin.upper())

        if jenis_kelamin.upper() == 'P':
            print("   - Gender Perempuan, mengisi form kehamilan...")
            try:
                wait.until(EC.presence_of_element_located((By.ID, 'status_kehamilan')))
                Select(driver.find_element(By.ID, 'status_kehamilan')).select_by_value('Ya')

                wait.until(EC.presence_of_element_located((By.ID, 'waktu_kunjungan_kehamilan')))
                Select(driver.find_element(By.ID, 'waktu_kunjungan_kehamilan')).select_by_value(kunjungan_kehamilan)
                print("   - Form kehamilan berhasil diisi.")
            except TimeoutException:
                print("   - ⚠️ Form kehamilan tidak muncul tepat waktu.")

        pekerjaan_dari_excel = row['pekerjaan']
        if pd.isna(pekerjaan_dari_excel) or str(pekerjaan_dari_excel).strip() == '-':
            pekerjaan_dipilih = "Tidak Bekerja"
        elif 'pensiunan' in str(pekerjaan_dari_excel).lower():
            pekerjaan_dipilih = "Pensiunan"
        elif 'pelajar' in str(pekerjaan_dari_excel).lower() or 'mahasiswa' in str(pekerjaan_dari_excel).lower():
            pekerjaan_dipilih = "Mahasiswa"
        else:
            pekerjaan_dipilih = "Bekerja"

        try:
            print(f"   - Memilih pekerjaan: '{pekerjaan_dipilih}'")
            Select(driver.find_element(By.NAME, 'pekerjaan')).select_by_value(pekerjaan_dipilih)
        except NoSuchElementException:
            print(f"   - ⚠️ Opsi pekerjaan '{pekerjaan_dipilih}' tidak ditemukan.")

        pendidikan_dari_excel = str(row['pendidikan']).lower()
        if 'sma' in pendidikan_dari_excel or 'smk' in pendidikan_dari_excel:
            pendidikan_dipilih = "SMA Atau Sederajat"
        elif 'smp' in pendidikan_dari_excel:
            pendidikan_dipilih = "SMP Atau Sederajat"
        elif 'sd' in pendidikan_dari_excel:
            pendidikan_dipilih = "SD Atau Sederajat"
        elif 'sarjana' in pendidikan_dari_excel:
            pendidikan_dipilih = "D3/S1/S2/S3/Pasca Sarjana"
        else:
            pendidikan_dipilih = "SMA Atau Sederajat"

        try:
            print(f"   - Memilih pendidikan: '{pendidikan_dipilih}'")
            Select(driver.find_element(By.NAME, 'pendidikan')).select_by_value(pendidikan_dipilih)
        except NoSuchElementException:
            print(f"   - ⚠️ Opsi pendidikan '{pendidikan_dipilih}' tidak ditemukan.")

        print("   - Memilih institusi: Puskesmas Contoh")
        Select(driver.find_element(By.NAME, 'penyelenggara')).select_by_value('Puskesmas')
        driver.find_element(By.ID, 'nama_institusi').send_keys('Puskesmas Contoh')

        print("   - Status menikah: Menikah")
        Select(driver.find_element(By.ID, 'status_nikah')).select_by_value("Menikah")

        print("   - Mengisi alamat domisili...")
        kabupaten_dari_excel = "TANGERANG SELATAN"
        kelurahan_dari_excel = "SERPONG"

        if "TANGERANG SELATAN" in kabupaten_dari_excel:
            print("   - Skenario: Tangerang Selatan.")
            try:
                Select(driver.find_element(By.ID, 'propinsi_domisili')).select_by_value(provinsi_map["BANTEN"])
                time.sleep(1)

                wait.until(EC.presence_of_element_located((By.XPATH, "//select[@id='kabupaten_domisili']/option[contains(text(), 'TANGERANG SELATAN')]")))
                Select(driver.find_element(By.ID, 'kabupaten_domisili')).select_by_visible_text("KOTA TANGERANG SELATAN")
                time.sleep(1)

                wait.until(EC.presence_of_element_located((By.XPATH, "//select[@id='kecamatan_domisili']/option[contains(text(), 'SERPONG')]")))
                Select(driver.find_element(By.ID, 'kecamatan_domisili')).select_by_visible_text("SERPONG")
                time.sleep(1)

                wait.until(EC.presence_of_element_located((By.XPATH, "//select[@id='kelurahan_domisili']/option[contains(text(), 'SERPONG')]")))
                Select(driver.find_element(By.ID, 'kelurahan_domisili')).select_by_visible_text("SERPONG")
                time.sleep(1)
            except TimeoutException:
                print("      - ⚠️ Gagal memuat dropdown alamat Tangsel.")
        else:
            print("   - Skenario: Alamat acak.")
            try:
                provinsi_acak_nama = random.choice(["BANTEN", "JAWA BARAT", "DKI JAKARTA"])
                Select(driver.find_element(By.ID, 'propinsi_domisili')).select_by_value(provinsi_map[provinsi_acak_nama])
                print(f"   - Provinsi acak: {provinsi_acak_nama}")
                time.sleep(1)

                wait.until(EC.numberOfElementsToBeMoreThan((By.XPATH, "//select[@id='kabupaten_domisili']/option"), 1))
                kabupaten_select = Select(driver.find_element(By.ID, 'kabupaten_domisili'))
                kabupaten_acak = random.choice(kabupaten_select.options[1:])
                kabupaten_select.select_by_visible_text(kabupaten_acak.text)
                print(f"   - Kabupaten acak: {kabupaten_acak.text}")
                time.sleep(1)

                wait.until(EC.numberOfElementsToBeMoreThan((By.XPATH, "//select[@id='kecamatan_domisili']/option"), 1))
                kecamatan_select = Select(driver.find_element(By.ID, 'kecamatan_domisili'))
                kecamatan_acak = random.choice(kecamatan_select.options[1:])
                kecamatan_select.select_by_visible_text(kecamatan_acak.text)
                print(f"   - Kecamatan acak: {kecamatan_acak.text}")
                time.sleep(1)

                wait.until(EC.numberOfElementsToBeMoreThan((By.XPATH, "//select[@id='kelurahan_domisili']/option"), 1))
                kelurahan_select = Select(driver.find_element(By.ID, 'kelurahan_domisili'))
                kelurahan_acak = random.choice(kelurahan_select.options[1:])
                kelurahan_select.select_by_visible_text(kelurahan_acak.text)
                print(f"   - Kelurahan acak: {kelurahan_acak.text}")
            except TimeoutException:
                print("      - ⚠️ Gagal memuat dropdown alamat acak.")

        no_telp_dari_excel = str(row['no_telp']).strip()
        if pd.notna(row['no_telp']) and no_telp_dari_excel != '-':
            no_telp_dikirim = ''.join(filter(str.isdigit, no_telp_dari_excel))
        else:
            no_telp_dikirim = "0"
        print(f"   - No. Telp: '{no_telp_dikirim}'")
        driver.find_element(By.ID, 'no_telepon').send_keys(no_telp_dikirim)

        answer_template = TEMPLATES[random.choice(list(TEMPLATES.keys()))]
        print(f"   - Template kuesioner: {answer_template} (Skor: {sum(answer_template)})")
        fill_questionnaire(driver, answer_template)

        time.sleep(2)
        print("📝 Menyimpan data...")
        try:
            simpan_btn = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-dark-blue[type='submit']"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", simpan_btn)
            simpan_btn.click()
            print("✅ Data berhasil disimpan.")
            time.sleep(5)
            driver.get(URL_FORM)
        except Exception as e:
            print(f"❌ ERROR saat klik tombol Simpan: {e}")

except FileNotFoundError:
    print(f"❌ ERROR: File '{FILE_DATA_PATH}' tidak ditemukan.")
except KeyError as e:
    print(f"❌ ERROR 'KeyError': Kolom {e} tidak ditemukan.")
except Exception as e:
    print(f"❌ ERROR: {e}")