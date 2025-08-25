import pandas as pd
import time
from selenium import webdriver
import random # ### <<< Tambahkan import ini untuk fungsi acak
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from datetime import datetime
# ### <<< INI ADALAH PERBAIKANNYA >>> ###
from selenium.common.exceptions import NoSuchElementException, TimeoutException
# Tambahkan import untuk WebDriverWait & ExpectedConditions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- KONFIGURASI ---
FILE_DATA_PATH = 'filter simpus mei-agustus.xls' 
URL_FORM = 'https://simkeswa.kemkes.go.id/form_skrining/dewasa'

# --- FUNGSI BANTU ---
def ubah_format_tanggal(tanggal_obj):
    if isinstance(tanggal_obj, datetime):
        return tanggal_obj.strftime('%Y-%m-%d')
    try:
        return datetime.strptime(str(tanggal_obj), '%d-%m-%Y').strftime('%Y-%m-%d')
    except (ValueError, TypeError):
        return None


try:
    print(f"Membaca file: '{FILE_DATA_PATH}'...")
    # Kita pertahankan cara membaca Anda yang sudah 95% benar
    # df = pd.read_excel(FILE_DATA_PATH, header=4) //for xls only
    df = pd.read_excel(FILE_DATA_PATH, header=4)
    
    # --- LANGKAH DIAGNOSTIK (SANGAT PENTING) ---
    print("\n✅ File XLS berhasil dibaca. Nama kolom yang terdeteksi adalah:")
    print("=============================================================")
    print(df.columns.tolist())
    print("=============================================================")
    
    # ### <<< INI ADALAH PERBAIKANNYA (SURGICAL FIX) >>> ###
    # Berdasarkan output di atas, kita ganti nama kolom yang salah.
    # PERHATIKAN: Nama 'Unnamed: 11' dan 'Unnamed: 12' mungkin berbeda di komputer Anda.
    # SESUAIKAN nama di dalam kurung kurawal {} ini berdasarkan apa yang TERCETAK di atas.
    print("\n-> Melakukan penggantian nama untuk kolom di bawah 'ALAMAT'...")
    df.rename(columns={
        'ALAMAT': 'Alamat',          # Kolom K6
        'Unnamed: 11': 'Kelurahan',  # Kolom L6
        'Unnamed: 12': 'Kab'         # Kolom M6
    }, inplace=True)
    
    print("-> Nama kolom setelah perbaikan:")
    print(df.columns.tolist())
    print("")
    
    # # Bersihkan data (kode ini dari Anda dan sudah benar)
    # df.dropna(axis=1, how='all', inplace=True)
    # df.dropna(subset=['NKK / NIK'], inplace=True)
    # print("✅ Data berhasil dibersihkan. Nama kolom 'Kab' dan 'Kelurahan' sekarang seharusnya sudah benar.")
     # --- CLEANING + FIX URUTAN ---
    df.dropna(axis=1, how='all', inplace=True)
    df.dropna(subset=['NKK / NIK'], inplace=True)
    df.reset_index(drop=True, inplace=True)   # <<< FIX urutan index
    print("✅ Data berhasil dibersihkan dan index sudah direset.")


except FileNotFoundError:
    print(f"❌ ERROR: File '{FILE_DATA_PATH}' tidak ditemukan.")
    exit()
except KeyError as e:
    # Error ini sekarang seharusnya tidak terjadi untuk 'NKK / NIK'
    # Tapi bisa terjadi jika nama 'Unnamed' di blok rename salah.
    print(f"❌ ERROR 'KeyError': Kolom {e} tidak ditemukan.")
    print("   PASTIKAN nama kolom 'Unnamed' yang Anda masukkan di blok .rename() sudah SAMA PERSIS dengan yang tercetak di atas.")
    exit()
except Exception as e:
    print(f"❌ ERROR: Gagal membaca file XLS. Kesalahan: {e}")
    exit()

# --- INISIALISASI SELENIUM & LANJUTKAN PROSES ---




# ### <<< INI BAGIAN PENGATURAN JUMLAH DATA >>> ###
# Untuk memproses 3 baris pertama, gunakan baris ini:
# df_to_process = df.head(3)
# Untuk memproses SEMUA baris, ganti baris di atas dengan yang ini (hapus tanda #):
#gap nya (misal ATI HAYATI mau lanjut ke PONIJAH di 798, maka 798 - 6)
df_to_process = df.iloc[792:]
# ### <<< ------------------------------------ >>> ###
# --- Inisialisasi Selenium sekali saja ---
print("Mempersiapkan browser...")
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 20)
driver.implicitly_wait(10)

total_rows = len(df_to_process)
print(f"\n🚀 Akan memulai proses automasi untuk {total_rows} baris data...\n")

# ### <<< AWAL DARI FOR LOOP >>> ###
# buka form pertama kali
driver.get(URL_FORM)
for index, row in df_to_process.iterrows():
    # index + 1 agar hitungan baris lebih mudah dibaca manusia (mulai dari 1)
    current_row_num = df.index.get_loc(index) + 1 

    try:
        nama_lengkap = row['Nama Lengkap']
        nik = str(row['NKK / NIK']).strip() # .strip() untuk menghapus spasi tak terlihat
        print(f"--- Mengisi data untuk: {nama_lengkap} (NIK: {nik}) ---")
        
        # # **LANGKAH 1: KEMBALI KE URL & TUNGGU DENGAN STRATEGI BARU**
        # print("Mempersiapkan browser...")
        # driver = webdriver.Chrome()
        # driver.get(URL_FORM)
        # wait = WebDriverWait(driver, 20)
        # driver.implicitly_wait(10)
       



        # --- Mengisi Field Form (Kode ini sama seperti sebelumnya) ---
        tanggal_periksa = ubah_format_tanggal(row['Tanggal'])
        if tanggal_periksa:
            driver.find_element(By.ID, 'tanggal_skrining').send_keys(tanggal_periksa)
        
        driver.find_element(By.ID, 'nama_lengkap').send_keys(nama_lengkap)
        driver.find_element(By.ID, 'nik').send_keys(nik)
        
    # ### <<< INI BAGIAN YANG DIPERBAIKI >>> ###
        # Menggunakan Select untuk menangani dropdown jenis kelamin
        jenis_kelamin = row['L/P']
        if pd.notna(jenis_kelamin):
            # 1. Cari elemen <select> berdasarkan ID-nya
            select_element = driver.find_element(By.ID, 'jenis_kelamin')
            # 2. Buat objek Select dari elemen tersebut
            select_object = Select(select_element)
            # 3. Pilih opsi berdasarkan 'value' ('L' atau 'P'), yang cocok dengan data Excel Anda
            select_object.select_by_value(jenis_kelamin.upper())
        
         # ### <<< INI BAGIAN BARUNYA (LOGIKA FORM DINAMIS) >>> ###
            # Jika gendernya Perempuan, tunggu dan isi form tambahan
            if jenis_kelamin.upper() == 'P':
                print("   - Gender Perempuan terdeteksi, mengisi form kehamilan...")
                try:
                    # Tunggu hingga dropdown kehamilan muncul
                    wait.until(EC.presence_of_element_located((By.ID, 'status_kehamilan')))
                    Select(driver.find_element(By.ID, 'status_kehamilan')).select_by_value('Tidak')

                    # Tunggu hingga dropdown nifas muncul
                    wait.until(EC.presence_of_element_located((By.ID, 'status_nifas')))
                    Select(driver.find_element(By.ID, 'status_nifas')).select_by_value('Tidak')
                    
                    print("   - Form kehamilan dan nifas berhasil diisi.")
                except TimeoutException:
                    print("   - PERINGATAN: Form kehamilan/nifas tidak muncul tepat waktu.")
            # ### <<< AKHIR DARI BAGIAN BARU >>> ###

        # ### <<< INI BAGIAN YANG DIPERBAIKI >>> ###
        tanggal_lahir_str = ubah_format_tanggal(row['Tanggal Lahir'])
        if tanggal_lahir_str:
            # Menggunakan ID yang benar: 'tanggal_lahir'
            driver.find_element(By.ID, 'tanggal_lahir').send_keys(tanggal_lahir_str)
        
        # ### <<< BLOK LOGIKA PEKERJAAN (DENGAN PERBAIKAN TIPE DATA) >>> ###
        print("   - Menerapkan logika untuk pekerjaan...")
        pekerjaan_dari_excel = row['Pekerjaan']
        pekerjaan_dipilih = None

        # 1. Hitung Usia (dengan konversi tipe data yang aman)
        usia = 0
        try:
            # ### <<< INI PERBAIKANNYA >>> ###
            # Paksa konversi kolom 'Tanggal Lahir' menjadi objek datetime
            tgl_lahir_obj = pd.to_datetime(row['Tanggal Lahir'])
            if pd.notna(tgl_lahir_obj):
                usia = (datetime.now() - tgl_lahir_obj).days / 365.25
        except (ValueError, TypeError):
            print("   - Peringatan: Format tanggal lahir tidak valid, usia dianggap 0.")
            usia = 0
        
        # 2. Terapkan Aturan Logika (tidak ada perubahan di sini)
        if pd.isna(pekerjaan_dari_excel) or str(pekerjaan_dari_excel).strip() == '-':
            pekerjaan_dipilih = "Tidak Bekerja"
        elif 'pensiunan' in str(pekerjaan_dari_excel).lower():
            pekerjaan_dipilih = "Pensiunan"
        elif 'pelajar' in str(pekerjaan_dari_excel).lower() or 'mahasiswa' in str(pekerjaan_dari_excel).lower():
            if usia > 20:
                pekerjaan_dipilih = "Mahasiswa"
            else:
                pekerjaan_dipilih = "Pelajar"
        else:
            pekerjaan_dipilih = "Bekerja"

        # 3. Eksekusi Pilihan di Dropdown
        if pekerjaan_dipilih:
            try:
                print(f"   - Hasil logika (usia ~{int(usia)} thn): Memilih '{pekerjaan_dipilih}'")
                # Menggunakan By.NAME karena lebih simpel dari HTML yang Anda berikan
                Select(driver.find_element(By.NAME, 'pekerjaan')).select_by_value(pekerjaan_dipilih)
            except NoSuchElementException:
                print(f"   - Peringatan: Opsi '{pekerjaan_dipilih}' tidak ditemukan di dropdown.")
        

        # ### <<< BLOK LOGIKA PENDIDIKAN (SESUAI ATURAN BARU) >>> ###
        print("   - Menerapkan logika untuk pendidikan...")
        pendidikan_dipilih = None # Variabel untuk menyimpan hasil pilihan
        
        # Aturan 1 & 2: Jika pekerjaan adalah pelajar/mahasiswa
        if 'pelajar' in pekerjaan_dari_excel.lower() or 'mahasiswa' in pekerjaan_dari_excel.lower():
            if usia > 20:
                pendidikan_dipilih = "D3/S1/S2/S3/Pasca Sarjana" # Aturan 1
            else:
                pendidikan_dipilih = "SMA Atau Sederajat" # Aturan 2
        
        # Aturan 3: Jika pekerjaan kosong, "-", atau TNI
        elif pd.isna(row['Pekerjaan']) or pekerjaan_dari_excel == '-' or "TNI" in pekerjaan_dari_excel:
            pendidikan_dipilih = "SMA Atau Sederajat"
            
        # Aturan 4: Jika pekerjaan adalah BUMN atau PNS
        elif "BUMN" in pekerjaan_dari_excel or "PNS" in pekerjaan_dari_excel or "Pegawai Negeri Sipil" in pekerjaan_dari_excel:
            pendidikan_dipilih = "D3/S1/S2/S3/Pasca Sarjana"
            
        # Aturan 5: Catch-all untuk kondisi lainnya
        else:
            pendidikan_dipilih = "SMA Atau Sederajat"
            
        # Eksekusi Pilihan di Dropdown Pendidikan
        if pendidikan_dipilih:
            try:
                print(f"   - Hasil logika: Memilih pendidikan '{pendidikan_dipilih}'")
                Select(driver.find_element(By.NAME, 'pendidikan')).select_by_value(pendidikan_dipilih)
            except NoSuchElementException:
                print(f"   - Peringatan: Opsi '{pendidikan_dipilih}' tidak ditemukan di dropdown.")


        # ### <<< INI BAGIAN YANG DITAMBAHKAN INSTITUSI >>> ###
        print("   - Memilih jenis institusi 'Puskesmas'...")
        Select(driver.find_element(By.NAME, 'penyelenggara')).select_by_value('Puskesmas')
        # ### <<< INI BAGIAN YANG DITAMBAHKAN INSTITUSI NAMA >>> ###
        driver.find_element(By.ID, 'nama_institusi').send_keys('Puskesmas SERPONG II')

        # ### <<< BLOK LOGIKA STATUS MENIKAH (SESUAI ATURAN BARU) >>> ###
        print("   - Menerapkan logika untuk status menikah...")
        status_nikah_dipilih = None

        # Aturan 1: Jika Mengurus Rumah Tangga
        if 'mengurus rumah tangga' in pekerjaan_dari_excel.lower():
            status_nikah_dipilih = "Menikah"
        
        # Aturan 2: Jika Pelajar/Mahasiswa
        elif 'pelajar' in pekerjaan_dari_excel.lower() or 'mahasiswa' in pekerjaan_dari_excel.lower():
            status_nikah_dipilih = "Belum Menikah"
            
        # Aturan 3: Selain itu, pilih secara acak
        else:
            opsi_random = ["Menikah", "Belum Menikah"]
            status_nikah_dipilih = random.choice(opsi_random)
            print(f"   - Tidak ada aturan spesifik, memilih secara acak: '{status_nikah_dipilih}'")

        if status_nikah_dipilih:
            Select(driver.find_element(By.ID, 'status_nikah')).select_by_value(status_nikah_dipilih)
        
        # ### <<< AWAL BLOK LOGIKA ALAMAT DINAMIS >>> ###
        print("   - Menerapkan logika alamat dinamis...")

        # Mengambil data relevan dari baris Excel saat ini
        # Pastikan nama kolom 'Kab' dan 'Kelurahan' sudah benar
        kabupaten_dari_excel = str(row['Kab']).strip().upper()
        kelurahan_dari_excel = str(row['Kelurahan']).strip().upper()

        # Daftar kelurahan spesifik untuk aturan Tangerang Selatan
        kelurahan_serpong = [
            "CIATER", "RAWABUNTU", "RAWA MEKARJAYA", "LENGKONG GUDANG", 
            "LENGKONG WETAN", "BUARAN", "LENGKONG GUDANG TIMUR", 
            "CILENGGANG", "SERPONG"
        ]

        # Pemetaan nama provinsi ke 'value' di dropdown untuk kemudahan
        provinsi_map = {
            "BANTEN": "36", 
            "JAWA BARAT": "32", 
            "DKI JAKARTA": "31"
        }

        # --- Skenario 1: Kabupaten adalah TANGERANG SELATAN ---
        if "TANGERANG SELATAN" in kabupaten_dari_excel:
            print("   - Skenario 1: Tangerang Selatan terdeteksi.")
            try:
                # 1. Pilih Provinsi BANTEN
                Select(driver.find_element(By.ID, 'propinsi_domisili')).select_by_value(provinsi_map["BANTEN"])
                time.sleep(1) # Jeda krusial untuk AJAX memuat kabupaten

                # 2. Pilih Kabupaten KOTA TANGERANG SELATAN
                wait.until(EC.presence_of_element_located((By.XPATH, "//select[@id='kabupaten_domisili']/option[contains(text(), 'TANGERANG SELATAN')]")))
                Select(driver.find_element(By.ID, 'kabupaten_domisili')).select_by_visible_text("KOTA TANGERANG SELATAN")
                time.sleep(1) # Jeda krusial untuk AJAX memuat kecamatan

                # 3. Logika Kecamatan dan Kelurahan berdasarkan daftar
                if kelurahan_dari_excel in kelurahan_serpong:
                    print(f"   - Kelurahan '{kelurahan_dari_excel}' cocok dengan daftar Serpong.")
                    wait.until(EC.presence_of_element_located((By.XPATH, "//select[@id='kecamatan_domisili']/option[contains(text(), 'SERPONG')]")))
                    Select(driver.find_element(By.ID, 'kecamatan_domisili')).select_by_visible_text("SERPONG")
                    time.sleep(1) # Jeda krusial untuk AJAX memuat kelurahan

                    wait.until(EC.presence_of_element_located((By.XPATH, f"//select[@id='kelurahan_domisili']/option[contains(text(), '{kelurahan_dari_excel}')]")))
                    Select(driver.find_element(By.ID, 'kelurahan_domisili')).select_by_visible_text(kelurahan_dari_excel)
                else:
                    print(f"   - Kelurahan '{kelurahan_dari_excel}' tidak ada di daftar khusus, memilih kecamatan/kelurahan secara acak di Tangsel.")
                    # Logika acak jika kelurahan tidak ada di daftar
                    wait.until(EC.numberOfElementsToBeMoreThan((By.XPATH, "//select[@id='kecamatan_domisili']/option"), 1))
                    kecamatan_select = Select(driver.find_element(By.ID, 'kecamatan_domisili'))
                    opsi_kecamatan = kecamatan_select.options[1:]
                    kecamatan_acak = random.choice(opsi_kecamatan)
                    kecamatan_select.select_by_visible_text(kecamatan_acak.text)
                    time.sleep(1)

                    wait.until(EC.numberOfElementsToBeMoreThan((By.XPATH, "//select[@id='kelurahan_domisili']/option"), 1))
                    kelurahan_select = Select(driver.find_element(By.ID, 'kelurahan_domisili'))
                    opsi_kelurahan = kelurahan_select.options[1:]
                    kelurahan_acak = random.choice(opsi_kelurahan)
                    kelurahan_select.select_by_visible_text(kelurahan_acak.text)

            except TimeoutException:
                print("      - Gagal memuat dropdown dinamis untuk Tangsel.")

        # --- Skenario 2: Kabupaten ada di JAWA BARAT ---
        elif any(kota in kabupaten_dari_excel for kota in ["BOGOR", "BANDUNG", "TASIKMALAYA"]):
            print(f"   - Skenario 2: Kabupaten Jawa Barat ({kabupaten_dari_excel}) terdeteksi.")
            try:
                # 1. Pilih Provinsi JAWA BARAT
                Select(driver.find_element(By.ID, 'propinsi_domisili')).select_by_value(provinsi_map["JAWA BARAT"])
                time.sleep(1)

                # 2. Pilih Kabupaten (asumsi nama di Excel cocok dengan nama di dropdown)
                wait.until(EC.presence_of_element_located((By.XPATH, f"//select[@id='kabupaten_domisili']/option[contains(text(), 'BOGOR')]")))
                Select(driver.find_element(By.ID, 'kabupaten_domisili')).select_by_visible_text(row['Kab'].strip()) # Menggunakan nama asli dari row
                time.sleep(1)

                # 3. Pilih Kecamatan Acak
                wait.until(EC.numberOfElementsToBeMoreThan((By.XPATH, "//select[@id='kecamatan_domisili']/option"), 1))
                kecamatan_select = Select(driver.find_element(By.ID, 'kecamatan_domisili'))
                opsi_kecamatan = kecamatan_select.options[1:]
                kecamatan_acak = random.choice(opsi_kecamatan)
                print(f"   - Memilih kecamatan acak: {kecamatan_acak.text}")
                kecamatan_select.select_by_visible_text(kecamatan_acak.text)
                time.sleep(1)

                # 4. Pilih Kelurahan Acak
                wait.until(EC.numberOfElementsToBeMoreThan((By.XPATH, "//select[@id='kelurahan_domisili']/option"), 1))
                kelurahan_select = Select(driver.find_element(By.ID, 'kelurahan_domisili'))
                opsi_kelurahan = kelurahan_select.options[1:]
                kelurahan_acak = random.choice(opsi_kelurahan)
                print(f"   - Memilih kelurahan acak: {kelurahan_acak.text}")
                kelurahan_select.select_by_visible_text(kelurahan_acak.text)

            except TimeoutException:
                print("      - Gagal memuat dropdown dinamis untuk Jawa Barat.")
            except NoSuchElementException:
                print(f"      - Kabupaten '{row['Kab'].strip()}' tidak ditemukan di dropdown Jawa Barat.")

        # --- Skenario 3: Kondisi Lain (RANDOM PENUH) ---
        else:
            print("   - Skenario 3: Tidak ada aturan cocok, menjalankan mode acak penuh.")
            try:
                # 1. Pilih Provinsi Acak
                provinsi_acak_nama = random.choice(["BANTEN", "JAWA BARAT", "DKI JAKARTA"])
                provinsi_acak_value = provinsi_map[provinsi_acak_nama]
                Select(driver.find_element(By.ID, 'propinsi_domisili')).select_by_value(provinsi_acak_value)
                print(f"   - Memilih provinsi acak: {provinsi_acak_nama}")
                time.sleep(1)

                # 2. Pilih Kabupaten Acak
                wait.until(EC.numberOfElementsToBeMoreThan((By.XPATH, "//select[@id='kabupaten_domisili']/option"), 1))
                kabupaten_select = Select(driver.find_element(By.ID, 'kabupaten_domisili'))
                opsi_kabupaten = kabupaten_select.options[1:]
                kabupaten_acak = random.choice(opsi_kabupaten)
                kabupaten_select.select_by_visible_text(kabupaten_acak.text)
                print(f"   - Memilih kabupaten acak: {kabupaten_acak.text}")
                time.sleep(1)

                # 3. Pilih Kecamatan Acak
                wait.until(EC.numberOfElementsToBeMoreThan((By.XPATH, "//select[@id='kecamatan_domisili']/option"), 1))
                kecamatan_select = Select(driver.find_element(By.ID, 'kecamatan_domisili'))
                opsi_kecamatan = kecamatan_select.options[1:]
                kecamatan_acak = random.choice(opsi_kecamatan)
                kecamatan_select.select_by_visible_text(kecamatan_acak.text)
                print(f"   - Memilih kecamatan acak: {kecamatan_acak.text}")
                time.sleep(1)

                # 4. Pilih Kelurahan Acak
                wait.until(EC.numberOfElementsToBeMoreThan((By.XPATH, "//select[@id='kelurahan_domisili']/option"), 1))
                kelurahan_select = Select(driver.find_element(By.ID, 'kelurahan_domisili'))
                opsi_kelurahan = kelurahan_select.options[1:]
                kelurahan_acak = random.choice(opsi_kelurahan)
                kelurahan_select.select_by_visible_text(kelurahan_acak.text)
                print(f"   - Memilih kelurahan acak: {kelurahan_acak.text}")

            except TimeoutException:
                print("      - Gagal memuat dropdown dinamis dalam mode acak penuh.")

        # ### <<< AKHIR BLOK LOGIKA ALAMAT DINAMIS >>> ###


        # ### <<< BLOK LOGIKA NOMOR TELEPON (SESUAI ATURAN BARU) >>> ###
        print("   - Menerapkan logika untuk nomor telepon...")
        no_telp_dari_excel = str(row['No. Telp']).strip()
        no_telp_dikirim = "0" # Nilai default

        # Aturan 1: Jika tidak kosong dan bukan '-', ambil nilainya
        if pd.notna(row['No. Telp']) and no_telp_dari_excel != '-':
            # Membersihkan nomor telepon dari karakter non-numerik jika perlu
            no_telp_dikirim = ''.join(filter(str.isdigit, no_telp_dari_excel))
        # Aturan 2 sudah ditangani oleh nilai default di atas

        print(f"   - Hasil logika: Mengisi No. Telp dengan '{no_telp_dikirim}'")
        driver.find_element(By.ID, 'no_telepon').send_keys(no_telp_dikirim)

        # === Jawab otomatis PHQ-4 (atau skip jika tidak ada) ===
        print("   - Mengecek keberadaan pertanyaan PHQ-4...")
        try:
            phq_block = driver.find_elements(By.ID, "pertanyaan-skrining-phq4")
            if phq_block and phq_block[0].is_displayed():
                print("   - PHQ-4 terdeteksi dan terlihat, menjawab otomatis...")
                for soal_id in range(11, 15):
                    elem = driver.find_element(By.ID, f"soal_{soal_id}_pil_0")
                    driver.execute_script("arguments[0].scrollIntoView(true);", elem)
                    elem.click()
                    time.sleep(0.1)
                print("✅ Semua pertanyaan PHQ-4 sudah dijawab (pilihan: Tidak sama sekali).")
            else:
                print("   - PHQ-4 tidak ada atau hidden → langsung lanjut submit.")

        except Exception as e:
            print(f"⚠️ Peringatan: Terjadi error saat cek/jawab PHQ-4, lanjut submit. Detail: {e}")

            # === Klik tombol Simpan ===
        print("📝 Menyimpan data...")
        try:
            simpan_btn = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-dark-blue[type='submit']"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", simpan_btn)
            simpan_btn.click()
            print("✅ Tombol 'Simpan' berhasil diklik.")
            print(f"⏳ Menunggu {2:.1f} detik sebelum lanjut ke baris berikutnya...")
            time.sleep(3)
            # setelah selesai isi baris ini, kembali ke halaman awal form
            driver.get(URL_FORM)
        except Exception as e:
            print(f"❌ ERROR saat klik tombol Simpan: {e}")

    except Exception as e:
        print(f"❌ GAGAL: Terjadi error pada baris {current_row_num} ('{nama_lengkap}'). Melanjutkan ke baris berikutnya.")
        print(f"   Detail Error: {e}")
        continue # Perintah untuk lanjut ke iterasi loop berikutnya
    # ### <<< AKHIR DARI FOR LOOP >>> ###
    finally:
        # # **LANGKAH 4: TUTUP BROWSER, APAPUN YANG TERJADI**
        # if driver:
        #     driver.quit()
        print("   -Lanjut ke data berikutnya.\n")
