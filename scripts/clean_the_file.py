import os
import pandas as pd

# Path dihitung relatif ke root project
_ROOT      = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INPUT_FILE = os.path.join(_ROOT, 'data', 'Data_Identitas_Ibu_Hamil1.xlsx')
OUTPUT_FILE = os.path.join(_ROOT, 'data', 'file_bersih.xlsx')

df = pd.read_excel(INPUT_FILE)

# Buang semua baris yang kolom Nama-nya kosong
df_bersih = df[df['Nama'].notna()]

df_bersih.to_excel(OUTPUT_FILE, index=False)
print(f"Selesai. {len(df_bersih)} baris disimpan ke: {OUTPUT_FILE}")
