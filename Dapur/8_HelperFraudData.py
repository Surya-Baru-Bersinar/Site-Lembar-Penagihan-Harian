import configparser
import sys
import pandas as pd

config = configparser.ConfigParser()
config.read("piutang.conf")

if not config.has_section("AR") or not config.has_option("AR", "ar_data_fraud"):
    print("--> Konfigurasi [AR] -> ar_data_fraud tidak ditemukan. Proses dihentikan.")
    sys.exit()

status_fraud = config.get("AR", "ar_data_fraud").strip().lower()

if status_fraud != "ya":
    print("--> Nilai ar_data_fraud adalah 'No'. Proses pembersihan dilewati.")
    sys.exit()

nama_file = "Piutang_clean_temp.xlsx"
df = pd.read_excel(nama_file)

kolom_target = "Nama Penjual"

if kolom_target in df.columns:
    df_hasil = df[
        ~df[kolom_target].astype(str).str.contains("FRAUD", case=False, na=False)
    ]
    df_hasil.to_excel(nama_file, index=False)
    print(f"--> Pembersihan selesai. File {nama_file} berhasil diperbarui.")
else:
    print(f"--> Kolom '{kolom_target}' tidak ditemukan di dalam file Excel.")