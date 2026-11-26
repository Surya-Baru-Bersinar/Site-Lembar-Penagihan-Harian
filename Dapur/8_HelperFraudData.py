import configparser
import sys
import pandas as pd

try:
    import python_calamine
    DEFAULT_EXCEL_ENGINE = 'calamine'
except ImportError:
    DEFAULT_EXCEL_ENGINE = 'openpyxl'

try:
    pd.set_option('io.excel.xlsx.reader', DEFAULT_EXCEL_ENGINE)
except Exception:
    pass

def read_excel_auto(file_path, **kwargs):
    if DEFAULT_EXCEL_ENGINE == 'calamine':
        try:
            return pd.read_excel(file_path, engine='calamine', **kwargs)
        except Exception:
            return pd.read_excel(file_path, engine='openpyxl', **kwargs)
    return pd.read_excel(file_path, engine='openpyxl', **kwargs)

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
df = read_excel_auto(nama_file)

kolom_target = "Nama Penjual"

if kolom_target in df.columns:
    df_hasil = df[
        ~df[kolom_target].astype(str).str.contains("CASE", case=False, na=False)
    ]
    df_hasil.to_excel(nama_file, index=False)
    print(f"--> Pembersihan selesai. File {nama_file} berhasil diperbarui.")
else:
    print(f"--> Kolom '{kolom_target}' tidak ditemukan di dalam file Excel.")