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

config = configparser.ConfigParser(
    inline_comment_prefixes=('#', ';'),
    allow_no_value=True,
    strict=False
)
config.read("piutang.conf")

if not config.has_section("AR") or not config.has_option("AR", "ar_data_fraud"):
    print("--> Konfigurasi [AR] -> ar_data_fraud tidak ditemukan. Proses dihentikan.")
    sys.exit()

status_fraud = config.get("AR", "ar_data_fraud", fallback="").strip().lower()

if status_fraud != "ya":
    print("--> Nilai ar_data_fraud adalah 'No'. Proses pembersihan dilewati.")
    sys.exit()

keyword_conf = config.get("AR", "ar_fraud_keyword", fallback="CASE|FRAUD").strip()
fraud_pattern = keyword_conf if keyword_conf else "CASE|FRAUD"

nama_file = "Piutang_clean_temp.xlsx"
df = read_excel_auto(nama_file)

kolom_target = "Nama Penjual"

if kolom_target in df.columns:
    mask_fraud = df[kolom_target].astype(str).str.contains(fraud_pattern, case=False, na=False, regex=True)
    total_dihapus = mask_fraud.sum()

    if total_dihapus > 0:
        df_hasil = df[~mask_fraud]
        df_hasil.to_excel(nama_file, index=False)
        print(f"--> Pembersihan selesai. Berhasil menghapus {total_dihapus} baris dengan kata kunci '{fraud_pattern}'. File {nama_file} diperbarui.")
    else:
        print(f"--> Tidak ditemukan baris yang memuat '{fraud_pattern}' pada kolom '{kolom_target}'. File tidak diubah.")
else:
    print(f"--> Kolom '{kolom_target}' tidak ditemukan di dalam file Excel.")