import pandas as pd
from datetime import datetime

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

df = read_excel_auto('Piutang_clean_temp.xlsx')

mapping_bulan = {
    1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'Mei', 6: 'Jun',
    7: 'Jul', 8: 'Agu', 9: 'Sep', 10: 'Okt', 11: 'Nov', 12: 'Des'
}

indo_ke_eng = {
    'Peb': 'Feb',
    'Mei': 'May',
    'Agu': 'Aug',
    'Agt': 'Aug',
    'Ags': 'Aug',
    'Okt': 'Oct',
    'Nop': 'Nov',
    'Des': 'Dec'
}

def ubah_format(nilai):
    if pd.isna(nilai):
        return nilai
    
    if isinstance(nilai, (pd.Timestamp, datetime)):
        return f"{nilai.day:02d} {mapping_bulan[nilai.month]} {nilai.year}"
    
    val_str = str(nilai).strip()
    
    for indo, eng in indo_ke_eng.items():
        val_str = val_str.replace(indo, eng)
        
    try:
        dt = pd.to_datetime(val_str)
        return f"{dt.day:02d} {mapping_bulan[dt.month]} {dt.year}"
    except Exception:
        return nilai

df['Tgl Faktur'] = df['Tgl Faktur'].apply(ubah_format)
df['Jatuh Tempo'] = df['Jatuh Tempo'].apply(ubah_format)

df.to_excel('Piutang_clean_temp.xlsx', index=False)

print("--> Format tanggal pada kolom Tgl Faktur dan Jatuh Tempo berhasil diubah.")