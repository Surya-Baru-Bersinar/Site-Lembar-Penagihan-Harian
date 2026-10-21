import configparser
import pandas as pd
import gspread
from datetime import datetime

def sinkronkan_data_sheets():
    try:
        print("--> Membaca file konfigurasi piutang.conf...")
        piutang = configparser.ConfigParser()
        piutang.read('piutang.conf')

        url_ss = piutang['URL']['url_ss'].strip()
        nama_sheet = piutang['URL']['url_ss_sheet'].strip()

        print("--> Melakukan otentikasi dengan credentials.json...")
        gc = gspread.service_account(filename='credentials.json')

        print("--> Membuka Google Spreadsheet...")
        spreadsheet = gc.open_by_url(url_ss)
        worksheet = spreadsheet.worksheet(nama_sheet)

        print("--> Menghapus data lama dari baris 2 (Kolom A2:K)...")
        worksheet.batch_clear(['A2:K'])

        print("--> Membaca data dari Piutang_clean_temp.xlsx...")
        df = pd.read_excel('Piutang_clean_temp.xlsx')

        df_subset = df.iloc[:, :11]
        df_subset = df_subset.fillna('')
        data_baru = df_subset.values.tolist()

        if data_baru:
            jumlah_baris = len(data_baru)
            print(f"--> Menuliskan {jumlah_baris} baris data baru mulai sel A2...")
            worksheet.update(range_name='A2', values=data_baru, value_input_option='USER_ENTERED')
            
            bulan_indo = {
                1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'Mei', 6: 'Jun',
                7: 'Jul', 8: 'Ags', 9: 'Sep', 10: 'Okt', 11: 'Nov', 12: 'Des'
            }
            sekarang = datetime.now()
            nama_bulan = bulan_indo[sekarang.month]
            waktu_pembaruan = sekarang.strftime(f"%d {nama_bulan} %Y - %H.%M")

            print("--> Menuliskan waktu pembaruan ke sel L2...")
            worksheet.update(range_name='L2', values=[[waktu_pembaruan]], value_input_option='USER_ENTERED')

            print(f"--> Selesai! Data berhasil diperbarui di Google Sheets pada: {waktu_pembaruan}")
        else:
            print("--> File Excel tidak memiliki data untuk ditulis.")

    except Exception as e:
        print(f"--> Terjadi kesalahan: {e}")

if __name__ == '__main__':
    sinkronkan_data_sheets()