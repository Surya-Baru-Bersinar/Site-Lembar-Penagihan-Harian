import configparser
import os
import sys
import webbrowser

NAMA_FILE_CONFIG = "piutang.conf"

def baca_url_apps(file_path):
    if not os.path.exists(file_path):
        print(f"--> Gagal File '{file_path}' nggak ditemukan.")
        return None

    config = configparser.ConfigParser()
    config.read(file_path, encoding="utf-8")

    if "URL" not in config:
        print(f"--> Gagal Bagian [URL] nggak ditemukan di '{file_path}'.")
        return None

    url = config.get("URL", "url_apps", fallback="").strip()
    return url if url else None

def main():
    url_target = baca_url_apps(NAMA_FILE_CONFIG)

    if not url_target:
        print("--> Info 'url_apps' kosong atau nggak valid. Proses dibatalkan.")
        sys.exit(0)

    print(f"--> Membuka antarmuka penagihan di peramban: {url_target}")
    webbrowser.open(url_target)

if __name__ == "__main__":
    main()