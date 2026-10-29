import os
import sys
import shutil
import subprocess


def main():
    current_dir = os.path.dirname(os.path.abspath(__file__))

    if os.path.basename(current_dir) == 'Dapur':
        dapur_dir = current_dir
    else:
        dapur_dir = os.path.join(current_dir, 'Dapur')

    print("--> OTOMATISASI PEMBARUAN DATABASE AR SITUS")

    if not os.path.exists(dapur_dir):
        print(f"--> [ERROR] Folder '{dapur_dir}' tidak ditemukan!")
        input("Program telah selesai di jalankan dan tekan enter untuk keluar.")
        return

    required_files = [
        '__init__.py',
        '0_DownloaderMasterData.py',
        '1_CleanerAcc.py',
        '2_CleanerAccGiro.py',
        '3_MergedMaster2Main.py',
        '4_CleanZeroAR.py',
        '5_CompareGiro.py',
        '6_AdjDateFormat.py',
        '7_HelperSSData.py',
        '8_HelperFraudData.py',
        '9_ARInjectionToSS.py',
        '10_AutoOpenBrow.py',
        'piutang.conf',
        'credentials.json',
    ]

    missing_files = [
        f
        for f in required_files
        if not os.path.exists(os.path.join(dapur_dir, f))
    ]

    if missing_files:
        print(
            "--> [ERROR] Proses digagalkan! Berkas berikut tidak ditemukan di folder Dapur:"
        )
        for mf in missing_files:
            print(f"--> - {mf}")
        input("Program telah selesai di jalankan dan tekan enter untuk keluar.")
        return

    print("--> [OK] Semua berkas persyaratan lengkap.")

    piutang_path = os.path.join(current_dir, 'Piutang.xls')
    if not os.path.exists(piutang_path):
        print("--> [ERROR] Berkas wajib 'Piutang.xls' nggak ditemukan di folder utama!")
        input("Program telah selesai di jalankan dan tekan enter untuk keluar.")
        return

    shutil.copy2(piutang_path, os.path.join(dapur_dir, 'Piutang.xls'))
    print("--> [OK] Berhasil menyalin Piutang.xls ke folder Dapur.")

    giro_path = os.path.join(current_dir, 'Giro.xls')
    if os.path.exists(giro_path):
        shutil.copy2(giro_path, os.path.join(dapur_dir, 'Giro.xls'))
        print("--> [OK] Berhasil menyalin Giro.xls ke folder Dapur.")
    else:
        print("--> [INFO] Berkas opsional 'Giro.xls' nggak ditemukan, proses tetap dilanjutkan.")

    scripts_to_run = [
        '0_DownloaderMasterData.py',
        '1_CleanerAcc.py',
        '2_CleanerAccGiro.py',
        '3_MergedMaster2Main.py',
        '4_CleanZeroAR.py',
        '5_CompareGiro.py',
        '6_AdjDateFormat.py',
        '7_HelperSSData.py',
        '8_HelperFraudData.py',
        '9_ARInjectionToSS.py',
        '10_AutoOpenBrow.py',
    ]

    for script_name in scripts_to_run:
        print(f"--> MENJALANKAN: {script_name}")
        
        result = subprocess.run([sys.executable, script_name], cwd=dapur_dir)

        if result.returncode != 0:
            print(
                f"--> [ERROR] Skrip {script_name} berhenti dengan kesalahan (Code: {result.returncode})."
            )
            break

    input("Program telah selesai di jalankan dan tekan enter untuk keluar.")


if __name__ == '__main__':
    main()