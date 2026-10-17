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
        '1_CleanerAcc.py',
        '1_CleanerAccGiro.py',
        '1B_DownloaderMasterData.py',
        '1C_MergedMaster2Main.py',
        '1D_CleanZeroAR.py',
        '2_CompareGiro.py',
        '3_AdjDateFormat.py',
        '4_HelperSSData.py',
        '5_ARInjectionToSS.py',
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

    source_files = ['Piutang.xls', 'Giro.xls']
    missing_source_files = [
        f for f in source_files if not os.path.exists(os.path.join(current_dir, f))
    ]

    if missing_source_files:
        print("--> [ERROR] Berkas sumber berikut tidak ditemukan di folder utama:")
        for mf in missing_source_files:
            print(f"--> - {mf}")
        input("Program telah selesai di jalankan dan tekan enter untuk keluar.")
        return

    for f in source_files:
        src = os.path.join(current_dir, f)
        dst = os.path.join(dapur_dir, f)
        shutil.copy2(src, dst)
        print(f"--> [OK] Berhasil menyalin {f} ke folder Dapur.")

    scripts_to_run = [
        '1_CleanerAcc.py',
        '1_CleanerAccGiro.py',
        '1B_DownloaderMasterData.py',
        '1C_MergedMaster2Main.py',
        '1D_CleanZeroAR.py',
        '2_CompareGiro.py',
        '3_AdjDateFormat.py',
        '4_HelperSSData.py',
        '5_ARInjectionToSS.py',
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