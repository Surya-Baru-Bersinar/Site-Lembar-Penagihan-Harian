# 📄 Site Lembar Penagihan Harian (Pembuat Tagihan LHT - SBB)

![Python](https://img.shields.io/badge/Python-3.8%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-Data_Pipeline-150458?style=for-the-badge&logo=pandas&logoColor=white)
![Google Sheets](https://img.shields.io/badge/Google_Sheets-API_&_Apps_Script-34A853?style=for-the-badge&logo=googlesheets&logoColor=white)
![TailwindCSS](https://img.shields.io/badge/Tailwind_CSS-UI_Generator-06B6D4?style=for-the-badge&logo=tailwindcss&logoColor=white)

Sistem otomasi terintegrasi untuk pengolahan data **Piutang Dagang (*Accounts Receivable* / AR)** dan **Giro**, sinkronisasi otomatis ke **Google Spreadsheets**, serta aplikasi web **Pembuat Tagihan LHT (Laporan Hasil Tagihan)** berbasis Google Apps Script.

Proyek ini dirancang untuk mempercepat alur kerja tim Admin dan Finance Depo: mulai dari pembersihan file ekspor sistem akuntansi (`Piutang.xls` & `Giro.xls`), pencocokan data pelanggan master, pemotongan faktur yang sudah klop dengan Giro, hingga pembuatan lembar tagihan harian siap cetak (*Continuous Form / Letter* maupun *A4 Landscape*) dengan fitur penjadwalan otomatis (**SmartFill**).

---

## 📋 Daftar Isi

* [Fitur Utama](#-fitur-utama)
* [Arsitektur & Alur Kerja Sistem](#-arsitektur--alur-kerja-sistem)
* [Teknologi yang Digunakan](#-teknologi-yang-digunakan)
* [Struktur Direktori](#-struktur-direktori)
* [Prasyarat Sistem](#-prasyarat-sistem)
* [Panduan Instalasi & Konfigurasi](#-panduan-instalasi--konfigurasi)
  * [1. Persiapan Lingkungan Python](#1-persiapan-lingkungan-python)
  * [2. Konfigurasi Google Service Account (`credentials.json`)](#2-konfigurasi-google-service-account-credentialsjson)
  * [3. Konfigurasi Pipeline (`piutang.conf`)](#3-konfigurasi-pipeline-piutangconf)
  * [4. Pemasangan Web App / Add-on Google Sheets](#4-pemasangan-web-app--add-on-google-sheets)
* [Cara Penggunaan](#-cara-penggunaan)
  * [A. Memperbarui Data Piutang (Pipeline Python)](#a-memperbarui-data-piutang-pipeline-python)
  * [B. Membuat & Mencetak Lembar LHT (Web App)](#b-membuat--mencetak-lembar-lht-web-app)
* [Referensi Konfigurasi](#-referensi-konfigurasi)
* [Penanganan Masalah Umum (Troubleshooting)](#-penanganan-masalah-umum-troubleshooting)
* [Kontribusi](#-kontribusi)
* [Lisensi & Kontak](#-lisensi--kontak)

---

## ✨ Fitur Utama

### 1. Pipeline ETL Piutang & Giro Otomatis (`Dapur/`)
* **Deteksi Header Dinamis**: Membaca file ekspor akuntansi (`Piutang.xls`) dan mendeteksi posisi baris judul kolom secara cerdas (hingga 150 baris pertama) beserta variasi penamaan kolom (*fallback*).
* **Pembersihan & Normalisasi Data**: Menghapus baris kosong, header berulang, dan baris subtotal; mengonversi format tanggal Bahasa Indonesia (`Mei`, `Agu`, `Okt`, `Nop`, `Des`, `Peb`, `Ags`, `Agt`); serta menstandarkan format angka desimal/ribuan.
* **Sinkronisasi Master Pelanggan**: Mengunduh data master pelanggan dari Google Sheets (`MIRROR PCMO`) secara otomatis dan memperbarui nama pelanggan pada data piutang berdasarkan kode pelanggan (`NOPEL`).
* **Eliminasi Piutang Nol**: Menghapus faktur yang memiliki `Sisa Piutang = 0` secara otomatis.
* **Rekonsiliasi Giro Otomatis (`Compare Giro`)**: Mengakumulasi nilai penerimaan pada `Giro.xls` per pelanggan dan nomor faktur, lalu menghapus faktur dari daftar tagihan apabila selisih antara `Nilai Faktur` dan total giro berada di bawah batas toleransi (`giro_cut`).
* **Injeksi Langsung ke Google Sheets**: Mengunggah data bersih ke sheet `Piutang` (kolom `A2:K`) dan mencatat *timestamp* pembaruan terakhir di sel `L2`.

### 2. Aplikasi Web Pembuat Tagihan LHT (`Addon/`)
* **✨ SmartFill Jadwal Penagihan**: Memilih pelanggan secara otomatis berdasarkan **Nama Penagih (Sales)**, **Minggu (1–4)**, dan **Hari Kunjungan (Senin–Minggu)** sesuai aturan yang diatur pada sheet `conf`.
* **🔍 Pencarian & Pengurutan A–Z**: Filter pelanggan secara *real-time*, pengurutan otomatis berdasarkan Nama Pelanggan dan Nomor Faktur, kalkulasi otomatis umur faktur (hari) dan jumlah terbayar, serta fitur tambah/hapus baris kosong.
* **🖨️ Cetak Multi-Format & PDF Vektor Presisi**:
  * Mendukung kertas **Continuous Form / Letter (279 x 216 mm)** untuk printer *dot-matrix* dan **A4 Landscape (297 x 210 mm)**.
  * Paginasi otomatis maksimal **17 baris faktur per halaman** lengkap dengan header, rekap total tagihan, dan kolom tanda tangan.
  * Tersedia 5 mode aksi: **Simpan Spreadsheet**, **Cetak Langsung (Browser)**, **Cetak Langsung (PDF Vektor)**, **Unduh PDF Saja**, serta **Simpan & Unduh PDF**.
* **⚙️ Mode Advance (Kustomisasi Tata Letak & Font)**:
  * Pilihan **10 Font Faktur & Akuntansi** berfitur *tabular-nums* (`Inter`, `Roboto`, `Open Sans`, `IBM Plex Sans`, `Source Sans 3`, `PT Sans`, `Inconsolata`, `Roboto Mono`, `IBM Plex Mono`, `Tinos`) dan opsi **Paksa Semua Huruf Tebal (Bold)**.
  * Pengaturan detail ukuran font (px), tinggi baris (px), dan proporsi lebar 15 kolom tabel (%) yang tersimpan otomatis di `localStorage` browser serta dapat diekspor/diimpor sebagai file preset `.json`.
* **💾 Integrasi Arsip LHT**: Menyisipkan baris tagihan langsung ke Spreadsheet LHT tujuan tepat di atas baris `TOTAL` dengan mempertahankan format tabel.

---

## 🔄 Arsitektur & Alur Kerja Sistem

Sistem bekerja melalui 9 tahapan berurutan yang dikendalikan oleh `Perbarui Data AR.py`, dilanjutkan oleh antarmuka Google Apps Script:

```text
[Piutang.xls] & [Giro.xls] (Folder Utama)
       │
       ▼  (Divalidasi & disalin oleh Perbarui Data AR.py)
┌─────────────────────────────────────────────────────────────────────────┐
│                        PIPELINE FOLDER DAPUR/                           │
├─────────────────────────────────────────────────────────────────────────┤
│ 0_DownloaderMasterData.py ──► Unduh Master Data Google Sheets (.xlsx)   │
│ 1_CleanerAcc.py           ──► Bersihkan Piutang.xls -> Piutang_clean    │
│ 2_CleanerAccGiro.py       ──► Bersihkan Giro.xls -> Giro_temp.xlsx      │
│ 3_MergedMaster2Main.py    ──► Sinkronisasi Nama Pelanggan dari Master   │
│ 4_CleanZeroAR.py          ──► Hapus baris dengan Sisa Piutang == 0      │
│ 5_CompareGiro.py          ──► Hapus faktur yang sudah klop dengan Giro  │
│ 6_AdjDateFormat.py        ──► Format tanggal menjadi "DD MMM YYYY"      │
│ 7_HelperSSData.py         ──► Hapus kolom spasi bantu (SS & SS.1)       │
│ 8_ARInjectionToSS.py      ──► Injeksi ke Google Sheets (A2:K & Jam L2)  │
└─────────────────────────────────────────────────────────────────────────┘
       │
       ▼
[Google Spreadsheet Database (Sheet: "Piutang" & "conf")]
       │
       ▼
┌─────────────────────────────────────────────────────────────────────────┐
│             ADDON GOOGLE APPS SCRIPT (Pembuat Tagihan LHT)              │
├─────────────────────────────────────────────────────────────────────────┤
│ • Code.gs      : Backend pengambil data Piutang, conf, & simpan LHT     │
│ • smartfill.gs : Mesin pencocokan jadwal kunjungan Sales per Minggu/Hari│
│ • Index.html   : UI pembuatan LHT, kustomisasi font, & cetak PDF Vektor │
└─────────────────────────────────────────────────────────────────────────┘
```

---

## 🛠 Teknologi yang Digunakan

* **Bahasa Pemrograman**: Python 3.x, JavaScript (Google Apps Script / ES6), HTML5, CSS3
* **Pustaka Python (*Libraries*)**:
  * `pandas` & `numpy` — Manipulasi dataframe, pembersihan teks/angka, dan rekonsiliasi data Giro
  * `openpyxl` & `xlsxwriter` & `xlrd` — Pembacaan file `.xls`/`.xlsx`, pemformatan lebar kolom, dan penghapusan baris/kolom
  * `gspread` — Integrasi dan injeksi data ke Google Sheets API menggunakan Service Account
  * `requests` & `configparser` — Pengunduhan file spreadsheet eksternal dan manajemen konfigurasi `.conf`
* **Frontend & Cetakan (`Addon/Index.html`)**:
  * **Tailwind CSS (CDN)** — Desain antarmuka responsif
  * **jsPDF (v2.5.1) & jsPDF-AutoTable (v3.5.31)** — Pembuatan dokumen PDF vektor presisi di sisi klien
  * **Google Fonts** — 10 jenis huruf khusus dokumen finansial dan printer *dot-matrix*

---

## 📂 Struktur Direktori

```text
Site-Lembar-Penagihan-Harian-main/
├── Addon/                          # Kode sumber Web App / Google Sheets Add-on
│   ├── Code.gs                     # Backend utama Apps Script (doGet, ambil data, simpan LHT)
│   ├── Index.html                  # Antarmuka pengguna (UI), pratinjau cetak, & generator PDF
│   └── smartfill.gs                # Modul parser jadwal otomatis (SmartFill) dari sheet 'conf'
├── Dapur/                          # Modul internal pipeline pemrosesan data Python
│   ├── __init__.py                 # Penanda paket Python (Author: me@ihfandica.com)
│   ├── 0_DownloaderMasterData.py   # Tahap 0: Mengunduh spreadsheet Master Pelanggan
│   ├── 1_CleanerAcc.py             # Tahap 1: Membersihkan dan menata file Piutang.xls
│   ├── 2_CleanerAccGiro.py         # Tahap 2: Membersihkan dan mengonversi file Giro.xls
│   ├── 3_MergedMaster2Main.py      # Tahap 3: Memperbarui Nama Pelanggan dari Master_temp.xlsx
│   ├── 4_CleanZeroAR.py            # Tahap 4: Menghapus baris piutang dengan sisa tagihan 0
│   ├── 5_CompareGiro.py            # Tahap 5: Mencocokkan piutang dengan akumulasi Giro
│   ├── 6_AdjDateFormat.py          # Tahap 6: Menyeragamkan format tanggal (DD MMM YYYY)
│   ├── 7_HelperSSData.py           # Tahap 7: Menghapus kolom bantu 'SS' dan 'SS.1'
│   ├── 8_ARInjectionToSS.py        # Tahap 8: Mengunggah hasil akhir ke Google Sheets
│   ├── credentials.json            # Kunci autentikasi Google Cloud Service Account
│   └── piutang.conf                # Konfigurasi parameter Giro, URL Sheets, dan Master Data
├── Ekspor Data.png                 # Panduan visual format ekspor data dari aplikasi akuntansi
├── Perbarui Data AR.py             # Skrip utama (entry-point) untuk menjalankan seluruh pipeline
└── README.md                       # Dokumentasi proyek
```

---

## ⚙️ Prasyarat Sistem

Sebelum menjalankan proyek ini, pastikan Anda memiliki:

1. **Python 3.8 atau lebih baru** terpasang di komputer Anda.
2. **Akun Google Cloud Platform (GCP)** dengan **Google Sheets API** dan **Google Drive API** yang telah diaktifkan.
3. **Kredensial Service Account (`.json`)** yang telah diberi akses *Editor* pada Google Spreadsheet target.
4. File ekspor dari aplikasi akuntansi:
   * `Piutang.xls` *(Wajib ada di folder utama)*
   * `Giro.xls` *(Opsional, diperlukan jika fitur pemotongan Giro diaktifkan)*

---

## 🚀 Panduan Instalasi & Konfigurasi

### 1. Persiapan Lingkungan Python

1. **Clone repositori ini** atau ekstrak arsip `.zip` ke komputer Anda:
   ```bash
   git clone [https://github.com/username-anda/Site-Lembar-Penagihan-Harian.git](https://github.com/username-anda/Site-Lembar-Penagihan-Harian.git)
   cd Site-Lembar-Penagihan-Harian
   ```

2. **Buat dan aktifkan *virtual environment* (opsional namun disarankan):**
   ```bash
   python -m venv venv
   # Untuk Windows:
   venv\Scripts\activate
   # Untuk macOS / Linux:
   source venv/bin/activate
   ```

3. **Instal seluruh pustaka Python yang dibutuhkan:**
   ```bash
   pip install pandas numpy openpyxl xlsxwriter xlrd python-calamine requests gspread
   ```

### 2. Konfigurasi Google Service Account (`credentials.json`)

1. Buka [Google Cloud Console](https://console.cloud.google.com/).
2. Buat *Service Account* baru dan unduh kunci privat dalam format **JSON**.
3. Ubah nama file JSON tersebut menjadi `credentials.json` dan letakkan di dalam folder `Dapur/credentials.json`.
4. Salin alamat email *Service Account* (`client_email` di dalam `credentials.json`), lalu **Bagikan (Share)** Google Spreadsheet database Anda ke email tersebut dengan hak akses **Editor**.

### 3. Konfigurasi Pipeline (`piutang.conf`)

Buka file `Dapur/piutang.conf` menggunakan teks editor dan sesuaikan nilainya:

```ini
[GIRO]
giro_stats = Ya
giro_cut = 1000

[URL]
url_ss = [https://docs.google.com/spreadsheets/d/ID_SPREADSHEET_DATABASE_ANDA/edit](https://docs.google.com/spreadsheets/d/ID_SPREADSHEET_DATABASE_ANDA/edit)
url_ss_sheet = Piutang

[MASTER]
masdatus = Ya
url = [https://docs.google.com/spreadsheets/d/ID_SPREADSHEET_MASTER_ANDA/edit](https://docs.google.com/spreadsheets/d/ID_SPREADSHEET_MASTER_ANDA/edit)
mas_sheet = MIRROR PCMO
mas_col_key = NOPEL
mas_col_ret = NAMA PELANGGAN
```

### 4. Pemasangan Web App / Add-on Google Sheets

1. Buka Google Spreadsheet database utama Anda.
2. Pastikan terdapat 2 lembar kerja (*sheet*) wajib:
   * **`Piutang`**: Tempat data piutang diinjeksi oleh Python (baris 1 berisi header, `A2:K` berisi data piutang, dan `L2` berisi waktu pembaruan).
   * **`conf`**: Berisi konfigurasi aplikasi dengan susunan kolom sebagai berikut:
     * Kolom A–E: Daftar **Admin**, **Penagih**, **Depo**, **Perusahaan**, dan **Divisi**.
     * Kolom **Tautan LHT** & **LHT Sheet**: URL Google Spreadsheet tujuan penyimpanan rekap LHT beserta nama *sheet*-nya.
     * Kolom SmartFill: **Nama Penagih**, **Minggu** (`Minggu 1` s/d `Minggu 4`), **Hari** (`Senin` s/d `Minggu`), dan **No. Pelanggan** (daftar kode pelanggan dipisahkan karakter pipa `|`, contoh: `CUST001|CUST002|CUST003`).
3. Pada menu Google Sheets, klik **Ekstensi (Extensions) > Apps Script**.
4. Buat/sesuaikan 3 berkas berikut di editor Apps Script dan salin isinya dari folder `Addon/`:
   * `Code.gs`
   * `smartfill.gs`
   * `Index.html`
5. Klik **Terapkan (Deploy) > Deployment baru (New deployment)**, pilih jenis **Aplikasi Web (Web app)**, lalu salin URL Web App yang dihasilkan untuk digunakan oleh tim Admin.

---

## 💻 Cara Penggunaan

### A. Memperbarui Data Piutang (Pipeline Python)

1. Ekspor data piutang dan giro terbaru dari software akuntansi Anda sesuai panduan pada gambar `Ekspor Data.png`.
2. Simpan file hasil ekspor di **folder utama** (sejajar dengan `Perbarui Data AR.py`) dengan nama:
   * `Piutang.xls` *(Wajib)*
   * `Giro.xls` *(Opsional)*
3. Jalankan skrip otomatisasi utama dengan klik ganda pada `Perbarui Data AR.py` atau melalui terminal:
   ```bash
   python "Perbarui Data AR.py"
   ```
4. Program akan memeriksa kelengkapan file di folder `Dapur/`, menyalin `Piutang.xls` dan `Giro.xls`, lalu menjalankan `0_DownloaderMasterData.py` hingga `8_ARInjectionToSS.py` secara otomatis.
5. Setelah muncul pesan `Selesai! Data berhasil diperbarui di Google Sheets`, tekan **Enter** untuk menutup jendela terminal.

### B. Membuat & Mencetak Lembar LHT (Web App)

1. Buka URL **Web App Pembuat Tagihan LHT**.
2. Pilih identitas dokumen pada panel atas: **Nama Depo**, **Penagih (Sales)**, **Perusahaan**, **Divisi**, **Tanggal**, dan **Admin**.
3. **Pemilihan Pelanggan**:
   * Saat memilih **Penagih** dan **Tanggal**, fitur **✨ SmartFill** akan otomatis mencentang daftar toko/pelanggan sesuai jadwal kunjungan minggu dan hari tersebut.
   * Anda juga dapat mengubah mode minggu secara manual (`Paksa Minggu 1` s/d `Minggu 4`) atau mencari pelanggan lewat kolom pencarian.
4. **Pengaturan Baris Faktur**:
   * Centang atau hapus centang pada faktur tertentu di tabel *Daftar Faktur Piutang Dipilih*.
   * Gunakan tombol **+ Tambah Baris Kosong** / **− Hapus Baris Kosong** jika membutuhkan baris tambahan untuk tulis tangan saat penagihan di lapangan.
5. **Kustomisasi Cetakan (Opsional)**:
   * Pilih ukuran kertas (**Continuous Form / Letter** atau **A4 Landscape**).
   * Buka **⚙️ Mode Advance** jika ingin mengganti jenis font, mengatur lebar kolom, atau memuat preset tata letak (`.json`).
6. **Simpan & Cetak**:
   * Klik **💾 Simpan Spreadsheet** untuk mengirim data ke sheet rekap LHT.
   * Klik **🖨️ Cetak Langsung (PDF)** atau **📥 Simpan & Unduh PDF** untuk mengarsipkan sekaligus mengunduh dokumen LHT.

---

## 🔧 Referensi Konfigurasi

Berikut adalah penjelasan parameter yang terdapat pada berkas `Dapur/piutang.conf`:

| Seksi | Parameter | Contoh Nilai | Deskripsi |
| :--- | :--- | :--- | :--- |
| `[GIRO]` | `giro_stats` | `Ya` / `Tidak` | Mengaktifkan (`Ya`) atau melewati (`Tidak`) pemrosesan `Giro.xls` dan rekonsiliasi giro. |
| `[GIRO]` | `giro_cut` | `1000` | Batas toleransi selisih (Rupiah) antara Nilai Faktur dan Nilai Terima Giro agar faktur dianggap lunas/klop. |
| `[URL]` | `url_ss` | `https://docs.google.com/...` | URL lengkap Google Spreadsheet database tujuan injeksi data piutang. |
| `[URL]` | `url_ss_sheet` | `Piutang` | Nama *worksheet* target di dalam Spreadsheet untuk menyimpan data piutang bersih. |
| `[MASTER]` | `masdatus` | `Ya` / `Tidak` | Mengaktifkan (`Ya`) atau melewati (`Tidak`) pengunduhan dan pencocokan nama pelanggan dari Master Data. |
| `[MASTER]` | `url` | `https://docs.google.com/...` | URL Google Spreadsheet Master Data pelanggan (harus dapat diakses publik/ekspor `.xlsx`). |
| `[MASTER]` | `mas_sheet` | `MIRROR PCMO` | Nama *sheet* pada file Master Data yang berisi tabel referensi pelanggan. |
| `[MASTER]` | `mas_col_key` | `NOPEL` | Nama kolom kunci kode pelanggan pada file Master Data. |
| `[MASTER]` | `mas_col_ret` | `NAMA PELANGGAN` | Nama kolom nilai nama pelanggan yang akan diambil dari file Master Data. |

---

## 🛠️ Penanganan Masalah Umum (Troubleshooting)

* **`[ERROR] Berkas wajib 'Piutang.xls' nggak ditemukan di folder utama!`**
  Pastikan file `Piutang.xls` diletakkan di direktori akar proyek (sejajar dengan `Perbarui Data AR.py`), bukan hanya di dalam folder `Dapur/`.
* **`Error: Kolom No. Faktur tidak ditemukan.` pada `1_CleanerAcc.py`**
  Pastikan format ekspor `Piutang.xls` sesuai dengan standar pada `Ekspor Data.png` dan memiliki kolom `No. Faktur` pada 150 baris pertama.
* **`Gagal mengunduh Master_temp.xlsx` pada `0_DownloaderMasterData.py`**
  Pastikan `url` pada seksi `[MASTER]` di `piutang.conf` telah diisi dengan tautan Google Spreadsheet yang valid dan memiliki izin berbagi (*Anyone with the link can view*).
* **`gspread.exceptions.SpreadsheetNotFound` atau `PermissionError` pada `8_ARInjectionToSS.py`**
  Pastikan `url_ss` di `piutang.conf` sudah benar dan Google Spreadsheet tersebut telah dibagikan (*Share*) ke alamat `client_email` yang tertera di dalam `Dapur/credentials.json`.

---

## 🤝 Kontribusi

Kontribusi untuk pengembangan fitur baru maupun perbaikan *bug* sangat terbuka:

1. Lakukan **Fork** pada repositori ini.
2. Buat *branch* fitur baru (`git checkout -b fitur/nama-fitur`).
3. Lakukan *commit* perubahan Anda (`git commit -m 'Menambahkan fitur X'`).
4. *Push* ke *branch* Anda (`git push origin fitur/nama-fitur`).
5. Buat **Pull Request** baru.

---

## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal perusahaan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)*
