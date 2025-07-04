# Otomatisasi Scan dan Ekstrak Data Faktur dengan Google Apps Script

Repositori menggunakan Google Apps Script untuk otomatiskan seluruh alur kerja pemrosesan faktur, mulai dari pemindaian dokumen hingga pencatatan data terstruktur di Google Sheet dan pengarsipan dokumen.

## Fitur Utama

- **Multi-Format:** Mampu memproses file faktur dalam format PDF, gambar (JPG, PNG), dan `.docx` secara otomatis.
- **Dua Mekanisme Ekstraksi Data:**
    - Untuk PDF dan gambar, menggunakan **OCR API** (dalam contoh kasus ini, OCR.space API) untuk mengekstrak teks.
    - Untuk `.docx`, melakukan konversi aman ke Google Doc untuk membaca teks secara langsung, memastikan akurasi dan kecepatan maksimal.
- **Dua Logics Ekstrak Data:** logika parser terpisah dan dioptimalkan untuk setiap jenis sumber data (dalam hal ini, biasanya hasil OCR cenderung tidak terstruktur vs. pembacaan langsung (direct) `.docx` lebih bersih sehingga butuh logika berbeda agar keduanya dapat diekstrak dengan baik).
- **Pengarsipan Otomatis:**
    - Membuat struktur folder dinamis untuk dokumentasi pembayaran dengan format `Tahun/Nama Penerima/Bulan`.
    - Menyimpan copy teks mentah dari setiap dokumen yang diproses dengan format nama `DDMMYYYY-HHmmss-NamaFileAsli.txt` di folder terpisah.
    - Memindahkan file faktur yang telah diproses ke folder arsip.
- **Pencatatan di Spreadsheet:** Secara otomatis menambahkan baris baru di Google Sheet yang berisi data terstruktur, status pemrosesan, dan tautan langsung ke folder dokumentasi.
- **Tangguh dan Andal:** Dilengkapi mekanisme coba lagi (retry) saat koneksi ke API eksternal gagal dan metode konversi file yang stabil.

## Struktur Proyek

```
.
├── Kode.gs         # File utama berisi semua logika Google Apps Script.
└── appsscript.json # File manifest (dibuat otomatis oleh Google).
└── README.md       # Dokumentasi ini.
```

## Persiapan dan Konfigurasi

Sebelum menjalankan skrip, pastikan Anda telah melakukan persiapan berikut di lingkungan Google Anda.

### 1. Konfigurasi Google Drive
Buat 4 folder berikut di Google Drive Anda dan catat ID-nya (ID bisa dilihat dari URL folder):
- **Folder Faktur Masuk:** Tempat Anda meletakkan file faktur baru untuk diproses.
- **Folder Faktur Terproses:** Tempat skrip akan memindahkan file setelah selesai diproses.
- **Folder Dokumen Pembayaran:** Folder utama tempat struktur `Tahun/Nama Penerima/Bulan` akan dibuat.
- **Folder Teks Mentah:** Folder khusus untuk menyimpan semua file `.txt` hasil ekstraksi teks.

### 2. Konfigurasi Google Sheet
- Buat file Google Sheet baru.
- Ubah nama "Sheet1" menjadi nama yang Anda inginkan (misalnya, "Contoh").
- Buat header di baris pertama dengan urutan persis seperti ini:
  `Sumber Faktur`, `Sumber API`, `Waktu Proses`, `Nomor Faktur`, `Tanggal Faktur`, `Penerima Faktur`, `Total Pembayaran`, `Status`, `Dokumen Pembayaran`

### 3. Dapatkan Kunci API
- Skrip ini menggunakan OCR API (dalam contoh script ini **OCR.space**) untuk memproses PDF/gambar.
- Apabila Anda tidak ingin memodifikasi lebih lanjut script ini, daftar di [situs web OCR.space](https://ocr.space/ocrapi/free) untuk dapat `apikey`.

### 4. Konfigurasi Skrip (`Kode.gs`)
Buka file `Kode.gs` dan isi semua variabel di bagian `KONFIGURASI PENGGUNA` dengan ID Folder dan Kunci API yang telah Anda siapkan.

```javascript
// ===============================================================
// KONFIGURASI PENGGUNA (Universal)
// ===============================================================

// -- Konfigurasi API (Saat ini == OCR.space) --
const OCR_API_KEY = 'MASUKKAN_API_KEY_ANDA_DI_SINI'; 
const OCR_API_ENDPOINT = '[https://api.ocr.space/parse/image](https://api.ocr.space/parse/image)';

// -- Konfigurasi Google Drive & Sheets --
const SOURCE_FOLDER_ID = 'MASUKKAN_ID_FOLDER_SUMBER_ANDA';
const PROCESSED_FOLDER_ID = 'MASUKKAN_ID_FOLDER_TERPROSES_ANDA';
const PAYMENT_DOCS_ROOT_FOLDER_ID = 'MASUKKAN_ID_FOLDER_DOKUMEN_ANDA';
const RAW_TEXT_FOLDER_ID = 'MASUKKAN_ID_FOLDER_TEKS_MENTAH_ANDA';
const SHEET_NAME = 'Contoh'; // Sesuaikan dengan nama sheet Anda
```

### 5. Aktifkan Layanan Tingkat Lanjut
Skrip ini memerlukan **Drive API Service** untuk mengonversi file `.docx`.
- Di Editor Apps Script, klik menu **Services +**.
- Cari dan pilih **Drive API**.
- Klik **Add**. Pastikan Identifier-nya adalah `Drive`.

## Instalasi dan Menjalankan Skrip

1.  Buka Google Sheet yang telah Anda siapkan.
2.  Buka menu **Extensions > Apps Script**.
3.  Salin seluruh konten dari file `Kode.gs` dan tempelkan ke editor, menggantikan semua kode yang ada.
4.  Lakukan **Konfigurasi Skrip** dan **Aktifkan Layanan Tingkat Lanjut** seperti yang dijelaskan di atas.
5.  Simpan proyek.
6.  **Otorisasi Awal:** Jalankan fungsi `prosesFakturUniversal` secara manual dari editor untuk pertama kalinya. Google akan meminta serangkaian izin, kemudian setujui semua izin.
7.  **Atur Pemicu (Trigger):**
    - Di editor, klik menu **Triggers** (ikon jam).
    - Klik **Add Trigger**.
    - Pilih fungsi `prosesFakturUniversal` untuk dijalankan.
    - Pilih sumber acara `Time-driven`.
    - Atur intervalnya (misalnya, `Hour timer` untuk berjalan setiap jam).
    - Simpan trigger.

## Cara Penggunaan

Cukup letakkan file faktur (PDF, JPG, PNG, atau DOCX) ke dalam folder "Faktur Masuk" di Google Drive Anda. Skrip akan secara otomatis memprosesnya pada interval waktu berikutnya sesuai dengan trigger yang Anda atur.

---

Kontribusi dan saran untuk perbaikan sangat diterima. Silakan buat *issue* atau *pull request*.
