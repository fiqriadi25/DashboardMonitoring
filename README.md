# Dashboard Digital Enablement

Web app Google Apps Script untuk memantau proses **Rekon Bitrans** dan
**SPJ Non-Rekon**. Repository ini juga menyediakan script Google Colab untuk
mengolah file operasional dan memperbarui database Google Sheets yang digunakan
dashboard.

## Fitur utama

- Monitoring Rekon harian berdasarkan bulan, customer, OP, TOH/WOH, produk, dan status.
- Tracking SPJ Non-Rekon beserta status adjustment (`Done`, `Process`, atau `No Adjustment`).
- Pencarian, filter, pagination, refresh data, dan download laporan Non-Rekon.
- Dashboard Looker Studio dalam satu navigasi.
- Administration Center untuk mapping Operating Point, TOH/WOH, email PIC, area, dan konfigurasi dashboard.
- Kontrol akses berdasarkan domain perusahaan, allowlist, pengajuan akses, persetujuan admin, dan verifikasi email 6 digit.
- Cache bertingkat untuk mempercepat pembacaan data Google Sheets.
- Email pengingat Non-Rekon otomatis kepada TOH/WOH pada hari kerja.
- History log dan upload log untuk audit pembaruan database.

## Struktur repository

| File | Keterangan |
|---|---|
| `Code.gs` | Backend Google Apps Script: akses, API dashboard, cache, laporan email, trigger, dan administrasi. |
| `index.html` | Antarmuka dashboard responsif untuk Rekon, SPJ, Looker Studio, dan Administration Center. |
| `generate_data_monitoring_rekon.py` | Script Google Colab untuk menggabungkan file input, memperbarui database Rekon, dan membangun tracker SPJ Non-Rekon. |
| `requirements.txt` | Dependensi Python untuk menjalankan generator di luar environment standar Colab. |

## Alur sistem

1. Pengguna mengunggah satu atau beberapa file CSV/XLSX ke Google Colab.
2. Script menormalkan header dan tanggal, lalu memisahkan data Rekon dan Non-Rekon.
3. Data agregat diperbarui ke sheet `Database`; perubahan dicatat pada `History_Log` dan `Upload_Log`.
4. Data SPJ individual diperbarui ke sheet `Data_Rekon`, termasuk lookup TOH/WOH dan status adjustment.
5. Google Apps Script membaca kedua sumber tersebut, melakukan caching, lalu menyajikannya melalui web app.
6. Trigger perubahan spreadsheet memperbarui token/cache sehingga dashboard memperoleh data terbaru.

## Sumber data dan sheet

### Database Rekon

- `Database`: data Rekon yang sudah diagregasi.
- `Master_OP_OH`: mapping OP, customer, TOH/WOH, dan email PIC.
- `History_Log`: audit insert, update, delete, dan perubahan produk.
- `Upload_Log`: ringkasan setiap proses upload.
- `Email_Log`: histori email laporan yang berhasil dikirim.
- `Access_Allowlist`: daftar email eksternal yang diizinkan.
- `Access_Request`: antrean permintaan akses.
- `Dashboard_Config`: konfigurasi dashboard yang dikelola admin.

### Tracker SPJ

- `Data_Rekon`: data SPJ Non-Rekon per transaksi.
- `Master_OP_OH`: mapping OP untuk sumber tracker.
- `Config`: waktu pembaruan terakhir.

### Data adjustment

- `DATA_REKAP`: referensi status adjustment berdasarkan nomor SPJ.

## Kolom input minimum

Generator mengenali variasi penamaan header, kemudian memetakannya ke kolom berikut:

- `Kode OP`
- `Nama OP`
- `Customer`
- `SPJ`
- `Produk`
- `Tanggal SLA`
- `UJP`
- `Charging`
- `Status Rekon`

Baris tanpa tanggal SLA yang valid tidak diproses.

## Menjalankan generator di Google Colab

1. Buka Google Colab dan buat notebook baru.
2. Unduh script dari repository atau clone repository:

   ```python
   !git clone https://github.com/fiqriadi25/DashbboardDigital.git
   %cd DashbboardDigital
   !pip install -r requirements.txt
   ```

3. Jalankan script:

   ```python
   %run generate_data_monitoring_rekon.py
   ```

4. Login ke akun Google yang mempunyai akses edit ke spreadsheet sumber.
5. Upload file CSV/XLSX saat dialog upload muncul.
6. Periksa summary proses serta sheet `History_Log` dan `Upload_Log`.

> Script menggunakan ID/URL spreadsheet pada bagian konfigurasi. Pastikan nilai
> tersebut mengarah ke spreadsheet environment yang benar sebelum dijalankan.

## Setup Google Apps Script

1. Buat atau buka project Google Apps Script.
2. Salin `Code.gs` dan `index.html` ke project tersebut.
3. Periksa konstanta konfigurasi di bagian atas `Code.gs`, terutama:
   - ID spreadsheet Rekon dan SPJ;
   - domain/email yang diizinkan;
   - email administrator;
   - URL Looker Studio dan pengaturan dashboard.
4. Atur timezone project ke `Asia/Jakarta`.
5. Jalankan `setupProjectTriggers()` satu kali dari editor dan berikan izin yang diminta.
6. Jika laporan email harian diperlukan, buat time-driven trigger untuk
   `sendDailyNonRekonEmail` sesuai jadwal operasional. Fungsi otomatis melewati
   Sabtu dan Minggu.
7. Deploy melalui **Deploy → New deployment → Web app**, lalu pilih cakupan akses
   yang sesuai kebijakan perusahaan.

Setelah mengubah kode, buat deployment version baru agar perubahan tampil pada
URL web app produksi.

## Kontrol akses

- Akun dengan domain perusahaan yang terdaftar dapat langsung masuk.
- Email eksternal harus tersedia pada `Access_Allowlist` atau mempunyai permintaan
  berstatus `APPROVED` di `Access_Request`.
- Bila identitas Google tidak tersedia pada sesi web app, sistem mengirim kode
  verifikasi 6 digit ke email yang telah disetujui.
- Menu Administration hanya tersedia untuk email admin yang dikonfigurasi.

Jangan menyimpan password, token, service-account JSON, atau credential lain di
repository. Batasi akses spreadsheet dan deployment sesuai kebutuhan operasional.

## Pemeriksaan setelah deployment

- Tab Rekon menampilkan periode, statistik, filter, dan tabel monitoring.
- Tab Non-Rekon menampilkan data SPJ, status adjustment, dan pagination.
- Tombol refresh memperbarui data tanpa error.
- Admin dapat membuka Administration Center dan menyimpan mapping/config.
- `setupProjectTriggers()` telah menghasilkan trigger edit/change tanpa duplikasi.
- Email harian tercatat di `Email_Log` dan tidak terkirim ganda untuk periode yang sama.

## Troubleshooting singkat

| Masalah | Pemeriksaan |
|---|---|
| Dashboard kosong | Pastikan nama sheet dan kolom sesuai, lalu cek izin spreadsheet. |
| Akses ditolak | Periksa domain, `Access_Allowlist`, dan status pada `Access_Request`. |
| Data belum berubah | Jalankan refresh, periksa sheet `Config`, kemudian reset cache dari menu admin. |
| Mapping TOH/WOH kosong | Lengkapi `Master_OP_OH` dengan kombinasi Kode OP, Nama OP, dan Customer yang sesuai. |
| Email tidak terkirim | Periksa trigger, kuota MailApp, email PIC, dan `Email_Log`. |
| Generator gagal membaca file | Pastikan format CSV/XLSX dan seluruh kolom minimum tersedia. |

## Catatan

Data operasional dan file hasil pengolahan tidak disimpan di repository. Gunakan
repository ini hanya untuk source code dan dokumentasi sistem.
