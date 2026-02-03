# Aplikasi Permohonan Geser Meter - PLN ULP Dinoyo

Aplikasi web untuk pencatatan dan pengelolaan permohonan geser meter di PLN ULP Dinoyo. Dibangun pakai Streamlit + Google Sheets.

🔗 **Live:** [permohonan-geser-meter.streamlit.app](https://permohonan-geser-meter.streamlit.app/)

---

## Kenapa Aplikasi Ini Dibuat?

Sebelumnya, pencatatan survey dan eksekusi geser meter masih dilakukan manual oleh petugas lapangan. Masalahnya:

- **Data material sering tercecer** — catatan kertas hilang atau tidak terbaca
- **Rekapitulasi lambat** — harus input ulang dari catatan lapangan ke spreadsheet
- **Susah tracking** — tidak jelas mana yang sudah di-survey, mana yang sudah dieksekusi

Aplikasi ini dibuat supaya petugas bisa langsung input data dari HP di lapangan, dan data langsung masuk ke sistem pusat. Tidak perlu catat manual lagi.

---

## Alur Kerja

### 1. Pelanggan Isi Form
Pelanggan isi data permohonan lewat Google Form. Data masuk ke Google Spreadsheet.

### 2. Survey Lapangan
Petugas survey ke lokasi pelanggan. Selesai survey, buka website lalu isi:
- Kuantitas barang yang dibutuhkan
- Foto hasil survey

![Halaman Input Barang & Survey](assets/screenshot_proses.png)

### 3. Simpan Data
Klik **Simpan**. Data masuk ke bagian *"Data Survey yang Sudah Tersimpan"*.

![Input Kuantitas Barang](assets/screenshot_input_barang.png)

### 4. Export Rekap
Pilih ID Pelanggan, lalu export. Sistem generate 2 rekap:
- **Rekap Pelanggan** — rincian material + biaya
- **Rekap Vendor** — kalkulasi harga vendor

Data di-copy dari template yang sudah ada formula untuk kolom Material, Jasa, dan Total.

### 5. Tanggal Survey Otomatis Terisi
Setelah export, kolom **Tanggal Survey** di spreadsheet otomatis terisi (`DD/MM/YYYY HH:MM:SS`).

### 6. Proses Admin
Admin klik **Sync To Rekap_Material** untuk sinkronisasi.

### 7. Eksekusi Lapangan
Petugas eksekusi ke alamat pelanggan. Setelah selesai, buka menu **Eksekusi** di website:
- Pilih ID Pelanggan
- Upload foto bukti pengerjaan

![Halaman Eksekusi](assets/screenshot_eksekusi.png)

### 8. Tanggal Eksekusi Otomatis Terisi
Kolom **Tanggal Eksekusi** di spreadsheet otomatis terisi (`DD/MM/YYYY`).

---

## Tech Stack

- **Streamlit** — framework web
- **Google Sheets API** — database utama
- **Google Drive API** — simpan foto
- **Cloudflare R2** — upload foto survey/eksekusi
- **Python**

---

## Cara Jalankan

```bash
pip install -r requirements.txt
streamlit run app.py
```

---

## Pengembang

Dibuat oleh mahasiswa **Universitas Brawijaya** untuk PLN ULP Dinoyo.

---

*© PLN Indonesia*
