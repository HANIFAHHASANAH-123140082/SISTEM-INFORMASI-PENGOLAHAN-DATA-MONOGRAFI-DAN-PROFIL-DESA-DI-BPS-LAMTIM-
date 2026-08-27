# Sistem Informasi Pengolahan Data Monografi dan Profil Desa

Sistem berbasis web untuk mengotomatisasi pengolahan data kuesioner Pendataan Potensi Desa (Podes) menjadi dokumen Monografi Desa dan Profil Desa, dikembangkan sebagai bagian dari Praktek Kerja Lapangan di Badan Pusat Statistik Kabupaten Lampung Timur.

## Tentang Project

Sistem ini dibangun untuk menjawab permasalahan proses penyusunan dokumen Monografi Desa dan Profil Desa yang sebelumnya dilakukan secara manual, sehingga rentan terhadap data yang tidak lengkap dan proses pengerjaan yang memakan waktu lama. Sistem memungkinkan petugas untuk:

- Mengelola data desa (tambah, edit, hapus)
- Mengisi jawaban 143 item kuesioner Podes secara manual
- Mengunggah berkas kuesioner (.docx atau .pdf) untuk pengisian jawaban secara otomatis
- Mengunduh dokumen Profil Desa dan Monografi Desa dalam format Microsoft Word (.docx) yang siap digunakan dan tetap dapat disunting

## Teknologi yang Digunakan

- **Backend:** Django 5.2 (Python 3.11)
- **Basis Data:** SQLite
- **Pembuatan Dokumen Word:** python-docx
- **Pengolahan Gambar Sampul:** Pillow
- **Pembacaan Berkas PDF:** pdfplumber
- **Frontend:** Bootstrap 5.3

## Instalasi dan Menjalankan Sistem

### Prasyarat
- Python 3.11 atau lebih baru sudah terinstal
- Git sudah terinstal

### Langkah Instalasi

1. Clone repository ini
git clone https://github.com/HANIFAHHASANAH-123140082/SISTEM-INFORMASI-PENGOLAHAN-DATA-MONOGRAFI-DAN-PROFIL-DESA-DI-BPS-LAMTIM-.git
cd SISTEM-INFORMASI-PENGOLAHAN-DATA-MONOGRAFI-DAN-PROFIL-DESA-DI-BPS-LAMTIM-

2. Buat dan aktifkan virtual environment
python -m venv venv

Windows (PowerShell):
venv\Scripts\Activate.ps1

Kalau muncul error izin PowerShell, jalankan dulu (sekali saja):
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned

3. Install seluruh library yang dibutuhkan
pip install -r requirements.txt

4. Jalankan migrasi basis data
python manage.py migrate

5. Isi data master (11 bab dan 143 item kuesioner)
python manage.py seed_data

6. Buat akun administrator
python manage.py createsuperuser

7. Jalankan server
python manage.py runserver

8. Buka browser ke `http://127.0.0.1:8000/`

## Cara Penggunaan

### 1. Menambahkan Desa Baru
Klik tombol **"Tambah Desa"** di halaman utama, isi data dasar desa (nama desa, kecamatan, kabupaten, koordinat, sejarah desa, nama kepala desa, foto kepala desa, peta wilayah). Setelah disimpan, sistem otomatis mengarahkan ke halaman input jawaban kuesioner.

### 2. Mengisi Jawaban Kuesioner
Terdapat dua cara mengisi jawaban 143 item kuesioner:

**a. Input Manual**
Isi langsung tiap kolom jawaban pada tabel yang tersedia. Kolom "Pilihan Jawaban" menampilkan referensi pilihan sesuai kuesioner asli sebagai panduan pengisian. Gunakan tombol panah atas/bawah pada keyboard untuk berpindah antar baris dengan cepat.

**b. Upload Otomatis**
Klik tombol **"Upload Berkas Kuesioner untuk Isi Otomatis"** pada halaman input jawaban, unggah berkas kuesioner yang sudah terisi (format .docx atau .pdf). Sistem akan otomatis membaca dan mengisi jawaban ke dalam formulir. Periksa kembali hasilnya sebelum menyimpan.

### 3. Mengunduh Dokumen
Pada halaman daftar desa, klik tombol unduh pada desa yang datanya sudah diisi untuk mengunduh:
- **Profil Desa** — dokumen lengkap dengan sampul, kata pengantar, daftar isi, dan data per bab (A–K)
- **Monografi Desa** — dokumen dengan 4 halaman CV kosong (Kepala Desa, Ketua BPD, Ketua LPMD, Ketua PKK) dan data per bidang

Bab yang tidak memiliki data akan otomatis ditampilkan sebagai kalimat deskriptif, dan baris tabel yang kosong akan otomatis disembunyikan.

### 4. Mengelola Data Desa
Gunakan tombol edit dan hapus pada halaman daftar desa untuk menyunting atau menghapus data desa yang sudah tersimpan.

## Struktur Project
sistem-monografi-desa/
├── desa/
│ ├── models.py # Struktur basis data (6 entitas)
│ ├── views.py # Logika aplikasi
│ ├── forms.py # Formulir input data
│ ├── document_generator.py # Generator dokumen Word otomatis
│ ├── kuesioner_parser.py # Pembaca berkas kuesioner (.docx/.pdf)
│ ├── validators.py # Validasi data (opsional)
│ ├── templates/desa/ # Tampilan halaman web
│ ├── management/commands/ # Script pengisian data master
│ └── assets/ # Logo dan gambar sampul
├── monografi_desa/ # Konfigurasi project Django
├── requirements.txt # Daftar library yang dibutuhkan
└── manage.py


## Konteks Pengembangan

Sistem ini dikembangkan sebagai bagian dari Praktek Kerja Lapangan (Mata Kuliah IF4004) mahasiswa Program Studi Teknik Informatika, Institut Teknologi Sumatera, yang dilaksanakan di Badan Pusat Statistik Kabupaten Lampung Timur.

**Penulis:** Hanifah Hasanah (123140082)
**Instansi:** Badan Pusat Statistik Kabupaten Lampung Timur
**Periode:** Juni – Agustus 2026
