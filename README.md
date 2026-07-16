# Portal Akademik Smart

Aplikasi manajemen akademik sekolah berbasis **Google Apps Script** — satu portal untuk Admin, Guru, dan Siswa.

## ✨ Fitur Utama

- **Admin:** Dashboard, Manajemen Akun (CRUD + Import Excel), Kelola Kelas, Log Aktivitas
- **Guru:** Presensi Siswa (dengan koreksi & indikator visual), Materi & Tugas (upload ke Google Drive, jadwal tayang per kelas), Rekap Nilai & Kehadiran, Asisten AI (Gemini) untuk bantu tulis deskripsi & analisis rekap kelas
- **Siswa:** Lihat Materi & kumpulkan Tugas, Rekap Kehadiran & status pengumpulan tugas
- Autentikasi berbasis NISN/Username, tersimpan di Google Sheets sebagai database

## 🗂️ Struktur Proyek

| File | Keterangan |
|---|---|
| `Code.gs` | Seluruh logika backend (Apps Script / server-side) |
| `Index.html` | Seluruh tampilan & logika frontend (HTML, Tailwind CSS, JavaScript) |

## 🚀 Cara Deploy

1. Buat Google Spreadsheet baru sebagai database aplikasi.
2. Buka **Extensions > Apps Script** dari spreadsheet tersebut.
3. Salin isi `Code.gs` ke file `Code.gs` di editor Apps Script.
4. Buat file HTML baru bernama `Index`, lalu salin isi `Index.html` ke dalamnya.
5. Klik **Deploy > New deployment** → pilih tipe **Web app**.
   - Execute as: **Me**
   - Who has access: sesuaikan kebutuhan (mis. **Anyone within [domain]** untuk sekolah Google Workspace)
6. (Opsional) Jika ingin mengaktifkan fitur Asisten AI, tambahkan **Script Property** bernama `GEMINI_API_KEY` (Project Settings > Script Properties) dengan API key dari [Google AI Studio](https://aistudio.google.com).
7. Buka URL Web App yang dihasilkan, lalu tambahkan akun Admin pertama secara manual di sheet "Users" pada spreadsheet.

## 🛠️ Tech Stack

- Google Apps Script (backend)
- Google Sheets (database)
- Google Drive (penyimpanan file materi/tugas)
- Tailwind CSS + Font Awesome (tampilan)
- SweetAlert2 (notifikasi)
- SheetJS, jsPDF (export Excel/PDF)
- Google Gemini API (asisten AI)

## 📌 Catatan

Proyek ini dikembangkan secara bertahap untuk kebutuhan sekolah. Modul Bank Soal & CBT (Ujian Online) masih dalam tahap pengembangan.
