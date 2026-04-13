# 🏥 RME Data Cleaner — Panduan Deployment

Aplikasi Streamlit untuk membersihkan data Rekam Medis Elektronik (RME) Puskesmas.

---

## 🚀 Cara Deploy ke Streamlit Cloud (Gratis)

### Langkah 1 — Siapkan Repository GitHub
1. Buat akun di [github.com](https://github.com) jika belum punya
2. Buat repository baru (klik tombol **New**)
3. Upload dua file berikut ke repository tersebut:
   - `app.py`
   - `requirements.txt`

### Langkah 2 — Deploy di Streamlit Cloud
1. Buka [share.streamlit.io](https://share.streamlit.io)
2. Login dengan akun GitHub
3. Klik **New app**
4. Pilih repository yang sudah dibuat
5. Pastikan **Main file path** diisi: `app.py`
6. Klik **Deploy!**

Aplikasi akan otomatis aktif dalam 1–2 menit. ✅

---

## 💻 Cara Jalankan di Komputer Lokal

```bash
# Install dependencies
pip install -r requirements.txt

# Jalankan aplikasi
streamlit run app.py
```

Buka browser dan akses: `http://localhost:8501`

---

## ✨ Fitur Pembersihan

| Fitur | Deskripsi |
|-------|-----------|
| 🗑️ Hapus Duplikat | Deteksi dan hapus baris yang persis sama |
| ✂️ Trailing Koma | Hapus ` ,` di akhir sel (No RM, NIK, No Penjamin) |
| 📅 Format Tanggal | Ubah datetime ke format `dd/mm/yyyy` |
| 📝 Standarisasi Keterangan | Ubah `-` dan `=` menjadi teks bermakna |
| 🔠 Normalisasi Teks | UPPERCASE Nama, Title Case Desa |
| 📋 Isi Nilai Kosong | Isi kolom kosong dengan nilai default yang bisa dikonfigurasi |
| ❌ Hapus Kolom | Pilih kolom mana yang ingin dihapus |
| 🔃 Urutkan Data | Sortir berdasarkan kolom pilihan |

---

## 📁 Format yang Didukung

- **Input:** `.xlsx`, `.xls`, `.csv`
- **Output:** `.xlsx` (dengan formatting rapi) atau `.csv`

---

## 📋 Struktur File

```
├── app.py            ← Aplikasi utama Streamlit
├── requirements.txt  ← Daftar library Python
└── README.md         ← Panduan ini
```
