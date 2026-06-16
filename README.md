# 📊 Sistem Rekap Pembuatan & Revisi Software

Aplikasi berbasis web yang dibangun menggunakan **Streamlit** dan **Pandas** untuk mempermudah manajemen, pelacakan, dan perekapan data revisi maupun pembuatan software baru. Aplikasi ini terintegrasi langsung dengan file database utama berbasis spreadsheet (`.xlsx`) secara *real-time*.

---

## ✨ Fitur Utama

- **📝 Form Input Data Baru (Sidebar):** Input data tugas/revisi baru secara praktis dengan validasi form serta penanganan otomatis untuk menjaga konsistensi jumlah kolom database (*Anti-Error Column Length*).
- **🔍 Multi-Filter Pencarian Data (Halaman Utama):** Penyaringan data rekap yang dinamis dan responsif menggunakan 4 parameter sekaligus:
  - **Filter Developer** (Pilihan ganda berdasarkan data master)
  - **Filter Aplikasi** (Pilihan ganda berdasarkan data master)
  - **Filter Jenis** (Pilihan ganda berdasarkan jenis pengerjaan)
  - **Cari Keterangan** (Pencarian berbasis teks/kata kunci, *case-insensitive*)
- **🧹 Auto-Cleansing & Exact Matching:** Sistem otomatis membersihkan spasi hantu (*trimming*) dan menyamakan kapitalisasi teks data Excel sebelum diproses, sehingga hasil filter dijamin 100% akurat.
- **📥 Export Data Hasil Filter:** Fitur unduh data yang telah difilter secara langsung ke dalam format **.CSV** atau **.XLSX (Excel)** untuk kebutuhan reporting eksternal.
- **🔒 Safe Save Permission Handling:** Dilengkapi pengaman jika file database sedang terkunci atau dibuka oleh program lain (Excel/WPS) agar aplikasi tidak *crash*.

---

## 🗂️ Master Data Validasi

Aplikasi ini mengunci dropdown form dan filter menggunakan data master yang terstandarisasi untuk mencegah duplikasi data:
- **Developer:** `RANDY`, `WILIAM`, `ILMAN`, `JONATHAN`, `CHASTRO`, `AGIS`, `EDI`
- **Aplikasi:** `PRODUKSI`, `QC`, `KARYAWAN`, `LOTUS EMPLOYEE`, `GUDANG`, `RND`, `INVENTORY`, `HRCM`, `PAYROLL`, `MARKETING`, `OPERATOR`, `NOTIFICATION`
- **Jenis Pengerjaan:** `REVISI MINOR`, `REVISI MAJOR`, `PEMBUATAN SOFTWARE`

---

## 🚀 Teknologi yang Digunakan

- **Python 3.13+**
- **Streamlit** (Framework UI Web)
- **Pandas** (Pengolahan & Analisis Data Dataframe)
- **OpenPyXL** (Engine pembaca & penulisan file Excel)

---

## 🛠️ Cara Menjalankan Aplikasi di Lokal

### 1. Clone Repositori
```bash
git clone [https://github.com/username-lo/nama-repo-lo.git](https://github.com/username-lo/nama-repo-lo.git)
cd nama-repo-lo
