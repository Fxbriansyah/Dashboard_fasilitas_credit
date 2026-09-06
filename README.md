# 📊 Dashboard Pencairan - All Segment Bisnis

Dashboard interaktif berbasis **Streamlit** untuk memantau data pencairan pinjaman
(Plafond, Outstanding, NPL, status loan, dll) di berbagai segmen bisnis
(Retail & Micro Loan, Supply Chain Financing, Consumer, SME Commercial).

## ✨ Fitur

- Upload file Excel (`.xlsx` / `.xls`) sebagai sumber data — tidak ada data yang
  disimpan permanen di server, semua diproses per sesi.
- Alur bertahap: Upload → Validasi & Proses → Ringkasan → Dashboard.
- Filter interaktif: **Tahun Pencairan**, **Bulan Pencairan**, **Status Loan**.
- Kartu ringkasan per segmen bisnis: Total Account, Total Plafond, Total
  Outstanding, Active Account, dan NPL %.
- Chart: trend pencairan, perbandingan outstanding antar segmen, komposisi
  Group Produk (donut chart), dan perbandingan Plafond vs Outstanding.
- Tabel detail data hasil filter + download CSV (kolom identitas sensitif
  seperti nomor rekening/nasabah sengaja tidak ditampilkan).

## 🗂️ Struktur File

```
.
├── app.py             # Kode utama aplikasi Streamlit
├── style.css          # Semua styling custom (terpisah dari app.py)
├── requirements.txt   # Daftar library Python yang dibutuhkan
└── README.md          # Dokumentasi ini
```

> ⚠️ `app.py` membaca `style.css` dari folder yang sama menggunakan path
> relatif. Pastikan kedua file ini selalu berada dalam satu folder/repo.

## 📋 Format Data yang Dibutuhkan

File Excel yang diupload wajib memiliki kolom-kolom berikut (nama kolom harus
persis sama, tidak case-sensitive untuk isinya tapi nama kolom harus sama):

```
PRODUCT_OWNER, GROUP_PRODUCT, PLAN_CODE, JENIS_FASILITAS,
PLAFOND, OUTSTANDING, STATUS_LOAN,
TAHUN PENCAIRAN, BULAN_PENCAIRAN, TANGGAL_PENCAIRAN
```

Catatan format nilai yang didukung:
- **BULAN_PENCAIRAN**: boleh berupa angka (1-12) atau nama bulan dalam teks,
  Bahasa Indonesia maupun Inggris, lengkap atau singkatan (mis. "Januari",
  "Jan", "January", "March", "Mei", dst).
- **TANGGAL_PENCAIRAN**: boleh format tanggal standar atau format angka
  `YYYYMMDD`.
- **STATUS_LOAN**: kode status pinjaman (mis. "1" = Active, "2" = Lunas).
  Nama tampilan kode ini diatur di kamus `STATUS_LOAN_LABEL` dalam `app.py`.

## 🚀 Menjalankan Secara Lokal

1. Pastikan Python 3.9+ sudah terpasang.
2. Install semua dependency:
   ```bash
   pip install -r requirements.txt
   ```
3. Jalankan aplikasi:
   ```bash
   streamlit run app.py
   ```
4. Buka browser ke alamat yang muncul di terminal (biasanya
   `http://localhost:8501`).

## ☁️ Deploy ke Streamlit Community Cloud (Gratis)

1. Upload seluruh isi folder ini (`app.py`, `style.css`, `requirements.txt`)
   ke sebuah repository GitHub.
   - Kalau datanya sensitif, gunakan repo **private**.
2. Buka [share.streamlit.io](https://share.streamlit.io) dan login dengan
   akun GitHub.
3. Klik **"New app"**, pilih repository, branch, dan file utama `app.py`.
4. Klik **Deploy**. Setelah proses build selesai, kamu akan mendapat link
   publik berbentuk `https://nama-app-kamu.streamlit.app` yang bisa dibagikan.

### Alternatif platform deploy lain
- **Hugging Face Spaces** — gratis, mendukung Streamlit langsung.
- **Railway** / **Render** — gratis dengan batasan resource, mendukung
  custom domain.
- **Server sendiri (VPS)** — kontrol penuh, direkomendasikan kalau data yang
  diproses sangat sensitif dan tidak boleh melalui pihak ketiga.

## 🔒 Catatan Keamanan

- Aplikasi ini **tidak menyimpan file Excel yang diupload** secara permanen —
  data hanya ada di memori sesi (`st.session_state`) selama browser tab
  dashboard masih terbuka.
- Kolom identitas sensitif (nomor rekening, dll) sengaja **tidak** dimasukkan
  ke `SAFE_DETAIL_COLUMNS` di `app.py`, sehingga tidak muncul di tabel
  dashboard maupun file CSV hasil download.
- Kalau data yang diproses bersifat rahasia/internal, pastikan:
  - Repository GitHub-nya **private**.
  - Akses ke link dashboard dibatasi (misalnya lewat fitur private app di
    Streamlit Cloud, atau deploy di jaringan internal/VPN).
