# =========================================================
# DASHBOARD PENCAIRAN - ALL SEGMENT BISNIS
# =========================================================
# Ringkasan alur file ini (biar gampang dicari kalau mau revisi):
#   1. Konfigurasi halaman & konstanta global
#   2. Helper format angka / uang / persen
#   3. Pipeline baca & bersihkan data Excel
#   4. CSS & komponen UI kecil (stepper, dsb)
#   5. Halaman: Upload -> Proses -> Ringkasan -> Dashboard
#   6. Router utama (main)
#
# Perubahan pada revisi ini:
#   - Fitur/kolom DPD (GROUP_DPD, DPD, NPL) DIHAPUS total.
#   - Chart donat yang tadinya "Komposisi Group DPD" sekarang jadi
#     "Komposisi Group Produk" (baik di kartu segmen maupun chart utama).
#   - Filter di sidebar: Tahun Pencairan, Bulan Pencairan, & Status Loan
#     (Segment/Produk/Branch masih belum dipakai sebagai filter).
#   - Kartu segmen: metrik "Active Account" kini menyesuaikan Status
#     Loan yang dipilih di filter (label & jumlahnya berubah sesuai
#     status yang aktif difilter; kalau filter kosong, tetap pakai
#     definisi lama STATUS_LOAN == "1").
#   - CSS dirapikan (spacing & konsistensi lebih dijaga).
#   - Ditambahkan komentar di hampir setiap blok kode.
# =========================================================

import io
import time
import hashlib
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

# ---------------------------------------------------------
# Konfigurasi dasar halaman Streamlit.
# layout="wide" supaya dashboard bisa memakai lebar penuh browser.
# ---------------------------------------------------------
st.set_page_config(
    page_title="Dashboard Pencairan - All Segment Bisnis",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================
# KONSTANTA
# =========================================================

# Kolom WAJIB ada di file Excel yang diupload user.
# Catatan: GROUP_DPD & DPD sengaja TIDAK dimasukkan lagi ke sini
# karena fitur DPD sudah dihapus dari dashboard ini.
REQUIRED_COLUMNS = [
    "PRODUCT_OWNER", "GROUP_PRODUCT", "PLAN_CODE", "JENIS_FASILITAS",
    "PLAFOND", "OUTSTANDING", "STATUS_LOAN",
    "TAHUN PENCAIRAN", "BULAN_PENCAIRAN", "TANGGAL_PENCAIRAN",
]

# Kolom yang ditampilkan pada tabel "Detail Data" di halaman dashboard.
# Kolom identitas sensitif (mis. nomor rekening/nasabah) sengaja tidak
# dimasukkan ke sini demi keamanan data.
SAFE_DETAIL_COLUMNS = [
    "PRODUCT_OWNER", "GROUP_PRODUCT", "PLAN_CODE", "JENIS_FASILITAS",
    "TAHUN PENCAIRAN", "BULAN_PENCAIRAN", "TANGGAL_PENCAIRAN",
    "PLAFOND", "OUTSTANDING", "STATUS_LOAN", "BRANCH_CODE",
]

# Urutan tampil kartu segmen bisnis di dashboard.
# Kalau ada segmen baru di data yang tidak ada di daftar ini,
# dia akan otomatis ditambahkan di belakang (lihat page_dashboard()).
SEGMENT_ORDER = [
    "Retail & Micro Loan",
    "Supply Chain Financing",
    "Consumer",
    "SME Commercial",
]

# Tema warna + ikon untuk tiap segmen (dipakai di header kartu segmen
# dan warna bar chart perbandingan antar segmen).
# Mau ganti warna? tinggal edit hex code di sini.
SEGMENT_THEME = {
    "Retail & Micro Loan":    {"icon": "👥", "color": "#2563eb", "bg": "#eff6ff", "border": "#bfdbfe"},
    "Supply Chain Financing": {"icon": "🔗", "color": "#059669", "bg": "#ecfdf5", "border": "#a7f3d0"},
    "Consumer":               {"icon": "🧑", "color": "#d97706", "bg": "#fffbeb", "border": "#fde68a"},
    "SME Commercial":         {"icon": "💼", "color": "#7c3aed", "bg": "#f5f3ff", "border": "#ddd6fe"},
}
# Tema default kalau ada segmen di luar 4 di atas (biar tidak error).
DEFAULT_THEME = {"icon": "📁", "color": "#334155", "bg": "#f8fafc", "border": "#e2e8f0"}

# Palet warna untuk chart donat "Komposisi Group Produk".
# Digabung dua palet plotly supaya tetap cukup warna walau jumlah
# GROUP_PRODUCT di data banyak/berbeda-beda tiap file.
PRODUCT_COLOR_SEQUENCE = px.colors.qualitative.Set2 + px.colors.qualitative.Pastel

# Label nama bulan (untuk ditampilkan di filter, bukan disimpan ke data).
BULAN_LABEL = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
    9: "September", 10: "Oktober", 11: "November", 12: "Desember",
}

# Label nama status untuk kode STATUS_LOAN (untuk ditampilkan di filter &
# kartu segmen, bukan disimpan ke data). Kode yang belum ada di kamus ini
# akan tetap ditampilkan apa adanya (kode aslinya) - lihat fmt_status_loan().
# TODO: lengkapi nama untuk kode "7" dan "8".
STATUS_LOAN_LABEL = {
    "1": "Active",
    "2": "Lunas",
    "7": "7",   # TODO: ganti dengan nama yang sesuai
    "8": "8",   # TODO: ganti dengan nama yang sesuai
}

# Kamus untuk mengubah teks nama bulan (dari file Excel) jadi angka 1-12.
# Mendukung nama lengkap & singkatan, huruf besar/kecil bebas.
# Kamus untuk mengubah teks nama bulan (dari file Excel) jadi angka 1-12.
# Mendukung nama lengkap & singkatan, huruf besar/kecil bebas.
BULAN_TEXT_TO_NUM = {
    "januari": 1, "jan": 1, "january": 1,
    "februari": 2, "feb": 2, "february": 2,
    "maret": 3, "mar": 3, "march": 3,
    "april": 4, "apr": 4,
    "mei": 5, "may": 5,
    "juni": 6, "jun": 6, "june": 6,
    "juli": 7, "jul": 7, "july": 7,
    "agustus": 8, "agu": 8, "aug": 8, "august": 8,
    "september": 9, "sep": 9, "sept": 9,
    "oktober": 10, "okt": 10, "oct": 10, "october": 10,
    "november": 11, "nov": 11,
    "desember": 12, "des": 12, "dec": 12, "december": 12,
}
# Daftar tahapan proses upload -> dashboard, dipakai oleh render_stepper().
STEPS = [
    ("1", "Upload Data", "📤"),
    ("2", "Validasi & Cek Struktur", "✅"),
    ("3", "Proses & Optimasi", "⚙️"),
    ("4", "Cache Data", "💾"),
    ("5", "Tampil Dashboard", "📈"),
]


# =========================================================
# HELPER FORMAT
# Fungsi-fungsi kecil untuk merapikan tampilan angka ala format
# Indonesia (titik untuk ribuan, koma untuk desimal).
# =========================================================
def fmt_num(x):
    """Format angka biasa -> '12.345' (titik sebagai pemisah ribuan)."""
    return f"{x:,.0f}".replace(",", ".")


def fmt_money(x):
    """
    Format nilai uang menjadi singkatan yang enak dibaca:
    Triliun (T), Miliar (M), Juta (Jt), atau nilai penuh kalau kecil.
    Contoh: 1250000000 -> 'Rp 1,25 M'
    """
    x = float(x or 0)
    ax = abs(x)
    if ax >= 1_000_000_000_000:
        return f"Rp {x/1_000_000_000_000:,.2f} T".replace(",", "X").replace(".", ",").replace("X", ".")
    if ax >= 1_000_000_000:
        return f"Rp {x/1_000_000_000:,.2f} M".replace(",", "X").replace(".", ",").replace("X", ".")
    if ax >= 1_000_000:
        return f"Rp {x/1_000_000:,.2f} Jt".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"Rp {x:,.0f}".replace(",", ".")


def fmt_pct(x):
    """Format angka -> '12,34%' (gaya Indonesia, koma sebagai desimal)."""
    return f"{x:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_status_loan(code):
    """Ubah kode STATUS_LOAN (mis. '1', '2') jadi nama yang enak dibaca (mis. 'Active', 'Lunas').
    Kalau kodenya belum ada di kamus STATUS_LOAN_LABEL, tampilkan kode aslinya apa adanya."""
    return STATUS_LOAN_LABEL.get(str(code), str(code))

def fmt_bulan(m):
    """Ubah angka bulan (1-12) jadi nama bulan Indonesia untuk tampilan filter."""
    try:
        return BULAN_LABEL.get(int(m), str(m))
    except (TypeError, ValueError):
        return str(m)

def fmt_tahun(y):
    """Ubah nilai tahun (kadang berupa float, mis. 2024.0) jadi string bersih '2024'."""
    try:
        return str(int(y))
    except (TypeError, ValueError):
        return str(y)


# =========================================================
# DATA PIPELINE
# Bagian ini yang bertanggung jawab: baca file Excel -> validasi
# kolom wajib -> bersihkan tipe data -> simpan ke cache Streamlit.
# =========================================================
def normalize_columns(df):
    """Rapikan nama kolom (buang spasi berlebih di awal/akhir nama kolom)."""
    df.columns = [str(c).strip() for c in df.columns]
    return df


def validate_file(df):
    """Cek apakah semua kolom wajib (REQUIRED_COLUMNS) ada di file. Return list kolom yang hilang."""
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    return missing


def prepare_data(raw):
    """
    Bersihkan & konversi tipe data dari file mentah:
    - PLAFOND, OUTSTANDING -> numeric
    - TAHUN PENCAIRAN, BULAN_PENCAIRAN -> numeric
    - TANGGAL_PENCAIRAN -> datetime (dengan fallback format YYYYMMDD)
    - Kolom teks -> string rapi, nilai kosong diisi "(Kosong)"
    """
    df = normalize_columns(raw.copy())

    # Kolom angka nominal.
    for c in ["PLAFOND", "OUTSTANDING"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    # Kolom TAHUN PENCAIRAN cukup numeric biasa.
    if "TAHUN PENCAIRAN" in df.columns:
        df["TAHUN PENCAIRAN"] = pd.to_numeric(df["TAHUN PENCAIRAN"], errors="coerce")

    # Kolom BULAN_PENCAIRAN: sumbernya bisa berupa angka (1-12) ATAU nama
    # bulan dalam teks (mis. "Januari", "Jan", "January"). Coba jadikan
    # angka dulu; untuk baris yang gagal (NaN), coba cocokkan lewat kamus
    # BULAN_TEXT_TO_NUM (huruf besar/kecil & spasi diabaikan).
    if "BULAN_PENCAIRAN" in df.columns:
        col = df["BULAN_PENCAIRAN"]
        as_num = pd.to_numeric(col, errors="coerce")
        mask = as_num.isna()
        if mask.any():
            mapped = (
                col[mask]
                .astype(str).str.strip().str.lower()
                .map(BULAN_TEXT_TO_NUM)
            )
            as_num.loc[mask] = mapped
        df["BULAN_PENCAIRAN"] = as_num

    # Tanggal pencairan: coba parse langsung dulu, kalau gagal (NaT) coba
    # anggap formatnya angka YYYYMMDD (format umum dari sistem lama).
    if "TANGGAL_PENCAIRAN" in df.columns:
        s = df["TANGGAL_PENCAIRAN"]
        parsed = pd.to_datetime(s, errors="coerce")
        mask = parsed.isna()
        if mask.any():
            numeric = pd.to_numeric(s[mask], errors="coerce")
            parsed2 = pd.to_datetime(numeric.astype("Int64").astype(str), format="%Y%m%d", errors="coerce")
            parsed.loc[mask] = parsed2
        df["TANGGAL_PENCAIRAN"] = parsed

    # Kolom teks/kategori -> rapikan jadi string, isi kosong diberi label "(Kosong)".
    # Catatan penting: kalau di file Excel ada baris kosong (NaN) di kolom
    # ini, pandas kadang membaca SELURUH kolom sebagai angka desimal
    # (float) walaupun isinya sebenarnya kode status seperti "1" atau "2".
    # Akibatnya nilai 1 jadi 1.0, dan setelah diubah ke teks jadi "1.0"
    # (bukan "1") - ini bikin perbandingan status (mis. STATUS_LOAN == "1")
    # jadi salah/tidak ketemu. Baris di bawah ini merapikan angka desimal
    # bulat semacam itu ("1.0", "2.0", dst) supaya jadi "1", "2" lagi.
    for c in ["PRODUCT_OWNER", "GROUP_PRODUCT", "PLAN_CODE", "JENIS_FASILITAS",
              "STATUS_LOAN", "BRANCH_CODE"]:
        if c in df.columns:
            df[c] = df[c].fillna("(Kosong)").astype(str).str.strip()
            df[c] = df[c].str.replace(r"^(-?\d+)\.0+$", r"\1", regex=True)

    return df


@st.cache_data(show_spinner=False)
def read_excel_bytes(file_bytes):
    """Baca file Excel dari bytes. Di-cache supaya file yang sama tidak dibaca berulang kali."""
    return pd.read_excel(io.BytesIO(file_bytes))


@st.cache_data(show_spinner=False)
def build_dashboard_cache(file_bytes, _digest):
    """
    Fungsi utama pipeline (langkah 3 & 4 di stepper):
    baca file -> validasi kolom wajib -> bersihkan data -> hasil di-cache
    berdasarkan hash file (_digest) supaya tidak diproses ulang tiap rerun.
    """
    raw = read_excel_bytes(file_bytes)
    missing = validate_file(raw)
    if missing:
        return None, missing
    df = prepare_data(raw)
    return df, []


# =========================================================
# UI HELPERS
# =========================================================
CSS_PATH = Path(__file__).parent / "style.css"


def inject_css():
    """
    Semua styling custom (di luar tema bawaan Streamlit) sekarang
    ditaruh di file terpisah "style.css" (folder yang sama dengan
    app.py ini), supaya CSS tidak lagi bercampur dengan kode Python.
    Kalau mau ganti warna utama aplikasi, cukup edit variabel di
    bagian ":root" pada style.css.
    """
    css_text = CSS_PATH.read_text(encoding="utf-8")
    st.markdown(f"<style>{css_text}</style>", unsafe_allow_html=True)


def render_stepper(current_index):
    """Gambar 5 kotak langkah proses di bagian atas halaman. current_index: 0..4."""
    cols = st.columns(len(STEPS))
    for i, (num, label, icon) in enumerate(STEPS):
        state = "done" if i < current_index else ("active" if i == current_index else "")
        with cols[i]:
            st.markdown(
                f"""
                <div class="step-box {state}">
                    <span class="step-num">{'✓' if i < current_index else num}</span>
                    <span class="step-title">{icon} {label}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )
    st.write("")


def go_to(stage):
    """Pindah ke halaman/tahap lain lalu paksa Streamlit rerun."""
    st.session_state.stage = stage
    st.rerun()


def reset_all():
    """Reset semua state ke awal (dipakai tombol 'Upload Ulang')."""
    for k in ["stage", "data", "file_hash", "file_name", "processed_at"]:
        st.session_state.pop(k, None)
    st.session_state.stage = "upload"
    st.rerun()


# =========================================================
# HALAMAN 1: UPLOAD
# =========================================================
def page_upload():
    render_stepper(0)

    st.markdown('<div class="main-title main-title--page indent-sm">📂 Halaman Upload Data</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-description indent-sm">Unggah file Excel sumber untuk mulai membangun dashboard.</div>', unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown(
            """
            <div class="upload-hint">
            📥 Drag &amp; drop file Excel di sini, atau klik tombol di bawah untuk memilih file.<br>
            Format yang didukung: <b>.xlsx</b> / <b>.xls</b>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.write("")
        uploaded = st.file_uploader(
            "Pilih File Excel",
            type=["xlsx", "xls"],
            help="Data diproses hanya untuk sesi dashboard ini. File sumber tidak dimasukkan ke source code.",
        )

        st.write("")
        st.markdown("**📋 Pastikan file memiliki kolom wajib berikut:**")
        st.code(", ".join(REQUIRED_COLUMNS))

    if uploaded is not None:
        # Hitung hash file supaya bisa mendeteksi apakah user upload
        # ulang file yang SAMA (kalau iya, tidak perlu proses ulang).
        raw_bytes = uploaded.getvalue()
        digest = hashlib.sha256(raw_bytes).hexdigest()

        if digest == st.session_state.get("file_hash"):
            go_to("summary")
            return

        # Simpan sementara file yang baru diupload, lalu pindah ke
        # halaman proses (page_processing) untuk divalidasi & dibersihkan.
        st.session_state["_pending_bytes"] = raw_bytes
        st.session_state["_pending_digest"] = digest
        st.session_state["_pending_name"] = uploaded.name
        go_to("processing")


# =========================================================
# HALAMAN 2+3: VALIDASI & PROSES (langkah 2, 3, 4 di stepper)
# =========================================================
def page_processing():
    render_stepper(1)

    st.markdown('<div class="main-title main-title--page indent-sm">⚙️ Proses Data</div>', unsafe_allow_html=True)
    st.write("")
    status_box = st.empty()
    progress = st.progress(0, text="Memproses data...")

    # Checklist visual saja (animasi loading), progres asli tetap
    # ditentukan oleh build_dashboard_cache() di bawah.
    checklist_steps = [
        ("Membaca file Excel", 15),
        ("Validasi struktur kolom", 35),
        ("Membersihkan data", 55),
        ("Mengubah tipe data", 70),
        ("Membuat agregasi", 85),
        ("Menyimpan ke cache", 100),
    ]

    raw_bytes = st.session_state.get("_pending_bytes")
    digest = st.session_state.get("_pending_digest")
    name = st.session_state.get("_pending_name")

    if raw_bytes is None:
        go_to("upload")
        return

    done_lines = []
    for label, pct in checklist_steps:
        done_lines.append(f"✔️ {label}")
        status_box.markdown(
            "<div class='upload-hint'>" + "<br>".join(done_lines) + "</div>",
            unsafe_allow_html=True,
        )
        progress.progress(pct, text=f"{label}... {pct}%")
        time.sleep(0.25)

    # Proses data sesungguhnya: baca, validasi kolom wajib, bersihkan.
    try:
        df, missing = build_dashboard_cache(raw_bytes, digest)
    except Exception as e:
        df, missing = None, [f"Gagal membaca file: {e}"]

    if missing or df is None:
        st.error("❌ File tidak sesuai template.")
        st.write("Kolom wajib yang belum ditemukan / masalah lain:")
        st.code("\n".join(missing))
        if st.button("⬅️ Kembali ke Upload"):
            for k in ["_pending_bytes", "_pending_digest", "_pending_name"]:
                st.session_state.pop(k, None)
            go_to("upload")
        return

    # Simpan hasil olahan ke session_state supaya bisa dipakai di
    # halaman-halaman berikutnya tanpa membaca ulang file.
    st.session_state.data = df
    st.session_state.file_hash = digest
    st.session_state.file_name = name
    st.session_state.processed_at = datetime.now()
    for k in ["_pending_bytes", "_pending_digest", "_pending_name"]:
        st.session_state.pop(k, None)

    go_to("summary")


# =========================================================
# HALAMAN 4: RINGKASAN DATA (sebelum tampil dashboard penuh)
# =========================================================
def page_summary():
    render_stepper(3)

    df = st.session_state.get("data")
    if df is None:
        go_to("upload")
        return

    st.markdown('<div class="main-title main-title--page indent-sm">📊 Ringkasan Data</div>', unsafe_allow_html=True)
    st.write("")
    st.markdown("<div class='success-banner'>✅ Data berhasil diproses dan siap ditampilkan!</div>", unsafe_allow_html=True)

    total_records = len(df)
    total_kolom = df.shape[1]
    segmen = df["PRODUCT_OWNER"].nunique() if "PRODUCT_OWNER" in df.columns else 0
    if "TAHUN PENCAIRAN" in df.columns and df["TAHUN PENCAIRAN"].notna().any():
        periode = f"{fmt_tahun(df['TAHUN PENCAIRAN'].min())} - {fmt_tahun(df['TAHUN PENCAIRAN'].max())}"
    else:
        periode = "-"

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Records", fmt_num(total_records))
    c2.metric("Total Kolom", fmt_num(total_kolom))
    c3.metric("Segmen Bisnis", fmt_num(segmen))
    c4.metric("Periode Pencairan", periode)

    st.write("")
    b1, b2 = st.columns(2)
    if b1.button("🚀 Tampilkan Dashboard", use_container_width=True, type="primary"):
        go_to("dashboard")
    if b2.button("🔄 Upload Ulang", use_container_width=True):
        reset_all()

    st.caption(f"File: {st.session_state.get('file_name', '-')}")


# =========================================================
# HALAMAN 5: DASHBOARD UTAMA
# =========================================================
def segment_card(df, segment, selected_status=None):
    """
    Render satu kartu ringkasan untuk 1 segmen bisnis:
    - Metrik: Total Account, Total Plafond, Total Outstanding,
      % Outstanding/Plafond, Active Account
    - Chart donat: komposisi GROUP_PRODUCT di segmen tersebut
      (dulunya komposisi Group DPD, sekarang diganti Group Produk).

    Layout metrik diubah dari 3+2 kolom (terlalu sempit sehingga label
    & nominal terpotong "...") menjadi 2 kolom per baris + 1 baris
    penuh untuk rasio, supaya tiap kotak metrik punya ruang lebih
    lebar dan nominalnya tetap terbaca utuh di kartu segmen manapun.

    `selected_status` adalah daftar Status Loan yang sedang dipilih user
    di filter sidebar (bisa kosong/None kalau user tidak memilih apa-apa,
    artinya "Semua Status"). Kartu metrik ke-4 ("Active Account") akan
    menyesuaikan:
    - Kalau user memilih status tertentu di filter -> `df` yang masuk ke
      sini sudah difilter sesuai status tersebut, jadi kartu menghitung
      & menamai metrik sesuai status yang dipilih (mis. "Account (Aktif)").
    - Kalau tidak ada status dipilih (Semua) -> tetap pakai definisi lama,
      yaitu menghitung account dengan STATUS_LOAN == "1" sebagai "Active Account".
    """
    theme = SEGMENT_THEME.get(segment, DEFAULT_THEME)

    s = df[df["PRODUCT_OWNER"] == segment]
    accounts = len(s)
    plafond = s["PLAFOND"].sum()
    outstanding = s["OUTSTANDING"].sum()
    ratio = (outstanding / plafond * 100) if plafond else 0

    if selected_status:
        # Filter Status Loan sedang aktif -> `s` sudah otomatis terbatas
        # pada status-status yang dipilih (karena `df` difilter di
        # page_dashboard sebelum sampai ke sini).
        active = len(s)
        if len(selected_status) == 1:
            active_label = f"Status Loan ({fmt_status_loan(selected_status[0])})"
        else:
            active_label = "Account (Status Terpilih)"
    else:
        # Tidak ada filter Status Loan -> tampilkan total SEMUA status
        # (semua akun di segmen ini, apapun STATUS_LOAN-nya).
        active = len(s)
        active_label = "Total Status Loan"
        
    with st.container(border=True):
        # Header kartu berwarna sesuai tema segmen.
        # Warna disuntikkan lewat CSS variable (--seg-bg/--seg-border/--seg-color)
        # supaya style tetap di style.css, hanya nilainya yang dinamis per segmen.
        st.markdown(
            f"""
            <div class="segment-header"
                 style="--seg-bg:{theme['bg']}; --seg-border:{theme['border']}; --seg-color:{theme['color']};">
                <span class="segment-header__icon">{theme['icon']}</span>
                <span class="segment-header__label">{segment.upper()}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # Baris 1: Total Account & Total Plafond (2 kolom, lebih lebar dari 3 kolom).
        m1, m2 = st.columns(2)
        m1.metric("Total Account", fmt_num(accounts))
        m2.metric("Total Plafond", fmt_money(plafond))

        # Baris 2: Total Outstanding & Active Account.
        m3, m4 = st.columns(2)
        m3.metric("Total Outstanding", fmt_money(outstanding))
        m4.metric(active_label, fmt_num(active))

        # Baris 3: % Outstanding/Plafond dapat lebar penuh karena labelnya
        # paling panjang -> paling rawan terpotong kalau dipaksa berbagi kolom.
        st.metric("% Outstanding / Plafond", fmt_pct(ratio))

        # Chart donat: komposisi GROUP_PRODUCT untuk segmen ini.
        produk = (
            s.groupby("GROUP_PRODUCT", dropna=False)
            .size()
            .rename("Jumlah")
            .reset_index()
            .sort_values("Jumlah", ascending=False)
        )
        if not produk.empty:
            st.markdown(
                f"<div class='segment-chart-title' style='--seg-color:{theme['color']};'>Komposisi Group Produk</div>",
                unsafe_allow_html=True,
            )
            fig = px.pie(
                produk, names="GROUP_PRODUCT", values="Jumlah", hole=0.6,
                color_discrete_sequence=PRODUCT_COLOR_SEQUENCE,
            )
            fig.update_traces(textinfo="none")
            fig.update_layout(
                height=250, margin=dict(l=5, r=5, t=5, b=5),
                legend=dict(font=dict(size=9), orientation="v"),
                showlegend=True,
            )
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def page_dashboard():
    df0 = st.session_state.get("data")
    if df0 is None:
        go_to("upload")
        return

    # --- Header halaman ---
    header_l, header_r = st.columns([3, 1])
    with header_l:
        st.markdown(
            """
            <div class="dashboard-title-card">
                <div class="main-title indent-sm">DASHBOARD PENCAIRAN</div>
                <div class="subtitle indent-sm">ALL SEGMENT BISNIS</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with header_r:
        processed_at = st.session_state.get("processed_at")
        ts = processed_at.strftime('%d %b %Y %H:%M:%S') if processed_at else "-"
        st.markdown(
            f"""
            <div class="data-chip-wrapper">
                <div class="data-chip">
                    <span class="data-chip__icon">🕒</span>
                    <div>
                        Data Terakhir
                        <b>{ts}</b>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # -----------------------------------------------------
    # SIDEBAR: FILTER
    # Hanya 2 filter yang tersedia: Tahun Pencairan & Bulan Pencairan.
    # Kalau nanti mau menambah filter lain lagi, tambahkan multiselect
    # baru di sini lalu tambahkan juga logika filternya di bawah (df = df[...]).
    # -----------------------------------------------------
    with st.sidebar:
        st.header("🔎 Filter Data")

        years = sorted(df0["TAHUN PENCAIRAN"].dropna().unique().tolist())
        months = sorted(df0["BULAN_PENCAIRAN"].dropna().unique().tolist())

        selected_year = st.multiselect(
            "Tahun Pencairan", years, default=[],
            format_func=fmt_tahun, placeholder="Semua",
        )
        selected_month = st.multiselect(
            "Bulan Pencairan", months, default=[],
            format_func=fmt_bulan, placeholder="Semua",
        )

        statuses = sorted(df0["STATUS_LOAN"].dropna().unique().tolist())
        selected_status = st.multiselect(
            "Status Loan", statuses, default=[],
            placeholder="Semua",
        )

        st.write("")
        fcol1, fcol2 = st.columns(2)
        with fcol1:
            if st.button("🔍 Terapkan", use_container_width=True, type="primary", key="btn_terapkan_filter"):
                st.rerun()
        with fcol2:
            if st.button("↺ Reset", use_container_width=True, key="btn_reset_filter"):
                st.rerun()

        st.divider()
        st.markdown("**ℹ️ Informasi Data**")
        st.caption(f"Total Records: {fmt_num(len(df0))}")
        st.caption(f"Total Account: {fmt_num(len(df0))}")
        if df0["TAHUN PENCAIRAN"].notna().any():
            st.caption(f"Periode Pencairan: {fmt_tahun(df0['TAHUN PENCAIRAN'].min())} - {fmt_tahun(df0['TAHUN PENCAIRAN'].max())}")
        st.caption("Kolom identitas sensitif tidak ditampilkan pada tabel dashboard.")

        st.divider()
        if st.button("🔄 Upload Ulang", use_container_width=True):
            reset_all()

    # Terapkan filter tahun & bulan ke data.
    df = df0.copy()
    if selected_year:
        df = df[df["TAHUN PENCAIRAN"].isin(selected_year)]
    if selected_month:
        df = df[df["BULAN_PENCAIRAN"].isin(selected_month)]
    if selected_status:
        df = df[df["STATUS_LOAN"].isin(selected_status)]

    # -----------------------------------------------------
    # KARTU PER SEGMEN
    # Di layar lebar tetap 4 kartu sebaris (maks. sesuai SEGMENT_ORDER).
    # Di layar sempit, CSS breakpoint di inject_css() akan otomatis
    # menyusun ulang kolom ini jadi 2 kartu/baris lalu 1 kartu/baris
    # supaya tiap kartu tidak terlalu sempit.
    # -----------------------------------------------------
    available_segments = [s for s in SEGMENT_ORDER if s in set(df["PRODUCT_OWNER"])]
    available_segments += [s for s in df["PRODUCT_OWNER"].unique() if s not in available_segments]

    cols = st.columns(min(4, max(1, len(available_segments))))
    for i, seg in enumerate(available_segments[:4]):
        with cols[i]:
            segment_card(df, seg, selected_status)

    st.divider()

    # -----------------------------------------------------
    # BARIS CHART: Trend, Perbandingan Segmen, Komposisi Produk
    # -----------------------------------------------------
    c1, c2, c3 = st.columns(3)

    with c1:
        st.subheader("Trend Pencairan (Total Plafond)")
        trend = (
            df.dropna(subset=["TAHUN PENCAIRAN"])
            .groupby("TAHUN PENCAIRAN", as_index=False)["PLAFOND"]
            .sum()
            .sort_values("TAHUN PENCAIRAN")
        )
        if not trend.empty:
            fig = px.line(trend, x="TAHUN PENCAIRAN", y="PLAFOND", markers=True)
            fig.update_traces(line_color="#2563eb", line_width=3, marker=dict(size=6, color="#2563eb"))
            fig.update_layout(height=340, margin=dict(l=10, r=10, t=10, b=10), yaxis_title="Plafond", xaxis_title="")
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        else:
            st.info("Tidak ada data untuk trend.")

    with c2:
        st.subheader("Perbandingan Outstanding per Segment")
        comp = (
            df.groupby("PRODUCT_OWNER", as_index=False)["OUTSTANDING"]
            .sum()
            .sort_values("OUTSTANDING", ascending=False)
        )
        if not comp.empty:
            colors = [SEGMENT_THEME.get(p, DEFAULT_THEME)["color"] for p in comp["PRODUCT_OWNER"]]
            fig = px.bar(comp, x="PRODUCT_OWNER", y="OUTSTANDING", text_auto=".2s")
            fig.update_traces(marker_color=colors)
            fig.update_layout(height=340, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="", yaxis_title="Outstanding")
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        else:
            st.info("Tidak ada data untuk perbandingan.")

    with c3:
        # Chart donat ini sebelumnya "Komposisi Outstanding berdasarkan Group DPD",
        # sekarang diganti jadi komposisi berdasarkan Group Produk.
        st.subheader("Komposisi Outstanding berdasarkan Group Produk")
        produk_out = (
            df.groupby("GROUP_PRODUCT", as_index=False)["OUTSTANDING"]
            .sum()
            .sort_values("OUTSTANDING", ascending=False)
        )
        if not produk_out.empty:
            fig = px.pie(
                produk_out, names="GROUP_PRODUCT", values="OUTSTANDING", hole=0.6,
                color_discrete_sequence=PRODUCT_COLOR_SEQUENCE,
            )
            fig.update_traces(textinfo="none")
            fig.update_layout(height=340, margin=dict(l=5, r=5, t=10, b=5), legend=dict(font=dict(size=9)))
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        else:
            st.info("Tidak ada data untuk komposisi produk.")

    # -----------------------------------------------------
    # CHART: Plafond vs Outstanding per Segment (bar berdampingan)
    # -----------------------------------------------------
    st.subheader("Plafond vs Outstanding per Segment")
    compare = df.groupby("PRODUCT_OWNER", as_index=False)[["PLAFOND", "OUTSTANDING"]].sum()
    long = compare.melt("PRODUCT_OWNER", var_name="Metric", value_name="Value")
    if not long.empty:
        fig = px.bar(
            long, x="PRODUCT_OWNER", y="Value", color="Metric", barmode="group",
            color_discrete_map={"PLAFOND": "#2563eb", "OUTSTANDING": "#93c5fd"},
        )
        fig.update_layout(height=340, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="", yaxis_title="Nilai")
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    # -----------------------------------------------------
    # TABEL DETAIL DATA + tombol download CSV
    # -----------------------------------------------------
    st.divider()
    st.subheader("Detail Data (Hasil Filter)")
    existing_safe = [c for c in SAFE_DETAIL_COLUMNS if c in df.columns]

    display_df = df[existing_safe].copy()
    if "TANGGAL_PENCAIRAN" in display_df.columns:
        display_df["TANGGAL_PENCAIRAN"] = display_df["TANGGAL_PENCAIRAN"].dt.strftime("%d/%m/%Y")

    top_row = st.columns([1, 1, 1, 2])
    top_row[0].metric("Filtered Records", fmt_num(len(df)))
    top_row[1].metric("Total Plafond", fmt_money(df["PLAFOND"].sum()))
    top_row[2].metric("Total Outstanding", fmt_money(df["OUTSTANDING"].sum()))
    with top_row[3]:
        max_rows = st.number_input("Jumlah baris yang ditampilkan", min_value=10, max_value=5000, value=100, step=10)

    st.dataframe(display_df.head(max_rows), use_container_width=True, hide_index=True)

    csv = display_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ Download Hasil Filter (CSV)",
        data=csv,
        file_name="hasil_filter_dashboard.csv",
        mime="text/csv",
    )

    st.caption(
        f"Dashboard diperbarui {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} • "
        f"{fmt_num(len(df))} record setelah filter."
    )


# =========================================================
# MAIN / ROUTER
# Menentukan halaman mana yang ditampilkan berdasarkan
# st.session_state.stage ("upload" -> "processing" -> "summary" -> "dashboard").
# =========================================================
def main():
    inject_css()

    # Inisialisasi state kalau belum ada (pertama kali app dibuka).
    if "stage" not in st.session_state:
        st.session_state.stage = "upload"
    if "data" not in st.session_state:
        st.session_state.data = None
    if "file_hash" not in st.session_state:
        st.session_state.file_hash = None
    if "file_name" not in st.session_state:
        st.session_state.file_name = None

    stage = st.session_state.stage

    if stage == "upload":
        page_upload()
    elif stage == "processing":
        page_processing()
    elif stage == "summary":
        page_summary()
    elif stage == "dashboard":
        page_dashboard()
    else:
        # Fallback kalau stage tidak dikenal -> balik ke upload.
        st.session_state.stage = "upload"
        st.rerun()


if __name__ == "__main__":
    main()