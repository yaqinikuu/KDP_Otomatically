import hashlib
import io
import os
import re
import shutil
import tempfile
import time
import zipfile
from xml.etree import ElementTree as ET

import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import load_workbook

PNM = "PERMODALAN NASIONAL MADANI"

# Produk yang otomatis tercentang (sesuai script awal), user bisa ubah
DEFAULT_DIHAPUS = [
    "KMK PEN",
    "Kredit Konstruksi dan Pengadaan Barang/Jasa",
    "Kredit Mikro BAV",
    "Kredit OTO",
    "Kredit Super Mikro",
    "Kredit Usaha Rakyat",
    "Kredit Usaha Rakyat 2015",
    "Produk Mekar Plus",
]

MODE_TIMPA = "Timpa sheet yang sudah ada"
MODE_BARU = "Tambah sheet baru"

KOLOM_KDP = [
    "laporan", "id_dc_wilayah_kerja", "id_opmt_transaksi_penjaminan", "Tgl_Surat",
    "nama_dc_produk", "Penerima_jaminan", "Unit_Kerja", "Jumlah_Klaim", "penjaminan",
]
KOLOM_KEPUTUSAN = [
    "laporan", "id_dc_wilayah_kerja", "id_opmt_transaksi_penjaminan", "status_keputusan",
]
NAMA_JENIS = {"kdp": "KDP Klaim", "keputusan": "Keputusan Klaim", "lima": "5 Keputusan All"}


# =====================================================================
# Logika data (sama dengan versi Tkinter)
# =====================================================================
def buat_id(d):
    return (
        d["laporan"].astype(str)
        + "_" + d["id_dc_wilayah_kerja"].astype(str)
        + "_" + d["id_opmt_transaksi_penjaminan"].astype(str)
    ).str.strip()


def baca_sheet(xls, nama):
    """Baca sheet dan deteksi baris header yang berisi 'id' dan 'status'."""
    mentah = xls.parse(nama, dtype=str, header=None)
    for i in range(min(10, len(mentah))):
        baris = mentah.iloc[i].astype(str).str.strip().str.lower()
        if {"id", "status"}.issubset(set(baris)):
            s = mentah.iloc[i + 1:].copy()
            s.columns = baris
            return s
    return None


# ---- pembaca xlsx cepat: langsung dari zip, hanya baris awal ----
_NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
_M = "{%s}" % _NS["m"]
_RID = "{%s}id" % _NS["r"]


def _peta_sheet(z):
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    target = {}
    for rel in rels:
        t = rel.get("Target")
        target[rel.get("Id")] = t.lstrip("/") if t.startswith("/") else "xl/" + t
    return [(s.get("name"), target[s.get(_RID)]) for s in wb.find("m:sheets", _NS)]


def _shared_strings(z, perlu):
    if not perlu or "xl/sharedStrings.xml" not in z.namelist():
        return {}
    hasil, maks, i = {}, max(perlu), 0
    with z.open("xl/sharedStrings.xml") as f:
        for _, el in ET.iterparse(f, events=("end",)):
            if el.tag == _M + "si":
                if i in perlu:
                    hasil[i] = "".join(t.text or "" for t in el.iter(_M + "t"))
                i += 1
                el.clear()
                if i > maks:
                    break
    return hasil


def _baris_awal(z, path_sheet, max_baris):
    mentah, perlu = [], set()
    with z.open(path_sheet) as f:
        for _, el in ET.iterparse(f, events=("end",)):
            if el.tag != _M + "row":
                continue
            sel = []
            for c in el.findall(_M + "c"):
                t = c.get("t")
                v = c.find(_M + "v")
                if t == "s":
                    if v is not None and v.text:
                        idx = int(v.text)
                        perlu.add(idx)
                        sel.append(("s", idx))
                elif t == "inlineStr":
                    sel.append(("t", "".join(x.text or "" for x in c.iter(_M + "t"))))
                elif v is not None and v.text is not None:
                    sel.append(("t", v.text))
            mentah.append(sel)
            el.clear()
            if len(mentah) >= max_baris:
                break
    ss = _shared_strings(z, perlu)
    return [[ss.get(x, "") if k == "s" else x for k, x in sel] for sel in mentah]


def daftar_sheet(path):
    try:
        with zipfile.ZipFile(path) as z:
            return [n for n, _ in _peta_sheet(z)]
    except Exception:
        wb = load_workbook(path, read_only=True)
        try:
            return list(wb.sheetnames)
        finally:
            wb.close()


def _hitung_sheet_valid(path):
    """(total_sheet, 0|1): apakah ada sheet berkolom 'id' & 'status'."""
    try:
        with zipfile.ZipFile(path) as z:
            peta = _peta_sheet(z)
            for _, p in reversed(peta):
                for baris in _baris_awal(z, p, 10):
                    if {"id", "status"} <= {str(c).strip().lower() for c in baris}:
                        return len(peta), 1
            return len(peta), 0
    except Exception:
        wb = load_workbook(path, read_only=True)
        try:
            total = len(wb.worksheets)
            for ws in reversed(wb.worksheets):
                for baris in ws.iter_rows(min_row=1, max_row=10, values_only=True):
                    if {"id", "status"} <= {str(c).strip().lower() for c in baris if c is not None}:
                        return total, 1
            return total, 0
        finally:
            wb.close()


def _header_sheet_pertama(path):
    try:
        with zipfile.ZipFile(path) as z:
            peta = _peta_sheet(z)
            baris = _baris_awal(z, peta[0][1], 1)
            return {str(c) for c in baris[0]} if baris else set()
    except Exception:
        wb = load_workbook(path, read_only=True)
        try:
            ws = wb.worksheets[0]
            baris = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
            return {str(c) for c in baris if c is not None}
        finally:
            wb.close()


def status_nama(jenis, nama_file):
    """'sesuai' kalau nama file memuat nama standar jenisnya, selain itu 'tidak_biasa'."""
    nama = re.sub(r"[\s_\-]+", " ", os.path.splitext(os.path.basename(nama_file))[0].lower())
    return "sesuai" if NAMA_JENIS[jenis].lower() in nama else "tidak_biasa"


def cek_isi(jenis, path, nama_asli):
    """Validasi isi file. Return (ok, pesan)."""
    if os.path.basename(nama_asli).startswith("~$"):
        return False, "Ini file sementara Excel (~$...). Tutup Excel lalu pilih file aslinya."
    if not nama_asli.lower().endswith((".xlsx", ".xlsm")):
        return False, "Gunakan file Excel berformat .xlsx (atau .xlsm)."
    try:
        if jenis in ("kdp", "keputusan"):
            header = _header_sheet_pertama(path)
            wajib = KOLOM_KDP if jenis == "kdp" else KOLOM_KEPUTUSAN
            kurang = [k for k in wajib if k not in header]
            if not kurang:
                return True, "File sesuai."
            lain_jenis = "keputusan" if jenis == "kdp" else "kdp"
            lain = KOLOM_KEPUTUSAN if jenis == "kdp" else KOLOM_KDP
            if all(k in header for k in lain):
                return False, (f"Ini sepertinya file {NAMA_JENIS[lain_jenis]}, "
                               f"bukan {NAMA_JENIS[jenis]}.")
            return False, "Kolom tidak ditemukan di sheet pertama: " + ", ".join(kurang)

        total, valid = _hitung_sheet_valid(path)
        if valid == 0:
            header = _header_sheet_pertama(path)
            if "Jumlah_Klaim" in header:
                return False, "Ini sepertinya file KDP Klaim, bukan 5 Keputusan All."
            if "status_keputusan" in header:
                return False, "Ini sepertinya file Keputusan Klaim, bukan 5 Keputusan All."
            return False, "Tidak ada sheet dengan kolom 'id' dan 'status'."
        return True, f"File sesuai ({total} sheet terdeteksi)."
    except Exception as e:
        return False, f"File tidak bisa dibaca ({type(e).__name__}): {e}"


class FileSalahError(Exception):
    """File input yang dipilih bukan file yang seharusnya."""


def _baca_excel(path, jenis):
    try:
        return pd.read_excel(path)
    except (zipfile.BadZipFile, ValueError) as e:
        raise FileSalahError(f"{NAMA_JENIS[jenis]} bukan file Excel yang valid.") from e


def _wajib_kolom(df, jenis, wajib):
    kurang = [k for k in wajib if k not in df.columns]
    if kurang:
        raise FileSalahError(
            f"{NAMA_JENIS[jenis]} tidak sesuai, kolom tidak ditemukan: {', '.join(kurang)}.")


def tahap_cleaning(path_kdp, path_keputusan, path_5keputusan, sheet_target, lapor=None):
    """Tahap 1: cleaning sampai dapat data 'belum ada putusan'.
    'lapor(fraksi, teks)' (opsional) dipanggil di tiap tahap untuk progress bar.
    Kalau gagal setelah file 5 Keputusan ditulis, isinya dikembalikan dari backup."""
    def _l(fraksi, teks):
        if lapor:
            lapor(fraksi, teks)

    bak = path_5keputusan + ".bak"
    sudah_menulis = False
    try:
        # 1-3. File KDP, buat Id, sort, buang duplikat
        _l(0.02, "Membaca KDP Klaim...")
        df = _baca_excel(path_kdp, "kdp")
        _wajib_kolom(df, "kdp", KOLOM_KDP)
        df["Id"] = buat_id(df)
        df = df.sort_values(by="Tgl_Surat", ascending=False)
        df_clean = df.drop_duplicates(subset=["Id"], keep="first").copy()

        # 4. Keputusan Klaim -> tulis ke sheet tujuan di file 5 Keputusan
        _l(0.25, "Membaca Keputusan Klaim...")
        df2 = _baca_excel(path_keputusan, "keputusan")
        _wajib_kolom(df2, "keputusan", KOLOM_KEPUTUSAN)
        df2["Id"] = buat_id(df2)
        id_df2 = df2[["Id", "status_keputusan"]].rename(columns={"status_keputusan": "status"})

        _l(0.40, "Memeriksa file 5 Keputusan All...")
        try:
            _, valid = _hitung_sheet_valid(path_5keputusan)
        except (zipfile.BadZipFile, KeyError) as e:
            raise FileSalahError("5 Keputusan All bukan file Excel yang valid.") from e
        if valid == 0:
            raise FileSalahError(
                "5 Keputusan All tidak sesuai, tidak ada sheet dengan kolom 'id' dan 'status'.")

        _l(0.45, "Menulis sheet ke 5 Keputusan All...")
        shutil.copy2(path_5keputusan, bak)
        sudah_menulis = True
        with pd.ExcelWriter(path_5keputusan, engine="openpyxl", mode="a",
                            if_sheet_exists="replace") as writer:
            id_df2.to_excel(writer, sheet_name=sheet_target, index=False)

        # 5. Baca semua sheet 5 Keputusan
        _l(0.60, "Membuka file 5 Keputusan All...")
        try:
            xls = pd.ExcelFile(path_5keputusan, engine="calamine")
        except Exception:
            xls = pd.ExcelFile(path_5keputusan, engine="openpyxl")

        frames, dilewati = [], []
        n_sheet = max(1, len(xls.sheet_names))
        for i, nama in enumerate(xls.sheet_names):
            _l(0.65 + 0.25 * i / n_sheet, f"Membaca sheet '{nama}' ({i + 1}/{n_sheet})...")
            s = baca_sheet(xls, nama)
            if s is None:
                dilewati.append(nama)
                continue
            t = s[["id", "status"]].copy()
            t.columns = ["ID", "Status"]
            t["ID"] = t["ID"].astype(str).str.strip().str.lower()
            t = t[t["ID"] != "nan"]
            frames.append(t)

        if not frames:
            raise ValueError("Tidak ada sheet yang punya kolom 'id' dan 'status'.")

        # 6. Lookup
        _l(0.92, "Mencocokkan status keputusan...")
        lookup = pd.concat(frames, ignore_index=True)
        lookup = lookup.drop_duplicates(subset="ID", keep="first")
        lookup["Status"] = (
            lookup["Status"].replace(r"^\s*$", pd.NA, regex=True).fillna("ada, status kosong")
        )

        df_clean["status"] = (
            df_clean["Id"].str.lower()
            .map(lookup.set_index("ID")["Status"])
            .fillna("belum ada putusan")
        )
        df_bersih = df_clean[df_clean["status"] == "belum ada putusan"].copy()
        return df_bersih, dilewati
    except Exception:
        if sudah_menulis and os.path.exists(bak):
            shutil.copy2(bak, path_5keputusan)
        raise


def tahap_filter(df_bersih, produk_dihapus, hapus_pnm, path_simpan):
    """Tahap 2: filter produk, PNM, hitung KDP, simpan."""
    d = df_bersih[~df_bersih["nama_dc_produk"].isin(produk_dihapus)]

    # Kredit Mikro dari BRI selalu dibuang
    d = d[~((d["nama_dc_produk"] == "Kredit Mikro")
            & (d["Penerima_jaminan"] == "BANK RAKYAT INDONESIA"))]

    if hapus_pnm:
        d = d[d["Penerima_jaminan"] != PNM]

    d = d[d["Unit_Kerja"] != "Khusus Jakarta"].copy()

    d["KDP"] = np.where(
        d["Jumlah_Klaim"] == 0,
        d["penjaminan"],
        np.minimum(d["Jumlah_Klaim"], d["penjaminan"]),
    )
    d.to_excel(path_simpan, index=False)
    return len(d)


def validasi_nama_sheet(nama, sudah_ada):
    if not nama:
        return "Nama sheet baru belum diisi."
    if len(nama) > 31:
        return "Nama sheet maksimal 31 karakter."
    if re.search(r"[\\/?*\[\]:]", nama):
        return "Nama sheet tidak boleh mengandung \\ / ? * [ ] :"
    if nama.lower() in {s.lower() for s in sudah_ada}:
        return f"Sheet '{nama}' sudah ada. Pilih 'Timpa sheet yang sudah ada' atau ganti namanya."
    return None


def saran_sheet_baru(sheets):
    angka = [int(s) for s in sheets if s.strip().isdigit()]
    return str(max(angka) + 1) if angka else ""


def format_waktu(detik):
    total_ms = int(round(detik * 1000))
    jam, sisa = divmod(total_ms, 3_600_000)
    menit, sisa = divmod(sisa, 60_000)
    d, ms = divmod(sisa, 1000)
    return f"{jam:02d}:{menit:02d}:{d:02d}.{ms:03d}"


# =====================================================================
# Helper Streamlit
# =====================================================================
def folder_kerja():
    if "folder" not in st.session_state or not os.path.isdir(st.session_state["folder"]):
        st.session_state["folder"] = tempfile.mkdtemp(prefix="kdp_")
    return st.session_state["folder"]


def simpan_upload(jenis, up):
    """Simpan upload ke folder sementara. Return (path, hash). Tidak menulis ulang kalau isi sama."""
    data = up.getvalue()
    h = hashlib.md5(data).hexdigest()
    path = os.path.join(folder_kerja(), jenis, up.name)
    meta = st.session_state.setdefault("meta_file", {})
    if meta.get(jenis) != (path, h) or not os.path.exists(path):
        # bersihkan file lama jenis ini
        shutil.rmtree(os.path.join(folder_kerja(), jenis), ignore_errors=True)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "wb") as f:
            f.write(data)
        meta[jenis] = (path, h)
        # file 5 Keputusan berganti -> hasil lama tidak berlaku
        st.session_state.pop("df_bersih", None)
        st.session_state.pop("hasil_5", None)
        st.session_state.pop("hasil_kdp", None)
    return path, h


@st.cache_data(show_spinner=False)
def _cek_cached(jenis, h, path, nama):
    return cek_isi(jenis, path, nama)


def blok_upload(jenis):
    """Widget upload + validasi + konfirmasi. Return (path, hash) kalau siap dipakai, else None."""
    up = st.file_uploader(NAMA_JENIS[jenis], type=["xlsx", "xlsm"], key=f"up_{jenis}")
    if up is None:
        return None
    path, h = simpan_upload(jenis, up)
    ok, pesan = _cek_cached(jenis, h, path, up.name)
    if not ok:
        st.error("✗ " + pesan)
        return None
    if status_nama(jenis, up.name) == "sesuai":
        st.success("✓ Nama file sesuai. " + pesan)
        yakin = st.checkbox(f"Saya yakin isinya {NAMA_JENIS[jenis]}", value=True, key=f"yakin_{jenis}")
    else:
        st.warning(f"Nama file tidak seperti biasanya ({up.name}). {pesan}")
        yakin = st.checkbox(f"Saya yakin isinya {NAMA_JENIS[jenis]}", value=False, key=f"yakin_{jenis}")
    return (path, h) if yakin else None


# =====================================================================
# UI
# =====================================================================
st.set_page_config(page_title="Otomatisasi KDP Klaim", layout="centered")

CSS_GLASS = """
<style>
/* latar berwarna lembut supaya efek kaca terlihat */
.stApp {
    background:
        radial-gradient(900px 600px at 10% -10%, rgba(120, 170, 255, 0.45), transparent 60%),
        radial-gradient(800px 600px at 100% 10%, rgba(255, 150, 210, 0.35), transparent 60%),
        radial-gradient(900px 700px at 50% 110%, rgba(130, 235, 200, 0.40), transparent 60%),
        #eef2fb;
    background-attachment: fixed;
}
header[data-testid="stHeader"] { background: transparent; }

/* card kaca */
div[data-testid="stVerticalBlockBorderWrapper"] {
    background: linear-gradient(135deg, rgba(255,255,255,0.62), rgba(255,255,255,0.30));
    -webkit-backdrop-filter: blur(28px) saturate(180%);
    backdrop-filter: blur(28px) saturate(180%);
    border: 1px solid rgba(255, 255, 255, 0.70) !important;
    border-radius: 26px !important;
    padding: 0.6rem 0.9rem;
    box-shadow:
        0 10px 40px rgba(60, 80, 160, 0.16),
        0 2px 8px rgba(60, 80, 160, 0.08),
        inset 0 1px 0 rgba(255, 255, 255, 0.95),
        inset 0 -1px 0 rgba(255, 255, 255, 0.35);
    transition: transform .25s ease, box-shadow .25s ease;
}
div[data-testid="stVerticalBlockBorderWrapper"]:hover {
    transform: translateY(-2px);
    box-shadow:
        0 16px 48px rgba(60, 80, 160, 0.22),
        0 4px 12px rgba(60, 80, 160, 0.10),
        inset 0 1px 0 rgba(255, 255, 255, 0.95),
        inset 0 -1px 0 rgba(255, 255, 255, 0.35);
}

/* area upload */
div[data-testid="stFileUploaderDropzone"] {
    background: rgba(255, 255, 255, 0.45);
    border: 1px dashed rgba(90, 110, 190, 0.35);
    border-radius: 18px;
}

/* tombol */
.stButton > button, .stDownloadButton > button {
    border-radius: 999px;
    border: 1px solid rgba(255, 255, 255, 0.75);
    background: linear-gradient(180deg, rgba(255,255,255,0.85), rgba(255,255,255,0.45));
    box-shadow: 0 4px 14px rgba(60, 80, 160, 0.15), inset 0 1px 0 rgba(255,255,255,0.9);
    -webkit-backdrop-filter: blur(10px);
    backdrop-filter: blur(10px);
    transition: transform .15s ease, box-shadow .15s ease;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    transform: translateY(-1px);
    box-shadow: 0 8px 20px rgba(60, 80, 160, 0.22), inset 0 1px 0 rgba(255,255,255,0.9);
}
.stButton > button[kind="primary"] {
    background: linear-gradient(180deg, #5b9cff, #3a78f2);
    color: #fff;
    border: 1px solid rgba(255, 255, 255, 0.45);
    box-shadow: 0 6px 18px rgba(58, 120, 242, 0.40), inset 0 1px 0 rgba(255,255,255,0.45);
}

/* input / pilihan */
div[data-baseweb="select"] > div, div[data-baseweb="input"] > div {
    border-radius: 14px !important;
    background: rgba(255, 255, 255, 0.55) !important;
}
div[data-testid="stAlert"] { border-radius: 16px; }
</style>
"""
st.markdown(CSS_GLASS, unsafe_allow_html=True)

# Kartu progres upload (persen, kecepatan, sisa waktu). Memantau request upload bawaan
# uploader di browser, jadi tidak mengubah cara kerja file_uploader.
JS_UPLOAD = """
<script>
(function () {
  var P = window.parent;
  if (!P || P.__kdpUploadPatched) return;
  P.__kdpUploadPatched = true;
  var D = P.document;

  var box = D.createElement('div');
  box.style.cssText = 'position:fixed;right:20px;bottom:20px;z-index:99999;min-width:270px;' +
    'padding:14px 16px;border-radius:20px;font:14px system-ui,sans-serif;color:#1f2a44;' +
    'background:rgba(255,255,255,.65);backdrop-filter:blur(24px) saturate(180%);' +
    '-webkit-backdrop-filter:blur(24px) saturate(180%);border:1px solid rgba(255,255,255,.8);' +
    'box-shadow:0 10px 40px rgba(60,80,160,.22);display:none';
  box.innerHTML =
    '<div class="t" style="font-weight:600;margin-bottom:8px"></div>' +
    '<div style="height:8px;border-radius:99px;background:rgba(90,110,190,.18);overflow:hidden">' +
    '<div class="b" style="height:100%;width:0;border-radius:99px;transition:width .2s;' +
    'background:linear-gradient(90deg,#5b9cff,#3a78f2)"></div></div>' +
    '<div class="s" style="margin-top:8px;font-size:12px;color:#55607e"></div>';
  D.body.appendChild(box);
  var elT = box.querySelector('.t'), elB = box.querySelector('.b'), elS = box.querySelector('.s');

  var batch = [], speed = 0, lastBytes = 0, lastTime = 0, timer = null, hideTimer = null;

  function mb(x) { return (x / 1048576).toFixed(1); }
  function fmtTime(s) {
    if (!isFinite(s) || s < 0) return '--:--';
    s = Math.round(s);
    var m = Math.floor(s / 60), d = s % 60;
    return (m < 10 ? '0' : '') + m + ':' + (d < 10 ? '0' : '') + d;
  }
  function sums() {
    var l = 0, t = 0, done = true;
    batch.forEach(function (r) { l += r.loaded; t += r.total; if (!r.done) done = false; });
    return { l: l, t: t, done: done };
  }
  function render() {
    var s = sums(), pct = s.t > 0 ? Math.min(100, s.l / s.t * 100) : 0;
    elB.style.width = pct.toFixed(1) + '%';
    if (s.done) {
      elT.textContent = 'Upload selesai (' + mb(s.t) + ' MB)';
      elS.textContent = '100%';
    } else {
      elT.textContent = 'Mengunggah... ' + pct.toFixed(0) + '%';
      var eta = speed > 0 ? (s.t - s.l) / speed : Infinity;
      elS.textContent = mb(s.l) + ' / ' + mb(s.t) + ' MB  \\u2022  ' +
        (speed / 1048576).toFixed(2) + ' MB/s  \\u2022  sisa ' + fmtTime(eta);
    }
  }
  function tick() {
    var now = Date.now(), s = sums();
    var inst = (s.l - lastBytes) / ((now - lastTime) / 1000);
    speed = speed ? speed * 0.6 + inst * 0.4 : inst;
    lastBytes = s.l; lastTime = now;
    render();
    if (s.done) {
      clearInterval(timer); timer = null;
      hideTimer = setTimeout(function () { box.style.display = 'none'; batch = []; }, 3000);
    }
  }
  function mulai(rec) {
    if (hideTimer) { clearTimeout(hideTimer); hideTimer = null; }
    if (sums().done) { batch = []; speed = 0; lastBytes = 0; }   // batch baru
    batch.push(rec);
    box.style.display = 'block';
    if (!timer) { lastTime = Date.now(); lastBytes = sums().l; timer = setInterval(tick, 400); }
    render();
  }

  var XHR = P.XMLHttpRequest.prototype;
  var oOpen = XHR.open, oSend = XHR.send;
  XHR.open = function (m, u) { this.__kdpUrl = String(u); return oOpen.apply(this, arguments); };
  XHR.send = function (body) {
    try {
      if (this.__kdpUrl && this.__kdpUrl.indexOf('upload_file') !== -1 && body) {
        var rec = { loaded: 0, total: (body.size || 0), done: false };
        mulai(rec);
        this.upload.addEventListener('progress', function (e) {
          rec.loaded = e.loaded;
          if (e.lengthComputable) rec.total = e.total;
        });
        this.addEventListener('loadend', function () {
          rec.done = true;
          if (rec.total) rec.loaded = rec.total;
        });
      }
    } catch (err) {}
    return oSend.apply(this, arguments);
  };
})();
</script>
"""
components.html(JS_UPLOAD, height=0)

st.title("Otomatisasi KDP Klaim")

# ---------- Langkah 1 ----------
st.header("Langkah 1 - Pilih file input")
siap = {}
for j in ("kdp", "keputusan", "lima"):
    with st.container(border=True):
        siap[j] = blok_upload(j)

# file yang sama tidak boleh dipakai di dua kolom
hashes = [v[1] for v in siap.values() if v]
if len(hashes) != len(set(hashes)):
    st.error("Ada file yang sama dipilih di lebih dari satu kolom.")
    semua_siap = False
else:
    semua_siap = all(siap.values())

sheets_ada = []
sheet = ""
konfirmasi_timpa = True
with st.container(border=True):
    st.subheader("Simpan data Keputusan Klaim ke file 5 Keputusan All")
    if siap["lima"]:
        try:
            sheets_ada = daftar_sheet(siap["lima"][0])
        except Exception as e:
            st.error(f"Gagal membaca daftar sheet: {type(e).__name__}: {e}")
    else:
        st.caption("Upload file 5 Keputusan All dulu agar daftar sheet muncul. "
                   "Contoh ganti tahun: tambah sheet baru '2027'.")

    mode = st.radio("Mode", [MODE_TIMPA, MODE_BARU], horizontal=True, label_visibility="collapsed")
    if mode == MODE_TIMPA:
        angka = [s for s in sheets_ada if s.strip().isdigit()]
        idx = sheets_ada.index(max(angka, key=int)) if angka else (len(sheets_ada) - 1 if sheets_ada else 0)
        sheet = st.selectbox("Sheet yang ditimpa", sheets_ada, index=idx if sheets_ada else None,
                             disabled=not sheets_ada)
        sheet = sheet or ""
        konfirmasi_timpa = st.checkbox(
            f"Saya paham isi sheet '{sheet}' di 5 Keputusan All akan ditimpa "
            "(hasilnya diunduh sebagai file baru, file asli di komputer Anda tidak berubah)",
            value=False, disabled=not sheet)
    else:
        sheet = st.text_input("Nama sheet baru", value=saran_sheet_baru(sheets_ada)).strip()

if st.button("Proses Cleaning", type="primary", disabled=not semua_siap):
    err_sheet = None
    if mode == MODE_TIMPA:
        if not sheet:
            err_sheet = "Pilih sheet yang mau ditimpa."
        elif not konfirmasi_timpa:
            err_sheet = "Centang konfirmasi penimpaan sheet dulu."
    else:
        err_sheet = validasi_nama_sheet(sheet, sheets_ada)

    if err_sheet:
        st.warning(err_sheet)
    else:
        t0 = time.time()
        # kerja pada salinan supaya upload asli tetap utuh kalau user mengulang proses
        kerja = os.path.join(folder_kerja(), "kerja")
        shutil.rmtree(kerja, ignore_errors=True)
        os.makedirs(kerja)
        p_kdp, p_kep, p_lima = (
            shutil.copy2(siap[j][0], kerja) for j in ("kdp", "keputusan", "lima")
        )
        bar = st.progress(0.0, text="Memulai...")

        def lapor(fraksi, teks):
            bar.progress(min(max(fraksi, 0.0), 1.0),
                         text=f"{teks}  •  {format_waktu(time.time() - t0)}")

        try:
            try:
                df_bersih, dilewati = tahap_cleaning(p_kdp, p_kep, p_lima, sheet, lapor=lapor)
            finally:
                bar.empty()
            with open(p_lima, "rb") as f:
                st.session_state["hasil_5"] = (os.path.basename(p_lima), f.read())
            st.session_state["df_bersih"] = df_bersih
            st.session_state["dilewati"] = dilewati
            st.session_state["sheet_dipakai"] = sheet
            st.session_state["durasi_cleaning"] = time.time() - t0
            produk = sorted(df_bersih["nama_dc_produk"].dropna().unique())
            st.session_state["produk_dihapus"] = [p for p in produk if p in DEFAULT_DIHAPUS]
            st.session_state.pop("hasil_kdp", None)
        except FileSalahError as e:
            st.session_state.pop("df_bersih", None)
            st.error(f"Maaf, file tidak bisa diproses.\nFile salah: {e}")
        except Exception as e:
            st.session_state.pop("df_bersih", None)
            st.error(f"Maaf, file tidak bisa diproses.\nError: {type(e).__name__}: {e}")
            st.exception(e)

# ---------- Langkah 2 ----------
if "df_bersih" in st.session_state:
    df_bersih = st.session_state["df_bersih"]
    st.success(
        f"Cleaning selesai dalam {format_waktu(st.session_state['durasi_cleaning'])}: "
        f"{len(df_bersih)} baris 'belum ada putusan'. "
        f"Sheet '{st.session_state['sheet_dipakai']}' sudah diperbarui."
    )
    if st.session_state.get("dilewati"):
        st.info("Sheet dilewati (tanpa kolom id/status): " + ", ".join(st.session_state["dilewati"]))

    nama5, data5 = st.session_state["hasil_5"]
    st.download_button("⬇ Unduh 5 Keputusan All (sudah diperbarui)", data5, file_name=nama5,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.header("Langkah 2 - Pilih produk yang mau dihapus")
    produk = sorted(df_bersih["nama_dc_produk"].dropna().unique())
    with st.container(border=True):
        c1, c2, _ = st.columns([1, 1, 3])
        if c1.button("Centang semua"):
            st.session_state["produk_dihapus"] = list(produk)
            st.rerun()
        if c2.button("Kosongkan"):
            st.session_state["produk_dihapus"] = []
            st.rerun()

        produk_dihapus = st.multiselect("Produk yang dihapus", produk, key="produk_dihapus")
        st.caption(f"{len(produk_dihapus)} dari {len(produk)} produk dicentang")
        hapus_pnm = st.checkbox(f"Hapus Penerima Jaminan: {PNM.title()}", value=True)
        st.caption("Otomatis juga dibuang: Kredit Mikro BRI dan Unit Kerja 'Khusus Jakarta'.")

    if st.button("Proses & Simpan", type="primary"):
        t0 = time.time()
        out = os.path.join(folder_kerja(), "KDP Bersih.xlsx")
        try:
            with st.spinner("Memfilter dan menyimpan hasil..."):
                n = tahap_filter(df_bersih, produk_dihapus, hapus_pnm, out)
            with open(out, "rb") as f:
                st.session_state["hasil_kdp"] = (n, f.read(), time.time() - t0)
        except Exception as e:
            st.session_state.pop("hasil_kdp", None)
            st.error(f"Maaf, hasil tidak bisa disimpan.\nError: {type(e).__name__}: {e}")
            st.exception(e)

    if "hasil_kdp" in st.session_state:
        n, data, dur = st.session_state["hasil_kdp"]
        st.success(f"Selesai dalam {format_waktu(dur)}. {n} baris siap diunduh.")
        st.download_button("⬇ Unduh KDP Bersih.xlsx", data, file_name="KDP Bersih.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
