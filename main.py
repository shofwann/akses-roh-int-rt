"""Ambil data ROH / INT / RT dari HDKS lalu tulis ke sheet di shofwans_tools_v3.xlsx.

Ketiga modul memakai endpoint yang sama persis; yang membedakan cuma parameter
sys=. Yang perlu diperbarui tiap hari hanya HDKS_LOGIN_URL di .env, karena
sSession diambil dari situ.
"""

import os
import re
import sys
from datetime import date, datetime
from io import StringIO
from pathlib import Path
from urllib.parse import parse_qsl, urlencode, urlparse

import openpyxl
import pandas as pd
import requests
import urllib3
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

BASE_URL = os.getenv("HDKS_BASE_URL", "").strip()
TANGGAL = os.getenv("HDKS_TANGGAL", "").strip() or date.today().isoformat()
EXCEL_PATH = BASE_DIR / os.getenv("EXCEL_PATH", "shofwans_tools_v3.xlsx")
RAW_DIR = BASE_DIR / os.getenv("RAW_DIR", "raw")
TULIS_WORKBOOK = os.getenv("TULIS_KE_WORKBOOK", "true").strip().lower() in ("1", "true", "yes")
VERIFY_SSL = os.getenv("HDKS_VERIFY_SSL", "false").strip().lower() in ("1", "true", "yes")

# nama sheet tujuan -> kode modul untuk parameter sys=
DATASET = {
    "ROH": "LS_Rencana",
    "INT": "LS_Rencana_Revisi",
    "RT": "LS_Realisasi",
}

# Tabel HDKS bersarang: read_html ikut mengembalikan tabel pembungkus yang
# seluruh isi tabel dalamnya tergencet jadi satu sel raksasa. Tabel data yang
# asli selnya pendek-pendek, jadi itu yang dipakai penyaring.
MAKS_PANJANG_SEL = 100
MIN_BARIS = 5
MIN_KOLOM = 5

if not VERIFY_SSL:
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def sesi_dari_url(url: str) -> str:
    """Ambil nilai sSession dari sebuah URL hasil paste."""
    return dict(parse_qsl(urlparse(url).query, keep_blank_values=True)).get("sSession", "")


SESSION = os.getenv("HDKS_SESSION", "").strip() or sesi_dari_url(
    os.getenv("HDKS_LOGIN_URL", "").strip()
)


def simpan_login_url(url: str) -> None:
    """Tulis balik URL ke .env supaya run berikutnya cukup tekan Enter."""
    berkas = BASE_DIR / ".env"
    isi = berkas.read_text(encoding="utf-8")
    baris_baru = f"HDKS_LOGIN_URL={url}"
    if re.search(r"(?m)^HDKS_LOGIN_URL=.*$", isi):
        # lambda: cegah "\g" / "" di dalam URL ditafsirkan re.sub
        isi = re.sub(r"(?m)^HDKS_LOGIN_URL=.*$", lambda _m: baris_baru, isi, count=1)
    else:
        isi = isi.rstrip() + "\n" + baris_baru + "\n"
    berkas.write_text(isi, encoding="utf-8")


def minta_url_sekali() -> bool:
    """Minta user paste URL sekali. True kalau session berhasil dibaca."""
    global SESSION
    try:
        url = input("URL> ").strip()
    except EOFError:
        return False
    if not url:
        return False
    sesi = sesi_dari_url(url)
    if not sesi:
        print("  URL itu tidak mengandung sSession.")
        return False
    simpan_login_url(url)
    SESSION = sesi
    print(f"  session terbaca: ...{sesi[-8:]}  (tersimpan di .env)")
    return True


def bangun_url(kode_sys: str) -> str:
    """Susun URL modul: sama untuk semua, beda hanya di sys=."""
    params = {
        "fnp": "1",
        "setdate": TANGGAL,
        "sSession": SESSION,
        "sys": kode_sys,
        "sort": "rev_recdate",
        "page": "1",
        "subsys": "View",
        "id": TANGGAL,
    }
    return f"{BASE_URL}?{urlencode(params)}"


def sesi_kedaluwarsa(html: str) -> bool:
    """Halaman login balasan HDKS menandai dirinya lewat isUSER=0 / session TEMP-."""
    return bool(
        re.search(r"var\s+isUSER\s*=\s*'0'", html)
        or re.search(r"var\s+sSession\s*=\s*'TEMP-", html)
    )


def cek_sesi() -> bool:
    """Cek cepat apakah session masih hidup.

    Cukup baca 8KB pertama halaman modul: penanda halaman login ada di bagian
    atas HTML, jadi tidak perlu mengunduh halaman penuh yang ratusan KB.
    """
    if not SESSION:
        return False
    url = bangun_url(next(iter(DATASET.values())))
    try:
        with requests.get(url, timeout=30, verify=VERIFY_SSL, stream=True) as r:
            awal = r.raw.read(8192, decode_content=True)
    except Exception as e:
        print(f"  tidak bisa menghubungi HDKS: {e}")
        return False
    return not sesi_kedaluwarsa(awal.decode("latin-1", "ignore"))


def sel_terpanjang(t: pd.DataFrame) -> int:
    return max((len(str(v)) for v in t.values.ravel()), default=0)


def pilih_tabel(tabel: list[pd.DataFrame]) -> pd.DataFrame | None:
    """Ambil tabel data terbesar, abaikan tabel pembungkus."""
    kandidat = [
        t
        for t in tabel
        if t.shape[0] >= MIN_BARIS
        and t.shape[1] >= MIN_KOLOM
        and sel_terpanjang(t) <= MAKS_PANJANG_SEL
    ]
    if not kandidat:
        return None
    return max(kandidat, key=lambda t: t.shape[0] * t.shape[1])


def ke_angka(df: pd.DataFrame) -> pd.DataFrame:
    """Ubah sel berisi angka menjadi numerik supaya bisa dihitung di Excel.

    Kolom pertama (LOKASI) dilewati. Sel yang bukan angka - label grup seperti
    'PT. INDONESIA POWER' dan penanda MW/MX di RT - dibiarkan apa adanya.
    """
    hasil = df.copy()
    for kolom in hasil.columns[1:]:
        asli = hasil[kolom]
        angka = pd.to_numeric(asli, errors="coerce")
        hasil[kolom] = angka.where(angka.notna() | asli.isna(), asli)
    return hasil


PENANDA_MW_MX = {"MW", "MX"}


def buang_kolom_mw_mx(df: pd.DataFrame) -> pd.DataFrame:
    """Khusus RT: buang baris MX lalu hapus kolom penandanya.

    Hanya jalan kalau kolom ke-3 memang cuma berisi MW/MX, jadi ROH dan INT
    tidak tersentuh. Baris tanpa penanda - label grup, pemisah, dan baris
    total - tetap dipertahankan karena strukturnya dipakai di sheet lain.
    """
    if df.shape[1] < 3:
        return df
    kol = df.columns[2]
    isi = set(df[kol].dropna().astype(str).str.strip().unique())
    if not isi or not isi <= PENANDA_MW_MX:
        return df
    hasil = df[df[kol].astype(str).str.strip() != "MX"].drop(columns=[kol])
    print(f"    kolom {kol!r} dihapus, {len(df) - len(hasil)} baris MX dibuang")
    return hasil.reset_index(drop=True)


def ambil_data(kode_sys: str) -> pd.DataFrame | None:
    """Unduh satu modul dan kembalikan tabel datanya."""
    html = requests.get(bangun_url(kode_sys), timeout=60, verify=VERIFY_SSL).text
    print(f"    panjang html: {len(html)}")
    if sesi_kedaluwarsa(html):
        raise RuntimeError("session kedaluwarsa - paste ulang HDKS_LOGIN_URL di .env")

    df = pilih_tabel(pd.read_html(StringIO(html)))
    # ROH dan INT punya <thead>, RT tidak - headernya cuma <tr> biasa sehingga
    # kolomnya terbaca 0,1,2,... dan label jam ikut dikonversi jadi angka
    # ("00.30" -> 0.3). Kalau itu yang terjadi, baca ulang dengan header=0.
    if df is not None and all(isinstance(c, int) for c in df.columns):
        print("    header tidak terdeteksi, baca ulang dengan header=0")
        df = pilih_tabel(pd.read_html(StringIO(html), header=0))
    return df


def nama_sheet_pertama(berkas: Path) -> str:
    """Nama sheet pertama di sebuah workbook (raw/*.xlsx dibuat user, biasanya Sheet1)."""
    wb = openpyxl.load_workbook(berkas, read_only=True)
    try:
        return wb.sheetnames[0]
    finally:
        wb.close()


def tulis_raw(sheet: str, df: pd.DataFrame) -> bool:
    """Tulis satu dataset ke raw/<sheet>.xlsx, timpa sheet pertamanya."""
    berkas = RAW_DIR / f"{sheet}.xlsx"
    if not berkas.exists():
        print(f"    lewat raw: {berkas} tidak ada")
        return False
    try:
        target = nama_sheet_pertama(berkas)
        with pd.ExcelWriter(
            berkas, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            df.to_excel(writer, sheet_name=target, index=False)
    except PermissionError:
        print(f"    GAGAL raw: {berkas.name} sedang dibuka di Excel")
        return False
    print(f"    raw/{berkas.name} [{target}] <- {df.shape[0]} baris x {df.shape[1]} kolom")
    return True


def tulis_workbook(hasil: dict[str, pd.DataFrame]) -> bool:
    """Tulis semua dataset ke sheet ROH/INT/RT di workbook utama."""
    if not EXCEL_PATH.exists():
        print(f"Lewat workbook: {EXCEL_PATH.name} tidak ditemukan.")
        return False
    try:
        with pd.ExcelWriter(
            EXCEL_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            for sheet, df in hasil.items():
                df.to_excel(writer, sheet_name=sheet, index=False)
                print(f"    {EXCEL_PATH.name} [{sheet}] <- {df.shape[0]} baris")
    except PermissionError:
        # bukan error fatal: data mentahnya sudah aman di raw/
        print(f"Lewat workbook: {EXCEL_PATH.name} sedang dibuka di Excel.")
        return False
    return True


def main() -> int:
    if "--tanya" in sys.argv:
        print("Mengecek session tersimpan...")
        if SESSION and cek_sesi():
            print(f"  session masih aktif (...{SESSION[-8:]}) - langsung jalan.")
        else:
            print("  session sudah tidak valid." if SESSION else "  belum ada session.")
            print("Paste URL halaman HDKS (setelah login), lalu tekan Enter.")
            for sisa in range(2, -1, -1):
                if minta_url_sekali() and cek_sesi():
                    print("  session baru valid.")
                    break
                if SESSION:
                    print("  session dari URL itu juga tidak valid.")
                if sisa:
                    print(f"  coba lagi ({sisa} kesempatan tersisa).")
            else:
                print("Tidak dapat session yang valid. Batal.")
                return 1
    elif not SESSION:
        print("Session kosong - paste HDKS_LOGIN_URL di .env")
        return 1
    sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"\ntanggal data: {TANGGAL}   diambil: {sekarang}   session: ...{SESSION[-8:]}")

    hasil: dict[str, pd.DataFrame] = {}
    for sheet, kode_sys in DATASET.items():
        print(f"[{sheet}] sys={kode_sys}")
        try:
            df = ambil_data(kode_sys)
        except Exception as e:
            print(f"    gagal: {e}")
            continue
        if df is None or df.empty:
            print("    tidak ada tabel data")
            continue
        print(f"    dipakai tabel shape={df.shape}")
        hasil[sheet] = buang_kolom_mw_mx(ke_angka(df))

    if not hasil:
        print("Tidak ada data yang bisa ditulis.")
        return 1

    print("\nmenulis hasil:")
    n_raw = sum(tulis_raw(sheet, df) for sheet, df in hasil.items())
    if TULIS_WORKBOOK:
        tulis_workbook(hasil)

    if n_raw == 0:
        print("\nTidak ada satu pun file raw yang berhasil ditulis.")
        return 1
    if n_raw < len(hasil):
        print(f"\nSebagian saja: {n_raw} dari {len(hasil)} file raw tertulis.")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
