# Ambil data ROH / INT / RT dari HDKS

Skrip Python yang menarik tiga modul dari HDKS PLN Jawa-Bali lalu menyimpannya
ke `raw/*.xlsx`. Analisisnya dikerjakan `shofwans_tools_v4.xlsx`, yang membaca
folder `raw/` lewat Power Query.

| Modul | `sys=` | Isi |
|-------|--------|-----|
| ROH | `LS_Rencana` | Rencana Operasi Harian |
| INT | `LS_Rencana_Revisi` | Rencana revisi (intraday) |
| RT | `LS_Realisasi` | Realisasi |

## Alur data

```
main.py  ->  raw/ROH.xlsx, raw/INT.xlsx, raw/RT.xlsx
                    |
                    |  Power Query (query `raw`, gabung 3 file)
                    v
        shofwans_tools_v4.xlsx  tabel `raw`  ->  calculator / monitor
```

Skrip **tidak** menulis ke workbook. Sesudah skrip selesai, buka v4 lalu
**Data -> Refresh All** (`Ctrl+Alt+F5`).

## Pemakaian harian

Dobel-klik `jalankan.bat`.

Kalau session masih hidup, langsung jalan. Kalau sudah mati, dia minta paste URL
halaman HDKS sesudah login, lalu menyimpannya sendiri ke `.env`.

Jendela tertutup sendiri kalau sukses, dan berhenti menunggu kalau ada error.

## Pasang di PC baru

Prasyarat: **Python 3.10+** dan Git.

1. Clone dan siapkan virtualenv:

   ```
   git clone https://github.com/shofwann/akses-roh-int-rt.git
   cd akses-roh-int-rt
   python -m venv .venv
   .venv\Scripts\pip install -r requirements.txt
   ```

2. Buat konfigurasi:

   ```
   copy .env.example .env
   ```

   Tidak perlu diisi apa-apa. `jalankan.bat` akan meminta URL lewat prompt
   `URL>` dan menuliskannya sendiri.

3. **Arahkan ulang Power Query.** Path folder `raw/` di dalam v4 tersimpan
   sebagai path absolut, jadi di PC lain pasti gagal. Lihat bagian berikut.

`.venv/` dan `.env` sengaja tidak ikut ter-commit. `raw/*.xlsx` ikut, karena
skrip hanya menimpa sheet pertamanya - kalau filenya tidak ada, modul dilewati.

## Membuat path Power Query portabel

Sekali dikerjakan, workbook bisa dipindah ke folder atau PC mana pun tanpa
disentuh lagi. Kalau hanya butuh sekali jalan, lompat ke "Cara cepat" di bawah.

### 1. Sel berisi lokasi folder

Buka `shofwans_tools_v4.xlsx`, ke sheet `catatan`, pilih satu sel kosong
(misal `A1`), isi:

```
=LEFT(CELL("filename"),FIND("[",CELL("filename"))-1)&"raw\"
```

Hasilnya path folder tempat workbook berada, ditambah `raw\`.

> `CELL("filename")` mengembalikan teks kosong kalau workbook belum pernah
> disimpan. Simpan dulu (`Ctrl+S`) sebelum lanjut.

### 2. Beri nama selnya

Sel itu masih terpilih -> **Formulas -> Define Name** -> isi Name dengan
`PathRaw` -> OK.

Namanya wajib, karena Power Query hanya bisa melihat tabel dan named range,
bukan alamat sel biasa.

### 3. Buat query pembaca sel

**Data -> Get Data -> From Other Sources -> Blank Query**, lalu
**Home -> Advanced Editor**, ganti seluruh isinya dengan:

```m
let
    Sumber = Excel.CurrentWorkbook(){[Name="PathRaw"]}[Content]{0}[Column1]
in
    Sumber
```

Beri nama query ini `PathRaw` juga, lalu **Close & Load To -> Only Create
Connection**.

### 4. Pakai di kedua query yang memuat path

Ini bagian yang paling sering terlewat: pathnya ada di **dua** query, bukan satu.

Di panel Queries, buka **Advanced Editor** untuk:

- query **`raw`**
- query **`Sample File`** - tersembunyi di grup **Helper Queries**

Di masing-masing, ganti barisnya:

```m
Source = Folder.Files("C:\Users\PLN\OneDrive - PLN\Apps\akses_roh_int_real\raw"),
```

menjadi:

```m
Source = Folder.Files(PathRaw),
```

### 5. Matikan Privacy Level

Menggabungkan nilai dari dalam workbook dengan sumber file akan ditolak
Power Query dengan `Formula.Firewall`. Matikan pemeriksaannya:

**Data -> Get Data -> Query Options -> Current Workbook -> Privacy ->**
centang **Ignore the Privacy Levels**.

### 6. Uji

**Data -> Refresh All**. Tabel `raw` harus terisi lagi dan kolom
`Source.Name` berisi `ROH`, `INT`, `RT`.

Lalu tes portabilitasnya: pindahkan seluruh folder proyek ke lokasi lain,
buka v4, Refresh. Kalau tetap jalan, selesai.

### Cara cepat (tanpa dibuat portabel)

Buka **Advanced Editor** untuk query `raw` **dan** `Sample File`, ganti string
pathnya ke folder `raw` yang baru di PC itu. Harus diulang tiap pindah PC.

## Konfigurasi

Semua di `.env`, keterangan tiap kunci ada di `.env.example`.

Dua yang perlu diperhatikan:

- **`TULIS_KE_WORKBOOK` biarkan `false`.** Sejak v4, workbook diisi Power Query,
  bukan ditulis skrip.
- **`EXCEL_PATH` jangan pernah diarahkan ke `shofwans_tools_v4.xlsx`.** Penulisan
  lewat openpyxl akan menghapus Power Query di dalamnya secara permanen, dan
  tidak bisa dibatalkan.

## Kalau ada masalah

| Gejala | Sebab & solusi |
|--------|----------------|
| `session kedaluwarsa` | Login ulang di browser, paste URL barunya saat diminta |
| `GAGAL raw: ... sedang dibuka di Excel` | Tutup file `raw/*.xlsx`, jalankan lagi |
| `DataSource.Error` saat Refresh | Path Power Query masih menunjuk PC lama - lihat bagian di atas |
| Tabel `raw` kosong sesudah Refresh | Sheet di `raw/*.xlsx` harus bernama `Sheet1`; query membacanya berdasarkan nama itu |
| Bar proses tampil sebagai `?????` | Konsol bukan UTF-8. Tidak merusak apa-apa - `jalankan.bat` sudah set `chcp 65001` |
| Angka di `calculator` tidak berubah | Belum Refresh. `Ctrl+Alt+F5` |
