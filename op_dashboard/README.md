# OpenProject Import/Export Dashboard

Dashboard web (FastAPI) yang jadi versi web dari 2 macro VBA:
- `CreateProjectAndImportSubitemsSimple` -> halaman **Import**
- `ExportProjectToExcel` -> halaman **Export**

Logikanya diport 1:1 (sub-project dari kolom `Project`, hierarki dari
kolom `ID`+`Parent` atau indentasi, milestone cuma pakai field `date`,
dst), tapi parsing JSON-nya pakai `response.json()` bawaan Python --
jadi nggak ada lagi masalah "ID ketuker" kayak yang kejadian di VBA
(regex nebak-nebak pola teks).

## Install & Jalankan

```bash
cd op_dashboard
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt

uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

Buka browser ke **http://localhost:8000**

## Konfigurasi

Buka halaman **Settings** (`/settings`), isi:
- **Base URL**: alamat OpenProject-mu, contoh `http://10.10.1.102:4002`
- **API Key**: dari OpenProject -> My Account -> Access tokens -> API

Tersimpan di `config.json` di folder project (bukan di kode), atau bisa
juga di-set lewat environment variable sebelum run:

```bash
export OPENPROJECT_BASE_URL="http://10.10.1.102:4002"
export OPENPROJECT_API_KEY="xxxxxxxx"
```

## Fitur Import (`/import`)

1. Ketik nama project baru.
2. Upload file `.xlsx` / `.xls` / `.csv`, kolom yang dikenali (header
   case-insensitive):
   - `Subject` (wajib)
   - `Type`, `Status`, `Assignee`, `Priority` (opsional)
   - `Start Date`, `Due Date` (opsional)
   - `Project` (opsional) -> tiap nilai unik jadi **sub-project asli** di
     OpenProject, nested di bawah project utama; task di baris itu masuk
     ke sub-project tersebut.
   - `ID` + `Parent` (opsional) -> hierarki dari situ (format Parent:
     `Task #99: Nama`). Kalau tidak ada, hierarki dibaca dari **indentasi
     4-spasi** di depan `Subject`.
3. Submit -> muncul ringkasan: berapa task berhasil/gagal, mapping
   sub-project, rincian jumlah task per project, dan detail error kalau
   ada yang gagal (pesan asli dari OpenProject, bukan generik).

## Fitur Export (`/export`)

1. Ketik nama project **persis sama** (case-insensitive, tapi bukan
   "mengandung kata itu" -- harus sama persis biar nggak ketuker sama
   project lain yang namanya mirip). Kalau ada 2+ project dengan nama
   sama persis, sistem kasih tau ID-nya masing-masing dan berhenti
   (tidak menebak).
2. Submit -> sistem otomatis cari semua sub-project di bawahnya
   (rekursif, berapa level pun), tarik semua task-nya, susun hierarki,
   dan sediakan link download file Excel dengan format kolom yang sama
   kayak yang dipakai Import (jadi bisa langsung diimport ulang kalau
   perlu).

## Struktur kode

```
app/
  main.py               - route FastAPI (dashboard, import, export, settings)
  importer.py           - logic import (port CreateProjectAndImportSubitemsSimple)
  exporter.py           - logic export (port ExportProjectToExcel)
  openproject_client.py - HTTP client tipis + helper ekstraksi error/ID
  config.py             - baca/simpan Base URL + API Key
  templates/            - halaman HTML (Jinja2 + Bootstrap CDN)
```

## Catatan

- Ini jalan sebagai proses lokal/server sederhana (belum ada login/
  auth untuk dashboard-nya sendiri) -- kalau mau dipasang di server yang
  bisa diakses banyak orang, tambahkan autentikasi dulu sebelum dipasang
  di jaringan yang lebih luas.
- File hasil export disimpan sementara di folder temp OS, nama file
  dikasih suffix acak per-request supaya nggak tabrakan antar pengguna.
