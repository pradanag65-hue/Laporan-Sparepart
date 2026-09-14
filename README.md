# Rekap Barang Bekas — Trans Jogja

Arsitektur:

```
Google Sheet (database)
      │  dibaca/ditulis oleh
Google Apps Script (Code.gs, deploy sebagai Web App = API)
      │  dipanggil lewat fetch() dari
Frontend React/Vite (repo ini) ── push ke GitHub ── deploy ke Vercel
```

Foto disimpan sebagai file di folder Google Drive "RekapBarangBekas_Foto";
URL publiknya disimpan di kolom sheet, bukan file base64 — supaya sheet tetap ringan.

## 1. Siapkan database (Google Sheet + Apps Script)

1. Buat Google Sheet baru, mis. "Rekap Barang Bekas - DB".
2. Buka **Extensions > Apps Script**.
3. Hapus isi default, tempel seluruh isi file `Code.gs` dari paket ini.
4. **Deploy > New deployment** → pilih tipe **Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Klik **Deploy**, izinkan (authorize) permintaan akses yang muncul.
6. Salin URL yang diakhiri `/exec` — itu URL API kamu.
7. Setiap kali `Code.gs` diubah, buat deployment baru (Manage deployments →
   edit → New version) supaya perubahan aktif di URL yang sama.

Sheet "Entries" beserta header-nya dibuat otomatis saat pertama kali dipanggil.

## 2. Jalankan/lihat lokal (opsional)

```bash
npm install
npm run dev
```

Saat pertama kali dibuka, aplikasi akan minta URL Apps Script — tempel URL
`/exec` dari langkah 1. URL disimpan di localStorage browser, jadi tidak perlu
env variable atau rebuild kalau nanti ganti sheet/deployment.

## 3. Push ke GitHub

```bash
git init
git add .
git commit -m "Rekap barang bekas: Sheet + Apps Script + Vite"
git branch -M main
git remote add origin https://github.com/<username>/<nama-repo>.git
git push -u origin main
```

## 4. Deploy ke Vercel

- Import repo GitHub di https://vercel.com/new, framework otomatis
  terdeteksi sebagai **Vite** — tidak perlu setting tambahan.
- Setelah deploy selesai, buka domain Vercel-nya, tempel URL Apps Script
  di layar "Hubungkan ke Google Sheet" saat pertama kali dibuka.

Tidak perlu menyimpan URL Apps Script sebagai environment variable di
Vercel — cukup diisi sekali lewat UI aplikasi (ikon gerigi di pojok kanan
atas juga bisa dipakai untuk mengubahnya kapan saja, misalnya kalau pindah
ke sheet/deployment baru).
