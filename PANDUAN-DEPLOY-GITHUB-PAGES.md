# Panduan Deploy Add-in Word ke GitHub Pages

Panduan **step-by-step dari nol** untuk Windows, tanpa Python sebelumnya.
Hasil akhir: add-in Anda online di `https://<username>.github.io/<repo-name>/` dan bisa dipakai di Word Online dari device mana saja.

---

## Prasyarat

- [x] Akun GitHub (sudah Anda punya)
- [x] Windows + akses PowerShell
- [ ] Python 3.x (akan diinstall di Step 1 — **HANYA jika ingin pakai script `serve.py` untuk testing lokal**). Untuk GitHub Pages, Python **tidak diperlukan**.
- [ ] Git for Windows (akan diinstall di Step 2)

---

## Step 1 — (Opsional) Install Python

> ⚠️ **Hanya perlu jika Anda mau coba add-in di localhost dulu sebelum upload.**
> Kalau langsung upload ke GitHub Pages, **SKIP step ini**.

1. Buka https://www.python.org/downloads/windows/
2. Download **Python 3.12.x** (atau versi stabil terbaru, 64-bit).
3. Jalankan installer. **PENTING**: centang kedua kotak:
   - ✅ **Add python.exe to PATH**
   - ✅ **Install py launcher for all users**
4. Klik **Install Now**.
5. Tutup PowerShell yang terbuka, buka baru.
6. Verifikasi:
   ```powershell
   python --version
   ```
   Harus muncul: `Python 3.12.x`

---

## Step 2 — Install Git for Windows

1. Buka https://git-scm.com/download/win
2. Download installer (otomatis terdeteksi 64-bit).
3. Jalankan installer — **biarkan semua default**, klik Next sampai Install.
4. Setelah selesai, tutup & buka PowerShell baru.
5. Verifikasi:
   ```powershell
   git --version
   ```
   Harus muncul: `git version 2.x.x.windows.x`

---

## Step 3 — Login ke GitHub & Buat Repo Baru

1. Buka https://github.com dan login.
2. Klik tombol **+** di kanan atas → **New repository**.
3. Isi:
   - **Repository name**: `word-addin-resize-gambar` (atau nama lain, **jangan pakai spasi**)
   - **Description**: `Word add-in untuk resize gambar otomatis`
   - **Visibility**: 🟢 **Public** (WAJIB public — GitHub Pages gratis hanya untuk public repo)
   - ☐ **Add a README file** — **BIARKAN KOSONG** (jangan centang, biar repo benar-benar kosong)
   - ☐ .gitignore — biarkan "None"
   - ☐ License — biarkan "None"
4. Klik **Create repository**.
5. Catat URL repo Anda. Contoh:
   ```
   https://github.com/AP26021214/word-addin-resize-gambar
   ```
   (ganti `AP26021214` dengan username GitHub Anda yang sebenarnya)

---

## Step 4 — Download & Siapkan Folder Add-in

### Opsi A — Kalau Anda sudah extract zip `word-addin-resize-gambar.zip` di Downloads

1. Buka folder `Downloads\word-addin-resize-gambar\` di File Explorer.
2. Hapus file berikut (tidak diperlukan untuk GitHub Pages):
   - `serve.py` (hanya untuk localhost)
   - `cert.pem` & `key.pem` (jika ada — hasil `serve.py`)
   - `manifest.xml.bak` (jika ada)
3. Lanjut ke **Step 5**.

### Opsi B — Belum extract

1. Extract `word-addin-resize-gambar.zip` ke `Downloads\`.
2. Lakukan Opsi A di atas.

---

## Step 5 — Ganti URL di manifest.xml

Anda perlu mengganti **5 kemunculan** `https://YOUR-DOMAIN` dengan URL GitHub Pages Anda.

URL GitHub Pages formatnya: `https://<username>.github.io/<repo-name>`

Contoh untuk username `AP26021214` dan repo `word-addin-resize-gambar`:
```
https://AP26021214.github.io/word-addin-resize-gambar
```

### Cara otomatis (pakai PowerShell) — RECOMMENDED

1. Buka PowerShell di folder add-in:
   - Di File Explorer, masuk ke folder `Downloads\word-addin-resize-gambar\`
   - Klik kanan di area kosong + tahan Shift → pilih **"Open PowerShell window here"**
     (atau buka PowerShell lalu `cd "$env:USERPROFILE\Downloads\word-addin-resize-gambar"`)

2. Salin & jalankan script `update-manifest.ps1` (saya sediakan terpisah di zip):
   ```powershell
   # Ganti URL berikut dengan URL GitHub Pages Anda:
   $url = "https://AP26021214.github.io/word-addin-resize-gambar"

   (Get-Content manifest.xml -Raw) -replace 'https://YOUR-DOMAIN', $url |
     Set-Content manifest.xml -NoNewline

   Write-Host "Manifest diupdate ke: $url" -ForegroundColor Green
   Select-String -Path manifest.xml -Pattern $url |
     ForEach-Object { Write-Host "  baris $($_.LineNumber): $($_.Line.Trim())" }
   ```
3. **Ganti `AP26021214` dengan username GitHub Anda yang sebenarnya!**
4. Output harus menampilkan 5 baris yang sudah terupdate.

### Cara manual (jika ragu pakai script)

1. Buka `manifest.xml` di Notepad.
2. Tekan Ctrl+H (Find & Replace).
3. Find: `https://YOUR-DOMAIN`
4. Replace: `https://<username>.github.io/<repo-name>` (sesuaikan!)
5. Klik **Replace All**. Harus ada 5 replacement.
6. Save (Ctrl+S).

---

## Step 6 — Upload Folder Add-in ke GitHub

### Cara termudah: drag-and-drop via web GitHub

1. Buka repo Anda di browser: `https://github.com/<username>/<repo-name>`
2. Klik **"uploading an existing file"** (link biru di layar kosong).
3. Di File Explorer, buka folder `Downloads\word-addin-resize-gambar\`.
4. Pilih **SEMUA file & folder** (`manifest.xml`, `taskpane.html`, `taskpane.css`, `taskpane.js`, `README.md`, dan folder `assets/`).
5. Drag-and-drop semuanya ke area browser GitHub.
6. Tunggu sampai semua file ter-upload (terutama `assets/` yang isinya 4 PNG).
7. Di bagian **"Commit changes"**:
   - Title: `Initial commit — add-in resize gambar v1.1`
   - Description: biarkan kosong
   - Pilih: ⦿ **Commit directly to the `main` branch**
8. Klik tombol hijau **Commit changes**.

### Cara alternatif: pakai Git CLI

```powershell
cd "$env:USERPROFILE\Downloads\word-addin-resize-gambar"
git init
git remote add origin https://github.com/<username>/word-addin-resize-gambar.git
git add .
git commit -m "Initial commit — add-in resize gambar v1.1"
git branch -M main
git push -u origin main
```
(Saat `git push`, browser akan muncul untuk login GitHub → authorize.)

---

## Step 7 — Aktifkan GitHub Pages

1. Di repo Anda (browser), klik tab **Settings** (kanan atas, bukan Settings akun).
2. Menu kiri → **Pages** (di section "Code and automation").
3. Di bawah **"Build and deployment"**:
   - **Source**: pilih **Deploy from a branch** (default).
   - **Branch**: pilih `main` dan folder `/ (root)`.
   - Klik **Save**.
4. Tunggu 1–3 menit. Refresh halaman. Akan muncul kotak hijau di atas dengan teks:
   ```
   Your site is live at https://<username>.github.io/<repo-name>/
   ```
5. **Catat URL ini** — ini adalah URL hosting add-in Anda.

---

## Step 8 — Verifikasi Hosting Berhasil

1. Buka URL berikut di browser (ganti username & repo):
   ```
   https://<username>.github.io/<repo-name>/taskpane.html
   ```
2. Seharusnya muncul halaman add-in (header biru dengan judul "Resize Gambar").
   (UI-nya akan muncul, tapi tombol tidak berfungsi karena ini bukan di dalam Word — itu normal.)
3. Cek juga ikon:
   ```
   https://<username>.github.io/<repo-name>/assets/icon-32.png
   ```
   Harus muncul ikon biru.

Jika muncul **404**, kemungkinan:
- GitHub Pages belum selesai di-deploy (tunggu 2-3 menit lagi).
- URL salah — cek case-sensitive (huruf besar/kecil berpengaruh).

---

## Step 9 — Sideload Add-in ke Word Online

1. Buka https://www.office.com/launch/word → login dengan akun Microsoft 365 Anda.
2. Buka atau buat dokumen Word baru.
3. Tab **Insert** (Sisipkan) → tombol **Add-ins** (di grup "Add-ins", ikon puzzle).
4. Di panel "Add-ins" yang muncul di kanan, scroll ke bawah → klik **"Upload My Add-in"** (atau "Upload Add-in" / "Unggah Add-in Saya").
5. Pilih file `manifest.xml` dari folder `Downloads\word-addin-resize-gambar\manifest.xml`.
6. Klik **Open / Upload**.
7. Akan muncul notifikasi "Add-in added successfully".
8. Add-in sekarang muncul di:
   - Tab **Home** → group **"Gambar Tools"** → tombol **"Resize Gambar"**.
   - Jika tidak muncul, klik tab **Insert** → **My Add-ins** → pilih "Resize Gambar" → klik tombol Add.

---

## Step 10 — Pakai Add-in

1. Klik tombol **Resize Gambar** di ribbon → panel terbuka di kanan.
2. Atur **Lebar Maksimal (cm)** — default 15.
3. Pilih **scope** yang diproses:
   - ✅ Body — Inline (default on)
   - ✅ Body — Floating (default on)
   - ⬜ Header & Footer (centang jika perlu)
   - ⬜ Textbox / Shape (centang jika perlu)
4. Klik **Preview (hitung saja)** dulu untuk simulasi tanpa ubah dokumen → cek laporan di kotak status.
5. Jika sudah yakin, klik **Resize Sekarang** → dokumen terupdate.

---

## Verifikasi & Troubleshooting

| Gejala | Solusi |
|---|---|
| Word Online bilang "Could not load add-in" saat upload manifest | URL di manifest.xml masih salah. Cek Step 5. Jangan lupa: **tidak ada slash di akhir** URL. |
| Tombol add-in tidak muncul di ribbon Word Online | Refresh F5. Kadang perlu 30–60 detik. Atau klik Insert → My Add-ins → cari "Resize Gambar". |
| Add-in terbuka tapi blank putih | Cek apakah `taskpane.html` bisa diakses langsung di browser (Step 8). Kalau 404 → GitHub Pages belum siap atau URL salah. |
| "0 gambar ditemukan" padahal ada gambar | Centang scope **Body — Floating** juga. Atau gambar ada di header/footer → centang scope tersebut. |
| Error "API not supported" | Scope Textbox butuh Word API 1.4+. Word Online selalu support. Coba uncheck Textbox dulu. |
| Add-in hilang setelah refresh Word Online | Normal di mode "Upload My Add-in". Upload ulang manifest via Insert → Add-ins → Upload My Add-in. |
| Ingin update kode add-in | Edit file di komputer → upload ulang ke GitHub (drag-and-drop, pilih file lama → Replace). Tunggu 1-2 menit → refresh Word Online. Tidak perlu upload manifest ulang kalau URL tidak berubah. |

---

## Update Add-in di Masa Depan

Karena URL hosting tetap sama (`https://<username>.github.io/<repo-name>/`), Anda **tidak perlu upload manifest.xml ulang** ke Word. Cukup:

1. Edit file (`taskpane.js`, `taskpane.html`, dll.) di komputer.
2. Upload ke GitHub (drag-and-drop di repo → "Add file" → "Upload files" → replace existing).
3. Tunggu 1-2 menit (GitHub Pages rebuild).
4. Refresh Word Online (F5) → add-in versi baru otomatis dimuat.

## Backup & Akses dari Device Lain

- Repo GitHub = backup permanen.
- Untuk pakai add-in di device lain: cukup buka Word Online di device itu → Insert → Add-ins → Upload My Add-in → upload `manifest.xml` yang sama. URL hosting-nya tetap, jadi add-in jalan di mana saja.
- Atau bagikan file `manifest.xml` ke kolega → mereka upload sendiri ke Word Online mereka.

---

## Bantuan

Kalau ada langkah yang macet, paste error message atau screenshot deskripsinya — saya bantu debug sampai jalan.
