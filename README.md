# Word Add-in: Resize Gambar

Add-in Word **sesungguhnya** (bukan Script Lab snippet) untuk me-resize otomatis gambar di dokumen Word. Lebar maksimal dapat diatur (default 15 cm), tinggi menyesuaikan proporsi asli.

**Versi 1.1** — kini memproses 4 area sekaligus (dapat dipilih lewat checkbox):

- ✅ **Body — Inline** (gambar menyatu dengan teks)
- ✅ **Body — Floating** (gambar mengambang / wrap text: square, in front, behind, tight, through)
- ⬜ **Header & Footer** (semua section, termasuk first-page & even-page)
- ⬜ **Textbox / Shape** (gambar di dalam textbox)

## Struktur folder

```
word-addin-resize-gambar/
├── manifest.xml         # Definisi add-in (didaftarkan ke Word)
├── taskpane.html        # UI panel samping
├── taskpane.css         # Styling panel
├── taskpane.js          # Logika Office.js (resize gambar)
├── serve.py             # HTTPS server lokal (untuk development)
└── assets/
    ├── icon-16.png      # Ikon add-in (4 ukuran)
    ├── icon-32.png
    ├── icon-80.png
    └── icon-128.png
```

## Cara pakai — pilih salah satu opsi

### OPSI A — Localhost (untuk development / coba-coba)

Cepat, gratis, tapi add-in hanya bisa dipakai dari komputer Anda sendiri.

1. **Jalankan HTTPS server lokal:**
   ```bash
   cd word-addin-resize-gambar
   python serve.py     # Windows
   python3 serve.py    # macOS/Linux
   ```
   Script otomatis:
   - Membuat self-signed certificate (`cert.pem`, `key.pem`).
   - Mengganti `https://YOUR-DOMAIN` di `manifest.xml` menjadi `https://localhost:8443`.
   - Menyerve folder di `https://localhost:8443/`.

2. **Accept sertifikat self-signed di browser:**
   - Buka `https://localhost:8443/taskpane.html` di Chrome/Edge.
   - Klik **Advanced** → **Proceed to localhost (unsafe)**.
   - Layar akan tampil putih (itu normal — UI baru muncul saat dimuat oleh Word).

3. **Sideload ke Word Online:**
   - Buka https://www.office.com/launch/word
   - Buka atau buat dokumen baru.
   - Tab **Insert** → **Add-ins** → **Upload My Add-in**.
   - Pilih file `manifest.xml` di folder ini.
   - Add-in muncul di tab **Home** → group **Gambar Tools** → tombol **Resize Gambar**.

4. **Pakai:**
   - Klik tombol **Resize Gambar** di ribbon.
   - Panel terbuka di kanan. Atur lebar maksimal (cm), klik **Resize Sekarang**.
   - Tombol **Preview (hitung saja)** untuk simulasi tanpa mengubah dokumen.

### OPSI B — Static hosting (untuk pemakaian permanen / multi-device)

Host file add-in di layanan HTTPS gratis (GitHub Pages, Netlify, Vercel). Add-in bisa dipakai dari mana saja.

**Langkah umum:**

1. **Ganti placeholder URL di `manifest.xml`.**
   Cari semua `https://YOUR-DOMAIN` dan ganti dengan URL HTTPS publik Anda, contoh:
   ```
   https://myname.github.io/word-addin-resize-gambar
   ```
   Ada **5 kemunculan** di `manifest.xml` (icon-32, icon-80, AppDomain, SourceLocation, dan di dalam `<Resources>`).

2. **Upload semua file** (kecuali `serve.py`, `cert.pem`, `key.pem`, `*.bak`) ke layanan hosting:
   - **GitHub Pages** (paling mudah jika sudah punya akun):
     - Buat repo baru → upload semua file → Settings → Pages → enable dari branch `main`.
     - URL-nya jadi: `https://username.github.io/nama-repo/`
   - **Netlify Drop** (paling cepat):
     - Buka https://app.netlify.com/drop → drag folder add-in → langsung jadi HTTPS URL.
   - **Vercel / Cloudflare Pages**: sama seperti GitHub Pages, via CLI atau GitHub integration.

3. **Sideload manifest ke Word Online:**
   - Sama seperti OPSI A langkah 3, tidak perlu accept cert (URL-nya sudah HTTPS publik valid).

4. **Untuk Word Desktop (Word for Windows / Mac):**
   - Sideload via **Shared Folder**:
     - Windows: taruh `manifest.xml` di network share, daftarkan via registry.
     - Mac: `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`
   - Atau pakai **Centralized Deployment** dari Microsoft 365 Admin Center (butuh admin tenant).

## Cara kerja add-in

1. Saat dibuka, add-in memanggil `Office.onReady()` dan menunggu host = Word.
2. Tombol **Resize Sekarang** memanggil `Word.run()` dan memproses tiap scope yang dipilih:
   - **Body inline**: `context.document.body.inlinePictures` → resize `width`/`height` tiap inline picture.
   - **Body floating**: `context.document.body.shapes` → filter `type === "Image"` → resize `width`/`height` shape.
   - **Header/footer**: `context.document.sections[i].getHeader/getFooter(...)` untuk primary, first-page, even-page → inline pictures + floating shapes di tiap area.
   - **Textbox**: `body.shapes` → filter `type === "TextBox"` → akses `shape.textFrame.textRange.inlinePictures` → resize tiap pic.
3. Logika resize (sama untuk semua scope):
   - Baca `width` & `height` (dalam points).
   - Jika `width > max`, hitung `rasio = max / width`, set `width = max` dan `height = height * rasio`.
   - `context.sync()` untuk menerapkan perubahan.
4. Konversi satuan: `1 cm = 72 / 2.54 = 28,346 pt`. Word.js memakai **points**.

## Batasan & catatan

- Add-in meminta permission `ReadWriteDocument` (perlu untuk mengubah konten).
- Manifest `Id` sudah diisi GUID statis — biarkan apa adanya agar Word mengenali add-in Anda di sesi berikutnya.
- Beberapa API shape (terutama `body.shapes`, `textFrame.textRange.inlinePictures`) butuh **Word API requirement set 1.4+**. Word Online dan Word 365 desktop selalu mendukung ini. Word versi lama mungkin sebagian scope tidak jalan — add-in akan menampilkan error message dan opsi lain tetap bisa dipakai.
- Grup gambar (grouped shapes) tidak otomatis diproses — Anda perlu ungroup dulu sebelum resize.
- Gambar yang sudah ≤ lebar maksimal **tidak diubah** (tidak diperbesar), untuk menjaga kualitas gambar kecil.

## Troubleshooting

| Masalah | Solusi |
|---|---|
| Word Online bilang "Could not load add-in" | URL di manifest salah atau belum HTTPS. Cek apakah `https://YOUR-DOMAIN/...` sudah diganti semua. |
| Browser warning cert (localhost) | Accept dulu (langkah 2 OPSI A). Setiap tab baru harus accept lagi. |
| Tombol add-in tidak muncul di ribbon | Refresh halaman Word Online (F5). Kadang butuh 10–30 detik. |
| "0 gambar ditemukan" padahal ada gambar | Centang scope yang sesuai (Body Floating, Header/Footer, atau Textbox). Jika gambar di group shape, ungroup dulu. |
| Error "The action isn't supported in this version of Word" | Beberapa API shape butuh Word API 1.4+. Word Online & Word 365 selalu OK; coba matikan scope Textbox dan coba lagi. |
| Error `Office is not defined` | Pastikan `office.js` dari CDN dimuat (cek koneksi internet). |
| Add-in hilang setelah refresh | Itu normal di sideload mode. Upload ulang manifest via Insert → Add-ins → Upload My Add-in. |

## Kustomisasi

- **Ganti lebar default**: edit `value="15"` di `taskpane.html`.
- **Tambahkan fitur**: edit `taskpane.js`. Misalnya: proses gambar di dalam komentar, footnote, atau group shape (perlu rekursi).
- **Ganti warna theme**: edit `taskpane.css` (variabel warna di bagian `header` dan `.btn-primary`).

## Lisensi

Bebas dipakai, dimodifikasi, didistribusikan ulang.
