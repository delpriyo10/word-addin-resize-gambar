# =============================================================
# update-manifest.ps1
# Script PowerShell untuk mengganti placeholder URL di manifest.xml
# dengan URL GitHub Pages Anda.
#
# Cara pakai:
#   1. Buka PowerShell di folder add-in (folder ini).
#      (Klik kanan + Shift di File Explorer -> Open PowerShell here)
#   2. Edit baris $url di bawah dengan URL GitHub Pages Anda.
#   3. Jalankan:  .\update-manifest.ps1
# =============================================================

# === GANTI URL INI DENGAN URL GITHUB PAGES ANDA ===
$url = "https://AP26021214.github.io/word-addin-resize-gambar"
# ===================================================

$manifest = "manifest.xml"

if (-not (Test-Path $manifest)) {
    Write-Host "ERROR: file $manifest tidak ditemukan di folder ini." -ForegroundColor Red
    Write-Host "Jalankan script ini dari dalam folder add-in." -ForegroundColor Yellow
    exit 1
}

# Backup dulu
$backup = $manifest + ".bak"
if (-not (Test-Path $backup)) {
    Copy-Item $manifest $backup
    Write-Host "Backup dibuat: $backup" -ForegroundColor DarkGray
}

# Baca, replace, tulis kembali (preserve encoding & newline style)
$content = Get-Content $manifest -Raw
$count = ([regex]::Matches($content, 'https://YOUR-DOMAIN')).Count

if ($count -eq 0) {
    Write-Host "Tidak ada placeholder 'https://YOUR-DOMAIN' di manifest.xml." -ForegroundColor Yellow
    Write-Host "Mungkin sudah diganti sebelumnya. Cek baris berikut:" -ForegroundColor Yellow
    Select-String -Path $manifest -Pattern 'https://.*github.io' |
      ForEach-Object { Write-Host "  baris $($_.LineNumber): $($_.Line.Trim())" -ForegroundColor Cyan }
    exit 0
}

$newContent = $content -replace 'https://YOUR-DOMAIN', $url
Set-Content -Path $manifest -Value $newContent -NoNewline -Encoding UTF8

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " Manifest diupdate!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "URL: $url" -ForegroundColor Cyan
Write-Host "Jumlah replacement: $count" -ForegroundColor Cyan
Write-Host ""
Write-Host "Baris yang berisi URL Anda sekarang:" -ForegroundColor White
Select-String -Path $manifest -Pattern ([regex]::Escape($url)) |
  ForEach-Object { Write-Host "  baris $($_.LineNumber): $($_.Line.Trim())" -ForegroundColor Cyan }
Write-Host ""
Write-Host "Langkah selanjutnya: upload folder ini ke GitHub repo Anda." -ForegroundColor Yellow
