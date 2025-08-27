# ==========================
# Script Download & Extract Exambrowser Client (32/64-bit)
# ==========================

# 1. Aktifkan TLS 1.2 agar semua versi PowerShell bisa download
try {
    # Untuk PowerShell versi baru
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
} catch {
    # Fallback untuk versi lama (angka 3072 = TLS1.2)
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072
}

# 2. Tentukan arsitektur & URL download
$DownloadUrl64  = "https://unp.kemdikbud.go.id/tes/64BitExambrowserClient_25.0525.rar"
$DownloadUrl32  = "https://unp.kemdikbud.go.id/tes/32BitExambrowserClient_25.0525.rar"

if ([Environment]::Is64BitOperatingSystem) {
    Write-Host "Sistem terdeteksi: 64-bit"
    $DownloadUrl = $DownloadUrl64
} else {
    Write-Host "Sistem terdeteksi: 32-bit"
    $DownloadUrl = $DownloadUrl32
}

# 3. Lokasi target
$ExtractPath  = "$env:USERPROFILE\Documents\exambrowser_client"
$DownloadPath = Join-Path $ExtractPath (Split-Path $DownloadUrl -Leaf)
$SevenZipPath = "C:\Program Files\7-Zip\7z.exe"
$SevenZipInstallerUrl = "https://www.7-zip.org/a/7z2408-x64.exe"   # versi terbaru

# 4. Pastikan folder tujuan ada
if (!(Test-Path $ExtractPath)) {
    # Kalau folder belum ada → buat baru
    New-Item -ItemType Directory -Path $ExtractPath | Out-Null
}
else {
    # Kalau folder sudah ada → cek apakah file download ada
    if (!(Test-Path $DownloadPath)) {
        Write-Host "Update terdeteksi, membersihkan isi folder $ExtractPath ..."
        Get-ChildItem -Path $ExtractPath -Recurse -Force | Remove-Item -Recurse -Force
    }
}


# 5. Cek apakah 7-Zip sudah terinstall
if (!(Test-Path $SevenZipPath)) {
    Write-Host "7-Zip belum ada. Download & install..."
    $InstallerPath = Join-Path $env:TEMP "7z_installer.exe"
    Invoke-WebRequest -Uri $SevenZipInstallerUrl -OutFile $InstallerPath
    Start-Process -FilePath $InstallerPath -ArgumentList "/S" -Wait
}

# 6. Download file RAR jika belum ada
if (!(Test-Path $DownloadPath)) {
    Write-Host "Downloading file..."
    Invoke-WebRequest -Uri $DownloadUrl -OutFile $DownloadPath
} else {
    Write-Host "File sudah ada: $DownloadPath"
}

# 7. Extract file menggunakan 7-Zip
Write-Host "Extracting file..."
& $SevenZipPath x $DownloadPath -o"$ExtractPath" -y

# 8. Rapikan hasil ekstrak (kalau ada subfolder utama)
$subFolders = Get-ChildItem -Path $ExtractPath -Directory
foreach ($folder in $subFolders) {
    if ($folder.Name -like "*ExambrowserClient*") {
        Write-Host "Memindahkan isi dari $($folder.FullName) ke $ExtractPath..."

        Get-ChildItem -Path $folder.FullName -Force | ForEach-Object {
            $targetPath = Join-Path $ExtractPath $_.Name

            if (Test-Path $targetPath) {
                if ($_.PSIsContainer) {
                    # Jika folder sudah ada, pindahkan isi ke dalamnya
                    Get-ChildItem -Path $_.FullName -Force | ForEach-Object {
                        Move-Item -Path $_.FullName -Destination $targetPath -Force
                        Write-Host "Gabung folder: $($_.FullName) -> $targetPath"
                    }
                } else {
                    # Jika file sudah ada, overwrite
                    Move-Item -Path $_.FullName -Destination $targetPath -Force
                    Write-Host "Replace file: $($_.Name)"
                }
            } else {
                # Kalau belum ada, langsung pindah
                Move-Item -Path $_.FullName -Destination $ExtractPath
                Write-Host "Pindah baru: $($_.Name)"
            }
        }

        # Hapus folder kosong setelah isinya dipindah
        Remove-Item -Path $folder.FullName -Recurse -Force
    }
}

Write-Host "Selesai. File diextract ke: $ExtractPath"

# 9. Cari file ExamBrowser.exe
$ExePath = Get-ChildItem -Path $ExtractPath -Recurse -Filter "ExamBrowser.exe" -ErrorAction SilentlyContinue | Select-Object -First 1

if ($ExePath) {
    Write-Host "ExamBrowser.exe ditemukan: $($ExePath.FullName)"

    # Lokasi shortcut di desktop
    $Desktop = [Environment]::GetFolderPath("Desktop")
    $ShortcutPath = Join-Path $Desktop "ANBK.lnk"

    # Buat shortcut
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $ExePath.FullName
    $Shortcut.WorkingDirectory = Split-Path $ExePath.FullName
    $Shortcut.WindowStyle = 1
    $Shortcut.IconLocation = $ExePath.FullName
    $Shortcut.Save()

    Write-Host "Shortcut berhasil dibuat di Desktop: $ShortcutPath"
} else {
    Write-Host "ExamBrowser.exe tidak ditemukan di $ExtractPath"
}
