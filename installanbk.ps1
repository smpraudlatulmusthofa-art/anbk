# ==========================
# Script Download & Extract Exambrowser Client
# ==========================

# 1. Aktifkan TLS 1.2 agar semua versi PowerShell bisa download
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# 2. Lokasi target
$DownloadUrl  = "https://unp.kemdikbud.go.id/tes/64BitExambrowserClient_25.0525.rar"
$ExtractPath  = "$env:USERPROFILE\Documents\exambrowser_client"
$DownloadPath = Join-Path $ExtractPath "64BitExambrowserClient_25.0525.rar"
$SevenZipPath = "C:\Program Files\7-Zip\7z.exe"
$SevenZipInstallerUrl = "https://www.7-zip.org/a/7z2408-x64.exe"   # versi terbaru

# 3. Pastikan folder tujuan ada
if (!(Test-Path $ExtractPath)) {
    New-Item -ItemType Directory -Path $ExtractPath | Out-Null
}

# 4. Cek apakah 7-Zip sudah terinstall
if (!(Test-Path $SevenZipPath)) {
    Write-Host "7-Zip belum ada. Download & install..."
    $InstallerPath = Join-Path $env:TEMP "7z_installer.exe"
    Invoke-WebRequest -Uri $SevenZipInstallerUrl -OutFile $InstallerPath
    Start-Process -FilePath $InstallerPath -ArgumentList "/S" -Wait
}

# 5. Download file RAR jika belum ada
if (!(Test-Path $DownloadPath)) {
    Write-Host "Downloading file..."
    Invoke-WebRequest -Uri $DownloadUrl -OutFile $DownloadPath
} else {
    Write-Host "File sudah ada: $DownloadPath"
}

# 6. Extract file menggunakan 7-Zip
Write-Host "Extracting file..."
& $SevenZipPath x $DownloadPath -o"$ExtractPath" -y

# 7. Rapikan hasil ekstrak (kalau ada subfolder utama)
$subFolders = Get-ChildItem -Path $ExtractPath -Directory
foreach ($folder in $subFolders) {
    if ($folder.Name -like "64BitExambrowserClient*") {
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

# 8. Cari file ExamBrowser.exe
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
