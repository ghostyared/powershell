# =========================================
# CONFIGURACIÓN
# =========================================
$TARGET_VERSION   = "12.11.3.0"
$INSTALLER        = "WG-MVPN-SSL_12_11_3.exe"
$LOCAL_DIR        = "C:\SISTEMAS\UPDATEVPN"
$SRC              = Join-Path $LOCAL_DIR $INSTALLER
$DOWNLOAD_URL     = "https://cdn.watchguard.com/SoftwareCenter/Files/MUVPN_SSL/12_11_3/WG-MVPN-SSL_12_11_3.exe"

$APP_PATH         = "C:\Program Files (x86)\WatchGuard\WatchGuard Mobile VPN with SSL\wgsslvpnsrc.exe"
$VERSION_FILE_WMI = "C:\Program Files (x86)\WatchGuard\WatchGuard Mobile VPN with SSL\wgsslvpnc.exe"
$dirPath          = "C:\Program Files (x86)\WatchGuard\WatchGuard Mobile VPN with SSL"

# =========================================
# PREPARAR DIRECTORIO LOCAL
# =========================================
if (!(Test-Path -Path $LOCAL_DIR)) {
    New-Item -Path $LOCAL_DIR -ItemType Directory | Out-Null
}


# =========================================
# SI NO EXISTE EL ARCHIVO PRINCIPAL -> INSTALAR
# =========================================
if (!(Test-Path -Path $VERSION_FILE_WMI)) {
    Write-Host "+----------------------------------+"
    Write-Host "|   NO SE DETECTA EL SOFTWARE      |"
    Write-Host "+----------------------------------+"

    # =========================================
    # VERIFICAR/DESCARGAR INSTALADOR
    # =========================================
    if (!(Test-Path -Path $SRC)) {
        Write-Host "[INFO] Instalador no encontrado en $LOCAL_DIR, descargando desde $DOWNLOAD_URL..."
        try {
            Invoke-WebRequest -Uri $DOWNLOAD_URL -OutFile $SRC -UseBasicParsing
            Write-Host "[INFO] Descarga completada."
        }
        catch {
            Write-Host "[ERROR] No se pudo descargar el instalador desde $DOWNLOAD_URL"
            exit 1
        }
    }

    Write-Host "[INFO 1] Iniciando instalacion desde $SRC ..."
    Start-Process $SRC -ArgumentList "/SP-","/VERYSILENT","/SUPPRESSMSGBOXES","/NORESTART" -Wait
    Start-Sleep -Seconds 2

    Write-Host "[INFO 2] Creando servicio..."
    sc.exe create "WatchGuard SSLVPN Service" binPath= "`"$APP_PATH`"" | Out-Null
    sc.exe config "WatchGuard SSLVPN Service" start= auto | Out-Null
    sc.exe start "WatchGuard SSLVPN Service" | Out-Null

    Write-Host "[INFO 3] Lanzando aplicacion..."
    Start-Process $VERSION_FILE_WMI
    Start-Sleep -Seconds 5
    exit 0
}

# =========================================
# SI EXISTE -> VALIDAR VERSION (sin WMIC)
# =========================================
$InstalledVersion = (Get-Item $VERSION_FILE_WMI).VersionInfo.ProductVersion
$InstalledVersion = $InstalledVersion -replace ',', '.' -replace '\s', ''

Write-Host "===================================="
Write-Host "Version instalada: $InstalledVersion"
Write-Host "===================================="

if ([version]$InstalledVersion -gt [version]$TARGET_VERSION) {
    Write-Host "[INFO] La version instalada $InstalledVersion es MAYOR a la requerida $TARGET_VERSION"
    exit 0
}
elseif ([version]$InstalledVersion -eq [version]$TARGET_VERSION) {
    Write-Host "[INFO] La version instalada $InstalledVersion es IGUAL a la requerida $TARGET_VERSION"
    exit 0
}
else {
    Write-Host "[INFO] La version instalada $InstalledVersion es MENOR a la requerida $TARGET_VERSION"


    # =========================================
    # VERIFICAR/DESCARGAR INSTALADOR
    # =========================================
    if (!(Test-Path -Path $SRC)) {
        Write-Host "[INFO] Instalador no encontrado en $LOCAL_DIR, descargando desde $DOWNLOAD_URL..."
        try {
            Invoke-WebRequest -Uri $DOWNLOAD_URL -OutFile $SRC -UseBasicParsing
            Write-Host "[INFO] Descarga completada."
        }
        catch {
            Write-Host "[ERROR] No se pudo descargar el instalador desde $DOWNLOAD_URL"
            exit 1
        }
    }



    Write-Host "[INFO] Instalando nueva version desde $SRC ..."
    Stop-Process -Name "wgsslvpnc","openvpn" -Force -ErrorAction SilentlyContinue
    sc.exe stop "WatchGuard SSLVPN Service" | Out-Null
    sc.exe delete "WatchGuard SSLVPN Service" | Out-Null

    Start-Process $SRC -ArgumentList "/SP-","/VERYSILENT","/SUPPRESSMSGBOXES","/NORESTART" -Wait
    Start-Sleep -Seconds 2

    Write-Host "[INFO] Creando servicio..."
    sc.exe create "WatchGuard SSLVPN Service" binPath= "`"$APP_PATH`"" | Out-Null
    sc.exe config "WatchGuard SSLVPN Service" start= auto | Out-Null
    sc.exe start "WatchGuard SSLVPN Service" | Out-Null

    Write-Host "[INFO] Lanzando aplicacion..."
    Start-Process $VERSION_FILE_WMI
    Start-Sleep -Seconds 3
    exit 0
}
