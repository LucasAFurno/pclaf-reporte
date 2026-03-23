#Requires -Version 5.1
<#
.SYNOPSIS
    PCLAF - Script de diagnostico tecnico v3.0
.DESCRIPTION
    Genera un reporte HTML completo del equipo.
    Modo "cliente": reporte legible, visual, sin datos sensibles.
    Modo "tecnico": reporte completo con todos los datos tecnicos.
.PARAMETER Modo
    "cliente" o "tecnico" (default: cliente)
.PARAMETER Tecnico
    Nombre del tecnico que realiza el diagnostico (default: PCLAF)
.PARAMETER MesesMantenimiento
    Meses hasta la proxima revision recomendada (default: 6)
.PARAMETER SistemaInstaladoPorPCLAF
    Switch: indica si el SO fue instalado por PCLAF
.PARAMETER ServicioRealizado
    Descripcion del trabajo realizado hoy (aparece en reporte cliente)
.PARAMETER PrecioServicio
    Precio cobrado por el servicio (aparece en reporte cliente)
#>
param(
    [ValidateSet("cliente","tecnico")]
    [string]$Modo = "cliente",
    [string]$Tecnico = "PCLAF",
    [int]$MesesMantenimiento = 6,
    [switch]$SistemaInstaladoPorPCLAF,
    [string]$ServicioRealizado = "",
    [string]$PrecioServicio = ""
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "SilentlyContinue"
$ScriptVersion = "3.0"
$BasePath = if ($PSScriptRoot) { $PSScriptRoot } else { $env:TEMP }
$FechaReporte = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

# --- HELPERS ----------------------------------------------------------------

function Update-Stage {
    param([int]$Percent, [string]$Message)
    Write-Progress -Activity "PCLAF Diagnostico v$ScriptVersion" -Status $Message -PercentComplete $Percent
    Write-Host ("[{0,3}%] {1}" -f $Percent, $Message) -ForegroundColor DarkCyan
}

function Safe {
    param($v, [string]$d = "N/D")
    if ($null -eq $v -or "$v".Trim() -eq "") { return $d }
    return $v
}

function To-DT {
    param($v)
    try { if ($v) { return [Management.ManagementDateTimeConverter]::ToDateTime($v) } } catch {}
    return $v
}

function Round2 { param($v) try { return [math]::Round([double]$v, 2) } catch { return "N/D" } }

# --- RECOLECCION DE DATOS ---------------------------------------------------

function Get-OsInfo {
    $os = $null; $cv = $null
    try { $os = Get-CimInstance Win32_OperatingSystem } catch {}
    try { $cv = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" } catch {}

    $releaseId = "N/D"
    try { $releaseId = if ($cv.DisplayVersion) { $cv.DisplayVersion } elseif ($cv.ReleaseId) { $cv.ReleaseId } else { "N/D" } } catch {}

    $soName = if ($cv -and $cv.ProductName) { "$($cv.ProductName) $($cv.DisplayVersion)".Trim() } elseif ($os -and $os.Caption) { $os.Caption } else { "Windows" }
    [PSCustomObject]@{
        Equipo         = $env:COMPUTERNAME
        Usuario        = $env:USERNAME
        SO             = $soName
        Version        = $releaseId
        ReleaseId      = Safe $cv.ReleaseId
        Build          = Safe $os.BuildNumber
        Arquitectura   = Safe $os.OSArchitecture
        InstaladoDesde = To-DT $os.InstallDate
        UltimoArranque = To-DT $os.LastBootUpTime
        Fabricante     = Safe $os.Manufacturer
        SerialSO       = Safe $os.SerialNumber
    }
}

function Get-SystemSummary {
    $cs = $null; $bios = $null; $board = $null; $cpu = $null; $ram = $null
    try { $cs   = Get-CimInstance Win32_ComputerSystem } catch {}
    try { $bios  = Get-CimInstance Win32_BIOS } catch {}
    try { $board = Get-CimInstance Win32_BaseBoard } catch {}
    try { $cpu   = Get-CimInstance Win32_Processor | Select-Object -First 1 } catch {}
    try { $ram   = Get-CimInstance Win32_PhysicalMemory } catch {}

    $ramGB = "N/D"
    try {
        $sum = ($ram | Measure-Object -Property Capacity -Sum).Sum
        if ($sum) { $ramGB = Round2 ($sum / 1GB) }
        elseif ($cs.TotalPhysicalMemory) { $ramGB = Round2 ($cs.TotalPhysicalMemory / 1GB) }
    } catch {}

    [PSCustomObject]@{
        Fabricante     = Safe $cs.Manufacturer
        Modelo         = Safe $cs.Model
        CPU            = Safe (($cpu.Name -replace '\s+', ' ').Trim())
        Cores          = Safe $cpu.NumberOfCores
        Hilos          = Safe $cpu.NumberOfLogicalProcessors
        RAM_Total_GB   = $ramGB
        Slots_RAM      = if ($ram) { ($ram | Measure-Object).Count } else { "N/D" }
        BIOS_Version   = Safe ($bios.SMBIOSBIOSVersion -join ", ")
        BIOS_Fecha     = To-DT $bios.ReleaseDate
        Motherboard    = "$(Safe $board.Manufacturer) $(Safe $board.Product)".Trim()
        NroSerieEquipo = Safe $bios.SerialNumber
    }
}

function Get-GpuInfo {
    try {
        $gpus = Get-CimInstance Win32_VideoController | Where-Object { 
            $_.Name -notmatch "Remote|Virtual|Basic|Microsoft|Hyper-V" -or $_.AdapterRAM -gt 0
        }
        if (-not $gpus) { $gpus = Get-CimInstance Win32_VideoController | Select-Object -First 1 }
        if (-not $gpus) { return [PSCustomObject]@{ GPU="No detectada"; VRAM_GB="N/D"; Driver="N/D"; Resolucion="N/D" } }
        $gpus | ForEach-Object {
            [PSCustomObject]@{
                GPU       = Safe $_.Name
                VRAM_GB   = if ($_.AdapterRAM) { Round2 ($_.AdapterRAM / 1GB) } else { "N/D" }
                Driver    = Safe $_.DriverVersion
                Resolucion = if ($_.CurrentHorizontalResolution) { "$($_.CurrentHorizontalResolution)x$($_.CurrentVerticalResolution)" } else { "N/D" }
            }
        }
    } catch {
        [PSCustomObject]@{ GPU="Error"; VRAM_GB="N/D"; Driver="N/D"; Resolucion="N/D" }
    }
}

function Get-RamDetail {
    try {
        $mems = Get-CimInstance Win32_PhysicalMemory
        if (-not $mems) { return [PSCustomObject]@{ Banco="N/D"; GB="N/D"; Tipo="N/D"; MHz="N/D"; Fab="N/D"; Part="N/D" } }
        $mems | ForEach-Object {
            [PSCustomObject]@{
                Banco = Safe $_.BankLabel
                GB    = if ($_.Capacity) { Round2 ($_.Capacity / 1GB) } else { "N/D" }
                Tipo  = Safe $_.SMBIOSMemoryType
                MHz   = Safe $_.Speed
                Fab   = Safe $_.Manufacturer
                Part  = Safe ($_.PartNumber.Trim())
            }
        }
    } catch {
        [PSCustomObject]@{ Banco="Error"; GB="N/D"; Tipo="N/D"; MHz="N/D"; Fab="N/D"; Part="N/D" }
    }
}

function Get-Temperatures {
    $results = @()

    # WMI thermal zones (funciona en la mayoria de los equipos modernos)
    try {
        $zones = Get-WmiObject MSAcpi_ThermalZoneTemperature -Namespace "root/wmi"
        foreach ($z in $zones) {
            $celsius = [math]::Round(($z.CurrentTemperature / 10) - 273.15, 1)
            $estado = if ($celsius -ge 90) { "CRITICO" } elseif ($celsius -ge 75) { "ALTO" } elseif ($celsius -ge 60) { "ELEVADO" } else { "NORMAL" }
            $results += [PSCustomObject]@{
                Zona    = $z.InstanceName -replace ".*\\", ""
                Celsius = $celsius
                Estado  = $estado
            }
        }
    } catch {}

    # CPU via OpenHardwareMonitor si esta instalado (opcional)
    try {
        $ohm = Get-WmiObject -Namespace "root/OpenHardwareMonitor" -Class Sensor -ErrorAction Stop |
               Where-Object { $_.SensorType -eq "Temperature" }
        foreach ($s in $ohm) {
            $celsius = [math]::Round($s.Value, 1)
            $estado = if ($celsius -ge 90) { "CRITICO" } elseif ($celsius -ge 75) { "ALTO" } elseif ($celsius -ge 60) { "ELEVADO" } else { "NORMAL" }
            $results += [PSCustomObject]@{
                Zona    = "$($s.Parent) / $($s.Name)"
                Celsius = $celsius
                Estado  = $estado
            }
        }
    } catch {}

    if ($results.Count -eq 0) {
        return [PSCustomObject]@{
            Zona    = "No disponible"
            Celsius = "N/D"
            Estado  = "Sin sensor"
        }
    }
    return $results | Sort-Object { try { [double]$_.Celsius } catch { 0 } } -Descending
}

function Get-DiskHealth {
    try {
        $disks = Get-CimInstance Win32_DiskDrive
        if (-not $disks) { return [PSCustomObject]@{ Disco="Sin discos"; Modelo="N/D"; Serial="N/D"; Tipo="N/D"; GB="N/D"; Estado="N/D"; Salud="N/D"; SMART="N/D"; Detalle="N/D" } }

        # SMART map
        $smartMap = @{}
        try {
            $smart = Get-WmiObject -Namespace root\wmi -Class MSStorageDriver_FailurePredictStatus
            foreach ($s in $smart) { $smartMap[$s.InstanceName] = $s }
        } catch {}

        # Physical disk map
        $pdMap = @{}
        try {
            Get-PhysicalDisk | ForEach-Object {
                if ($_.FriendlyName) { $pdMap[$_.FriendlyName] = $_ }
            }
        } catch {}

        $disks | ForEach-Object {
            $model  = ($_.Model -replace '\s+', ' ').Trim()
            $serial = ($_.SerialNumber -replace '\s+', ' ').Trim()
            $sizeGB = if ($_.Size) { Round2 ($_.Size / 1GB) } else { "N/D" }

            $pd = $pdMap[$model]
            $health = if ($pd) { $pd.HealthStatus } else { "Desconocido" }
            $opState = if ($pd) { ($pd.OperationalStatus -join ", ") } else { "Desconocido" }
            $mediaType = if ($pd) { $pd.MediaType } else { "Desconocido" }

            $predictFail = $false; $predictReason = 0
            foreach ($k in $smartMap.Keys) {
                if ($k -like "*$($_.PNPDeviceID)*" -or ($serial -and $k -like "*$serial*")) {
                    $predictFail = $smartMap[$k].PredictFailure
                    $predictReason = $smartMap[$k].Reason
                    break
                }
            }

            $estado = "BIEN"; $detalle = "Sin alertas detectadas"
            if ($predictFail) { $estado = "REEMPLAZAR"; $detalle = "SMART indica riesgo de falla inminente" }
            elseif ($health -match "Unhealthy|Warning") { $estado = "MAL"; $detalle = "Windows reporta problema de salud" }
            elseif ($health -eq "Desconocido") { $estado = "SIN DATOS"; $detalle = "El driver no expone datos de salud" }

            [PSCustomObject]@{
                Disco   = "DRIVE$($_.Index)"
                Modelo  = $model
                Serial  = if ($serial) { $serial } else { "N/D" }
                Tipo    = Safe $mediaType
                GB      = $sizeGB
                Estado  = $estado
                Salud   = Safe $health
                SMART   = if ($predictFail) { "FALLO" } else { "OK" }
                Detalle = $detalle
            }
        }
    } catch {
        [PSCustomObject]@{ Disco="Error"; Modelo="N/D"; Serial="N/D"; Tipo="N/D"; GB="N/D"; Estado="N/D"; Salud="N/D"; SMART="N/D"; Detalle="Error al consultar" }
    }
}

function Get-VolumeUsage {
    try {
        $vols = Get-Volume | Where-Object { $_.DriveLetter }
        if (-not $vols) { return [PSCustomObject]@{ Unidad="N/D"; Etiqueta="N/D"; FS="N/D"; GB_Total="N/D"; GB_Libre="N/D"; Uso_Pct="N/D"; Alerta="N/D" } }
        $vols | ForEach-Object {
            $total  = if ($_.Size)            { Round2 ($_.Size / 1GB) }            else { 0 }
            $free   = if ($_.SizeRemaining)   { Round2 ($_.SizeRemaining / 1GB) }   else { 0 }
            $usePct = if ($_.Size -gt 0)      { Round2 ((($_.Size - $_.SizeRemaining) / $_.Size) * 100) } else { 0 }
            $alert  = if ($usePct -ge 95)     { "CRITICO" } elseif ($usePct -ge 85) { "ALTO" } elseif ($usePct -ge 70) { "MODERADO" } else { "OK" }
            [PSCustomObject]@{
                Unidad   = "$($_.DriveLetter):"
                Etiqueta = Safe $_.FileSystemLabel
                FS       = Safe $_.FileSystem
                GB_Total = $total
                GB_Libre = $free
                Uso_Pct  = $usePct
                Alerta   = $alert
            }
        }
    } catch {
        [PSCustomObject]@{ Unidad="Error"; Etiqueta="N/D"; FS="N/D"; GB_Total="N/D"; GB_Libre="N/D"; Uso_Pct="N/D"; Alerta="N/D" }
    }
}

function Get-NetworkInfo {
    try {
        Get-NetAdapter | ForEach-Object {
            [PSCustomObject]@{
                Nombre    = Safe $_.Name
                Interface = Safe $_.InterfaceDescription
                Estado    = Safe $_.Status
                Velocidad = Safe $_.LinkSpeed
                MAC       = Safe $_.MacAddress
            }
        }
    } catch {
        [PSCustomObject]@{ Nombre="Error"; Interface="N/D"; Estado="N/D"; Velocidad="N/D"; MAC="N/D" }
    }
}

function Get-SecurityInfo {
    try {
        $sb   = try { Confirm-SecureBootUEFI } catch { "No soportado" }
        $tpm  = Get-Tpm
        $av   = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct
        [PSCustomObject]@{
            SecureBoot  = Safe $sb
            TPM         = Safe $tpm.TpmPresent
            TPM_Ready   = Safe $tpm.TpmReady
            TPM_Version = Safe $tpm.ManufacturerVersion
            Antivirus   = if ($av) { ($av.displayName -join ", ") } else { "N/D" }
        }
    } catch {
        [PSCustomObject]@{ SecureBoot="N/D"; TPM="N/D"; TPM_Ready="N/D"; TPM_Version="N/D"; Antivirus="N/D" }
    }
}

function Get-DefenderStatus {
    try {
        $mp = Get-MpComputerStatus
        [PSCustomObject]@{
            Activo             = $mp.AntivirusEnabled
            TiempoReal         = $mp.RealTimeProtectionEnabled
            Antispyware        = $mp.AntispywareEnabled
            Version_Motor      = $mp.AMEngineVersion
            Firmas_AV          = $mp.AntivirusSignatureVersion
            Firmas_AS          = $mp.AntispywareSignatureVersion
            Update_Firmas      = $mp.AntispywareSignatureLastUpdated
            Ultimo_QuickScan   = $mp.QuickScanEndTime
            Ultimo_FullScan    = $mp.FullScanEndTime
        }
    } catch {
        [PSCustomObject]@{ Activo="N/D"; TiempoReal="N/D"; Antispyware="N/D"; Version_Motor="N/D"; Firmas_AV="N/D"; Firmas_AS="N/D"; Update_Firmas="N/D"; Ultimo_QuickScan="N/D"; Ultimo_FullScan="N/D" }
    }
}

function Get-BatteryInfo {
    try {
        $bat = Get-CimInstance Win32_Battery
        if (-not $bat) { return [PSCustomObject]@{ Detectada="NO"; Estado="Desktop o sin bateria"; Carga_Pct="N/D"; EstadoBat="N/D" } }
        $bat | ForEach-Object {
            [PSCustomObject]@{
                Detectada   = "SI"
                Estado      = Safe $_.Status
                Carga_Pct   = Safe $_.EstimatedChargeRemaining
                EstadoBat   = Safe $_.BatteryStatus
            }
        }
    } catch {
        [PSCustomObject]@{ Detectada="N/D"; Estado="N/D"; Carga_Pct="N/D"; EstadoBat="N/D" }
    }
}

function Get-WindowsActivation {
    try {
        $lic = Get-CimInstance SoftwareLicensingProduct | Where-Object { $_.PartialProductKey -and $_.Name -like "*Windows*" } | Select-Object -First 1
        $statusMap = @{ 1="Activado"; 0="Sin licencia"; 2="Fuera de gracia"; 3="Periodo adicional"; 4="Sin llave"; 5="Notificacion" }
        [PSCustomObject]@{
            Estado          = if ($lic) { $statusMap[[int]$lic.LicenseStatus] } else { "N/D" }
            Producto        = if ($lic) { $lic.Name } else { "N/D" }
            Clave_Parcial   = if ($lic) { $lic.PartialProductKey } else { "N/D" }
        }
    } catch {
        [PSCustomObject]@{ Estado="N/D"; Producto="N/D"; Clave_Parcial="N/D" }
    }
}

function Get-PowerProfile {
    try {
        $active = powercfg /getactivescheme 2>$null
        $line = ($active | Out-String).Trim()
        $nameMatch = [regex]::Match($line, '\((.+)\)')
        $name = if ($nameMatch.Success) { $nameMatch.Groups[1].Value } else { $line }
        $perfMap = @{
            "Alto rendimiento" = "MAXIMO"
            "High performance" = "MAXIMO"
            "Equilibrado" = "EQUILIBRADO"
            "Balanced" = "EQUILIBRADO"
            "Economizador de energia" = "AHORRO"
            "Power saver" = "AHORRO"
        }
        $clase = "EQUILIBRADO"
        foreach ($k in $perfMap.Keys) { if ($name -like "*$k*") { $clase = $perfMap[$k]; break } }
        [PSCustomObject]@{ Perfil = $name; Clase = $clase }
    } catch {
        [PSCustomObject]@{ Perfil = "N/D"; Clase = "N/D" }
    }
}

function Get-OpenPorts {
    try {
        $conns = netstat -ano 2>$null | Select-String "LISTENING|ESTABLISHED" | Select-Object -First 30
        $conns | ForEach-Object {
            $parts = ($_.Line -split '\s+') | Where-Object { $_ }
            if ($parts.Count -ge 4) {
                [PSCustomObject]@{
                    Protocolo = Safe $parts[0]
                    Local     = Safe $parts[1]
                    Remoto    = Safe $parts[2]
                    Estado    = Safe $parts[3]
                    PID       = Safe $parts[4]
                }
            }
        }
    } catch {
        [PSCustomObject]@{ Protocolo="N/D"; Local="N/D"; Remoto="N/D"; Estado="N/D"; PID="N/D" }
    }
}

function Get-NonMsServices {
    try {
        $svcs = Get-CimInstance Win32_Service | Where-Object {
            $_.State -eq "Running" -and
            $_.PathName -and
            $_.PathName -notmatch "system32|SysWOW64|Windows\\|Microsoft|MpKsl|WdFilter" -and
            $_.StartMode -ne "Disabled"
        } | Select-Object -First 25
        if (-not $svcs) { return [PSCustomObject]@{ Servicio="Ninguno"; Estado="N/D"; Inicio="N/D"; Ruta="N/D" } }
        $svcs | ForEach-Object {
            [PSCustomObject]@{
                Servicio = Safe $_.DisplayName
                Estado   = Safe $_.State
                Inicio   = Safe $_.StartMode
                Ruta     = Safe ($_.PathName -replace '"','')
            }
        }
    } catch {
        [PSCustomObject]@{ Servicio="Error"; Estado="N/D"; Inicio="N/D"; Ruta="N/D" }
    }
}

function Get-CustomScheduledTasks {
    try {
        $tasks = Get-ScheduledTask | Where-Object {
            $_.TaskPath -notmatch "\\Microsoft\\|\\Windows\\" -and
            $_.State -ne "Disabled"
        } | Select-Object -First 20
        if (-not $tasks) { return [PSCustomObject]@{ Tarea="Ninguna personalizada"; Estado="N/D"; Autor="N/D" } }
        $tasks | ForEach-Object {
            [PSCustomObject]@{
                Tarea  = Safe $_.TaskName
                Estado = Safe $_.State
                Autor  = Safe $_.Author
            }
        }
    } catch {
        [PSCustomObject]@{ Tarea="Error"; Estado="N/D"; Autor="N/D" }
    }
}

function Get-DriversInfo {
    try {
        $drivers = Get-CimInstance Win32_PnPSignedDriver |
            Where-Object { $_.DeviceName -and $_.DriverVersion } |
            Sort-Object DriverDate |
            Select-Object -First 30 @{N='Dispositivo';E={$_.DeviceName}}, @{N='Version';E={$_.DriverVersion}}, @{N='Fecha';E={$_.DriverDate}}, @{N='Fabricante';E={$_.Manufacturer}}
        if (-not $drivers) { return [PSCustomObject]@{ Dispositivo="Sin datos"; Version="N/D"; Fecha="N/D"; Fabricante="N/D" } }
        $drivers
    } catch {
        [PSCustomObject]@{ Dispositivo="Error"; Version="N/D"; Fecha="N/D"; Fabricante="N/D" }
    }
}

function Get-WindowsUpdateStatus {
    try {
        $reboot = $false
        if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") { $reboot = $true }
        if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") { $reboot = $true }
        $svc = Get-Service wuauserv
        [PSCustomObject]@{ Servicio_WU = Safe $svc.Status; Reinicio_Pendiente = if ($reboot) { "SI" } else { "NO" } }
    } catch {
        [PSCustomObject]@{ Servicio_WU = "N/D"; Reinicio_Pendiente = "N/D" }
    }
}

function Get-TrimStatus {
    try {
        $trim = (fsutil behavior query DisableDeleteNotify) -join " "
        $estado = if ($trim -match "= 0") { "ACTIVO" } elseif ($trim -match "= 1") { "DESACTIVADO" } else { "N/D" }
        [PSCustomObject]@{ TRIM = $estado; Detalle = $trim }
    } catch {
        [PSCustomObject]@{ TRIM = "N/D"; Detalle = "N/D" }
    }
}

function Get-BootTime {
    try {
        $ev = Get-WinEvent -FilterHashtable @{ LogName='System'; Id=@(1,12,6005); StartTime=(Get-Date).AddDays(-7) } -MaxEvents 10 -ErrorAction Stop
        $boot = $ev | Where-Object { $_.Id -in @(12,6005) } | Sort-Object TimeCreated -Descending | Select-Object -First 3
        $ms = $ev | Where-Object { $_.Id -eq 1 } | Sort-Object TimeCreated -Descending | Select-Object -First 3

        $bootRows = @()
        if ($boot) {
            $bootRows = $boot | ForEach-Object {
                [PSCustomObject]@{ Fecha=$_.TimeCreated; Evento="Arranque del sistema ($($_.Id))"; Segundos="N/D" }
            }
        }
        if ($ms) {
            $bootRows += $ms | ForEach-Object {
                $sec = "N/D"
                try { $sec = [math]::Round([xml]$_.ToXml().OuterXml.SelectSingleNode("//Data[@Name='BootTime']").'#text' / 1000, 1) } catch {}
                [PSCustomObject]@{ Fecha=$_.TimeCreated; Evento="Boot performance"; Segundos=$sec }
            }
        }
        if (-not $bootRows) { return [PSCustomObject]@{ Fecha="N/D"; Evento="Sin eventos recientes"; Segundos="N/D" } }
        return $bootRows | Sort-Object Fecha -Descending | Select-Object -First 5
    } catch {
        [PSCustomObject]@{ Fecha="N/D"; Evento="No se pudo consultar"; Segundos="N/D" }
    }
}

function Get-CriticalEvents {
    $start = (Get-Date).AddDays(-30)
    $results = @()
    $queries = @(
        @{ Log="System"; Prov="Microsoft-Windows-WHEA-Logger"; Levels=@(1,2,3) }
        @{ Log="System"; Prov="Disk"; Levels=@(1,2,3) }
        @{ Log="System"; Prov="Microsoft-Windows-Kernel-Power"; Levels=@(1,2,3) }
        @{ Log="System"; Prov="Ntfs"; Levels=@(1,2,3) }
        @{ Log="Application"; Prov="Application Error"; Levels=@(1,2,3) }
        @{ Log="System"; Prov="Service Control Manager"; Levels=@(1,2,3) }
    )
    foreach ($q in $queries) {
        foreach ($lvl in $q.Levels) {
            try {
                $ev = Get-WinEvent -FilterHashtable @{ LogName=$q.Log; ProviderName=$q.Prov; Level=$lvl; StartTime=$start } -ErrorAction SilentlyContinue
                if ($ev) { $results += $ev | Select-Object TimeCreated, Id, ProviderName, LevelDisplayName, @{N='Message';E={ ($_.Message -split "`n")[0] }} }
            } catch {}
        }
    }
    if ($results.Count -eq 0) { return [PSCustomObject]@{ Fecha=""; Id=""; Origen="Sin eventos criticos"; Nivel=""; Mensaje="" } }
    $results | Sort-Object TimeCreated -Descending | Select-Object -First 40 `
        @{N='Fecha';E={$_.TimeCreated}}, @{N='Id';E={$_.Id}},
        @{N='Origen';E={$_.ProviderName}}, @{N='Nivel';E={$_.LevelDisplayName}},
        @{N='Mensaje';E={$_.Message}}
}

function Get-MemoryErrors {
    try {
        $errs = Get-WinEvent -FilterHashtable @{
            LogName='System'; ProviderName='Microsoft-Windows-MemoryDiagnostics-Results'; StartTime=(Get-Date).AddDays(-30)
        } -ErrorAction SilentlyContinue
        if (-not $errs) { return [PSCustomObject]@{ Fecha="Sin pruebas recientes"; Resultado="N/D" } }
        $errs | ForEach-Object { [PSCustomObject]@{ Fecha=$_.TimeCreated; Resultado=($_.Message -split "`n")[0] } }
    } catch {
        [PSCustomObject]@{ Fecha="No disponible"; Resultado="N/D" }
    }
}

function Get-CurrentPerformance {
    try {
        $cpu = (Get-Counter '\Processor(_Total)\% Processor Time').CounterSamples.CookedValue
        $os  = Get-CimInstance Win32_OperatingSystem
        $tot = Round2 ($os.TotalVisibleMemorySize / 1MB)
        $fre = Round2 ($os.FreePhysicalMemory / 1MB)
        $use = Round2 ($tot - $fre)
        $pct = if ($tot -gt 0) { Round2 (($use / $tot) * 100) } else { 0 }
        [PSCustomObject]@{ CPU_Pct=$([math]::Round($cpu,1)); RAM_Usada_GB=$use; RAM_Total_GB=$tot; RAM_Pct=$pct }
    } catch {
        [PSCustomObject]@{ CPU_Pct="N/D"; RAM_Usada_GB="N/D"; RAM_Total_GB="N/D"; RAM_Pct="N/D" }
    }
}

function Get-TopProcesses {
    try {
        Get-Process | Sort-Object CPU -Descending | Select-Object -First 15 `
            @{N='Proceso';E={$_.ProcessName}},
            @{N='CPU_s';E={Round2 $_.CPU}},
            @{N='RAM_MB';E={Round2 ($_.WorkingSet64/1MB)}},
            @{N='PID';E={$_.Id}}
    } catch {
        [PSCustomObject]@{ Proceso="N/D"; CPU_s="N/D"; RAM_MB="N/D"; PID="N/D" }
    }
}

function Get-StartupApps {
    try {
        $items = Get-CimInstance Win32_StartupCommand |
            Select-Object @{N='Nombre';E={$_.Name}}, @{N='Comando';E={$_.Command}}, @{N='Ubicacion';E={$_.Location}}, @{N='Usuario';E={$_.User}}
        if (-not $items) { return [PSCustomObject]@{ Nombre="Ninguno"; Comando="N/D"; Ubicacion="N/D"; Usuario="N/D" } }
        $items
    } catch {
        [PSCustomObject]@{ Nombre="Error"; Comando="N/D"; Ubicacion="N/D"; Usuario="N/D" }
    }
}

function Get-InstalledApps {
    $paths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*","HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")
    $items = @()
    foreach ($p in $paths) {
        try { $items += Get-ItemProperty $p | Where-Object { $_.DisplayName } | Select-Object @{N='Nombre';E={$_.DisplayName}},@{N='Version';E={$_.DisplayVersion}},@{N='Fab';E={$_.Publisher}},@{N='Fecha';E={$_.InstallDate}} } catch {}
    }
    if (-not $items) { return [PSCustomObject]@{ Nombre="N/D"; Version="N/D"; Fab="N/D"; Fecha="N/D" } }
    $items | Sort-Object Nombre -Unique
}

function Get-IntegrityStatus {
    $sfcLog = Join-Path $env:windir "Logs\CBS\CBS.log"
    $state = "Sin datos"; $hint = "No se ejecuto SFC recientemente"
    try {
        if (Test-Path $sfcLog) {
            $lines = (Select-String -Path $sfcLog -Pattern "\[SR\]" | Select-Object -Last 20).Line -join " "
            if ($lines -match "cannot repair|corrupted") { $state = "CORRUPCION"; $hint = "Conviene correr SFC + DISM" }
            elseif ($lines -match "Verify and Repair|completed") { $state = "OK (previas reparaciones)"; $hint = "Se realizaron reparaciones en el pasado" }
            else { $state = "OK"; $hint = "Sin indicios de corrupcion reciente" }
        }
    } catch {}
    [PSCustomObject]@{ Estado=$state; Observacion=$hint; Accion="sfc /scannow && DISM /Online /Cleanup-Image /RestoreHealth" }
}

function Get-HardwareFingerprint {
    param($Sys, $Disks)
    $d0 = $Disks | Sort-Object Disco | Select-Object -First 1
    $plain = "$($env:COMPUTERNAME)|$($Sys.NroSerieEquipo)|$($Sys.Motherboard)|$($Sys.CPU)|$($Sys.RAM_Total_GB)|$($d0.Modelo)|$($d0.Serial)"
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $hash = [System.BitConverter]::ToString($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($plain))).Replace("-","")
    [PSCustomObject]@{ Hash=$hash; Hostname=$env:COMPUTERNAME; Serial=$Sys.NroSerieEquipo; MB=$Sys.Motherboard; CPU=$Sys.CPU; RAM=$Sys.RAM_Total_GB; Disco_Modelo=$d0.Modelo; Disco_Serial=$d0.Serial }
}

function Get-HardwareAge {
    param($Sys, $Gpu, $Disks, $Ram)
    $cpu = [string]$Sys.CPU; $gpu = [string](($Gpu|Select-Object -First 1).GPU)
    $ramGB = 0; try { $ramGB = [double]$Sys.RAM_Total_GB } catch {}
    $ramMhz = 0; try { $ramMhz = [int](($Ram|Select-Object -First 1).MHz) } catch {}
    $diskType = [string](($Disks|Select-Object -First 1).Tipo)

    $cpuEst = "INTERMEDIO"; $cpuMsg = "Rendimiento aceptable para uso diario"
    if ($cpu -match "FX-|Phenom|Core\(TM\)2|i[357]-[23]\d{3}|i[357]-4\d{3}|Athlon II|Pentium D|Celeron [GE][0-9]") { $cpuEst="VIEJO"; $cpuMsg="CPU muy viejo. Evaluar cambio de plataforma si se usa intensivamente" }
    elseif ($cpu -match "i[3579]-[4567]\d{3}|Ryzen [357] [123]\d{3}") { $cpuEst="USABLE"; $cpuMsg="Todavia funciona bien, pero ya no es plataforma nueva" }
    elseif ($cpu -match "Ryzen [357] [456789]\d{3}|i[3579]-1[0-4]\d{3}|Ultra") { $cpuEst="NUEVO"; $cpuMsg="CPU moderno. No hace falta cambio" }

    $ramEst = "INTERMEDIA"; $ramMsg = "RAM aceptable"
    if ($ramGB -lt 8) { $ramEst="INSUFICIENTE"; $ramMsg="Menos de 8 GB es critico hoy. Ampliar urgente" }
    elseif ($ramGB -lt 16) { $ramEst="JUSTA"; $ramMsg="8 GB alcanza justo. Conviene llegar a 16 GB" }
    elseif ($ramGB -lt 32) { $ramEst="BUENA"; $ramMsg="16 GB esta bien para la mayoria de usos" }
    else { $ramEst="SOBRADA"; $ramMsg="32 GB o mas es suficiente para cualquier uso" }
    if ($ramMhz -gt 0 -and $ramMhz -le 1600) { $ramMsg += " (velocidad baja, plataforma vieja)" }

    $diskEst = "SIN DATOS"; $diskMsg = "No se pudo clasificar"
    if ($diskType -match "HDD") { $diskEst="LENTO (HDD)"; $diskMsg="Pasando a SSD se nota mucho la diferencia" }
    elseif ($diskType -match "SSD" -or ($Disks|Where-Object{$_.Modelo -match "NVMe|SSD"})) { $diskEst="RAPIDO (SSD)"; $diskMsg="Bien. SSD marca la diferencia en velocidad" }

    $gpuEst = "INTERMEDIA"; $gpuMsg = "GPU razonable"
    if ($gpu -match "HD [234]\d{3}|HD Graphics [23]\d{3}|GT [67]\d{2}|GTX [456789]\d{2}|R5 Graphics") { $gpuEst="VIEJA"; $gpuMsg="GPU vieja o integrada. Sirve para ofimatica, no para juegos o renders" }
    elseif ($gpu -match "RTX [234]\d{3}|RX [67]\d{3}|Arc") { $gpuEst="NUEVA"; $gpuMsg="GPU actual. Sin necesidad de cambio" }

    $overall = "EQUIPO USABLE"; $overallMsg = "No es urgente cambiar hardware"
    if ($cpuEst -eq "VIEJO" -and $ramGB -lt 12) { $overall="PLATAFORMA VIEJA"; $overallMsg="Conviene cambio de mother + CPU + RAM" }
    elseif ($cpuEst -eq "NUEVO" -and $diskEst -match "SSD" -and $ramGB -ge 16) { $overall="EQUIPO VIGENTE"; $overallMsg="Hardware moderno. Sin cambios necesarios" }
    elseif ($ramEst -eq "INSUFICIENTE" -or $ramEst -eq "JUSTA") { $overall="MEJORA RECOMENDADA"; $overallMsg="Ampliar RAM es la mejora de mayor impacto" }

    [PSCustomObject]@{
        CPU_Estado=$cpuEst; CPU_Msg=$cpuMsg
        RAM_Estado=$ramEst; RAM_Msg=$ramMsg
        Disco_Estado=$diskEst; Disco_Msg=$diskMsg
        GPU_Estado=$gpuEst; GPU_Msg=$gpuMsg
        Equipo_Estado=$overall; Equipo_Msg=$overallMsg
    }
}

function Get-PCLAFPaths {
    $root = "C:\ProgramData\PCLAF"
    [PSCustomObject]@{ Root=$root; Last=(Join-Path $root "last.json"); Current=(Join-Path $root "current.json") }
}

function Read-PreviousRecord { param($P) try { if (Test-Path $P.Last) { return (Get-Content $P.Last -Raw | ConvertFrom-Json) } } catch {}; return $null }

function Write-Record {
    param($P, $Rec)
    try {
        if (-not (Test-Path $P.Root)) { New-Item $P.Root -ItemType Directory -Force | Out-Null }
        $j = $Rec | ConvertTo-Json -Depth 10
        Set-Content $P.Current $j -Encoding UTF8
        Set-Content $P.Last $j -Encoding UTF8
        return $true
    } catch { return $false }
}

function Set-RegistryMark {
    param($FP, $Status, $Sys, $Disks)
    $rp = "HKCU:\SOFTWARE\PCLAF\Diagnostics"
    try {
        if (-not (Test-Path $rp)) { New-Item $rp -Force | Out-Null }
        $serialArr = $Disks | ForEach-Object { $_.Serial }; $serials = $serialArr -join ";"
        @{
            LastRunDate   = (Get-Date -Format "yyyy-MM-ddTHH:mm:ss")
            ScriptVersion = $ScriptVersion
            EstadoGeneral = [string]$Status.EstadoGeneral
            Fingerprint   = [string]$FP.Hash
            Equipo        = $env:COMPUTERNAME
            SerialEquipo  = [string]$Sys.NroSerieEquipo
            CPU           = [string]$Sys.CPU
            RAM_Total_GB  = [string]$Sys.RAM_Total_GB
            DiskSerials   = $serials
            Tecnico       = $Tecnico
            PCLAF_OS      = $(if($SistemaInstaladoPorPCLAF){"SI"}else{"NO"})
        }.GetEnumerator() | ForEach-Object { Set-ItemProperty -Path $rp -Name $_.Key -Value $_.Value -Force }
        return $true
    } catch { return $false }
}

function Compare-Records {
    param($Prev, $FP, $Sys, $Disks, $Gpu)
    $r = [PSCustomObject]@{
        Marca_Previa="NO"; Fecha_Anterior="N/D"; Tecnico_Anterior="N/D"
        FP_Anterior="N/D"; FP_Actual=$FP.Hash; FP_Coincide="N/D"
        Cambio_Disco="N/D"; Cambio_RAM="N/D"; Cambio_CPU="N/D"; Cambio_GPU="N/D"
        Observacion="Sin revision anterior registrada"
    }
    if (-not $Prev) { return $r }
    try {
        $pDiskArr = $Prev.Discos | ForEach-Object { $_.Serial }; $pDisks = $pDiskArr -join ";"
        $cDiskArr = $Disks | ForEach-Object { $_.Serial }; $cDisks = $cDiskArr -join ";"
        $pFP = [string]$Prev.Fingerprint.Hash
        $r.Marca_Previa    = "SI"
        $r.Fecha_Anterior  = Safe ([string]$Prev.Metadata.Fecha)
        $r.Tecnico_Anterior= Safe ([string]$Prev.Metadata.Tecnico)
        $r.FP_Anterior     = Safe $pFP
        $r.FP_Coincide     = if ($pFP -eq $FP.Hash) { "SI" } else { "NO - el equipo cambio hardware" }
        $r.Cambio_Disco    = if ($pDisks -and $pDisks -ne $cDisks) { "SI" } else { "NO" }
        $r.Cambio_RAM      = if ([string]$Prev.SystemSummary.RAM_Total_GB -ne [string]$Sys.RAM_Total_GB) { "SI" } else { "NO" }
        $r.Cambio_CPU      = if ([string]$Prev.SystemSummary.CPU -ne [string]$Sys.CPU) { "SI" } else { "NO" }
        $r.Cambio_GPU      = if ([string]($Prev.GPU|Select-Object -First 1).GPU -ne [string](($Gpu|Select-Object -First 1).GPU)) { "SI" } else { "NO" }
        $notes = @()
        if ($r.FP_Coincide -match "NO") { $notes += "Hardware cambiado desde la ultima revision" }
        if ($r.Cambio_Disco -eq "SI") { $notes += "Cambio de disco detectado" }
        if ($r.Cambio_RAM -eq "SI") { $notes += "Cambio de RAM detectado" }
        if ($r.Cambio_CPU -eq "SI") { $notes += "Cambio de CPU detectado" }
        if ($r.Cambio_GPU -eq "SI") { $notes += "Cambio de GPU detectado" }
        $r.Observacion = if ($notes) { $notes -join " / " } else { "Sin cambios de hardware desde la ultima revision" }
    } catch {}
    return $r
}

function Get-FinalAssessment {
    param($Disks, $Vols, $Events, $HwAge, $Perf, $Defender)
    $estado = "EXCELENTE"; $motivos = @()

    if ($Disks | Where-Object { $_.Estado -eq "REEMPLAZAR" }) { $estado="REQUIERE ATENCION URGENTE"; $motivos += "Disco con fallo inminente" }
    elseif ($Disks | Where-Object { $_.Estado -eq "MAL" }) { $estado="REQUIERE REVISION"; $motivos += "Disco con problemas" }

    if ($Vols | Where-Object { $_.Alerta -in @("CRITICO","ALTO") }) {
        if ($estado -eq "EXCELENTE") { $estado="CON OBSERVACIONES" }
        $motivos += "Poco espacio en disco"
    }

    $critEvts = $Events | Where-Object { $_.Nivel -in @("Error","Critical","Warning") -and $_.Origen -notmatch "Service Control" }
    if ($critEvts) {
        if ($estado -eq "EXCELENTE") { $estado="CON OBSERVACIONES" }
        $motivos += "Eventos de advertencia en Windows"
    }

    if ($HwAge.Equipo_Estado -eq "PLATAFORMA VIEJA") {
        if ($estado -eq "EXCELENTE") { $estado="CON OBSERVACIONES" }
        $motivos += "Hardware con antiguedad significativa"
    }

    try {
        if ($Perf.RAM_Pct -ne "N/D" -and [double]$Perf.RAM_Pct -ge 85) {
            if ($estado -eq "EXCELENTE") { $estado="CON OBSERVACIONES" }
            $motivos += "RAM muy exigida en uso actual"
        }
    } catch {}

    if ($Defender.Activo -eq $false -or $Defender.TiempoReal -eq $false) {
        if ($estado -eq "EXCELENTE") { $estado="CON OBSERVACIONES" }
        $motivos += "Proteccion de Windows Defender inactiva"
    }

    if (-not $motivos) { $motivos += "Sin alertas detectadas" }
    [PSCustomObject]@{ EstadoGeneral=$estado; Motivos=($motivos -join " / ") }
}

function Get-Recommendations {
    param($Disks, $Vols, $Events, $Perf, $HwAge, $Integrity, $Def, $Trim)
    $items = @()
    if ($Disks | Where-Object { $_.Estado -eq "REEMPLAZAR" }) { $items += [PSCustomObject]@{Prioridad="URGENTE";Accion="Reemplazar disco";Motivo="SMART indica falla inminente - riesgo de perder datos"} }
    if ($Disks | Where-Object { $_.Estado -eq "MAL" }) { $items += [PSCustomObject]@{Prioridad="ALTA";Accion="Revisar disco";Motivo="Windows reporta problema de salud en el disco"} }
    if ($Disks | Where-Object { $_.Tipo -match "HDD" }) { $items += [PSCustomObject]@{Prioridad="MEDIA";Accion="Migrar a SSD";Motivo="El sistema corre en disco mecanico - un SSD cambia radicalmente la velocidad"} }
    if ($Vols | Where-Object { $_.Alerta -eq "CRITICO" }) { $items += [PSCustomObject]@{Prioridad="ALTA";Accion="Liberar espacio urgente";Motivo="Menos del 5% libre - puede generar errores del sistema"} }
    elseif ($Vols | Where-Object { $_.Alerta -eq "ALTO" }) { $items += [PSCustomObject]@{Prioridad="MEDIA";Accion="Liberar espacio";Motivo="Poco espacio libre afecta rendimiento y actualizaciones"} }
    try { if ($Perf.RAM_Pct -ne "N/D" -and [double]$Perf.RAM_Pct -ge 85) { $items += [PSCustomObject]@{Prioridad="MEDIA";Accion="Ampliar RAM o revisar consumo";Motivo="La RAM esta muy exigida en uso normal"} } } catch {}
    if ($HwAge.RAM_Estado -eq "INSUFICIENTE") { $items += [PSCustomObject]@{Prioridad="ALTA";Accion="Ampliar RAM a 16 GB minimo";Motivo="Menos de 8 GB es critico para el uso moderno"} }
    elseif ($HwAge.RAM_Estado -eq "JUSTA") { $items += [PSCustomObject]@{Prioridad="MEDIA";Accion="Ampliar RAM a 16 GB";Motivo="8 GB alcanza justo - 16 GB da mucho mas fluidez"} }
    if ($HwAge.Equipo_Estado -eq "PLATAFORMA VIEJA") { $items += [PSCustomObject]@{Prioridad="MEDIA";Accion="Evaluar cambio de plataforma";Motivo=$HwAge.Equipo_Msg} }
    if ($Integrity.Estado -match "CORRUPCION") { $items += [PSCustomObject]@{Prioridad="ALTA";Accion="Correr SFC y DISM";Motivo="Se detectaron archivos del sistema danados"} }
    if ($Def.Activo -eq $false -or $Def.TiempoReal -eq $false) { $items += [PSCustomObject]@{Prioridad="ALTA";Accion="Activar Windows Defender";Motivo="La proteccion en tiempo real no esta activa"} }
    if ($Trim.TRIM -eq "DESACTIVADO") { $items += [PSCustomObject]@{Prioridad="BAJA";Accion="Activar TRIM";Motivo="TRIM desactivado puede reducir vida util del SSD"} }
    if (-not $items) { $items += [PSCustomObject]@{Prioridad="NINGUNA";Accion="Sin acciones urgentes";Motivo="El equipo esta en buen estado"} }
    $items | Sort-Object { @{"URGENTE"=0;"ALTA"=1;"MEDIA"=2;"BAJA"=3;"NINGUNA"=4}[$_.Prioridad] }
}

function Set-MaintenanceTask { param([int]$Meses)
    $name = "PCLAF Mantenimiento Preventivo"
    $cmd = 'powershell.exe -WindowStyle Hidden -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show(''Recordatorio PCLAF: Ya es momento de tu mantenimiento preventivo. Contactanos para agendar.'',''PCLAF'')"'
    try {
        schtasks /Delete /TN "$name" /F 2>$null | Out-Null
        schtasks /Create /SC MONTHLY /MO $Meses /TN "$name" /TR "$cmd" /ST "10:00" /RL HIGHEST /F 2>$null | Out-Null
        return [PSCustomObject]@{ Tarea=$name; Frecuencia="Cada $Meses meses"; Estado="Creada" }
    } catch {
        return [PSCustomObject]@{ Tarea=$name; Frecuencia="Cada $Meses meses"; Estado="No se pudo crear" }
    }
}

# --- HTML GENERATION --------------------------------------------------------

$CSS = @'
<style>
:root{--bg:#050505;--panel:#0b0b0b;--panel2:#111;--line:#222;--line2:#2a2a2a;--txt:#f0f0f0;--muted:#aaa;--red:#e10600;--red2:#ff2a1f;--ok:#4ade80;--warn:#facc15;--bad:#f87171;--blue:#60a5fa;--radius:14px}
*{box-sizing:border-box;margin:0;padding:0}
html,body{font-family:Arial,Helvetica,sans-serif;background:radial-gradient(circle at top right,rgba(225,6,0,.10),transparent 30%),linear-gradient(180deg,#020202,#070707);color:var(--txt);line-height:1.5}
.wrap{max-width:1440px;margin:0 auto;padding:24px}

/* HERO */
.hero{background:linear-gradient(135deg,rgba(225,6,0,.12) 0%,rgba(255,255,255,.02) 40%),var(--panel);border:1px solid var(--line2);border-top:3px solid var(--red);border-radius:20px;padding:28px;box-shadow:0 8px 40px rgba(0,0,0,.5);margin-bottom:28px}
.brand{display:flex;align-items:center;gap:18px;margin-bottom:20px}
.brand-logo{width:80px;height:80px;border-radius:16px;object-fit:cover}
.brand-title{font-size:38px;font-weight:900;letter-spacing:1px}
.brand-title .r{color:var(--red2)}.brand-title .w{color:#fff}
.brand-sub{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:2px;margin-top:4px}
.hero h1{font-size:28px;font-weight:900;margin-bottom:6px}
.hero .sub{color:var(--muted);font-size:14px;margin-bottom:20px}

/* SEMAFORO */
.traffic{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:20px}
.tl{display:flex;align-items:center;gap:10px;background:var(--panel2);border:1px solid var(--line2);border-radius:10px;padding:10px 14px;flex:1;min-width:160px}
.tl-dot{width:14px;height:14px;border-radius:50%;flex-shrink:0}
.tl-dot.ok{background:var(--ok);box-shadow:0 0 8px rgba(74,222,128,.6)}
.tl-dot.warn{background:var(--warn);box-shadow:0 0 8px rgba(250,204,21,.6)}
.tl-dot.bad{background:var(--bad);box-shadow:0 0 8px rgba(248,113,113,.6)}
.tl-label{font-size:12px;color:var(--muted)}
.tl-value{font-size:13px;font-weight:700;color:#fff}

/* KPIs */
.kgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px;margin-bottom:20px}
.kpi{background:var(--panel2);border:1px solid var(--line2);border-radius:12px;padding:16px}
.kpi-l{font-size:11px;text-transform:uppercase;color:var(--muted);letter-spacing:.08em;margin-bottom:6px}
.kpi-v{font-size:20px;font-weight:900}

/* BANNER */
.banner{display:inline-flex;align-items:center;gap:8px;padding:10px 18px;border-radius:10px;font-weight:900;font-size:15px;letter-spacing:.5px;margin-bottom:16px}
.banner.ok{background:rgba(74,222,128,.12);border:1px solid rgba(74,222,128,.4);color:var(--ok)}
.banner.warn{background:rgba(250,204,21,.12);border:1px solid rgba(250,204,21,.4);color:var(--warn)}
.banner.bad{background:rgba(248,113,113,.12);border:1px solid rgba(248,113,113,.4);color:var(--bad)}

/* BARRA DE RAM */
.bar-wrap{background:#1a1a1a;border-radius:8px;overflow:hidden;height:18px;margin:8px 0}
.bar-fill{height:100%;border-radius:8px;transition:width .3s}
.bar-fill.ok{background:var(--ok)}.bar-fill.warn{background:var(--warn)}.bar-fill.bad{background:var(--bad)}

/* SECTIONS */
section{margin-bottom:28px}
h2{font-size:18px;font-weight:800;color:#fff;border-left:4px solid var(--red);padding:8px 14px;background:linear-gradient(90deg,rgba(225,6,0,.15),rgba(225,6,0,.03));border-radius:0 8px 8px 0;margin-bottom:14px}
.section-sub{font-size:12px;color:var(--muted);margin:-10px 0 14px 4px}

/* CARDS */
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:14px}
.card{background:var(--panel);border:1px solid var(--line2);border-radius:14px;padding:18px}
.card-icon{font-size:28px;margin-bottom:8px}
.card-title{font-size:14px;font-weight:700;color:#fff;margin-bottom:4px}
.card-body{font-size:13px;color:var(--muted);line-height:1.6}
.card.ok{border-color:rgba(74,222,128,.3);background:rgba(74,222,128,.04)}
.card.warn{border-color:rgba(250,204,21,.3);background:rgba(250,204,21,.04)}
.card.bad{border-color:rgba(248,113,113,.3);background:rgba(248,113,113,.04)}

/* TABLES */
.tw{width:100%;overflow-x:auto;border-radius:var(--radius);border:1px solid var(--line2);margin-bottom:18px}
table{width:100%;border-collapse:collapse;background:var(--panel);table-layout:auto}
th{background:#141414;color:#fff;padding:10px 14px;text-align:left;font-size:12px;text-transform:uppercase;letter-spacing:.06em;border-bottom:1px solid var(--line2);white-space:nowrap}
td{padding:9px 14px;border-bottom:1px solid var(--line);font-size:13px;vertical-align:top;word-break:break-word}
tr:last-child td{border-bottom:none}
tr:hover td{background:rgba(255,255,255,.025)}

/* TAGS */
.tag{display:inline-block;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700;letter-spacing:.04em}
.tag.ok{background:rgba(74,222,128,.15);color:var(--ok)}
.tag.warn{background:rgba(250,204,21,.15);color:var(--warn)}
.tag.bad{background:rgba(248,113,113,.15);color:var(--bad)}
.tag.info{background:rgba(96,165,250,.15);color:var(--blue)}

/* PRIO */
.prio-urgente{color:var(--bad);font-weight:900}
.prio-alta{color:#fb923c;font-weight:700}
.prio-media{color:var(--warn);font-weight:600}
.prio-baja{color:var(--ok)}
.prio-ninguna{color:var(--muted)}

/* RESUMEN CLIENTE */
.resumen-box{background:var(--panel2);border:1px solid var(--line2);border-radius:14px;padding:22px;line-height:1.8;font-size:15px;color:var(--txt)}
.resumen-box strong{color:#fff}
.resumen-box .highlight{color:var(--ok);font-weight:700}
.resumen-box .alert{color:var(--bad);font-weight:700}
.resumen-box .note{color:var(--warn);font-weight:700}

/* TRABAJO HECHO */
.work-box{background:rgba(96,165,250,.06);border:1px solid rgba(96,165,250,.25);border-radius:14px;padding:20px}
.work-title{font-size:15px;font-weight:700;color:var(--blue);margin-bottom:10px}
.work-body{font-size:14px;color:var(--muted);line-height:1.7}

/* PROXIMA REVISION */
.next-box{background:rgba(74,222,128,.06);border:1px solid rgba(74,222,128,.25);border-radius:14px;padding:20px;display:flex;align-items:center;gap:16px}
.next-date{font-size:24px;font-weight:900;color:var(--ok)}
.next-msg{font-size:13px;color:var(--muted);line-height:1.6}

/* FOOTER */
.footer{text-align:center;color:var(--muted);font-size:11px;margin-top:40px;padding-top:20px;border-top:1px solid var(--line)}

@media(max-width:768px){
  .wrap{padding:14px}
  .hero{padding:16px}
  .brand-title{font-size:28px}
  .hero h1{font-size:22px}
  .kgrid,.cards,.traffic{grid-template-columns:1fr}
  td,th{padding:8px 10px;font-size:12px}
}
</style>
'@

function Tag {
    param([string]$text, [string]$cls="info")
    return "<span class='tag $cls'>$text</span>"
}

function StatusTag {
    param([string]$text)
    $cls = switch -Regex ($text) {
        "BIEN|ACTUAL|NUEVO|OK|BUENA|RAPIDO|ACTIVO|SI|VIGENTE|EXCELENTE|SANA|SOBRADA" { "ok" }
        "USABLE|JUSTA|INTERMEDIA|MODERADO|ELEVADO|EQUILIBRADO" { "warn" }
        "MAL|REEMPLAZAR|CRITICO|ALTO|VIEJO|VIEJA|INSUFICIENTE|FALLO|PLATAFORMA VIEJA|REQUIERE|URGENTE|CORRUPCION|LENTO" { "bad" }
        default { "info" }
    }
    return "<span class='tag $cls'>$([System.Web.HttpUtility]::HtmlEncode($text))</span>"
}

function To-HtmlTable {
    param($Data, [string]$Caption = "")
    if (-not $Data) { return "<p style='color:var(--muted)'>Sin datos</p>" }
    $rows = @($Data)
    if ($rows.Count -eq 0) { return "<p style='color:var(--muted)'>Sin datos</p>" }
    $cols = $rows[0].PSObject.Properties.Name
    $html = "<div class='tw'><table>"
    if ($Caption) { $html += "<caption style='text-align:left;padding:8px 14px;color:var(--muted);font-size:12px'>$Caption</caption>" }
    $html += "<thead><tr>"
    foreach ($c in $cols) { $html += "<th>$c</th>" }
    $html += "</tr></thead><tbody>"
    foreach ($r in $rows) {
        $html += "<tr>"
        foreach ($c in $cols) {
            $v = [string]$r.$c
            $cell = [System.Web.HttpUtility]::HtmlEncode($v)
            # Apply status tags to specific columns
            if ($c -in @("Estado","Salud","SMART","Alerta","TRIM","Clase","Activo","TiempoReal","Antispyware","Detectada","FP_Coincide","Cambio_Disco","Cambio_RAM","Cambio_CPU","Cambio_GPU","Marca_Previa") -or
                $v -in @("BIEN","MAL","REEMPLAZAR","OK","CRITICO","ALTO","MODERADO","ACTIVO","DESACTIVADO","FALLO","SIN DATOS","NORMAL","ELEVADO")) {
                $cell = StatusTag $v
            }
            if ($c -eq "Prioridad") {
                $cls = $v.ToLower() -replace " ","_"
                $cell = "<span class='prio-$cls'>$([System.Web.HttpUtility]::HtmlEncode($v))</span>"
            }
            $html += "<td>$cell</td>"
        }
        $html += "</tr>"
    }
    $html += "</tbody></table></div>"
    return $html
}

function Get-TrafficLight {
    param([string]$Label, [string]$Value, [string]$Estado)
    $cls = switch -Regex ($Estado) {
        "ok|BIEN|EXCELENTE|NORMAL|ACTIVO|SANA|VIGENTE|BUENA|RAPIDO|SOBRADA" { "ok" }
        "warn|USABLE|JUSTA|MODERADO|ELEVADO|INTERMEDIA|EQUILIBRADO|CON OBSERVACIONES" { "warn" }
        default { "bad" }
    }
    return "<div class='tl'><div class='tl-dot $cls'></div><div><div class='tl-label'>$Label</div><div class='tl-value'>$($Value -replace "&","&amp;" -replace "<","&lt;" -replace ">","&gt;" -replace '"','&quot;')</div></div></div>"
}

function Get-RamBar {
    param([double]$Pct)
    $cls = if ($Pct -ge 85) { "bad" } elseif ($Pct -ge 65) { "warn" } else { "ok" }
    $w   = [math]::Min($Pct, 100)
    return "<div class='bar-wrap'><div class='bar-fill $cls' style='width:$($w)%'></div></div>"
}

function Get-ClientSummary {
    param($Status, $HwAge, $Vols, $Disks, $Perf, $Temps, $Rec, $NextDate)
    $estado = $Status.EstadoGeneral
    $estadoCls = if ($estado -match "EXCELENTE") { "ok" } elseif ($estado -match "OBSERVACIONES|REVISION") { "warn" } else { "bad" }
    $emoji = if ($estado -match "EXCELENTE") { "&#9989;" } elseif ($estado -match "OBSERVACIONES") { "&#9888;" } else { "&#128308;" }

    $diskMsg = ""
    $badDisk = $Disks | Where-Object { $_.Estado -ne "BIEN" -and $_.Estado -ne "SIN DATOS" }
    if ($badDisk) { $diskMsg = "&#9888; Se detecto un problema en el almacenamiento." }
    else { $diskMsg = "&#9989; Los discos estan en buen estado." }

    $ramPctVal = 0; try { $ramPctVal = [double]$Perf.RAM_Pct } catch {}
    $ramMsg = if ($ramPctVal -ge 85) { "&#9888; La memoria RAM esta muy exigida." } else { "&#9989; La memoria RAM opera con normalidad." }

    $tempMsg = ""
    $hotZone = $Temps | Where-Object { $_.Estado -in @("ALTO","CRITICO") }
    if ($hotZone) { $tempMsg = "&#127777; Se detectaron temperaturas elevadas." }
    else { $tempMsg = "&#127777; Las temperaturas son normales." }

    $spaceMsg = ""
    $tightVol = $Vols | Where-Object { $_.Alerta -in @("ALTO","CRITICO") }
    if ($tightVol) { $spaceMsg = "&#128190; Poco espacio en disco. Conviene liberar archivos." }
    else { $spaceMsg = "&#128190; El espacio en disco esta bien." }

    $urgentRec = ($Rec | Where-Object { $_.Prioridad -in @("URGENTE","ALTA") })
    $recMsg = if ($urgentRec) { "Se identificaron " + ($urgentRec | Measure-Object).Count + " punto(s) que requieren atencion." } else { "No hay acciones urgentes pendientes." }

    return @"
<div class='resumen-box'>
<p>$emoji <strong>Estado general del equipo: <span class='$estadoCls'>$estado</span></strong></p>
<br>
<p>$diskMsg</p>
<p>$ramMsg</p>
<p>$tempMsg</p>
<p>$spaceMsg</p>
<br>
<p><strong>&#128295; Evaluacion del hardware:</strong> $(HtmlEnc $HwAge.Equipo_Msg)</p>
<br>
<p><strong>&#128203; Recomendaciones:</strong> $recMsg</p>
<br>
<p><strong>&#128197; Proxima revision recomendada:</strong> <span class='highlight'>$NextDate</span></p>
</div>
"@
}

function HtmlEnc { param($s) $t=[string]$s; return $t -replace "&","&amp;" -replace "<","&lt;" -replace ">","&gt;" -replace '"','&quot;' }

# --- EJECUCION ---------------------------------------------------------------

Update-Stage 5 "Relevando sistema operativo"
$osInfo     = Get-OsInfo

Update-Stage 12 "Relevando hardware (CPU, RAM, MB)"
$sysInfo    = Get-SystemSummary
$gpuInfo    = Get-GpuInfo
$ramInfo    = Get-RamDetail

Update-Stage 22 "Leyendo temperaturas"
$tempInfo   = Get-Temperatures

Update-Stage 30 "Analizando discos (SMART, salud)"
$diskInfo   = Get-DiskHealth
$volInfo    = Get-VolumeUsage
$trimInfo   = Get-TrimStatus

Update-Stage 40 "Revisando red y seguridad"
$netInfo    = Get-NetworkInfo
$secInfo    = Get-SecurityInfo
$defInfo    = Get-DefenderStatus
$batInfo    = Get-BatteryInfo
$wuInfo     = Get-WindowsUpdateStatus
$activInfo  = Get-WindowsActivation
$powerProf  = Get-PowerProfile

Update-Stage 50 "Analizando rendimiento actual"
$perfInfo   = Get-CurrentPerformance
$topProc    = Get-TopProcesses

Update-Stage 58 "Revisando arranque y eventos de Windows"
$bootInfo   = Get-BootTime
$critEvts   = Get-CriticalEvents
$memErrs    = Get-MemoryErrors
$integ      = Get-IntegrityStatus

Update-Stage 68 "Relevando software instalado y configuracion"
$startApps  = Get-StartupApps
$instApps   = Get-InstalledApps
$openPorts  = Get-OpenPorts
$nonMsSvcs  = Get-NonMsServices
$schedTasks = Get-CustomScheduledTasks
$driversInfo= Get-DriversInfo

Update-Stage 78 "Leyendo marca previa y comparando hardware"
$pclafPaths = Get-PCLAFPaths
$prevRec    = Read-PreviousRecord -P $pclafPaths
$fingerprint= Get-HardwareFingerprint -Sys $sysInfo -Disks $diskInfo
$hwAge      = Get-HardwareAge -Sys $sysInfo -Gpu $gpuInfo -Disks $diskInfo -Ram $ramInfo
$comparison = Compare-Records -Prev $prevRec -FP $fingerprint -Sys $sysInfo -Disks $diskInfo -Gpu $gpuInfo

Update-Stage 86 "Calculando estado final y recomendaciones"
$finalStatus= Get-FinalAssessment -Disks $diskInfo -Vols $volInfo -Events $critEvts -HwAge $hwAge -Perf $perfInfo -Defender $defInfo
$recs       = Get-Recommendations -Disks $diskInfo -Vols $volInfo -Events $critEvts -Perf $perfInfo -HwAge $hwAge -Integrity $integ -Def $defInfo -Trim $trimInfo
$maintTask  = Set-MaintenanceTask -Meses $MesesMantenimiento

Update-Stage 92 "Guardando marca PCLAF en el equipo"
$record = [PSCustomObject]@{
    Metadata    = [PSCustomObject]@{ Fecha=(Get-Date -Format "yyyy-MM-ddTHH:mm:ss"); Version=$ScriptVersion; Equipo=$env:COMPUTERNAME; Tecnico=$Tecnico; Modo=$Modo; PCLAF_OS=$(if($SistemaInstaladoPorPCLAF){"SI"}else{"NO"}); MesesMant=$MesesMantenimiento }
    FinalStatus = $finalStatus; Fingerprint=$fingerprint; Comparacion=$comparison
    SystemInfo=$osInfo; SystemSummary=$sysInfo; GPU=$gpuInfo; RAM=$ramInfo
    Discos=$diskInfo; Volumenes=$volInfo; Rendimiento=$perfInfo
    Temperaturas=$tempInfo; Seguridad=$secInfo; Defender=$defInfo
    HardwareAge=$hwAge; Recomendaciones=$recs
}
$markOk  = Set-RegistryMark -FP $fingerprint -Status $finalStatus -Sys $sysInfo -Disks $diskInfo
$writeOk = Write-Record -P $pclafPaths -Rec $record

# --- BUILD HTML -------------------------------------------------------------

Update-Stage 96 "Generando reporte HTML $Modo"

$nextDate   = (Get-Date).AddMonths($MesesMantenimiento).ToString("MM/yyyy")
$estadoCls  = if ($finalStatus.EstadoGeneral -match "EXCELENTE") { "ok" } elseif ($finalStatus.EstadoGeneral -match "OBSERVACIONES|REVISION") { "warn" } else { "bad" }
$bannerEmoji= if ($finalStatus.EstadoGeneral -match "EXCELENTE") { "[OK]" } elseif ($finalStatus.EstadoGeneral -match "OBSERVACIONES") { "[!]" } else { "[!]" }

# Traffic lights
$tlHw    = Get-TrafficLight "Hardware"    $hwAge.Equipo_Estado  (if($hwAge.Equipo_Estado -match "VIGENTE|USABLE"){if($hwAge.Equipo_Estado -match "USABLE"){"warn"}else{"ok"}}else{"bad"})
$tlDisk  = Get-TrafficLight "Discos"      ($diskInfo|Select-Object -First 1).Estado (if(($diskInfo|Select-Object -First 1).Estado -eq "BIEN"){"ok"}else{"bad"})
$tlRam   = Get-TrafficLight "RAM"         "$($perfInfo.RAM_Pct)% usada" (if($perfInfo.RAM_Pct -ne "N/D" -and [double]$perfInfo.RAM_Pct -ge 85){"bad"}elseif($perfInfo.RAM_Pct -ne "N/D" -and [double]$perfInfo.RAM_Pct -ge 65){"warn"}else{"ok"})
$tlTemp  = Get-TrafficLight "Temperatura" (($tempInfo|Select-Object -First 1).Estado) (if(($tempInfo|Select-Object -First 1).Estado -in @("CRITICO","ALTO")){"bad"}elseif(($tempInfo|Select-Object -First 1).Estado -eq "ELEVADO"){"warn"}else{"ok"})
$tlDef   = Get-TrafficLight "Antivirus"   (if($defInfo.Activo -eq $true){"Activo"}else{"Inactivo"}) (if($defInfo.Activo -eq $true){"ok"}else{"bad"})
$tlSpace = Get-TrafficLight "Espacio"     (($volInfo|Sort-Object Uso_Pct -Descending|Select-Object -First 1).Alerta) (if(($volInfo|Sort-Object Uso_Pct -Descending|Select-Object -First 1).Alerta -in @("CRITICO","ALTO")){"bad"}elseif(($volInfo|Sort-Object Uso_Pct -Descending|Select-Object -First 1).Alerta -eq "MODERADO"){"warn"}else{"ok"})

# RAM bar
$ramBarPct = 0; try { $ramBarPct = [double]$perfInfo.RAM_Pct } catch {}
$ramBar = Get-RamBar -Pct $ramBarPct

# Logo base64 (compact placeholder - replace with your actual logo)
$LogoB64 = "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAYAAADimHc4AAAACXBIWXMAAA7EAAAOxAGVKw4bAAADFUlEQVR4nO2cS27CMBCGZwMcgtNwCM7S03AWTsNNuAqH6KJISCiEkIQmJI+dxHbi2GM7TtL/S1aRKPb4+8fj8UQUBKIYAAAAAAAAAAAAAAAAAAAB8C7UpX0qAAAAAABJRU5ErkJggg=="

$html = @"
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Reporte PCLAF - $($env:COMPUTERNAME) - $Modo</title>
$CSS
</head>
<body>
<div class="wrap">

<!-- HERO -->
<div class="hero">
  <div class="brand">
    <img class="brand-logo" src="data:image/png;base64,$LogoB64" alt="PCLAF">
    <div>
      <div class="brand-title"><span class="w">PC</span><span class="r">LAF</span></div>
      <div class="brand-sub">Servicio Tecnico - Floresta, CABA - pclaf.com.ar</div>
    </div>
  </div>
  <h1>Diagnostico del equipo: $($env:COMPUTERNAME)</h1>
  <div class="sub">Fecha: $(Get-Date -Format "dd/MM/yyyy HH:mm") - Tecnico: $Tecnico - Modo: $($Modo.ToUpper())</div>

  <!-- BANNER ESTADO -->
  <div class="banner $estadoCls">$bannerEmoji $($finalStatus.EstadoGeneral)</div>
  <p style="color:var(--muted);font-size:13px;margin-bottom:20px">$($finalStatus.Motivos)</p>

  <!-- SEMAFORO -->
  <div class="traffic">
    $tlHw $tlDisk $tlRam $tlTemp $tlDef $tlSpace
  </div>

  <!-- KPIs -->
  <div class="kgrid">
    <div class="kpi"><div class="kpi-l">Sistema</div><div class="kpi-v" style="font-size:14px">$(HtmlEnc $osInfo.SO)</div></div>
    <div class="kpi"><div class="kpi-l">CPU</div><div class="kpi-v" style="font-size:13px">$(HtmlEnc $sysInfo.CPU)</div></div>
    <div class="kpi"><div class="kpi-l">RAM</div><div class="kpi-v">$($sysInfo.RAM_Total_GB) GB</div></div>
    <div class="kpi"><div class="kpi-l">Uso RAM actual</div><div class="kpi-v">$($perfInfo.RAM_Pct)%$ramBar</div></div>
    <div class="kpi"><div class="kpi-l">Disco principal</div><div class="kpi-v" style="font-size:13px">$(HtmlEnc (($diskInfo|Select-Object -First 1).Modelo))</div></div>
    <div class="kpi"><div class="kpi-l">Temp. CPU</div><div class="kpi-v">$(if(($tempInfo|Select-Object -First 1).Celsius -ne "N/D"){"$(($tempInfo|Select-Object -First 1).Celsius) C"}else{"Sin sensor"})</div></div>
  </div>
</div>

"@

# -- CLIENTE SECTION ----------------------------------------------------------
if ($ServicioRealizado -or $PrecioServicio) {
    $html += @"
<section>
<h2>[S] Trabajo realizado hoy</h2>
<div class="work-box">
  <div class="work-title">Servicio realizado por PCLAF</div>
  <div class="work-body">
    $(if($ServicioRealizado){"<p><strong>Descripcion:</strong> $(HtmlEnc $ServicioRealizado)</p>"})
    $(if($PrecioServicio){"<p><strong>Precio:</strong> $PrecioServicio</p>"})
    <p><strong>Fecha:</strong> $(Get-Date -Format "dd/MM/yyyy")</p>
    <p><strong>Tecnico:</strong> $Tecnico</p>
  </div>
</div>
</section>
"@
}

$html += @"
<section>
<h2>&#128203; Resumen del estado del equipo</h2>
$(Get-ClientSummary -Status $finalStatus -HwAge $hwAge -Vols $volInfo -Disks $diskInfo -Perf $perfInfo -Temps $tempInfo -Rec $recs -NextDate $nextDate)
</section>

<section>
<h2>&#9989; Recomendaciones</h2>
$(To-HtmlTable $recs)
</section>

<section>
<h2>&#127777; Temperaturas del sistema</h2>
<div class="section-sub">Temperaturas reportadas por los sensores internos del equipo</div>
$(To-HtmlTable $tempInfo)
</section>

<section>
<h2>&#128187; Hardware del equipo</h2>
<div class="cards">
  <div class="card $(if($hwAge.CPU_Estado -in @("VIEJO")){"bad"}elseif($hwAge.CPU_Estado -eq "USABLE"){"warn"}else{"ok"})">
    <div class="card-icon">&#128306;</div>
    <div class="card-title">Procesador (CPU)</div>
    <div class="card-body">$(HtmlEnc $sysInfo.CPU)<br><strong>$(if($hwAge.CPU_Estado){HtmlEnc $hwAge.CPU_Estado}else{"N/D"})</strong> - $(if($hwAge.CPU_Msg){HtmlEnc $hwAge.CPU_Msg}else{"Sin datos"})<br>Nucleos: $($sysInfo.Cores) - Hilos: $($sysInfo.Hilos)</div>
  </div>
  <div class="card $(if($hwAge.RAM_Estado -eq "INSUFICIENTE"){"bad"}elseif($hwAge.RAM_Estado -eq "JUSTA"){"warn"}else{"ok"})">
    <div class="card-icon">&#129513;</div>
    <div class="card-title">Memoria RAM</div>
    <div class="card-body">$(HtmlEnc $sysInfo.RAM_Total_GB) GB instalados<br><strong>$(if($hwAge.RAM_Estado){HtmlEnc $hwAge.RAM_Estado}else{"N/D"})</strong> - $(if($hwAge.RAM_Msg){HtmlEnc $hwAge.RAM_Msg}else{"Sin datos"})</div>
  </div>
  <div class="card $(if($hwAge.Disco_Estado -match "LENTO"){"warn"}elseif($hwAge.Disco_Estado -eq "SIN DATOS"){"info"}else{"ok"})">
    <div class="card-icon">&#128190;</div>
    <div class="card-title">Almacenamiento</div>
    <div class="card-body"><strong>$(if($hwAge.Disco_Estado){HtmlEnc $hwAge.Disco_Estado}else{"N/D"})</strong> - $(if($hwAge.Disco_Msg){HtmlEnc $hwAge.Disco_Msg}else{"Sin datos"})</div>
  </div>
  <div class="card $(if($hwAge.GPU_Estado -eq "VIEJA"){"warn"}else{"ok"})">
    <div class="card-icon">&#127918;</div>
    <div class="card-title">Placa de video (GPU)</div>
    <div class="card-body">$(HtmlEnc (($gpuInfo|Select-Object -First 1).GPU))<br><strong>$(if($hwAge.GPU_Estado){HtmlEnc $hwAge.GPU_Estado}else{"N/D"})</strong> - $(if($hwAge.GPU_Msg){HtmlEnc $hwAge.GPU_Msg}else{"Sin datos"})</div>
  </div>
</div>
</section>

<section>
<h2>&#128190; Estado de los discos</h2>
$(To-HtmlTable $diskInfo)
</section>

<section>
<h2>&#128193; Espacio en unidades</h2>
$(To-HtmlTable $volInfo)
</section>

<section>
<h2>&#128197; Proxima revision</h2>
<div class="next-box">
  <div class="next-date">$nextDate</div>
  <div class="next-msg">
    <strong>Mantenimiento preventivo recomendado</strong><br>
    El mantenimiento regular prolonga la vida util del equipo y previene fallas.<br>
    Contactanos cuando sea el momento: <strong>11 4175-8129</strong>
  </div>
</div>
</section>

"@

# -- TECNICO-ONLY SECTIONS ----------------------------------------------------
if ($Modo -eq "tecnico") {
    $html += @"

<section>
<h2>[S] Analisis de hardware - Detalle tecnico</h2>
$(To-HtmlTable @($hwAge))
</section>

<section>
<h2>&#128187; Sistema operativo</h2>
$(To-HtmlTable @($osInfo))
</section>

<section>
<h2>&#128297; Resumen de hardware</h2>
$(To-HtmlTable @($sysInfo))
</section>

<section>
<h2>&#128250; Placas de video</h2>
$(To-HtmlTable $gpuInfo)
</section>

<section>
<h2>&#129513; Modulos de RAM</h2>
$(To-HtmlTable $ramInfo)
</section>

<section>
<h2>&#128273; Activacion de Windows</h2>
$(To-HtmlTable @($activInfo))
</section>

<section>
<h2>[!] Perfil de energia</h2>
$(To-HtmlTable @($powerProf))
</section>

<section>
<h2>&#128274; Seguridad basica</h2>
$(To-HtmlTable @($secInfo))
</section>

<section>
<h2>&#128737; Windows Defender</h2>
$(To-HtmlTable @($defInfo))
</section>

<section>
<h2>&#128267; Bateria</h2>
$(To-HtmlTable $batInfo)
</section>

<section>
<h2>&#127760; Red</h2>
$(To-HtmlTable $netInfo)
</section>

<section>
<h2>[WIN] Windows Update</h2>
$(To-HtmlTable @($wuInfo))
</section>

<section>
<h2>[S] TRIM de SSD</h2>
$(To-HtmlTable @($trimInfo))
</section>

<section>
<h2>[CPU] Integridad del sistema (SFC)</h2>
$(To-HtmlTable @($integ))
</section>

<section>
<h2>&#128678; Rendimiento actual</h2>
$(To-HtmlTable @($perfInfo))
</section>

<section>
<h2>&#9203; Historial de arranques</h2>
$(To-HtmlTable $bootInfo)
</section>

<section>
<h2>&#128202; Top procesos</h2>
$(To-HtmlTable $topProc)
</section>

<section>
<h2>&#128225; Puertos abiertos</h2>
$(To-HtmlTable $openPorts)
</section>

<section>
<h2>&#9881; Servicios de terceros</h2>
$(To-HtmlTable $nonMsSvcs)
</section>

<section>
<h2>&#128198; Tareas programadas</h2>
$(To-HtmlTable $schedTasks)
</section>

<section>
<h2>&#128640; Programas de inicio</h2>
$(To-HtmlTable $startApps)
</section>

<section>
<h2>&#128268; Drivers antiguos</h2>
<div class="section-sub">Ordenados por fecha - los mas antiguos primero</div>
$(To-HtmlTable $driversInfo)
</section>

<section>
<h2>[!] Errores de memoria RAM</h2>
$(To-HtmlTable $memErrs)
</section>

<section>
<h2>&#128680; Eventos criticos (30 dias)</h2>
$(To-HtmlTable $critEvts)
</section>

<section>
<h2>&#128270; Historial PCLAF - Comparativa</h2>
$(To-HtmlTable @($comparison))
</section>

<section>
<h2>&#128302; Fingerprint del equipo</h2>
<div class="section-sub">Hash unico basado en CPU, MB, RAM, disco y hostname. Detecta cambios de hardware.</div>
$(To-HtmlTable @($fingerprint))
</section>

<section>
<h2>[CD] Aplicaciones instaladas</h2>
$(To-HtmlTable $instApps)
</section>

<section>
<h2>&#9989; Marca PCLAF en el equipo</h2>
$(To-HtmlTable @([PSCustomObject]@{
  Registro_WIndows = if($markOk){"OK - HKCU:\SOFTWARE\PCLAF\Diagnostics"}else{"No se pudo escribir"}
  JSON_Local = if($writeOk){"OK - C:\ProgramData\PCLAF\last.json"}else{"No se pudo escribir"}
  Tecnico = $Tecnico
  SO_Instalado_PCLAF = if($SistemaInstaladoPorPCLAF){"SI"}else{"NO"}
  Prox_Revision = $nextDate
}))
</section>

"@
}

$html += @"

<div class="footer">
  PCLAF - Servicio tecnico - Campana 51, Floresta, CABA - 11 4175-8129 - pclaf.com.ar - @servicepclaf<br>
  Reporte generado el $(Get-Date -Format "dd/MM/yyyy HH:mm") por $Tecnico - Script v$ScriptVersion ($Modo)
</div>

</div><!-- /wrap -->
</body>
</html>
"@

# --- GUARDAR -----------------------------------------------------------------

$outFile = Join-Path $BasePath "Reporte_${Modo}_$($env:COMPUTERNAME)_${FechaReporte}.html"
$html | Set-Content -Path $outFile -Encoding UTF8

Update-Stage 100 "Listo!"
Write-Progress -Activity "PCLAF Diagnostico" -Completed
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  PCLAF Diagnostico v$ScriptVersion - $($Modo.ToUpper())" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Reporte : $outFile" -ForegroundColor Green
Write-Host "  Estado  : $($finalStatus.EstadoGeneral)" -ForegroundColor $(if($finalStatus.EstadoGeneral -match "EXCELENTE"){"Green"}elseif($finalStatus.EstadoGeneral -match "OBSERVACIONES"){"Yellow"}else{"Red"})
Write-Host "  Marca   : $(if($markOk){"Guardada en registro"}else{"No se pudo guardar"})" -ForegroundColor $(if($markOk){"Green"}else{"Yellow"})
Write-Host ""

# Abrir el reporte automaticamente
try { Start-Process $outFile } catch {}
