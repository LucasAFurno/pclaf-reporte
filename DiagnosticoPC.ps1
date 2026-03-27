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
    [string]$PrecioServicio = "",
    [string]$ReparacionId = "",
    [string]$SupabaseUrl = "",
    [string]$SupabaseAnonKey = "",
    [ValidateSet("manual","antes","despues","tecnico")]
    [string]$MomentoReporte = "manual",
    [switch]$SubirASupabase
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

function Get-NumericValueOrNull {
    param($Value)
    try { return [double]$Value } catch { return $null }
}

function Infer-DiskType {
    param([string]$Model)
    $m = [string]$Model
    if ($m -match 'SSD|NVME|NVMe|M\.2|WDS|WD Green|WD Blue SA|KINGSTON|SAMSUNG|CRUCIAL|ADATA|XPG|HYNIX|SK hynix|KBG|KXG|MZV|MZ7') { return "SSD" }
    if ($m -match 'HDD|TOSHIBA MQ|WD\d{2,}|ST\d{3,}|HITACHI|HGST') { return "HDD" }
    return "N/D"
}

function Infer-DiskSizeGb {
    param([string]$Model)
    $m = [string]$Model
    if ($m -match '([0-9]{3,4})\s?GB') { return [int]$matches[1] }
    if ($m -match '([0-9]{2,4})G(?!Hz|b)') { return [int]$matches[1] }
    if ($m -match '([1248])\s?TB') { return [int]$matches[1] * 1024 }
    return "N/D"
}

function Get-ReportSlotInfo {
    param([string]$ModoActual, [string]$MomentoActual)
    if ($ModoActual -eq "tecnico") {
        return [PSCustomObject]@{
            SlotName       = "tecnico"
            FileName       = "auto_tecnico.html"
            HtmlCliente    = $null
            DisplayName    = "Reporte tecnico"
        }
    }

    switch ($MomentoActual) {
        "antes" {
            return [PSCustomObject]@{
                SlotName       = "cliente_antes"
                FileName       = "auto_cliente_antes.html"
                HtmlCliente    = $null
                DisplayName    = "Reporte cliente antes"
            }
        }
        "despues" {
            return [PSCustomObject]@{
                SlotName       = "cliente_despues"
                FileName       = "auto_cliente_despues.html"
                HtmlCliente    = $null
                DisplayName    = "Reporte cliente despues"
            }
        }
        default {
            return [PSCustomObject]@{
                SlotName       = "cliente_manual"
                FileName       = "auto_cliente_manual.html"
                HtmlCliente    = $null
                DisplayName    = "Reporte cliente"
            }
        }
    }
}

function Sanitize-UploadText {
    param($Text)
    if ($null -eq $Text) { return $null }

    $value = [string]$Text
    $value = $value -replace [char]0, ''
    $value = $value -replace '\uFEFF', ''
    $value = $value -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', ''
    return $value
}

function Write-UploadLog {
    param(
        [string]$Stage,
        [string]$Message
    )

    try {
        $logFile = Join-Path $env:TEMP "PCLAF_Upload.log"
        $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Stage, $Message
        Add-Content -Path $logFile -Value $line -Encoding UTF8
    } catch {}
}

function Invoke-SupabaseReportUpload {
    param(
        [string]$BaseUrl,
        [string]$AnonKey,
        [string]$RepairId,
        [string]$HtmlFull,
        [string]$HtmlClient,
        [string]$FileName
    )

    if ([string]::IsNullOrWhiteSpace($BaseUrl) -or [string]::IsNullOrWhiteSpace($AnonKey) -or [string]::IsNullOrWhiteSpace($RepairId)) {
        return [PSCustomObject]@{ Ok=$false; Message="Faltan parametros de Supabase o reparacion" }
    }

    $repairGuid = [guid]::Empty
    if (-not [guid]::TryParse($RepairId, [ref]$repairGuid)) {
        return [PSCustomObject]@{ Ok=$false; Message="ReparacionId invalido: $RepairId" }
    }

    $base = $BaseUrl.TrimEnd('/')
    $headers = @{
        apikey        = $AnonKey
        Authorization = "Bearer $AnonKey"
        "Content-Type" = "application/json"
        Prefer        = "return=representation"
    }

    $htmlFullClean = Sanitize-UploadText $HtmlFull
    $htmlClientClean = Sanitize-UploadText $HtmlClient

    $payload = @{
        reparacion_id = $RepairId
        html_content  = $htmlFullClean
        html_cliente  = $htmlClientClean
        filename      = $FileName
    } | ConvertTo-Json -Depth 10

    $payloadBytes = 0
    try { $payloadBytes = [System.Text.Encoding]::UTF8.GetByteCount($payload) } catch {}
    Write-UploadLog -Stage "PREP" -Message ("Archivo={0} Bytes={1}" -f $FileName, $payloadBytes)

    try {
        $repairFilter = [uri]::EscapeDataString($RepairId)
        $fileFilter = [uri]::EscapeDataString($FileName)
        $existingUrl = "$base/rest/v1/reportes?reparacion_id=eq.$repairFilter&filename=eq.$fileFilter&select=id&limit=1"
        $existing = Invoke-RestMethod -Method Get -Uri $existingUrl -Headers $headers
        if ($existing -and $existing.Count -gt 0) {
            $id = $existing[0].id
            $patchUrl = "$base/rest/v1/reportes?id=eq.$id"
            Write-UploadLog -Stage "PATCH" -Message ("Actualizando reporte existente {0}" -f $id)
            $null = Invoke-RestMethod -Method Patch -Uri $patchUrl -Headers $headers -Body $payload
            Write-UploadLog -Stage "OK" -Message ("Reporte actualizado en Supabase: {0}" -f $id)
            return [PSCustomObject]@{ Ok=$true; Message="Reporte actualizado en Supabase"; Id=$id }
        }

        $postUrl = "$base/rest/v1/reportes"
        Write-UploadLog -Stage "POST" -Message ("Creando reporte nuevo {0}" -f $FileName)
        $created = Invoke-RestMethod -Method Post -Uri $postUrl -Headers $headers -Body $payload
        $newId = $null
        try { $newId = $created[0].id } catch {}
        Write-UploadLog -Stage "OK" -Message ("Reporte subido a Supabase: {0}" -f $newId)
        return [PSCustomObject]@{ Ok=$true; Message="Reporte subido a Supabase"; Id=$newId }
    } catch {
        $msg = $_.Exception.Message
        try {
            $resp = $_.Exception.Response
            if ($resp) {
                $reader = New-Object System.IO.StreamReader($resp.GetResponseStream())
                $bodyText = $reader.ReadToEnd()
                if (-not [string]::IsNullOrWhiteSpace($bodyText)) {
                    $msg = "$msg | $bodyText"
                }
            }
        } catch {}
        Write-UploadLog -Stage "ERROR" -Message $msg
        return [PSCustomObject]@{ Ok=$false; Message=$msg }
    }
}

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

    $biosReg = $null; $cpuReg = $null; $compInfo = $null
    try { $biosReg = Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' } catch {}
    try { $cpuReg = Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\CentralProcessor\0' } catch {}
    try {
        Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue
        $compInfo = New-Object Microsoft.VisualBasic.Devices.ComputerInfo
    } catch {}

    $ramGB = "N/D"
    try {
        $sum = ($ram | Measure-Object -Property Capacity -Sum).Sum
        if ($sum) { $ramGB = Round2 ($sum / 1GB) }
        elseif ($cs.TotalPhysicalMemory) { $ramGB = Round2 ($cs.TotalPhysicalMemory / 1GB) }
        elseif ($compInfo.TotalPhysicalMemory) { $ramGB = Round2 ($compInfo.TotalPhysicalMemory / 1GB) }
    } catch {}

    [PSCustomObject]@{
        Fabricante     = Safe $(if($cs.Manufacturer){$cs.Manufacturer}else{$biosReg.SystemManufacturer})
        Modelo         = Safe $(if($cs.Model){$cs.Model}else{$biosReg.SystemProductName})
        CPU            = Safe $(if($cpu.Name){($cpu.Name -replace '\s+', ' ').Trim()}else{$cpuReg.ProcessorNameString})
        Cores          = Safe $(if($cpu.NumberOfCores){$cpu.NumberOfCores}elseif($env:NUMBER_OF_PROCESSORS){[math]::Max(1,[int]([math]::Floor([int]$env:NUMBER_OF_PROCESSORS / 2)))}else{"N/D"})
        Hilos          = Safe $(if($cpu.NumberOfLogicalProcessors){$cpu.NumberOfLogicalProcessors}elseif($env:NUMBER_OF_PROCESSORS){$env:NUMBER_OF_PROCESSORS}else{"N/D"})
        RAM_Total_GB   = $ramGB
        Slots_RAM      = if ($ram) { ($ram | Measure-Object).Count } else { "N/D" }
        BIOS_Version   = Safe $(if($bios.SMBIOSBIOSVersion){($bios.SMBIOSBIOSVersion -join ", ")}else{$biosReg.BIOSVersion})
        BIOS_Fecha     = $(if($bios.ReleaseDate){To-DT $bios.ReleaseDate}else{Safe $biosReg.BIOSReleaseDate})
        Motherboard    = "$(Safe $(if($board.Manufacturer){$board.Manufacturer}else{$biosReg.BaseBoardManufacturer})) $(Safe $(if($board.Product){$board.Product}else{$biosReg.BaseBoardProduct}))".Trim()
        NroSerieEquipo = Safe $(if($bios.SerialNumber){$bios.SerialNumber}elseif($biosReg.SystemSerialNumber){$biosReg.SystemSerialNumber}else{"N/D"})
    }
}

function Get-GpuInfo {
    try {
        $gpus = Get-CimInstance Win32_VideoController | Where-Object { 
            $_.Name -notmatch "Remote|Virtual|Basic|Microsoft|Hyper-V" -or $_.AdapterRAM -gt 0
        }
        if (-not $gpus) { $gpus = Get-CimInstance Win32_VideoController | Select-Object -First 1 }
        if (-not $gpus) { throw "Sin datos CIM de GPU" }
        $gpus | ForEach-Object {
            [PSCustomObject]@{
                GPU       = Safe $_.Name
                VRAM_GB   = if ($_.AdapterRAM) { Round2 ($_.AdapterRAM / 1GB) } else { "N/D" }
                Driver    = Safe $_.DriverVersion
                Resolucion = if ($_.CurrentHorizontalResolution) { "$($_.CurrentHorizontalResolution)x$($_.CurrentVerticalResolution)" } else { "N/D" }
            }
        }
    } catch {
        try {
            $gpuRegs = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Video\*\0000' -ErrorAction SilentlyContinue |
                Where-Object { $_.DriverDesc }
            if ($gpuRegs) {
                return $gpuRegs | ForEach-Object {
                    $mem = $null
                    try { $mem = $_.'HardwareInformation.MemorySize' } catch {}
                    [PSCustomObject]@{
                        GPU        = Safe $_.DriverDesc
                        VRAM_GB    = if ($mem) { Round2 ($mem / 1GB) } else { "N/D" }
                        Driver     = Safe $_.DriverVersion
                        Resolucion = "N/D"
                    }
                } | Sort-Object GPU -Unique
            }
        } catch {}
        [PSCustomObject]@{ GPU="No detectada"; VRAM_GB="N/D"; Driver="N/D"; Resolucion="N/D" }
    }
}

function Get-RamDetail {
    try {
        $mems = Get-CimInstance Win32_PhysicalMemory
        if (-not $mems) { throw "Sin modulos CIM de RAM" }
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
        try {
            Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue
            $ci = New-Object Microsoft.VisualBasic.Devices.ComputerInfo
            return [PSCustomObject]@{
                Banco = "Memoria total"
                GB    = Round2 ($ci.TotalPhysicalMemory / 1GB)
                Tipo  = "N/D"
                MHz   = "N/D"
                Fab   = "N/D"
                Part  = "Detectada por .NET"
            }
        } catch {}
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
        if (-not $disks) { throw "Sin discos CIM" }

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
        try {
            $enum = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Services\disk\Enum' -ErrorAction SilentlyContinue
            $items = @()
            if ($enum) {
                for ($i = 0; $i -lt [int]$enum.Count; $i++) {
                    $instance = $enum."$i"
                    if (-not $instance) { continue }
                    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\$instance"
                    $d = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
                    if (-not $d) { continue }
                    $model = if ($d.FriendlyName) { $d.FriendlyName } else { ($instance -split '\\')[1] -replace '&',' ' }
                    $dtype = Infer-DiskType $model
                    $sizeGuess = Infer-DiskSizeGb $model
                    $items += [PSCustomObject]@{
                        Disco   = "DRIVE$i"
                        Modelo  = Safe $model
                        Serial  = "N/D"
                        Tipo    = Safe $dtype
                        GB      = $sizeGuess
                        Estado  = "SIN DATOS"
                        Salud   = "Desconocido"
                        SMART   = "N/D"
                        Detalle = "Detectado por registro de Windows"
                    }
                }
            }
            if ($items) { return $items }
        } catch {}
        [PSCustomObject]@{ Disco="Sin discos"; Modelo="N/D"; Serial="N/D"; Tipo="N/D"; GB="N/D"; Estado="N/D"; Salud="N/D"; SMART="N/D"; Detalle="No se pudo consultar" }
    }
}

function Get-VolumeUsage {
    try {
        $vols = Get-Volume | Where-Object { $_.DriveLetter }
        if (-not $vols) { throw "Sin volumenes CIM" }
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
        try {
            $drives = [System.IO.DriveInfo]::GetDrives() | Where-Object { $_.DriveType -eq [System.IO.DriveType]::Fixed }
            if ($drives) {
                return $drives | ForEach-Object {
                    $total = if ($_.TotalSize) { Round2 ($_.TotalSize / 1GB) } else { "N/D" }
                    $free  = if ($_.AvailableFreeSpace -ge 0) { Round2 ($_.AvailableFreeSpace / 1GB) } else { "N/D" }
                    $usePct = if ($_.TotalSize -gt 0) { Round2 (((($_.TotalSize - $_.AvailableFreeSpace) / $_.TotalSize) * 100)) } else { "N/D" }
                    $alert  = if ($usePct -is [double] -or $usePct -is [int]) {
                        if ($usePct -ge 95) { "CRITICO" } elseif ($usePct -ge 85) { "ALTO" } elseif ($usePct -ge 70) { "MODERADO" } else { "OK" }
                    } else { "N/D" }
                    [PSCustomObject]@{
                        Unidad   = $_.Name
                        Etiqueta = Safe $_.VolumeLabel
                        FS       = Safe $_.DriveFormat
                        GB_Total = $total
                        GB_Libre = $free
                        Uso_Pct  = $usePct
                        Alerta   = $alert
                    }
                }
            }
        } catch {}
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

function Get-BsodEvents {
    $start = (Get-Date).AddDays(-60)
    $results = @()

    $queries = @(
        @{ Log='System'; Provider='BugCheck'; Ids=@(1001) }
        @{ Log='System'; Provider='Microsoft-Windows-WER-SystemErrorReporting'; Ids=@(1001) }
    )

    foreach ($q in $queries) {
        foreach ($id in $q.Ids) {
            try {
                $events = Get-WinEvent -FilterHashtable @{
                    LogName      = $q.Log
                    ProviderName = $q.Provider
                    Id           = $id
                    StartTime    = $start
                } -ErrorAction SilentlyContinue

                foreach ($ev in $events) {
                    $msg = [string]$ev.Message
                    $code = "N/D"
                    if ($msg -match 'bugcheck (was|code is)\s*[: ]?\s*(0x[0-9A-Fa-f]+)') { $code = $matches[2].ToUpper() }
                    elseif ($msg -match '0x[0-9A-Fa-f]{8,16}') { $code = $matches[0].ToUpper() }

                    $results += [PSCustomObject]@{
                        Fecha    = $ev.TimeCreated
                        Origen   = $ev.ProviderName
                        Id       = $ev.Id
                        Codigo   = $code
                        Mensaje  = (($msg -split "`r?`n")[0]).Trim()
                    }
                }
            } catch {}
        }
    }

    if (-not $results) {
        return [PSCustomObject]@{ Fecha="Sin BSOD recientes"; Origen="N/D"; Id="N/D"; Codigo="N/D"; Mensaje="No se detectaron pantallazos azules en los ultimos 60 dias" }
    }

    $results | Sort-Object Fecha -Descending | Select-Object -Unique Fecha, Origen, Id, Codigo, Mensaje | Select-Object -First 10
}

function Get-UnexpectedShutdowns {
    $start = (Get-Date).AddDays(-30)
    $results = @()
    $queries = @(
        @{ Log='System'; Provider='Microsoft-Windows-Kernel-Power'; Ids=@(41) }
        @{ Log='System'; Provider='EventLog'; Ids=@(6008) }
    )

    foreach ($q in $queries) {
        foreach ($id in $q.Ids) {
            try {
                $events = Get-WinEvent -FilterHashtable @{
                    LogName      = $q.Log
                    ProviderName = $q.Provider
                    Id           = $id
                    StartTime    = $start
                } -ErrorAction SilentlyContinue
                foreach ($ev in $events) {
                    $results += [PSCustomObject]@{
                        Fecha   = $ev.TimeCreated
                        Origen  = $ev.ProviderName
                        Id      = $ev.Id
                        Mensaje = (([string]$ev.Message -split "`r?`n")[0]).Trim()
                    }
                }
            } catch {}
        }
    }

    if (-not $results) {
        return [PSCustomObject]@{ Fecha="Sin reinicios inesperados"; Origen="N/D"; Id="N/D"; Mensaje="No se detectaron apagados o reinicios bruscos en los ultimos 30 dias" }
    }

    $results | Sort-Object Fecha -Descending | Select-Object -First 10
}

function Get-WheaSummary {
    $start = (Get-Date).AddDays(-60)
    try {
        $events = Get-WinEvent -FilterHashtable @{
            LogName='System'; ProviderName='Microsoft-Windows-WHEA-Logger'; StartTime=$start
        } -ErrorAction SilentlyContinue

        if (-not $events) {
            return [PSCustomObject]@{ Fecha="Sin errores WHEA"; Id="N/D"; Severidad="N/D"; Mensaje="No se detectaron errores de hardware reportados por Windows" }
        }

        $events | Sort-Object TimeCreated -Descending | Select-Object -First 10 `
            @{N='Fecha';E={$_.TimeCreated}},
            @{N='Id';E={$_.Id}},
            @{N='Severidad';E={$_.LevelDisplayName}},
            @{N='Mensaje';E={ (([string]$_.Message -split "`r?`n")[0]).Trim() }}
    } catch {
        [PSCustomObject]@{ Fecha="No disponible"; Id="N/D"; Severidad="N/D"; Mensaje="No se pudo consultar WHEA" }
    }
}

function Get-DiskEventSummary {
    $start = (Get-Date).AddDays(-30)
    $providers = @('Disk','Ntfs','storahci','stornvme','iaStorA','iaStorAC')
    $rows = @()

    foreach ($provider in $providers) {
        try {
            $events = Get-WinEvent -FilterHashtable @{
                LogName='System'; ProviderName=$provider; StartTime=$start
            } -ErrorAction SilentlyContinue
            foreach ($ev in $events) {
                $rows += [PSCustomObject]@{
                    Fecha   = $ev.TimeCreated
                    Origen  = $ev.ProviderName
                    Id      = $ev.Id
                    Nivel   = $ev.LevelDisplayName
                    Mensaje = (([string]$ev.Message -split "`r?`n")[0]).Trim()
                }
            }
        } catch {}
    }

    if (-not $rows) {
        return [PSCustomObject]@{ Fecha="Sin errores de disco"; Origen="N/D"; Id="N/D"; Nivel="N/D"; Mensaje="No se detectaron errores de disco o controlador en los ultimos 30 dias" }
    }

    $rows | Sort-Object Fecha -Descending | Select-Object -First 12
}

function Get-ProblemDevices {
    try {
        $devices = Get-CimInstance Win32_PnPEntity | Where-Object {
            $_.ConfigManagerErrorCode -ne $null -and $_.ConfigManagerErrorCode -ne 0
        } | Select-Object `
            @{N='Dispositivo';E={ Safe $_.Name }},
            @{N='Clase';E={ Safe $_.PNPClass }},
            @{N='ErrorCode';E={ Safe $_.ConfigManagerErrorCode }},
            @{N='Fabricante';E={ Safe $_.Manufacturer }},
            @{N='PNPDeviceID';E={ Safe $_.PNPDeviceID }}

        if (-not $devices) {
            return [PSCustomObject]@{ Dispositivo="Sin dispositivos con error"; Clase="N/D"; ErrorCode="0"; Fabricante="N/D"; PNPDeviceID="N/D" }
        }

        $devices
    } catch {
        [PSCustomObject]@{ Dispositivo="No disponible"; Clase="N/D"; ErrorCode="N/D"; Fabricante="N/D"; PNPDeviceID="N/D" }
    }
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
    $cpu = "N/D"
    try {
        $cpuRaw = (Get-Counter '\Processor(_Total)\% Processor Time' -ErrorAction Stop).CounterSamples.CookedValue
        if ($null -ne $cpuRaw) { $cpu = [math]::Round([double]$cpuRaw, 1) }
    } catch {}

    $tot = $null
    $fre = $null

    try {
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $tot = Round2 ($os.TotalVisibleMemorySize / 1MB)
        $fre = Round2 ($os.FreePhysicalMemory / 1MB)
        if (($tot -as [double]) -le 0) { throw "Memoria CIM invalida" }
    } catch {
        try {
            Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue
            $ci = New-Object Microsoft.VisualBasic.Devices.ComputerInfo
            $tot = Round2 ($ci.TotalPhysicalMemory / 1GB)
            $fre = Round2 ($ci.AvailablePhysicalMemory / 1GB)
        } catch {}
    }

    try {
        $totVal = [double]$tot
        $freVal = [double]$fre
        if ($totVal -le 0) { throw "Sin memoria total" }
        $use = Round2 ($totVal - $freVal)
        $pct = Round2 (($use / $totVal) * 100)
        [PSCustomObject]@{ CPU_Pct=$cpu; RAM_Usada_GB=$use; RAM_Total_GB=$totVal; RAM_Pct=$pct }
    } catch {
        [PSCustomObject]@{ CPU_Pct=$cpu; RAM_Usada_GB="N/D"; RAM_Total_GB="N/D"; RAM_Pct="N/D" }
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
    param($Disks, $Vols, $Events, $HwAge, $Perf, $Defender, $Bsod, $Shutdowns, $Whea, $DiskEvents, $ProblemDevices)
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

    if ($Bsod | Where-Object { $_.Fecha -notmatch "^Sin BSOD recientes" }) {
        $estado = "REQUIERE REVISION"
        $motivos += "Pantallazos azules recientes"
    }

    if ($Shutdowns | Where-Object { $_.Fecha -notmatch "^Sin reinicios inesperados" }) {
        if ($estado -eq "EXCELENTE") { $estado="CON OBSERVACIONES" }
        $motivos += "Reinicios o apagados inesperados"
    }

    if ($Whea | Where-Object { $_.Fecha -notmatch "^Sin errores WHEA" -and $_.Fecha -notmatch "^No disponible" }) {
        $estado = "REQUIERE REVISION"
        $motivos += "Errores de hardware reportados por Windows"
    }

    if ($DiskEvents | Where-Object { $_.Fecha -notmatch "^Sin errores de disco" }) {
        if ($estado -eq "EXCELENTE") { $estado="CON OBSERVACIONES" }
        $motivos += "Errores recientes de disco o controlador"
    }

    if ($ProblemDevices | Where-Object { $_.ErrorCode -ne "0" -and $_.ErrorCode -ne "N/D" }) {
        if ($estado -eq "EXCELENTE") { $estado="CON OBSERVACIONES" }
        $motivos += "Dispositivos con error o sin driver correcto"
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
    param($Disks, $Vols, $Events, $Perf, $HwAge, $Integrity, $Def, $Trim, $Bsod, $Shutdowns, $Whea, $DiskEvents, $ProblemDevices)
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
    if ($Bsod | Where-Object { $_.Codigo -ne "N/D" -or $_.Mensaje -notmatch "No se detectaron" }) { $items += [PSCustomObject]@{Prioridad="ALTA";Accion="Revisar causa de pantallazos azules";Motivo="Windows registro errores graves tipo BSOD en los ultimos 60 dias"} }
    if ($Shutdowns | Where-Object { $_.Fecha -notmatch "^Sin reinicios inesperados" }) { $items += [PSCustomObject]@{Prioridad="ALTA";Accion="Diagnosticar reinicios o apagados bruscos";Motivo="Se detectaron apagados inesperados que pueden venir de hardware, energia o drivers"} }
    if ($Whea | Where-Object { $_.Fecha -notmatch "^Sin errores WHEA" -and $_.Fecha -notmatch "^No disponible" }) { $items += [PSCustomObject]@{Prioridad="ALTA";Accion="Revisar hardware por errores WHEA";Motivo="Windows reporto errores de hardware de bajo nivel"} }
    if ($DiskEvents | Where-Object { $_.Fecha -notmatch "^Sin errores de disco" }) { $items += [PSCustomObject]@{Prioridad="ALTA";Accion="Revisar eventos de disco y controlador";Motivo="Se detectaron errores de disco, NTFS o controlador de almacenamiento"} }
    if ($ProblemDevices | Where-Object { $_.ErrorCode -ne "0" -and $_.ErrorCode -ne "N/D" }) { $items += [PSCustomObject]@{Prioridad="MEDIA";Accion="Corregir dispositivos con error";Motivo="Hay dispositivos de Windows con controlador faltante o estado anormal"} }
    if (-not $items) { $items += [PSCustomObject]@{Prioridad="NINGUNA";Accion="Sin acciones urgentes";Motivo="El equipo esta en buen estado"} }
    $items | Sort-Object { @{"URGENTE"=0;"ALTA"=1;"MEDIA"=2;"BAJA"=3;"NINGUNA"=4}[$_.Prioridad] }, Accion -Unique
}

function Get-ClientIssues {
    param($Disks, $Vols, $Perf, $Temps, $Integrity, $Def, $HwAge, $Bsod, $Shutdowns, $Whea, $DiskEvents, $ProblemDevices)
    $items = @()

    if ($Disks | Where-Object { $_.Estado -eq "REEMPLAZAR" }) {
        $items += [PSCustomObject]@{ Prioridad="URGENTE"; Problema="Disco con riesgo de falla"; Impacto="Puede perder archivos o dejar de iniciar"; Accion="Hacer backup y reemplazar el disco cuanto antes" }
    } elseif ($Disks | Where-Object { $_.Estado -eq "MAL" }) {
        $items += [PSCustomObject]@{ Prioridad="ALTA"; Problema="Disco con problemas de salud"; Impacto="Puede causar lentitud, cuelgues o errores"; Accion="Revisar estado del disco y planificar reemplazo" }
    }

    if ($DiskEvents | Where-Object { $_.Fecha -notmatch "^Sin errores de disco" }) {
        $items += [PSCustomObject]@{ Prioridad="ALTA"; Problema="Errores de disco o controlador"; Impacto="Puede provocar congelamientos, archivos corruptos o fallas al arrancar"; Accion="Revisar almacenamiento, cables, controladores o disco del sistema" }
    }

    if ($Bsod | Where-Object { $_.Fecha -notmatch "^Sin BSOD recientes" }) {
        $items += [PSCustomObject]@{ Prioridad="ALTA"; Problema="Pantallazos azules recientes"; Impacto="El sistema tuvo fallos graves y puede seguir inestable"; Accion="Diagnosticar drivers, RAM, disco y hardware asociado" }
    }

    if ($Shutdowns | Where-Object { $_.Fecha -notmatch "^Sin reinicios inesperados" }) {
        $items += [PSCustomObject]@{ Prioridad="ALTA"; Problema="Reinicios o apagados inesperados"; Impacto="Puede haber problema de energia, temperatura, drivers o hardware"; Accion="Revisar estabilidad electrica y componentes criticos" }
    }

    if ($Whea | Where-Object { $_.Fecha -notmatch "^Sin errores WHEA" -and $_.Fecha -notmatch "^No disponible" }) {
        $items += [PSCustomObject]@{ Prioridad="ALTA"; Problema="Errores de hardware reportados por Windows"; Impacto="Puede indicar fallas en CPU, RAM, motherboard, GPU o PCIe"; Accion="Hacer diagnostico tecnico de hardware" }
    }

    if ($ProblemDevices | Where-Object { $_.ErrorCode -ne "0" -and $_.ErrorCode -ne "N/D" }) {
        $items += [PSCustomObject]@{ Prioridad="MEDIA"; Problema="Dispositivos con error o drivers faltantes"; Impacto="Puede haber funciones que no trabajen bien"; Accion="Instalar o corregir controladores del hardware afectado" }
    }

    if ($Vols | Where-Object { $_.Alerta -eq "CRITICO" }) {
        $items += [PSCustomObject]@{ Prioridad="ALTA"; Problema="Disco casi lleno"; Impacto="Windows puede fallar, trabarse o no actualizar"; Accion="Liberar espacio de forma urgente" }
    } elseif ($Vols | Where-Object { $_.Alerta -eq "ALTO" }) {
        $items += [PSCustomObject]@{ Prioridad="MEDIA"; Problema="Poco espacio disponible"; Impacto="Afecta rendimiento y mantenimiento del sistema"; Accion="Liberar archivos o ampliar almacenamiento" }
    }

    try {
        if ($Perf.RAM_Pct -ne "N/D" -and [double]$Perf.RAM_Pct -ge 85) {
            $items += [PSCustomObject]@{ Prioridad="MEDIA"; Problema="Memoria RAM muy exigida"; Impacto="Puede generar lentitud y cuelgues"; Accion="Revisar consumo o ampliar memoria" }
        }
    } catch {}

    if ($HwAge.RAM_Estado -eq "INSUFICIENTE") {
        $items += [PSCustomObject]@{ Prioridad="ALTA"; Problema="RAM insuficiente para uso actual"; Impacto="El equipo puede sentirse lento aun sin errores visibles"; Accion="Ampliar la memoria a 16 GB o mas" }
    }

    $hotZone = $Temps | Where-Object { $_.Estado -in @("ALTO","CRITICO") }
    if ($hotZone) {
        $items += [PSCustomObject]@{ Prioridad="ALTA"; Problema="Temperaturas elevadas"; Impacto="Puede causar apagados, ruido y menor vida util"; Accion="Hacer limpieza interna y revisar refrigeracion" }
    }

    if ($Integrity.Estado -match "CORRUPCION") {
        $items += [PSCustomObject]@{ Prioridad="ALTA"; Problema="Archivos del sistema danados"; Impacto="Puede generar errores de Windows y fallos raros"; Accion="Reparar Windows con herramientas del sistema" }
    }

    if ($Def.Activo -eq $false -or $Def.TiempoReal -eq $false) {
        $items += [PSCustomObject]@{ Prioridad="ALTA"; Problema="Proteccion antivirus inactiva"; Impacto="El equipo queda mas expuesto a amenazas"; Accion="Activar o reconfigurar la proteccion de Windows" }
    }

    if ($HwAge.Equipo_Estado -eq "PLATAFORMA VIEJA") {
        $items += [PSCustomObject]@{ Prioridad="MEDIA"; Problema="Hardware antiguo"; Impacto="Aunque funcione, puede limitar rendimiento y compatibilidad"; Accion="Evaluar mejora de plataforma o reemplazo de equipo" }
    }

    if (-not $items) {
        return [PSCustomObject]@{ Prioridad="NINGUNA"; Problema="Sin errores importantes detectados"; Impacto="El equipo no muestra fallas relevantes en esta revision"; Accion="Mantener controles preventivos periodicos" }
    }

    $items | Sort-Object { @{"URGENTE"=0;"ALTA"=1;"MEDIA"=2;"BAJA"=3;"NINGUNA"=4}[$_.Prioridad] }, Problema -Unique
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
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=JetBrains+Mono:wght@400;500;700&display=swap');
:root{
  --red:#CC0000;
  --red-dark:#a80000;
  --red-soft:rgba(204,0,0,.12);
  --red-soft-2:rgba(204,0,0,.18);
  --black:#0a0a0a;
  --black2:#111111;
  --black3:#1a1a1a;
  --black4:#222222;
  --border:#2e2e2e;
  --border2:#3a3a3a;
  --text:#f0f0f0;
  --text2:#b4b4b4;
  --text3:#666;
  --green:#22c55e;
  --green-dim:rgba(34,197,94,.12);
  --amber:#f59e0b;
  --amber-dim:rgba(245,158,11,.12);
  --blue:#3b82f6;
  --blue-dim:rgba(59,130,246,.12);
  --mono:'JetBrains Mono',monospace;
  --sans:'Inter',system-ui,sans-serif;
  --radius:18px;
  --shadow:0 24px 60px rgba(0,0,0,.35);
}
*{box-sizing:border-box;margin:0;padding:0}
html{scroll-behavior:smooth}
body{
  font-family:var(--sans);
  background:
    radial-gradient(circle at top right,rgba(204,0,0,.14),transparent 28%),
    radial-gradient(circle at bottom left,rgba(204,0,0,.08),transparent 26%),
    linear-gradient(180deg,#070707,#0d0d0d 30%,#090909 100%);
  color:var(--text);
  line-height:1.6;
  font-size:14px;
  min-height:100vh;
}
body::before{
  content:'';
  position:fixed;inset:0;pointer-events:none;z-index:0;
  background-image:
    linear-gradient(rgba(255,255,255,.016) 1px, transparent 1px),
    linear-gradient(90deg, rgba(255,255,255,.016) 1px, transparent 1px);
  background-size:32px 32px;
  mask-image:linear-gradient(180deg,rgba(0,0,0,.8),transparent 88%);
}
.wrap{max-width:1320px;margin:0 auto;padding:28px 18px 42px;position:relative;z-index:1}
.hero{
  background:
    linear-gradient(140deg,rgba(255,255,255,.03),rgba(255,255,255,.01)),
    linear-gradient(180deg,var(--black2),var(--black));
  border:1px solid var(--border2);
  border-top:3px solid var(--red);
  border-radius:28px;
  padding:26px;
  margin-bottom:24px;
  box-shadow:var(--shadow);
  overflow:hidden;
  position:relative;
}
.hero::after{
  content:'';
  position:absolute;
  top:-120px;right:-60px;
  width:320px;height:320px;border-radius:50%;
  background:radial-gradient(circle,rgba(204,0,0,.18),transparent 68%);
  pointer-events:none;
}
.hero-grid{
  display:grid;
  grid-template-columns:minmax(0,1.2fr) minmax(360px,.92fr);
  gap:20px;
  align-items:start;
}
.hero-main,.hero-side{
  min-width:0;
  position:relative;
  z-index:1;
}
.hero-main{
  display:flex;
  flex-direction:column;
  gap:14px;
}
.brand{
  display:flex;
  align-items:center;
  gap:14px;
  margin-bottom:6px;
}
.brand-logo{
  width:72px;
  height:72px;
  flex-shrink:0;
  border-radius:14px;
  border:1px solid var(--border2);
  background:linear-gradient(180deg,var(--black3),var(--black2));
  display:grid;
  place-items:center;
  overflow:hidden;
  box-shadow:0 16px 34px rgba(0,0,0,.3);
}
.brand-logo svg{
  width:100%;
  height:100%;
  display:block;
}
.brand-title{
  font-size:34px;
  font-weight:900;
  line-height:1;
  letter-spacing:.04em;
}
.brand-title .r{color:var(--red)}
.brand-sub{
  margin-top:6px;
  font-family:var(--mono);
  font-size:11px;
  color:var(--text3);
  text-transform:uppercase;
  letter-spacing:.08em;
}
.page-kicker{
  display:inline-flex;
  align-items:center;
  gap:8px;
  padding:6px 10px;
  border-radius:999px;
  border:1px solid var(--border2);
  background:rgba(255,255,255,.02);
  color:var(--text2);
  font-family:var(--mono);
  font-size:11px;
  margin-bottom:12px;
}
.page-kicker::before{
  content:'';
  width:7px;height:7px;border-radius:50%;
  background:var(--red);
  box-shadow:0 0 0 4px rgba(204,0,0,.18);
}
.hero h1{
  font-size:28px;
  font-weight:900;
  line-height:1.08;
  margin-bottom:8px;
  max-width:18ch;
}
.hero .sub{
  display:flex;
  flex-wrap:wrap;
  gap:8px;
  margin-bottom:16px;
}
.hero .sub span{
  display:inline-flex;
  align-items:center;
  gap:8px;
  padding:8px 11px;
  border-radius:999px;
  background:rgba(255,255,255,.03);
  border:1px solid var(--border);
  color:var(--text2);
  font-size:12px;
}
.hero-summary{
  display:grid;
  grid-template-columns:repeat(3,minmax(0,1fr));
  gap:10px;
  margin-top:6px;
}
.summary-chip{
  background:var(--black3);
  border:1px solid var(--border);
  border-radius:14px;
  padding:14px;
}
.summary-chip .k{
  color:var(--text3);
  font-size:10px;
  text-transform:uppercase;
  letter-spacing:.08em;
  font-family:var(--mono);
  margin-bottom:6px;
}
.summary-chip .v{
  color:#fff;
  font-size:18px;
  font-weight:800;
  line-height:1.1;
}
.hero-side{
  display:flex;
  flex-direction:column;
  gap:14px;
}
.status-card{
  background:linear-gradient(180deg,rgba(255,255,255,.03),rgba(255,255,255,.015)),var(--black2);
  border:1px solid var(--border2);
  border-radius:22px;
  padding:17px 18px;
}
.status-top{
  display:flex;
  align-items:center;
  justify-content:space-between;
  gap:10px;
  margin-bottom:10px;
}
.status-label{
  font-family:var(--mono);
  font-size:11px;
  color:var(--text3);
  text-transform:uppercase;
  letter-spacing:.08em;
}
.banner{
  display:inline-flex;
  align-items:center;
  gap:8px;
  padding:8px 14px;
  border-radius:999px;
  font-weight:800;
  font-size:11px;
  text-transform:uppercase;
  letter-spacing:.08em;
}
.banner::before{
  content:'';
  width:8px;height:8px;border-radius:50%;
  background:currentColor;
}
.banner.ok{background:var(--green-dim);border:1px solid rgba(34,197,94,.28);color:var(--green)}
.banner.warn{background:var(--amber-dim);border:1px solid rgba(245,158,11,.28);color:var(--amber)}
.banner.bad{background:var(--red-soft);border:1px solid rgba(204,0,0,.28);color:#ff7575}
.status-reason{
  color:var(--text2);
  font-size:13px;
  line-height:1.7;
  margin-bottom:14px;
}
.traffic{
  display:grid;
  grid-template-columns:repeat(2,minmax(0,1fr));
  gap:8px;
}
.tl{
  background:var(--black3);
  border:1px solid var(--border);
  border-radius:14px;
  padding:12px 13px;
  min-width:0;
}
.tl-head{
  display:flex;
  align-items:center;
  gap:8px;
  margin-bottom:5px;
}
.tl-dot{
  width:9px;height:9px;border-radius:50%;flex-shrink:0;
}
.tl-dot.ok{background:var(--green);box-shadow:0 0 0 4px rgba(34,197,94,.14)}
.tl-dot.warn{background:var(--amber);box-shadow:0 0 0 4px rgba(245,158,11,.14)}
.tl-dot.bad{background:var(--red);box-shadow:0 0 0 4px rgba(204,0,0,.18)}
.tl-label{
  color:var(--text3);
  font-family:var(--mono);
  font-size:10px;
  text-transform:uppercase;
  letter-spacing:.08em;
}
.tl-value{
  color:#fff;
  font-size:12px;
  font-weight:700;
  word-break:break-word;
}
.hero-panel{
  margin-top:18px;
  border-top:1px solid var(--border);
  padding-top:18px;
}
.section-accent{
  display:flex;
  align-items:end;
  justify-content:space-between;
  gap:12px;
  margin-bottom:12px;
}
.section-accent .sa-kicker{
  color:var(--red);
  font-size:11px;
  font-weight:800;
  letter-spacing:.08em;
  text-transform:uppercase;
  font-family:var(--mono);
}
.section-accent .sa-title{
  color:#fff;
  font-size:19px;
  font-weight:900;
  line-height:1.05;
}
.section-accent .sa-line{
  flex:1;
  height:2px;
  border-radius:999px;
  background:linear-gradient(90deg,var(--red),transparent);
  opacity:.8;
}
.meter-grid{
  display:grid;
  grid-template-columns:repeat(2,minmax(0,1fr));
  gap:12px;
}
.meter{
  background:var(--black3);
  border:1px solid var(--border);
  border-radius:16px;
  padding:15px;
  min-height:116px;
}
.meter-top{
  display:flex;
  align-items:end;
  justify-content:space-between;
  gap:8px;
  margin-bottom:12px;
}
.meter-label{
  color:var(--text3);
  font-family:var(--mono);
  font-size:10px;
  text-transform:uppercase;
  letter-spacing:.08em;
}
.meter-value{
  color:#fff;
  font-size:18px;
  font-weight:900;
  line-height:1;
  text-align:right;
}
.meter-track{
  height:12px;
  border-radius:999px;
  overflow:hidden;
  background:#090909;
  border:1px solid var(--border);
}
.meter-fill{
  height:100%;
  border-radius:999px;
}
.meter-fill.ok{background:linear-gradient(90deg,#22c55e,#57e38a)}
.meter-fill.warn{background:linear-gradient(90deg,#f59e0b,#ffd166)}
.meter-fill.bad{background:linear-gradient(90deg,#CC0000,#ff5959)}
.meter-note{
  margin-top:9px;
  color:var(--text2);
  font-size:12px;
  line-height:1.45;
}
.kgrid{
  display:grid;
  grid-template-columns:repeat(3,minmax(0,1fr));
  gap:12px;
  margin-top:20px;
}
.kpi{
  background:var(--black3);
  border:1px solid var(--border);
  border-radius:16px;
  padding:15px;
  min-width:0;
  min-height:104px;
}
.kpi-l{
  color:var(--text3);
  font-size:10px;
  font-family:var(--mono);
  text-transform:uppercase;
  letter-spacing:.08em;
  margin-bottom:7px;
}
.kpi-v{
  color:#fff;
  font-size:17px;
  font-weight:800;
  line-height:1.25;
  word-break:break-word;
}
.bar-wrap{
  background:#090909;
  border:1px solid var(--border);
  border-radius:999px;
  overflow:hidden;
  height:8px;
  margin-top:10px;
}
.bar-fill{height:100%;border-radius:999px}
.bar-fill.ok{background:linear-gradient(90deg,#22c55e,#57e38a)}
.bar-fill.warn{background:linear-gradient(90deg,#f59e0b,#ffd166)}
.bar-fill.bad{background:linear-gradient(90deg,#CC0000,#ff5959)}
section{
  margin-bottom:18px;
  border:1px solid var(--border);
  border-radius:22px;
  background:linear-gradient(180deg,rgba(255,255,255,.018),rgba(255,255,255,.008)),var(--black2);
  box-shadow:0 16px 40px rgba(0,0,0,.18);
  overflow:hidden;
}
h2{
  display:flex;
  align-items:center;
  gap:10px;
  padding:18px 20px 14px;
  border-bottom:1px solid var(--border);
  font-size:17px;
  font-weight:800;
  color:#fff;
}
h2::before{
  content:'';
  width:4px;
  height:18px;
  border-radius:4px;
  background:var(--red);
  box-shadow:0 0 0 4px rgba(204,0,0,.16);
}
.section-sub{
  color:var(--text2);
  font-size:12px;
  padding:0 20px 14px;
}
.resumen-box{
  margin:18px;
  border:1px solid var(--border);
  background:linear-gradient(180deg,rgba(204,0,0,.06),rgba(255,255,255,.01)),var(--black3);
  border-radius:18px;
  padding:20px 22px;
  line-height:1.9;
  font-size:14px;
}
.resumen-box strong{color:#fff}
.resumen-box .highlight{color:var(--green);font-weight:800}
.resumen-box .alert{color:#ff8585;font-weight:800}
.resumen-box .note{color:var(--amber);font-weight:800}
.tw{
  width:calc(100% - 36px);
  margin:0 18px 18px;
  overflow:auto;
  border:1px solid var(--border);
  border-radius:16px;
  background:var(--black3);
}
table{
  width:100%;
  border-collapse:collapse;
}
th{
  background:rgba(204,0,0,.08);
  color:var(--text2);
  text-align:left;
  padding:12px 14px;
  border-bottom:1px solid var(--border2);
  font-family:var(--mono);
  font-size:10px;
  font-weight:700;
  text-transform:uppercase;
  letter-spacing:.08em;
  white-space:nowrap;
}
td{
  padding:12px 14px;
  border-bottom:1px solid rgba(255,255,255,.04);
  color:var(--text);
  font-size:13px;
  vertical-align:top;
}
tbody tr:last-child td{border-bottom:none}
tbody tr:nth-child(even) td{background:rgba(255,255,255,.016)}
tbody tr:hover td{background:rgba(204,0,0,.05)}
.tag{
  display:inline-flex;
  align-items:center;
  gap:6px;
  padding:4px 10px;
  border-radius:999px;
  font-size:10px;
  font-weight:800;
  letter-spacing:.08em;
  text-transform:uppercase;
  font-family:var(--mono);
}
.tag.ok{background:var(--green-dim);color:var(--green);border:1px solid rgba(34,197,94,.2)}
.tag.warn{background:var(--amber-dim);color:var(--amber);border:1px solid rgba(245,158,11,.2)}
.tag.bad{background:var(--red-soft);color:#ff7d7d;border:1px solid rgba(204,0,0,.24)}
.tag.info{background:var(--blue-dim);color:var(--blue);border:1px solid rgba(59,130,246,.2)}
.prio-urgente{color:#ff7d7d;font-weight:800}
.prio-alta{color:#ff9a57;font-weight:800}
.prio-media{color:var(--amber);font-weight:700}
.prio-baja{color:var(--green);font-weight:700}
.prio-ninguna{color:var(--text3);font-weight:700}
.cards{
  display:grid;
  grid-template-columns:repeat(4,minmax(0,1fr));
  gap:12px;
  padding:0 18px 18px;
}
.card{
  background:var(--black3);
  border:1px solid var(--border);
  border-radius:18px;
  padding:18px;
  position:relative;
}
.card::before{
  content:'';
  position:absolute;left:0;top:0;bottom:0;width:4px;border-radius:18px 0 0 18px;
}
.card.ok::before{background:var(--green)}
.card.warn::before{background:var(--amber)}
.card.bad::before{background:var(--red)}
.card.info::before{background:var(--blue)}
.card-icon{font-size:26px;margin-bottom:10px}
.card-title{font-size:14px;font-weight:800;color:#fff;margin-bottom:8px}
.card-body{font-size:12.5px;color:var(--text2);line-height:1.7}
.card-body strong{color:#fff}
.work-box,.next-box{
  margin:18px;
  border-radius:18px;
  padding:20px;
  border:1px solid var(--border);
}
.work-box{
  background:linear-gradient(180deg,rgba(59,130,246,.08),rgba(255,255,255,.01)),var(--black3);
}
.work-title{
  font-size:14px;
  font-weight:800;
  color:var(--blue);
  margin-bottom:10px;
}
.work-body{font-size:13px;color:var(--text2);line-height:1.8}
.work-body strong{color:#fff}
.next-box{
  display:flex;
  align-items:center;
  gap:18px;
  background:linear-gradient(180deg,rgba(34,197,94,.08),rgba(255,255,255,.01)),var(--black3);
}
.next-date{
  flex-shrink:0;
  font-family:var(--mono);
  font-size:28px;
  font-weight:800;
  color:var(--green);
}
.next-msg{
  color:var(--text2);
  font-size:13px;
  line-height:1.8;
}
.next-msg strong{color:#fff}
.footer{
  margin-top:28px;
  padding:18px 8px 0;
  border-top:1px solid var(--border);
  color:var(--text3);
  text-align:center;
  font-size:11px;
  line-height:1.8;
  font-family:var(--mono);
}
@media (max-width:1100px){
  .hero-grid{grid-template-columns:1fr}
  .meter-grid{grid-template-columns:repeat(2,minmax(0,1fr))}
  .kgrid{grid-template-columns:repeat(3,minmax(0,1fr))}
  .cards{grid-template-columns:repeat(2,minmax(0,1fr))}
}
@media (max-width:720px){
  .wrap{padding:14px 10px 28px}
  .hero{padding:18px}
  .brand{align-items:flex-start}
  .brand-title{font-size:28px}
  .hero h1{font-size:24px;max-width:none}
  .hero-summary,.traffic,.meter-grid,.kgrid,.cards{grid-template-columns:1fr}
  .next-box{flex-direction:column;align-items:flex-start}
  .tw{width:calc(100% - 20px);margin:0 10px 10px}
  .resumen-box,.work-box,.next-box{margin:10px}
  h2{padding:16px 14px 12px}
  .section-sub{padding:0 14px 12px}
}
@media print{
  body{background:#fff;color:#111}
  body::before,.hero::after{display:none}
  .hero,section,.meter,.kpi,.card,.resumen-box,.work-box,.next-box,.tw{
    background:#fff!important;
    border-color:#ddd!important;
    box-shadow:none!important;
  }
  .page-kicker,.banner,.tag{border-color:#ccc!important}
  h2,th,td,.kpi-v,.meter-value,.card-title,.brand-title,.hero h1,.status-reason,.tl-value,.summary-chip .v{color:#111!important}
  .brand-title .r{color:#a00!important}
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

function Get-MeterCard {
    param(
        [string]$Label,
        [double]$Percent,
        [string]$Display,
        [string]$Tone = "ok",
        [string]$Hint = ""
    )
    $pct = [math]::Max(0, [math]::Min(100, [math]::Round($Percent, 0)))
    $safeLabel = HtmlEnc $Label
    $safeDisplay = HtmlEnc $Display
    $safeHint = HtmlEnc $Hint
    return @"
<div class='meter $Tone'>
  <div class='meter-top'>
    <div class='meter-label'>$safeLabel</div>
    <div class='meter-value'>$safeDisplay</div>
  </div>
  <div class='meter-track'><div class='meter-fill $Tone' style='width:${pct}%'></div></div>
  <div class='meter-note'>$safeHint</div>
</div>
"@
}

function Get-ClientSummary {
    param($Status, $HwAge, $Vols, $Disks, $Perf, $Temps, $Rec, $NextDate, $Issues, $Bsod, $Shutdowns)
    $estado = $Status.EstadoGeneral
    $estadoCls = if ($estado -match "EXCELENTE") { "ok" } elseif ($estado -match "OBSERVACIONES|REVISION") { "warn" } else { "bad" }
    $emoji = if ($estado -match "EXCELENTE") { "&#9989;" } elseif ($estado -match "OBSERVACIONES") { "&#9888;" } else { "&#128308;" }

    $diskMsg = ""
    $badDisk = $Disks | Where-Object { $_.Estado -in @("REEMPLAZAR","MAL") }
    $unknownDisk = $Disks | Where-Object { $_.Estado -in @("SIN DATOS","N/D") }
    if ($badDisk) { $diskMsg = "&#9888; Se detecto un problema en el almacenamiento." }
    elseif ($unknownDisk) { $diskMsg = "&#128712; No fue posible leer la salud avanzada del disco, pero no hay una falla critica informada." }
    else { $diskMsg = "&#9989; Los discos estan en buen estado." }

    $ramPctVal = $null; try { $ramPctVal = [double]$Perf.RAM_Pct } catch {}
    $ramMsg = if ($null -eq $ramPctVal) { "&#128712; No fue posible medir el uso actual de RAM en esta lectura." } elseif ($ramPctVal -ge 85) { "&#9888; La memoria RAM esta muy exigida." } else { "&#9989; La memoria RAM opera con normalidad." }

    $tempMsg = ""
    $hotZone = $Temps | Where-Object { $_.Estado -in @("ALTO","CRITICO") }
    if ($hotZone) { $tempMsg = "&#127777; Se detectaron temperaturas elevadas." }
    else { $tempMsg = "&#127777; Las temperaturas son normales." }

    $spaceMsg = ""
    $tightVol = $Vols | Where-Object { $_.Alerta -in @("ALTO","CRITICO") }
    if ($tightVol) { $spaceMsg = "&#128190; Poco espacio en disco. Conviene liberar archivos." }
    else { $spaceMsg = "&#128190; El espacio en disco esta bien." }

    $urgentRec = ($Rec | Where-Object { $_.Prioridad -in @("URGENTE","ALTA") })
    $issueCount = ($Issues | Where-Object { $_.Prioridad -ne "NINGUNA" } | Measure-Object).Count
    $recMsg = if ($urgentRec) { "Se identificaron " + ($urgentRec | Measure-Object).Count + " punto(s) que requieren atencion." } else { "No hay acciones urgentes pendientes." }
    $crashMsg = if ($Bsod | Where-Object { $_.Fecha -notmatch "^Sin BSOD recientes" }) { "&#128308; Se registraron pantallazos azules recientes." } elseif ($Shutdowns | Where-Object { $_.Fecha -notmatch "^Sin reinicios inesperados" }) { "&#9888; Se detectaron reinicios o apagados inesperados." } else { "&#9989; No se detectaron fallos graves recientes del sistema." }

    return @"
<div class='resumen-box'>
<p>$emoji <strong>Estado general del equipo: <span class='$estadoCls'>$estado</span></strong></p>
<br>
<p>$diskMsg</p>
<p>$ramMsg</p>
<p>$tempMsg</p>
<p>$spaceMsg</p>
<p>$crashMsg</p>
<br>
<p><strong>&#128295; Evaluacion del hardware:</strong> $(HtmlEnc $HwAge.Equipo_Msg)</p>
<br>
<p><strong>&#128680; Problemas detectados:</strong> $issueCount</p>
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
$bsodInfo   = Get-BsodEvents
$shutdownInfo = Get-UnexpectedShutdowns
$wheaInfo   = Get-WheaSummary
$diskEvtInfo = Get-DiskEventSummary
$memErrs    = Get-MemoryErrors
$integ      = Get-IntegrityStatus

Update-Stage 68 "Relevando software instalado y configuracion"
$startApps  = Get-StartupApps
$instApps   = Get-InstalledApps
$openPorts  = Get-OpenPorts
$nonMsSvcs  = Get-NonMsServices
$schedTasks = Get-CustomScheduledTasks
$driversInfo= Get-DriversInfo
$problemDevices = Get-ProblemDevices

Update-Stage 78 "Leyendo marca previa y comparando hardware"
$pclafPaths = Get-PCLAFPaths
$prevRec    = Read-PreviousRecord -P $pclafPaths
$fingerprint= Get-HardwareFingerprint -Sys $sysInfo -Disks $diskInfo
$hwAge      = Get-HardwareAge -Sys $sysInfo -Gpu $gpuInfo -Disks $diskInfo -Ram $ramInfo
$comparison = Compare-Records -Prev $prevRec -FP $fingerprint -Sys $sysInfo -Disks $diskInfo -Gpu $gpuInfo

Update-Stage 86 "Calculando estado final y recomendaciones"
    $finalStatus= Get-FinalAssessment -Disks $diskInfo -Vols $volInfo -Events $critEvts -HwAge $hwAge -Perf $perfInfo -Defender $defInfo -Bsod $bsodInfo -Shutdowns $shutdownInfo -Whea $wheaInfo -DiskEvents $diskEvtInfo -ProblemDevices $problemDevices
$recs       = Get-Recommendations -Disks $diskInfo -Vols $volInfo -Events $critEvts -Perf $perfInfo -HwAge $hwAge -Integrity $integ -Def $defInfo -Trim $trimInfo -Bsod $bsodInfo -Shutdowns $shutdownInfo -Whea $wheaInfo -DiskEvents $diskEvtInfo -ProblemDevices $problemDevices
$clientIssues = Get-ClientIssues -Disks $diskInfo -Vols $volInfo -Perf $perfInfo -Temps $tempInfo -Integrity $integ -Def $defInfo -HwAge $hwAge -Bsod $bsodInfo -Shutdowns $shutdownInfo -Whea $wheaInfo -DiskEvents $diskEvtInfo -ProblemDevices $problemDevices
$maintTask  = Set-MaintenanceTask -Meses $MesesMantenimiento

Update-Stage 92 "Guardando marca PCLAF en el equipo"
$record = [PSCustomObject]@{
    Metadata    = [PSCustomObject]@{ Fecha=(Get-Date -Format "yyyy-MM-ddTHH:mm:ss"); Version=$ScriptVersion; Equipo=$env:COMPUTERNAME; Tecnico=$Tecnico; Modo=$Modo; PCLAF_OS=$(if($SistemaInstaladoPorPCLAF){"SI"}else{"NO"}); MesesMant=$MesesMantenimiento }
    FinalStatus = $finalStatus; Fingerprint=$fingerprint; Comparacion=$comparison
    SystemInfo=$osInfo; SystemSummary=$sysInfo; GPU=$gpuInfo; RAM=$ramInfo
    Discos=$diskInfo; Volumenes=$volInfo; Rendimiento=$perfInfo
    Temperaturas=$tempInfo; Seguridad=$secInfo; Defender=$defInfo
    HardwareAge=$hwAge; Recomendaciones=$recs; ProblemasCliente=$clientIssues
    BSOD=$bsodInfo; ReiniciosInesperados=$shutdownInfo; WHEA=$wheaInfo
    EventosDisco=$diskEvtInfo; DispositivosConProblemas=$problemDevices
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
$tlDisk  = Get-TrafficLight "Discos"      ($diskInfo|Select-Object -First 1).Estado (if(($diskInfo|Select-Object -First 1).Estado -eq "BIEN"){"ok"}elseif(($diskInfo|Select-Object -First 1).Estado -in @("SIN DATOS","N/D")){"warn"}else{"bad"})
$tlRam   = Get-TrafficLight "RAM"         $(if($perfInfo.RAM_Pct -eq "N/D"){"Uso no disponible"}else{"$($perfInfo.RAM_Pct)% usada"}) (if($perfInfo.RAM_Pct -eq "N/D"){"warn"}elseif([double]$perfInfo.RAM_Pct -ge 85){"bad"}elseif([double]$perfInfo.RAM_Pct -ge 65){"warn"}else{"ok"})
$tlTemp  = Get-TrafficLight "Temperatura" (($tempInfo|Select-Object -First 1).Estado) (if(($tempInfo|Select-Object -First 1).Estado -in @("CRITICO","ALTO")){"bad"}elseif(($tempInfo|Select-Object -First 1).Estado -eq "ELEVADO"){"warn"}else{"ok"})
$tlDef   = Get-TrafficLight "Antivirus"   (if($defInfo.Activo -eq $true){"Activo"}else{"Inactivo"}) (if($defInfo.Activo -eq $true){"ok"}else{"bad"})
$tlSpace = Get-TrafficLight "Espacio"     (($volInfo|Sort-Object Uso_Pct -Descending|Select-Object -First 1).Alerta) (if(($volInfo|Sort-Object Uso_Pct -Descending|Select-Object -First 1).Alerta -in @("CRITICO","ALTO")){"bad"}elseif(($volInfo|Sort-Object Uso_Pct -Descending|Select-Object -First 1).Alerta -in @("MODERADO","N/D")){"warn"}else{"ok"})

# RAM bar
$ramBarPct = 0
$ramPctDisplay = "N/D"
try {
    $ramBarPct = [double]$perfInfo.RAM_Pct
    $ramPctDisplay = "$([math]::Round($ramBarPct,0))%"
} catch {}
$ramBar = Get-RamBar -Pct $ramBarPct

$numericVols = @(
    $volInfo | Where-Object {
        $val = Get-NumericValueOrNull $_.Uso_Pct
        $null -ne $val
    }
)
$topVol = if ($numericVols.Count -gt 0) {
    $numericVols | Sort-Object { [double]$_.Uso_Pct } -Descending | Select-Object -First 1
} else {
    $volInfo | Select-Object -First 1
}
$diskUsePct = 0
$diskPctDisplay = "N/D"
try {
    $diskUsePct = [double]$topVol.Uso_Pct
    $diskPctDisplay = "$([math]::Round($diskUsePct,0))%"
} catch {}
$topTemp = $tempInfo | Select-Object -First 1
$tempPct = 0
try {
    if ($topTemp.Celsius -ne "N/D") { $tempPct = [math]::Min(100, [math]::Max(0, ([double]$topTemp.Celsius / 100) * 100)) }
} catch {}
$cpuPct = 0
$cpuPctDisplay = "N/D"
try {
    $cpuPct = [double]$perfInfo.CPU_Pct
    $cpuPctDisplay = "$([math]::Round($cpuPct,0))%"
} catch {}
$critCount = 0; try { $critCount = @($critEvts).Count } catch {}
$bsodCount = 0; try { $bsodCount = @($bsodInfo | Where-Object { $_.Fecha -notmatch "^Sin BSOD recientes" }).Count } catch {}

$clientMeterHtml = (
    (Get-MeterCard -Label "RAM actual" -Percent $ramBarPct -Display $ramPctDisplay -Tone $(if($ramPctDisplay -eq "N/D"){"warn"}elseif($ramBarPct -ge 85){"bad"}elseif($ramBarPct -ge 65){"warn"}else{"ok"}) -Hint "Uso actual de memoria del sistema") +
    (Get-MeterCard -Label "Disco principal" -Percent $diskUsePct -Display $diskPctDisplay -Tone $(if($diskPctDisplay -eq "N/D"){"warn"}elseif($diskUsePct -ge 95){"bad"}elseif($diskUsePct -ge 85){"warn"}else{"ok"}) -Hint "Ocupacion de la unidad mas cargada") +
    (Get-MeterCard -Label "Temperatura" -Percent $tempPct -Display $(if($topTemp.Celsius -ne "N/D"){"$($topTemp.Celsius) C"}else{"N/D"}) -Tone $(if($topTemp.Estado -in @("CRITICO","ALTO")){"bad"}elseif($topTemp.Estado -eq "ELEVADO"){"warn"}else{"ok"}) -Hint "Referencia visual para la temperatura reportada") +
    (Get-MeterCard -Label "Alertas graves" -Percent ([math]::Min(100, ($bsodCount*25)+($critCount*5))) -Display "$bsodCount BSOD / $critCount eventos" -Tone $(if($bsodCount -gt 0){"bad"}elseif($critCount -gt 0){"warn"}else{"ok"}) -Hint "Pantallazos azules y eventos relevantes recientes")
)

$techMeterHtml = (
    (Get-MeterCard -Label "CPU actual" -Percent $cpuPct -Display $cpuPctDisplay -Tone $(if($cpuPctDisplay -eq "N/D"){"warn"}elseif($cpuPct -ge 90){"bad"}elseif($cpuPct -ge 70){"warn"}else{"ok"}) -Hint "Uso instantaneo del procesador") +
    (Get-MeterCard -Label "RAM actual" -Percent $ramBarPct -Display $ramPctDisplay -Tone $(if($ramPctDisplay -eq "N/D"){"warn"}elseif($ramBarPct -ge 85){"bad"}elseif($ramBarPct -ge 65){"warn"}else{"ok"}) -Hint "Uso instantaneo de memoria") +
    (Get-MeterCard -Label "Disco / volumen" -Percent $diskUsePct -Display $diskPctDisplay -Tone $(if($diskPctDisplay -eq "N/D"){"warn"}elseif($diskUsePct -ge 95){"bad"}elseif($diskUsePct -ge 85){"warn"}else{"ok"}) -Hint "Volumen con mayor ocupacion") +
    (Get-MeterCard -Label "Temperatura" -Percent $tempPct -Display $(if($topTemp.Celsius -ne "N/D"){"$($topTemp.Celsius) C"}else{"N/D"}) -Tone $(if($topTemp.Estado -in @("CRITICO","ALTO")){"bad"}elseif($topTemp.Estado -eq "ELEVADO"){"warn"}else{"ok"}) -Hint "Sensor mas exigido informado"),
    (Get-MeterCard -Label "Eventos criticos" -Percent ([math]::Min(100, $critCount*8)) -Display "$critCount eventos" -Tone $(if($critCount -gt 8){"bad"}elseif($critCount -gt 0){"warn"}else{"ok"}) -Hint "Eventos criticos relevados en 30 dias"),
    (Get-MeterCard -Label "Defensa" -Percent $(if($defInfo.Activo -eq $true){100}else{25}) -Display $(if($defInfo.Activo -eq $true){"Activa"}else{"Inactiva"}) -Tone $(if($defInfo.Activo -eq $true){"ok"}else{"bad"}) -Hint "Estado de Windows Defender")
) -join ""

$heroMeterHtml = if ($Modo -eq "tecnico") { $techMeterHtml } else { $clientMeterHtml }

# Logo base64 (favicon/logo real de PCLAF)
$LogoB64 = "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"
$BrandMark = @"
<svg viewBox="0 0 42 42" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
  <rect width="42" height="42" rx="6" fill="#CC0000"></rect>
  <text x="21" y="19" text-anchor="middle" font-family="Arial Black,Arial,sans-serif" font-weight="900" font-size="18" fill="white">PC</text>
  <text x="21" y="38" text-anchor="middle" font-family="Arial Black,Arial,sans-serif" font-weight="900" font-size="14" fill="white">LAF</text>
</svg>
"@

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
  <div class="hero-grid">
    <div class="hero-main">
      <div class="brand">
        <div class="brand-logo">$BrandMark</div>
        <div>
          <div class="brand-title">PC<span class="r">LAF</span></div>
          <div class="brand-sub">Servicio tecnico • Floresta, CABA • pclaf.com.ar</div>
        </div>
      </div>
      <div class="page-kicker">Reporte $($Modo.ToUpper()) • Script v$ScriptVersion</div>
      <h1>Diagnostico completo del equipo $($env:COMPUTERNAME)</h1>
      <div class="sub">
        <span>$(Get-Date -Format "dd/MM/yyyy HH:mm")</span>
        <span>Tecnico: $Tecnico</span>
        <span>Modo: $($Modo.ToUpper())</span>
      </div>
      <div class="hero-summary">
        <div class="summary-chip">
          <div class="k">Sistema operativo</div>
          <div class="v">$(HtmlEnc $osInfo.SO)</div>
        </div>
        <div class="summary-chip">
          <div class="k">Memoria instalada</div>
          <div class="v">$($sysInfo.RAM_Total_GB) GB</div>
        </div>
        <div class="summary-chip">
          <div class="k">Revision sugerida</div>
          <div class="v">$nextDate</div>
        </div>
      </div>
    </div>
    <div class="hero-side">
      <div class="status-card">
        <div class="status-top">
          <div class="status-label">Estado general</div>
          <div class="banner $estadoCls">$($finalStatus.EstadoGeneral)</div>
        </div>
        <div class="status-reason">$(HtmlEnc $finalStatus.Motivos)</div>
        <div class="traffic">
          $tlHw $tlDisk $tlRam $tlTemp $tlDef $tlSpace
        </div>
      </div>
      <div class="status-card">
        <div class="section-accent">
          <div>
            <div class="sa-kicker">Panel visual</div>
            <div class="sa-title">Indicadores PCLAF</div>
          </div>
          <div class="sa-line"></div>
        </div>
        <div class="meter-grid">
          $heroMeterHtml
        </div>
      </div>
    </div>
  </div>
  <div class="kgrid">
    <div class="kpi"><div class="kpi-l">Sistema operativo</div><div class="kpi-v" style="font-size:13px">$(HtmlEnc $osInfo.SO)</div></div>
    <div class="kpi"><div class="kpi-l">Procesador</div><div class="kpi-v" style="font-size:12px">$(HtmlEnc $sysInfo.CPU)</div></div>
    <div class="kpi"><div class="kpi-l">Memoria RAM</div><div class="kpi-v">$($sysInfo.RAM_Total_GB) GB</div></div>
    <div class="kpi"><div class="kpi-l">Uso RAM actual</div><div class="kpi-v" style="font-size:17px">$ramPctDisplay$ramBar</div></div>
    <div class="kpi"><div class="kpi-l">Disco principal</div><div class="kpi-v" style="font-size:12px">$(HtmlEnc (($diskInfo|Select-Object -First 1).Modelo))</div></div>
    <div class="kpi"><div class="kpi-l">Temperatura CPU</div><div class="kpi-v">$(if(($tempInfo|Select-Object -First 1).Celsius -ne "N/D"){"$(($tempInfo|Select-Object -First 1).Celsius)°C"}else{"Sin sensor"})</div></div>
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
$(Get-ClientSummary -Status $finalStatus -HwAge $hwAge -Vols $volInfo -Disks $diskInfo -Perf $perfInfo -Temps $tempInfo -Rec $recs -NextDate $nextDate -Issues $clientIssues -Bsod $bsodInfo -Shutdowns $shutdownInfo)
</section>

<section>
<h2>&#128680; Problemas detectados y que conviene corregir</h2>
<div class="section-sub">Resumen claro de errores y condiciones que deberia revisar el tecnico</div>
$(To-HtmlTable $clientIssues)
</section>

<section>
<h2>&#128565; Fallos graves recientes del sistema</h2>
<div class="section-sub">Pantallazos azules y reinicios inesperados detectados por Windows</div>
$(To-HtmlTable $bsodInfo)
$(To-HtmlTable $shutdownInfo)
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
<div class="section-accent">
  <div>
    <div class="sa-kicker">Panel tecnico</div>
    <div class="sa-title">Tacometros y carga del sistema</div>
  </div>
  <div class="sa-line"></div>
</div>
<div class="meter-grid">
$techMeterHtml
</div>
</section>

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
<h2>&#128565; Pantallazos azules (BSOD)</h2>
$(To-HtmlTable $bsodInfo)
</section>

<section>
<h2>&#9888; Reinicios inesperados</h2>
$(To-HtmlTable $shutdownInfo)
</section>

<section>
<h2>&#129520; Errores WHEA de hardware</h2>
$(To-HtmlTable $wheaInfo)
</section>

<section>
<h2>&#128190; Errores de disco y controlador</h2>
$(To-HtmlTable $diskEvtInfo)
</section>

<section>
<h2>&#128421; Dispositivos con problemas</h2>
$(To-HtmlTable $problemDevices)
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

$uploadResult = $null
if ($SubirASupabase -and $ReparacionId -and $SupabaseUrl -and $SupabaseAnonKey) {
    $slotInfo = Get-ReportSlotInfo -ModoActual $Modo -MomentoActual $MomentoReporte
    $htmlClienteSubida = if ($Modo -eq "cliente") { $html } else { $null }
    Update-Stage 97 "Preparando reporte para subida"
    Update-Stage 98 "Validando datos para Supabase"
    Update-Stage 99 "Enviando archivo a Supabase"
    $uploadResult = Invoke-SupabaseReportUpload -BaseUrl $SupabaseUrl -AnonKey $SupabaseAnonKey -RepairId $ReparacionId -HtmlFull $html -HtmlClient $htmlClienteSubida -FileName $slotInfo.FileName
}

Update-Stage 100 "Listo!"
Write-Progress -Activity "PCLAF Diagnostico" -Completed
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  PCLAF Diagnostico v$ScriptVersion - $($Modo.ToUpper())" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Reporte : $outFile" -ForegroundColor Green
Write-Host "  Estado  : $($finalStatus.EstadoGeneral)" -ForegroundColor $(if($finalStatus.EstadoGeneral -match "EXCELENTE"){"Green"}elseif($finalStatus.EstadoGeneral -match "OBSERVACIONES"){"Yellow"}else{"Red"})
Write-Host "  Marca   : $(if($markOk){"Guardada en registro"}else{"No se pudo guardar"})" -ForegroundColor $(if($markOk){"Green"}else{"Yellow"})
if ($uploadResult) {
    Write-Host "  Subida  : $($uploadResult.Message)" -ForegroundColor $(if($uploadResult.Ok){"Green"}else{"Yellow"})
    Write-Host "  Log     : $env:TEMP\PCLAF_Upload.log" -ForegroundColor DarkCyan
}
Write-Host ""

if ($SubirASupabase -and (-not $uploadResult -or -not $uploadResult.Ok)) {
    Write-Error ("La subida automatica del reporte fallo. " + $(if ($uploadResult) { $uploadResult.Message } else { "No se obtuvo respuesta del intento de subida." }))
    exit 1
}

# Abrir el reporte automaticamente
try { Start-Process $outFile } catch {}
