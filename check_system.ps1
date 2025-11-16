#instalar por powershell : Install-Module -Name ImportExcel

$BasePath = 'BasePath'
$SucCode = Read-Host "Ingresar codigo de sucursal"

$FolderPath = Get-ChildItem -LiteralPath $BasePath -Directory | Where-Object { $_.Name -match "^$SucCode\s-\s" }
if (-not $FolderPath) {
    Write-Host "No se encontro la carpeta para la sucursal" $SucCode -ForegroundColor Red
    exit
}

$ExcelFile = Get-ChildItem -Path $FolderPath.FullName |
Where-Object { $_.Extension -like '*.xls*' }
if (-not $ExcelFile) {
    Write-Host "No se encontro ningun archivo Excel en la carpeta" $($FolderPath.FullName) -ForegroundColor Red
    exit
}


$pattern = '^[a-z][0-9]{4}sc[0-9]{4}$'

$ComputerList = Import-Excel -Path $ExcelFile.FullName | Select-Object -ExpandProperty 'Nombre del Sistema'

Write-Host "Equipos encontrados para la sucursal:" $SucCode -ForegroundColor Cyan
$ComputerList

$ComputerList = $ComputerList |
Where-Object { $_ } |
ForEach-Object { $_.ToString().Trim() } |
Where-Object { $_ -imatch $pattern } |
Sort-Object -Unique

<#
.SYNOPSIS
    Convierte el código numérico de form factor de memoria RAM a su etiqueta legible.

.DESCRIPTION
    Esta función toma el código numérico del form factor de memoria RAM (obtenido de Win32_PhysicalMemory)
    y lo convierte a una etiqueta legible como DIMM, SODIMM, etc.

.PARAMETER ff
    Código numérico del form factor de memoria RAM.

.EXAMPLE
    Get-FormFactorLabel -ff 8
    Retorna: 'DIMM'
#>
function Get-FormFactorLabel {
    param([int]$ff)
    switch ($ff) {
        8 { 'DIMM' } 12 { 'SODIMM' } 9 { 'RIMM' } 10 { 'SODIMM-RIMM' }
        0 { 'Desconocido' } default { "FormFactor:$ff" }
    }
}

<#
.SYNOPSIS
    Convierte el código SMBIOS de tipo de memoria RAM a su etiqueta DDR.

.DESCRIPTION
    Esta función toma el código SMBIOSMemoryType de Win32_PhysicalMemory y lo convierte
    a una etiqueta legible como DDR3, DDR4, DDR5, etc.

.PARAMETER smbiosType
    Código SMBIOS del tipo de memoria RAM.

.EXAMPLE
    Get-DDRTypeLabel -smbiosType 24
    Retorna: 'DDR3'
#>
function Get-DDRTypeLabel {
    param([int]$smbiosType)
    switch ($smbiosType) {
        20 { 'DDR' }
        21 { 'DDR2' }
        24 { 'DDR3' }
        26 { 'DDR4' }
        34 { 'DDR5' }
        0 { 'Desconocido' }
        default { "Tipo:$smbiosType" }
    }
}



<#
.SYNOPSIS
    Establece una sesión CIM remota con un equipo, intentando primero WSMAN y luego DCOM como fallback.

.DESCRIPTION
    Intenta crear una sesión CIM remota usando primero el protocolo WSMAN. Si falla,
    intenta con DCOM como protocolo alternativo. Esto mejora la compatibilidad con
    diferentes configuraciones de red y sistemas.

.PARAMETER Computer
    Nombre del equipo remoto al que se desea conectar.

.EXAMPLE
    $session = Test-NewCimSession -Computer "PC01"
    Crea una sesión CIM con el equipo PC01.
#>
function Test-NewCimSession {
    param(
        [string]$Computer
    )
    $TimeoutSeconds = 30
    try {
        $opt = New-CimSessionOption -OperationTimeoutSec $TimeoutSeconds
        return New-CimSession -ComputerName $Computer -SessionOption $opt -ErrorAction Stop
    }
    catch {
        try {
            $opt = New-CimSessionOption -Protocol Dcom -OperationTimeoutSec $TimeoutSeconds
            return New-CimSession -ComputerName $Computer -SessionOption $opt -ErrorAction Stop
        }
        catch { return $null }
    }
}


<#
.SYNOPSIS
    Obtiene información de los discos físicos de un equipo remoto.

.DESCRIPTION
    Intenta obtener información de los discos físicos usando Get-PhysicalDisk (si está disponible)
    o Win32_DiskDrive como fallback. Identifica si los discos son SSD o HDD basándose en
    el MediaType y el modelo del disco.

.PARAMETER Computer
    Nombre del equipo remoto.

.PARAMETER CimSession
    Sesión CIM existente para reutilizar la conexión.

.EXAMPLE
    $discos = Get-DisksInfo -Computer "PC01" -CimSession $session
    Obtiene información de los discos del equipo PC01.
#>
function Get-DisksInfo {
    param(
        [string]$Computer,
        [Microsoft.Management.Infrastructure.CimSession]$CimSession
    )

    if ($CimSession -and $CimSession.Protocol -eq 'WSMAN') {
        try {
            $pd = Get-PhysicalDisk -CimSession $CimSession -ErrorAction Stop |
            Select-Object FriendlyName, MediaType
            if ($pd) { return $pd }
        }
        catch { }
    }
    try {
        $dd = if ($CimSession) {
            Get-CimInstance -ClassName Win32_DiskDrive -CimSession $CimSession -ErrorAction Stop
        }
        else {
            Get-WmiObject -Class Win32_DiskDrive -ComputerName $Computer -ErrorAction Stop
        }
        return $dd |
        ForEach-Object {
            $tipo = if ((($_.MediaType -as [string]) -match '(?i)SSD' -or
                            ($_.Model -as [string]) -match '(?i)SSD|NVMe')) { 'SSD' }
            else { 'HDD/Desconocido' }
            [pscustomobject]@{
                FriendlyName = $_.Model
                MediaType    = $tipo
            }
        }
    }
    catch {
        return @()
    }
}

<#
.SYNOPSIS
    Normaliza el texto de estado de cumplimiento a un valor estándar.

.DESCRIPTION
    Convierte diferentes variaciones de texto de cumplimiento (Sí, Si, SÍ, etc.)
    a valores normalizados: 'Si', 'No', 'Error' o 'Desconocido'.

.PARAMETER Texto
    Texto del estado de cumplimiento a normalizar.

.EXAMPLE
    Resolve-EstadoCumplimiento -Texto "Sí"
    Retorna: 'Si'
#>
function Resolve-EstadoCumplimiento {
    param([string]$Texto)

    if (-not $Texto) { return 'Desconocido' }
    $t = $Texto.Trim()
    if ($t -match '^(?i)s[ií]$') { return 'Si' }
    if ($t -match '^(?i)no$') { return 'No' }
    if ($t -match '(?i)error|no\s*responde|rpc|unreachable|fall[oó]|timeout|offline|no\s*accesible') {
        return 'Error'
    }
    return 'Desconocido'
}

$FechaHoraActual = Get-Date -Format 'dd-MM-yyyy || HH:mm'

<#
.SYNOPSIS
    Calcula las fechas de ejecución y cumplimiento basándose en el estado actual y previo.

.DESCRIPTION
    Esta función determina las fechas de ejecución y cumplimiento según el estado actual:
    - Si cumple: mantiene la fecha previa de cumplimiento o establece la actual si es primera vez
    - Si no cumple: limpia la fecha de cumplimiento (null)
    - Si hay error: mantiene la fecha previa si existe (no se puede verificar)

.PARAMETER Prev
    Objeto con los datos previos del equipo (de ejecuciones anteriores).

.PARAMETER CumpleActual
    Estado actual de cumplimiento ('Sí', 'No', 'Error', etc.).

.PARAMETER FechaHoraActual
    Fecha y hora actual en formato 'dd-MM-yyyy || HH:mm'.

.EXAMPLE
    $fechas = Get-FechasEjecucion -Prev $prev -CumpleActual 'Sí' -FechaHoraActual $FechaHoraActual
    Calcula las fechas para un equipo que cumple.
#>
function Get-FechasEjecucion {
    param(
        [psobject]$Prev,
        [string]$CumpleActual,
        [string]$FechaHoraActual
    )

    $fechaUlt = $FechaHoraActual
    $fechaCumpl = $null

    $fechaPrev = $null
    if ($Prev -and $Prev.PSObject.Properties['FechaHoraCumplimiento'] -and $Prev.FechaHoraCumplimiento) {
        $fechaPrev = [string]$Prev.FechaHoraCumplimiento
    }

    $estado = Resolve-EstadoCumplimiento -Texto $CumpleActual

    switch ($estado) {
        'Si' {
            $fechaCumpl = $(if ($fechaPrev) { $fechaPrev } else { $FechaHoraActual })
        }
        'No' {
            $fechaCumpl = $null
        }
        'Error' {
            $fechaCumpl = $(if ($fechaPrev) { $fechaPrev } else { $null })
        }
        default {
            $fechaCumpl = $fechaPrev
        }
    }

    [pscustomobject]@{
        FechaHoraEjecucion    = $fechaUlt
        FechaHoraCumplimiento = $fechaCumpl
    }
}



$Separator = "=" * 80
$FolderName = $FolderPath.Name

$ResultadoExcel = "ResultadoExcelPath- $FolderName.xlsx"
$Hoja = 'Resultados'


$datos = if (Test-Path $ResultadoExcel) {
    Import-Excel -Path $ResultadoExcel -WorksheetName $Hoja
}
else { @() }

$porEquipo = @{}
foreach ($r in $datos) { $porEquipo[$r.Equipo] = $r }

$TotalEquipos = $ComputerList.Count
$EquipoActual = 0
$EquiposCumplen = 0
$EquiposNoCumplen = 0
$EquiposError = 0
$EquiposOmitidos = 0

Write-Host "`nTotal de equipos a procesar: $TotalEquipos" -ForegroundColor Cyan
Write-Host "$Separator`n" -ForegroundColor Cyan

foreach ($Computer in $ComputerList) {
    $EquipoActual++
    Write-Host "`n$Separator" -ForegroundColor Cyan
    Write-Host "[$EquipoActual/$TotalEquipos] Iniciando monitoreo de: $Computer" -ForegroundColor Yellow
    Write-Host "$Separator" -ForegroundColor Cyan


    $ComputerKey = ([string]$Computer).Trim()
    $prev = $porEquipo[$ComputerKey]

    if ($prev) {
        $valorCumple = ([string]$prev.Cumple).Trim()

        if ($valorCumple -match '^(s[ií])$') {
            Write-Host "Equipo $Computer ya verificado y CUMPLE. Se omite el analisis." -ForegroundColor Cyan
            $EquiposOmitidos++
            continue
        }
    }

    if (-not (Test-Connection -ComputerName $Computer -Count 1 -Quiet)) {
        Write-Host "Equipo $Computer no disponible. Se omite." -ForegroundColor DarkYellow
        $EquiposError++
        $prev = $porEquipo[$Computer]
        $fechas = Get-FechasEjecucion -Prev $prev -CumpleActual 'No responde' -FechaHoraActual $FechaHoraActual
        if ($prev) {

            $nuevo = $prev.PSObject.Copy()
            $nuevo.Cumple = 'No responde'
            $nuevo.Observacion = 'No responde al ping'
            $nuevo.FechaHoraEjecucion = $fechas.FechaHoraEjecucion
            $nuevo.FechaHoraCumplimiento = $fechas.FechaHoraCumplimiento
            if (-not $nuevo.FabricanteEquipo) { $nuevo.FabricanteEquipo = 'N/D' }
            if (-not $nuevo.TipoRAM) { $nuevo.TipoRAM = 'N/D' }
            if (-not $nuevo.VelocidadRAM) { $nuevo.VelocidadRAM = 'N/D' }
            $porEquipo[$Computer] = $nuevo
        }
        else {

            $porEquipo[$Computer] = [pscustomobject]@{
                Equipo                = $Computer
                Serial                = 'N/D'
                FabricanteEquipo      = 'N/D'
                Cumple                = 'No responde'
                RAM                   = 'N/D'
                TipoRAM               = 'N/D'
                VelocidadRAM          = 'N/D'
                SlotsRAM              = 'N/D'
                FabricanteRAM         = 'N/D'
                CapacidadPorSlot      = 'N/D'
                Discos                = 'N/D'
                TotalDiscoGB          = 'N/D'
                Observacion           = 'No responde al ping'
                FechaHoraEjecucion    = $fechas.FechaHoraEjecucion
                FechaHoraCumplimiento = $fechas.FechaHoraCumplimiento

            }
        }
        continue
    }

    $session = $null
    try {

        $DDRTypes = @()
        $VelocidadesRAM = @()
        $CapPorSlotArray = @()

        $session = Test-NewCimSession -Computer $Computer -TimeoutSeconds 30

        $serial = if ($session) {
            Get-CimInstance -ClassName Win32_BIOS -CimSession $session -ErrorAction Stop |
            Select-Object -ExpandProperty SerialNumber
        }
        else {
            Get-WmiObject -Class Win32_BIOS -ComputerName $Computer -ErrorAction Stop |
            Select-Object -ExpandProperty SerialNumber
        }

        $FabricanteEquipo = if ($session) {
            Get-CimInstance -ClassName Win32_ComputerSystem -CimSession $session -ErrorAction Stop |
            Select-Object -ExpandProperty Manufacturer
        }
        else {
            Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Computer -ErrorAction Stop |
            Select-Object -ExpandProperty Manufacturer
        }

        $PhysicalMemory = if ($session) {
            Get-CimInstance -ClassName Win32_PhysicalMemory -CimSession $session -ErrorAction Stop

        }
        else {
            Get-WmiObject -Class Win32_PhysicalMemory -ComputerName $Computer -ErrorAction Stop
        }


        $MemoryInfo = if ($session) {
            Get-CimInstance -ClassName Win32_OperatingSystem -CimSession $session -ErrorAction Stop
        }
        else {
            Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Computer -ErrorAction Stop
        }


        if ($null -eq $PhysicalMemory -or $PhysicalMemory.Count -eq 0) {
            Write-Host "INFO: Usando Win32_ComputerSystem para la RAM total (Fallback)."
            $PhysicalMemory = Get-WmiObject Win32_ComputerSystem -ComputerName $Computer -ErrorAction Stop

            $TotalBytes = $PhysicalMemory.TotalPhysicalMemory
            $DDRTypes = @() 
        }
        else {
            $TotalBytes = ($PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum
        }

        $TotalRAM_GB = [math]::Round($TotalBytes / 1GB, 2)





        $FreeRAM_GB = [math]::Round($MemoryInfo.FreePhysicalMemory / (1GB / 1KB), 2)


        $discos = Get-DisksInfo -Computer $Computer -CimSession $session

        $Disks = if ($session) {
            Get-CimInstance -ClassName Win32_LogicalDisk -CimSession $session -Filter "DriveType = 3" -ErrorAction Stop
        }
        else {
            Get-WmiObject -Class Win32_LogicalDisk -ComputerName $Computer -Filter "DriveType = 3" -ErrorAction Stop
        }

        Write-Host ">>> Nro serial:" $serial -ForegroundColor Green
        Write-Host ">>> Fabricante:" $FabricanteEquipo -ForegroundColor Green
        Write-Host "`n>>> MEMORIA RAM" -ForegroundColor Green
        Write-Host "RAM Total: $TotalRAM_GB GB"
        Write-Host "RAM Libre: $FreeRAM_GB GB"
        $PhysicalMemory |
        Where-Object { $_.Capacity } | 
        ForEach-Object {
            $Capacity_GB = [math]::Round($_.Capacity / 1GB, 2)
            $FormFactor = Get-FormFactorLabel $_.FormFactor
            $DDRType = Get-DDRTypeLabel $_.SMBIOSMemoryType
            $Speed = $_.Speed
            $CapPorSlotArray += $Capacity_GB
            if ($DDRType -and $DDRType -ne 'Desconocido' -and $DDRType -notmatch '^Tipo:') {
                $DDRTypes += $DDRType
            }
            if ($Speed -and $Speed -gt 0) {
                $VelocidadesRAM += $Speed
            }
            Write-Host " --- Slot [PID: $($_.Tag)] ---" -ForegroundColor White
            Write-Host "  Fabricante: $($_.Manufacturer)"
            Write-Host "  Capacidad: $Capacity_GB GB"
            Write-Host "  Velocidad: $Speed MHz"
            Write-Host "  Formato: $FormFactor"
            Write-Host "  Tipo DDR: $DDRType"
        }

        Write-Host "`n>>> DISCOS FISICOS (Tipo)" -ForegroundColor Green
        if ($discos) {
            foreach ($disco in $discos) {
                Write-Host " --- Disco Fisico ---" -ForegroundColor White
                Write-Host "  Modelo: $($disco.FriendlyName)"
                Write-Host "  Tipo: $($disco.MediaType)"
            }
        }
        else {
            Write-Host "No se pudieron obtener discos fisicos." -ForegroundColor DarkYellow
        }

        Write-Host "`n>>> DISCOS LOCALES" -ForegroundColor Green
        $TotalSize_GB_All = 0
        $FreeSpace_GB_All = 0
        if ($Disks) {
            foreach ($Disk in $Disks) {
                if ($Disk.Size) {
                    $TotalSize_GB = [math]::Round($Disk.Size / 1GB, 2)
                    $FreeSpace_GB = [math]::Round($Disk.FreeSpace / 1GB, 2)
                    $UsedPercent = [math]::Round(($Disk.Size - $Disk.FreeSpace) / $Disk.Size * 100, 2)
                    $TotalSize_GB_All += $TotalSize_GB
                    $FreeSpace_GB_All += $FreeSpace_GB
                    Write-Host " Unidad $($Disk.DeviceID)" -ForegroundColor White
                    Write-Host "  Etiqueta: $($Disk.VolumeName)"
                    Write-Host "  Capacidad Total: $TotalSize_GB GB"
                    Write-Host "  Espacio Libre: $FreeSpace_GB GB"
                    Write-Host "  Ocupacion (%): $UsedPercent%"
                }
            }
        }
        else {
            Write-Host "No se encontraron discos locales o son inaccesibles." -ForegroundColor DarkYellow
        }

        $CantidadSlots = $PhysicalMemory.Count
        $CapacidadPorSlotText = if ($CapPorSlotArray.Count -gt 0) {
            ($CapPorSlotArray | ForEach-Object { '{0:N2}' -f $_ }) -join ' ' + ' GB - '
        }
        else { 'N/D' }

        $FabricanteRAMInfo = ($PhysicalMemory | Select-Object -ExpandProperty Manufacturer | Sort-Object -Unique) -join ', '
        $TipoRAMInfo = if ($DDRTypes.Count -gt 0) {
            ($DDRTypes | Sort-Object -Unique) -join ', '
        }
        else { 'N/D' }
        $VelocidadRAMInfo = if ($VelocidadesRAM.Count -gt 0) {
            ($VelocidadesRAM | Sort-Object -Unique | ForEach-Object { "$_ MHz" }) -join ', '
        }
        else { 'N/D' }
        $DiscosInfo = if ($discos) {
            ($discos | ForEach-Object { "$($_.FriendlyName) [$($_.MediaType)]" }) -join ", "
        }
        else { 'N/D' }

        
        $haySSD = $false
        $SSDValido = $false
        $SSDSizeGB = 0
        $DiscoSO = $null
        
        if ($Disks) {
            $DiscoSO = $Disks | Where-Object { $_.DeviceID -eq $MemoryInfo.SystemDrive }
            if (-not $DiscoSO) {
                $DiscoSO = $Disks | Where-Object { $_.DeviceID -eq 'C:' } | Select-Object -First 1
            }
        }
        

        if ($discos) {
            foreach ($d in $discos) { 
                if ($d.MediaType -match '(?i)SSD') { 
                    $haySSD = $true
                    break 
                } 
            }
        }
        

        if ($DiscoSO) {
            $SSDSizeGB = [math]::Round($DiscoSO.Size / 1GB, 2)
        }
        
  
        if ($haySSD) {
            $SSDValido = $true
        }
        
 
        $SSDTamanoMinimo = 128  #
        $CumpleCondicion = ($TotalRAM_GB -ge 7 -and $SSDValido)
        
        if ($SSDSizeGB -gt 0 -and $SSDSizeGB -lt $SSDTamanoMinimo) {
            $CumpleCondicion = $false
        }
        $CumpleTexto = if ($CumpleCondicion) { 'Sí' } else { 'No' }
        

        if ($CumpleTexto -eq 'Sí') {
            $EquiposCumplen++
        }
        else {
            $EquiposNoCumplen++
        }


        $prev = $porEquipo[$Computer]
        $fechas = Get-FechasEjecucion -Prev $prev -CumpleActual $CumpleTexto -FechaHoraActual $FechaHoraActual


        $ObservacionDetallada = ''
        if (-not $CumpleCondicion) {
            $razones = @()
            if ($TotalRAM_GB -lt 7) { $razones += "RAM insuficiente ($TotalRAM_GB GB < 7 GB)" }
            if (-not $SSDValido) { $razones += "No tiene SSD o SSD no es disco del sistema" }
            if ($SSDSizeGB -gt 0 -and $SSDSizeGB -lt $SSDTamanoMinimo) { 
                $razones += "SSD muy pequeño ($SSDSizeGB GB < $SSDTamanoMinimo GB)" 
            }
            if ($razones.Count -gt 0) {
                $ObservacionDetallada = $razones -join '; '
            }
        }
        
        $fila = [pscustomobject]@{
            Equipo                = $Computer
            Serial                = $serial
            FabricanteEquipo      = $FabricanteEquipo
            Cumple                = $CumpleTexto
            RAM                   = ("{0:N2} GB" -f $TotalRAM_GB)
            TipoRAM               = $TipoRAMInfo
            VelocidadRAM          = $VelocidadRAMInfo
            SlotsRAM              = $CantidadSlots
            CapacidadPorSlot      = $CapacidadPorSlotText
            FabricanteRAM         = $FabricanteRAMInfo
            Discos                = $DiscosInfo
            TotalDiscoGB          = ("{0:N2} GB" -f $TotalSize_GB_All)
            Observacion           = $ObservacionDetallada
            FechaHoraEjecucion    = $fechas.FechaHoraEjecucion
            FechaHoraCumplimiento = $fechas.FechaHoraCumplimiento

        }
        $porEquipo[$Computer] = $fila

    }
    catch {
        Write-Host "`nERROR al conectarse o consultar '$Computer'." -ForegroundColor Red
        Write-Host "Mensaje de error:" $_.Exception.Message -ForegroundColor Red
        Write-Host "Asegurate de que la PC este encendida, en red y que tenes permisos." -ForegroundColor Red

        $EquiposError++
        $prev = $porEquipo[$Computer]
        $obs = $_.Exception.Message
        
  
        if ($prev) {
            $nuevo = $prev.PSObject.Copy()
            $nuevo.Cumple = 'Error'
            $nuevo.Observacion = $obs
            if (-not $nuevo.Serial) { $nuevo.Serial = 'N/D' }
            if (-not $nuevo.FabricanteEquipo) { $nuevo.FabricanteEquipo = 'N/D' }
            if (-not $nuevo.RAM) { $nuevo.RAM = 'N/D' }
            if (-not $nuevo.TipoRAM) { $nuevo.TipoRAM = 'N/D' }
            if (-not $nuevo.VelocidadRAM) { $nuevo.VelocidadRAM = 'N/D' }
            if (-not $nuevo.SlotsRAM) { $nuevo.SlotsRAM = 'N/D' }
            if (-not $nuevo.FabricanteRAM) { $nuevo.FabricanteRAM = 'N/D' }
            if (-not $nuevo.CapacidadPorSlot) { $nuevo.CapacidadPorSlot = 'N/D' }
            if (-not $nuevo.Discos) { $nuevo.Discos = 'N/D' }
            if (-not $nuevo.TotalDiscoGB) { $nuevo.TotalDiscoGB = 'N/D' }
            $nuevo.FechaHoraEjecucion = $FechaHoraActual
   

            $porEquipo[$Computer] = $nuevo
        }
        else {
            $porEquipo[$Computer] = [pscustomobject]@{
                Equipo                = $Computer
                Serial                = 'N/D'
                FabricanteEquipo      = 'N/D'
                Cumple                = 'Error'
                RAM                   = 'N/D'
                TipoRAM               = 'N/D'
                VelocidadRAM          = 'N/D'
                SlotsRAM              = 'N/D'
                FabricanteRAM         = 'N/D'
                CapacidadPorSlot      = 'N/D'
                Discos                = 'N/D'
                TotalDiscoGB          = 'N/D'
                Observacion           = $obs
                FechaHoraEjecucion    = $FechaHoraActual
                FechaHoraCumplimiento = $null  

            }
        }
    }
    finally {
        if ($session) {
            $session | Remove-CimSession -ErrorAction SilentlyContinue
        }
    }

    Write-Host "`n$Separator" -ForegroundColor Cyan
}

$final = $porEquipo.GetEnumerator() | ForEach-Object { $_.Value }

$final |
Export-Excel -Path $ResultadoExcel -WorksheetName $Hoja `
    -AutoSize -AutoFilter -TableName 'Resultados' -BoldTopRow -FreezeTopRow -ClearSheet

Write-Host "`n$Separator" -ForegroundColor Cyan
Write-Host "*** RESUMEN FINAL ***" -ForegroundColor Green
Write-Host "$Separator" -ForegroundColor Cyan
Write-Host "Total de equipos procesados: $TotalEquipos" -ForegroundColor White
Write-Host "  Equipos que CUMPLEN: $EquiposCumplen" -ForegroundColor Green
Write-Host "  Equipos que NO CUMPLEN: $EquiposNoCumplen" -ForegroundColor Yellow
Write-Host "   Equipos con ERROR/No responde: $EquiposError" -ForegroundColor Red
Write-Host "   Equipos omitidos (ya cumplían): $EquiposOmitidos" -ForegroundColor Cyan
Write-Host "$Separator" -ForegroundColor Cyan
Write-Host "n*** Proceso de Monitoreo Remoto Finalizado ***" -ForegroundColor Green
Write-Host "Archivo generado: $ResultadoExcel" -ForegroundColor Cyan
