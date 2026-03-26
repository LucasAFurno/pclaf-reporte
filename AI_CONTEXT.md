# pclaf-reporte

Contexto operativo y tecnico del script `DiagnosticoPC.ps1`, redactado para que una IA pueda entender rapido como funciona, que produce y como se integra con `pclaf-web`.

## Resumen corto

Este repo contiene una herramienta de diagnostico para Windows escrita en PowerShell 5.1.

Su flujo principal es:

1. La web descarga un `.bat` generado en cliente.
2. Ese `.bat` descarga `DiagnosticoPC.ps1` desde este repo.
3. El script corre localmente en la PC del cliente o del tecnico.
4. El script releva hardware, sistema, discos, seguridad, rendimiento y trazabilidad previa.
5. El script genera un reporte HTML.
6. El script guarda una huella local en JSON y en el registro de Windows.
7. Luego ese HTML puede subirse manualmente desde `pclaf-web` a la tabla `reportes`.

## Archivos del repo

- `DiagnosticoPC.ps1`: script principal. Contiene recoleccion de datos, evaluacion, persistencia local y render HTML.
- `DiagnosticoPC_Cliente.cmd`: wrapper local para ejecutar el script en modo `cliente`.
- `DiagnosticoPC_Tecnico.cmd`: wrapper local para ejecutar el script en modo `tecnico`.

## Rol de este repo en el sistema completo

Este repo no expone una API ni ejecuta nada en servidor.

La ejecucion real ocurre en la maquina Windows donde se abre el `.bat`. Por eso:

- las consultas WMI/CIM y del registro son locales;
- el HTML se genera localmente;
- la trazabilidad (`C:\ProgramData\PCLAF\last.json` y `HKCU:\SOFTWARE\PCLAF\Diagnostics`) queda en esa PC;
- la web solo actua como lanzador y como interfaz para subir y mostrar el reporte.

## Integracion con `pclaf-web`

La integracion actual depende de `admin.html` del repo web.

### Punto de entrada desde la web

La web define:

- `const TOOLS_BASE = 'https://raw.githubusercontent.com/LucasAFurno/pclaf-reporte/main'`
- `function lanzarScript(tipo)`

La funcion:

1. construye un `.bat` en memoria;
2. lo descarga en el navegador;
3. el usuario lo ejecuta como administrador;
4. el `.bat` descarga `DiagnosticoPC.ps1` desde `raw.githubusercontent.com`;
5. el `.bat` ejecuta `powershell -ExecutionPolicy Bypass -File "%TEMP%\DiagnosticoPC_PCLAF.ps1" -Modo cliente|tecnico`.

Consecuencia importante:

- cualquier cambio de parametros, nombre de archivo, requerimientos o convenciones de salida en este repo puede impactar directamente el flujo de `pclaf-web`.

### Flujo posterior al script

Despues de ejecutar el script:

1. se genera un HTML local;
2. en la web, el admin sube ese archivo manualmente;
3. `pclaf-web` sanitiza el HTML;
4. `pclaf-web` genera una version cliente removiendo secciones sensibles;
5. la web guarda `html_content` y `html_cliente` en Supabase;
6. `historial.html` muestra `html_cliente` al cliente final.

## Parametros de entrada

El script declara:

```powershell
param(
    [ValidateSet("cliente","tecnico")]
    [string]$Modo = "cliente",
    [string]$Tecnico = "PCLAF",
    [int]$MesesMantenimiento = 6,
    [switch]$SistemaInstaladoPorPCLAF,
    [string]$ServicioRealizado = "",
    [string]$PrecioServicio = ""
)
```

### Semantica de parametros

- `Modo`: define si se incluyen solo secciones generales o tambien secciones tecnicas sensibles.
- `Tecnico`: nombre visible en metadata, pie de reporte y marca local.
- `MesesMantenimiento`: se usa para calcular la proxima revision y para la tarea recordatoria.
- `SistemaInstaladoPorPCLAF`: bandera de trazabilidad.
- `ServicioRealizado`: texto opcional que aparece en la seccion de trabajo realizado.
- `PrecioServicio`: texto opcional visible en esa misma seccion.

## Compatibilidad y entorno esperado

- Windows.
- PowerShell 5.1 o compatible con los cmdlets usados.
- Permisos suficientes para consultar WMI/CIM, registro y algunas rutas del sistema.
- Mejor experiencia si se ejecuta como administrador, aunque parte del relevamiento puede funcionar sin elevacion.

Dependencias implicitas:

- `Get-CimInstance`
- `Get-WmiObject`
- acceso a `HKCU` y potencialmente `HKLM`
- `Get-PhysicalDisk`
- `schtasks`
- `sfc`
- `powershell.exe`

Algunos datos son opcionales o degradan a `"N/D"` si el equipo no expone sensores o el comando falla.

## Salidas del script

### 1. HTML principal

El archivo se genera en:

```text
$BasePath\Reporte_${Modo}_${COMPUTERNAME}_${FechaReporte}.html
```

Notas:

- `BasePath` es `PSScriptRoot` cuando existe; si no, usa `%TEMP%`.
- el script intenta abrir el HTML al finalizar con `Start-Process`.

### 2. Persistencia local JSON

Ruta principal:

```text
C:\ProgramData\PCLAF\last.json
```

Tambien escribe:

```text
C:\ProgramData\PCLAF\current.json
```

Uso:

- comparar contra el ultimo diagnostico;
- reconstruir historial local rapido;
- alimentar el flujo de "verificar historial" desde la web.

### 3. Registro de Windows

Clave principal:

```text
HKCU:\SOFTWARE\PCLAF\Diagnostics
```

La web tambien consulta la variante `HKLM\SOFTWARE\PCLAF\Diagnostics` cuando usa la opcion de verificar historial.

## Estructura logica del script

El archivo esta organizado en bloques grandes:

1. helpers;
2. recoleccion de datos del sistema;
3. trazabilidad PCLAF;
4. evaluacion final y recomendaciones;
5. generacion de HTML;
6. guardado final.

## Funciones clave

### Helpers

- `Update-Stage`: muestra progreso.
- `Safe`: normaliza valores nulos o vacios a `"N/D"`.
- `To-DT`: convierte fechas WMI.
- `Round2`: redondeo tolerante a errores.

### Recoleccion de sistema

- `Get-OsInfo`
- `Get-SystemSummary`
- `Get-GpuInfo`
- `Get-RamDetail`
- `Get-Temperatures`
- `Get-DiskHealth`
- `Get-VolumeUsage`
- `Get-NetworkInfo`
- `Get-SecurityInfo`
- `Get-DefenderStatus`
- `Get-BatteryInfo`
- `Get-WindowsActivation`
- `Get-PowerProfile`
- `Get-OpenPorts`
- `Get-NonMsServices`
- `Get-CustomScheduledTasks`
- `Get-DriversInfo`
- `Get-WindowsUpdateStatus`
- `Get-TrimStatus`
- `Get-BootTime`
- `Get-CriticalEvents`
- `Get-MemoryErrors`
- `Get-CurrentPerformance`
- `Get-TopProcesses`
- `Get-StartupApps`
- `Get-InstalledApps`
- `Get-IntegrityStatus`

### Trazabilidad y comparacion

- `Get-HardwareFingerprint`: genera un hash del equipo basado en componentes relevantes.
- `Get-HardwareAge`: clasifica CPU, RAM, disco y GPU con heuristicas de vigencia.
- `Get-PCLAFPaths`: centraliza rutas locales.
- `Read-PreviousRecord`: lee `last.json`.
- `Write-Record`: escribe `current.json` y `last.json`.
- `Set-RegistryMark`: guarda marca local con datos PCLAF.
- `Compare-Records`: compara diagnostico actual contra el previo.

### Evaluacion y salida cliente

- `Get-FinalAssessment`: decide el estado general del equipo.
- `Get-Recommendations`: emite recomendaciones accionables.
- `Set-MaintenanceTask`: prepara recordatorio futuro.
- `Get-ClientSummary`: arma el resumen legible para cliente.

### Render HTML

- `Tag`
- `StatusTag`
- `To-HtmlTable`
- `Get-TrafficLight`
- `Get-RamBar`
- `HtmlEnc`

## Pipeline principal de ejecucion

A alto nivel, el script hace esto:

1. recolecta informacion del sistema;
2. intenta leer la marca y el registro previo;
3. calcula fingerprint y comparacion;
4. calcula estado final y recomendaciones;
5. escribe persistencia local;
6. renderiza HTML con secciones comunes;
7. si `Modo = tecnico`, agrega secciones sensibles;
8. guarda el HTML;
9. muestra resumen por consola y abre el reporte.

## Objeto `record` persistido

La persistencia JSON se arma con un objeto que contiene al menos:

- `Metadata`
- `FinalStatus`
- `Fingerprint`
- `Comparacion`
- `SystemInfo`
- `SystemSummary`
- `GPU`
- `RAM`
- `Discos`
- `Volumenes`
- `Rendimiento`
- `Temperaturas`
- `Seguridad`
- `Defender`
- `HardwareAge`
- `Recomendaciones`

Esto es importante porque otras piezas del sistema dependen de nombres estables, por ejemplo:

- la opcion "verificar historial" de `pclaf-web` lee `Metadata.Fecha`, `FinalStatus.EstadoGeneral` y `SystemSummary`.

## Diferencias entre modo `cliente` y `tecnico`

### Siempre presentes

- hero con estado general;
- resumen del estado del equipo;
- recomendaciones;
- temperaturas;
- hardware;
- estado de discos;
- espacio en unidades;
- proxima revision;
- bloque opcional de "trabajo realizado hoy".

### Solo en modo `tecnico`

- detalle tecnico de hardware;
- sistema operativo;
- resumen detallado de hardware;
- GPU y modulos RAM;
- activacion de Windows;
- perfil de energia;
- seguridad;
- Defender;
- bateria;
- red;
- Windows Update;
- TRIM;
- integridad del sistema;
- rendimiento;
- historial de arranques;
- top procesos;
- puertos abiertos;
- servicios de terceros;
- tareas programadas;
- programas de inicio;
- drivers antiguos;
- errores de memoria;
- eventos criticos;
- comparativa con diagnostico previo;
- fingerprint;
- aplicaciones instaladas;
- bloque de marca PCLAF.

## Acoplamientos importantes con la web

Estos acoplamientos conviene tratarlos como contrato:

### 1. Nombre del archivo remoto

La web descarga especificamente:

```text
DiagnosticoPC.ps1
```

Si cambia el nombre, la web deja de poder lanzarlo.

### 2. Parametro `-Modo`

La web invoca `-Modo cliente` o `-Modo tecnico`.

Si cambia el `ValidateSet`, la ejecucion desde la web se rompe.

### 3. Persistencia local esperada por "verificar historial"

La web intenta leer:

- registro `PCLAF\Diagnostics`
- archivo `C:\ProgramData\PCLAF\last.json`

Si cambian estas rutas o la forma del JSON, la opcion de verificacion pierde valor o se rompe.

### 4. Estructura HTML con `<section>` y `<h2>`

`pclaf-web` genera una version cliente removiendo secciones sensibles por coincidencia de texto en titulos `h2`.

Si cambian mucho los titulos o la estructura HTML:

- puede filtrarse informacion sensible al cliente;
- o puede eliminarse contenido que deberia mostrarse.

### 5. Contenido sensible

La web enmascara algunas celdas por heuristicas de texto, no por un esquema tipado.

Eso significa que cambios en el formato del HTML pueden requerir actualizar tambien la logica de `generarReporteCliente(...)`.

## Riesgos conocidos de mantenimiento

- El script mezcla logica de negocio, adquisicion de datos y render HTML en un solo archivo grande.
- La version cliente en la web depende de texto visible en los encabezados, no de metadatos estructurados.
- La persistencia local y la visualizacion web comparten una forma de datos informal, no validada por esquema.
- Algunas consultas dependen de comandos o namespaces que pueden no existir en todos los equipos.
- Hay tolerancia a errores alta (`SilentlyContinue`), lo que favorece robustez operativa pero puede ocultar fallos.

## Que deberia revisar una IA antes de modificar este repo

1. Si el cambio afecta parametros de entrada usados por la web.
2. Si el cambio altera nombres de archivos, rutas locales o claves del registro.
3. Si el cambio modifica titulos `h2` o la estructura `section` del HTML.
4. Si el cambio agrega datos sensibles que la version cliente actual no elimina.
5. Si el cambio modifica la estructura del JSON persistido en `last.json`.
6. Si el cambio requiere tambien ajustes en `pclaf-web/admin.html` o `pclaf-web/historial.html`.

## Casos tipicos de cambio

### Agregar una nueva fuente de datos tecnica

Pasos recomendados:

1. crear o ampliar una funcion `Get-*`;
2. incorporarla al pipeline principal;
3. decidir si va al `record` persistido;
4. decidir si aparece en modo cliente, tecnico o ambos;
5. revisar si la web necesita ocultarla en `html_cliente`.

### Cambiar la salida HTML

Antes de hacerlo:

1. revisar `generarReporteCliente(fullHtml)` en `pclaf-web`;
2. revisar la sanitizacion del reporte;
3. mantener secciones y encabezados estables salvo que tambien se actualice la web.

### Cambiar trazabilidad local

Si se tocan:

- `Get-PCLAFPaths`
- `Write-Record`
- `Set-RegistryMark`

tambien hay que revisar la herramienta de verificacion que arma la web.

## Ejemplos de ejecucion

### Cliente

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\DiagnosticoPC.ps1 -Modo cliente
```

### Tecnico

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\DiagnosticoPC.ps1 -Modo tecnico -Tecnico "Lucas / PCLAF" -MesesMantenimiento 6 -SistemaInstaladoPorPCLAF
```

### Con detalle del servicio

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\DiagnosticoPC.ps1 -Modo cliente -ServicioRealizado "Limpieza interna y cambio de pasta termica" -PrecioServicio "$45.000"
```

## Checklist rapido para una IA

- Este repo genera diagnosticos locales, no servicios remotos.
- La web lo lanza descargando `DiagnosticoPC.ps1` desde GitHub raw.
- El HTML tecnico se sube despues manualmente desde la web.
- `html_cliente` no lo genera este repo; lo genera `pclaf-web`.
- `last.json` y el registro de Windows forman parte del contrato operativo.
- Los encabezados de las secciones importan porque la web los usa para ocultar datos sensibles.

## Repos relacionados

- `pclaf-web`: frontend estatico que descarga el launcher `.bat`, sube el HTML y muestra la version cliente.

