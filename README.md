# pclaf-reporte

Script de diagnostico para Windows usado por PCLAF para generar reportes tecnicos y reportes resumidos para cliente.

## Que hace

`DiagnosticoPC.ps1` releva informacion del equipo, evalua su estado general y genera un reporte HTML.

Segun el modo elegido puede producir:

- un reporte `cliente`, mas resumido y facil de leer;
- un reporte `tecnico`, con mucho mas detalle para uso interno.

Ademas, guarda una marca local del diagnostico para poder comparar ejecuciones futuras.

## Archivos del repo

- `DiagnosticoPC.ps1`: script principal.
- `DiagnosticoPC_Cliente.cmd`: ejecuta el script en modo cliente.
- `DiagnosticoPC_Tecnico.cmd`: ejecuta el script en modo tecnico.
- `AI_CONTEXT.md`: documentacion tecnica orientada a IA y mantenimiento.

## Como se usa

### Opcion 1: desde los `.cmd`

Para uso local simple:

- abrir `DiagnosticoPC_Cliente.cmd` para generar reporte cliente;
- abrir `DiagnosticoPC_Tecnico.cmd` para generar reporte tecnico.

En ambos casos conviene ejecutarlo como administrador.

### Opcion 2: ejecutar PowerShell manualmente

Modo cliente:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\DiagnosticoPC.ps1 -Modo cliente
```

Modo tecnico:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\DiagnosticoPC.ps1 -Modo tecnico -Tecnico "Lucas / PCLAF" -MesesMantenimiento 6 -SistemaInstaladoPorPCLAF
```

Ejemplo con detalle del servicio:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\DiagnosticoPC.ps1 -Modo cliente -ServicioRealizado "Limpieza interna y cambio de pasta termica" -PrecioServicio "$45.000"
```

## Parametros principales

- `-Modo cliente|tecnico`: define el tipo de reporte.
- `-Tecnico`: nombre del tecnico que aparece en el reporte.
- `-MesesMantenimiento`: meses hasta la proxima revision sugerida.
- `-SistemaInstaladoPorPCLAF`: marca si el sistema fue instalado por PCLAF.
- `-ServicioRealizado`: texto opcional del trabajo hecho.
- `-PrecioServicio`: texto opcional del valor cobrado.

## Que informacion incluye

El script puede relevar:

- sistema operativo;
- procesador, RAM, GPU y motherboard;
- estado y salud de discos;
- uso de volumenes;
- temperatura;
- seguridad y Defender;
- rendimiento actual;
- procesos y programas de inicio;
- eventos criticos;
- fingerprint del hardware;
- comparacion con un diagnostico anterior.

El modo tecnico incluye mas secciones que el modo cliente.

## Salidas generadas

### Reporte HTML

El reporte se guarda como un archivo `.html` en la carpeta desde donde corre el script, o en `%TEMP%` cuando corresponde.

El nombre sigue este formato:

```text
Reporte_{modo}_{equipo}_{fecha}.html
```

### Persistencia local

Tambien guarda informacion en:

- `C:\ProgramData\PCLAF\last.json`
- `C:\ProgramData\PCLAF\current.json`
- `HKCU:\SOFTWARE\PCLAF\Diagnostics`

Eso permite tener trazabilidad local del equipo y comparar cambios entre diagnosticos.

## Relacion con `pclaf-web`

Este repo se integra con la web de PCLAF, pero no corre en servidor.

El flujo actual es:

1. la web descarga un `.bat`;
2. ese `.bat` baja `DiagnosticoPC.ps1` desde el sitio publicado de PCLAF;
3. el script corre localmente en Windows;
4. se genera el HTML;
5. luego ese HTML se puede subir manualmente desde la web.

Importante:

- la copia publicada del script debe mantenerse sincronizada en `pclaf-web/tools/DiagnosticoPC.ps1`;
- si cambia el nombre `DiagnosticoPC.ps1`, la web deja de poder descargarlo;
- si cambian los modos `cliente` y `tecnico`, la web tambien se ve afectada.

## Requisitos

- Windows.
- PowerShell 5.1 o compatible.
- Recomendado: ejecutar como administrador.

Algunos chequeos dependen del hardware, drivers o permisos del equipo, por lo que en ciertos casos pueden mostrarse valores como `N/D`.

## Recomendaciones para mantenimiento

- Mantener estable el nombre `DiagnosticoPC.ps1`.
- Tener cuidado si se modifican rutas locales o claves del registro.
- Si se cambia mucho la estructura del HTML, revisar tambien `pclaf-web`.
- Para contexto tecnico mas profundo, leer `AI_CONTEXT.md`.
