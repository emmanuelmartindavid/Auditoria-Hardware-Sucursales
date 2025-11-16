# Script de Verificación de Equipos - Estado Sólido

## Descripción

Este script de PowerShell está diseñado para verificar remotamente el estado de los equipos en las sucursalesde una empresa, específicamente para comprobar que los técnicos hayan realizado las actualizaciones requeridas:

- **Memoria RAM**: Verificar que los equipos tengan al menos 7 GB de RAM
- **Disco SSD**: Verificar que los equipos tengan un disco SSD instalado (mínimo 128 GB)

El script genera un archivo Excel con los resultados detallados de cada equipo, incluyendo información sobre memoria, discos, fabricante, y estado de cumplimiento.

## Requisitos Previos

### Módulos de PowerShell

Antes de ejecutar el script, es necesario instalar el módulo `ImportExcel`:

```powershell
Install-Module -Name ImportExcel
```

### Permisos

- Permisos de lectura en la ruta base donde se encuentran las carpetas de sucursales
- Permisos de red para conectarse remotamente a los equipos (WMI/CIM)
- Permisos de escritura en la ruta donde se guardará el archivo Excel de resultados

### Configuración de Red

- Los equipos deben estar accesibles en la red
- El servicio WMI/CIM debe estar habilitado en los equipos remotos
- Firewall configurado para permitir conexiones WMI/CIM

## Configuración del Script

Antes de ejecutar el script, es necesario configurar las siguientes variables:

1. **`$BasePath`**: Ruta base donde se encuentran las carpetas de las sucursales
   ```powershell
   $BasePath = 'BasePath'
   ```

2. **`$ResultadoExcel`**: Ruta donde se guardará el archivo Excel con los resultados
   ```powershell
   $ResultadoExcel = "ResultadoExcelPath- $FolderName.xlsx"
   ```

## Uso

1. Abrir PowerShell (como Administrador si es necesario)

2. Ejecutar el script:
   ```powershell
   .\check_system.ps1
   ```

3. Ingresar el código de sucursal cuando se solicite:
   ```
   Ingresar codigo de sucursal: 1234
   ```

4. El script procesará automáticamente todos los equipos encontrados en el archivo Excel de la sucursal.

## Funcionamiento

### 1. Lectura de Datos Iniciales

- El script busca la carpeta de la sucursal en la ruta base usando el código ingresado
- Busca archivos Excel (`.xls`, `.xlsx`) en la carpeta de la sucursal
- Lee la columna `'Nombre del Sistema'` del archivo Excel
- Filtra los equipos que coinciden con el patrón: `^[a-z][0-9]{4}sc[0-9]{4}$`

### 2. Procesamiento de Equipos

Para cada equipo encontrado:

1. **Verificación de Estado Previo**: Si el equipo ya fue verificado y cumple, se omite el análisis
2. **Test de Conectividad**: Se verifica que el equipo responda al ping
3. **Conexión Remota**: Se establece una sesión CIM/WMI remota
4. **Recolección de Información**:
   - Serial del equipo
   - Fabricante del equipo (HP, Lenovo, Dell, etc.)
   - Información de memoria RAM (total, tipo DDR, velocidad, slots, fabricante)
   - Información de discos (tipo, tamaño, modelo)
5. **Validación de Cumplimiento**:
   - RAM >= 7 GB
   - Presencia de SSD
   - Tamaño mínimo de SSD >= 128 GB

### 3. Manejo de Errores

- **Equipos que no responden al ping**: Se marca como "No responde"
- **Errores de conexión (RPC no disponible, equipo apagado)**: Se marca como "Error" pero **NO se modifica la fecha de cumplimiento previa** (si existe)
- **Errores durante la verificación**: Se capturan y se registran en la columna Observación

### 4. Generación de Resultados

El script genera un archivo Excel con las siguientes columnas:

- **Equipo**: Nombre del equipo
- **Serial**: Número de serie del equipo
- **FabricanteEquipo**: Marca del equipo (HP, Lenovo, etc.)
- **Cumple**: Sí/No según si cumple los requisitos
- **RAM**: Cantidad total de RAM en GB
- **TipoRAM**: Tipo de memoria (DDR3, DDR4, DDR5, etc.)
- **VelocidadRAM**: Velocidad de la memoria en MHz
- **SlotsRAM**: Cantidad de slots de memoria utilizados
- **CapacidadPorSlot**: Capacidad de cada slot en GB
- **FabricanteRAM**: Fabricante de los módulos de memoria
- **Discos**: Información de los discos físicos
- **TotalDiscoGB**: Capacidad total de discos en GB
- **Observacion**: Razones por las que no cumple (si aplica)
- **FechaHoraEjecucion**: Fecha y hora de la última ejecución
- **FechaHoraCumplimiento**: Fecha y hora en que se verificó que cumple (solo si cumple)

## Funciones Principales

### `Get-FormFactorLabel`
Convierte el código numérico de form factor de memoria RAM a su etiqueta legible (DIMM, SODIMM, etc.).

### `Get-DDRTypeLabel`
Convierte el código SMBIOS de tipo de memoria RAM a su etiqueta DDR (DDR3, DDR4, DDR5, etc.).

### `Test-NewCimSession`
Establece una sesión CIM remota con un equipo, intentando primero WSMAN y luego DCOM como fallback.

### `Get-DisksInfo`
Obtiene información de los discos físicos de un equipo remoto, identificando si son SSD o HDD.

### `Resolve-EstadoCumplimiento`
Normaliza el texto de estado de cumplimiento a un valor estándar (Si, No, Error, Desconocido).

### `Get-FechasEjecucion`
Calcula las fechas de ejecución y cumplimiento basándose en el estado actual y previo del equipo.

## Criterios de Cumplimiento

Un equipo **CUMPLE** cuando:

- ✅ Tiene al menos **7 GB de RAM** instalada
- ✅ Tiene un **disco SSD** instalado
- ✅ El disco SSD tiene al menos **128 GB** de capacidad (si se puede determinar)

Un equipo **NO CUMPLE** cuando:

- ❌ Tiene menos de 7 GB de RAM
- ❌ No tiene disco SSD
- ❌ El disco SSD es menor a 128 GB

## Resumen Final

Al finalizar la ejecución, el script muestra un resumen con:

- Total de equipos procesados
- Equipos que cumplen
- Equipos que no cumplen
- Equipos con error/no responden
- Equipos omitidos (ya cumplían previamente)

## Notas Importantes

1. **Fecha de Cumplimiento**: Solo se actualiza cuando se verifica exitosamente que el equipo cumple. Si hay un error de conexión, se mantiene la fecha previa (si existe) o queda vacía.

2. **Equipos Omitidos**: Los equipos que ya cumplían en una ejecución anterior se omiten automáticamente para ahorrar tiempo.

3. **Persistencia de Datos**: El script lee el archivo Excel de resultados previo (si existe) y actualiza solo los equipos que necesitan ser re-verificados.

4. **Timeout**: Las conexiones remotas tienen un timeout de 30 segundos. Si un equipo no responde en ese tiempo, se marca como error.

## Solución de Problemas

### Error: "No se encontró la carpeta para la sucursal"
- Verificar que el código de sucursal sea correcto
- Verificar que la ruta base (`$BasePath`) sea correcta
- Verificar que exista una carpeta con el formato: `[CODIGO] - [NOMBRE]`

### Error: "No se encontró ningún archivo Excel"
- Verificar que exista un archivo Excel (`.xls` o `.xlsx`) en la carpeta de la sucursal
- Verificar que el archivo tenga la columna `'Nombre del Sistema'`

### Error: "RPC no está disponible"
- El equipo puede estar apagado
- El servicio WMI puede estar deshabilitado
- Problemas de firewall o red
- El equipo responde al ping pero no permite conexiones WMI/CIM

### Error: "No se pudo instalar el módulo ImportExcel"
- Ejecutar PowerShell como Administrador
- Verificar conexión a internet
- Intentar: `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser`



## Versión

1.0

