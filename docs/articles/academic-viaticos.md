# Viáticos Académicos: Vistas Dinámicas y Detección de Conflictos

Este artículo describe la lógica de negocio y los componentes de UI implementados para la gestión de viáticos en el módulo académico.

## 1. El Modelo de Agrupación (ViaticoBloque)

Para facilitar la liquidación de viáticos, las sesiones individuales se agrupan en **bloques**. Un bloque representa un conjunto de sesiones que ocurren en días consecutivos o el mismo día para un docente en una sede específica.

### Algoritmo de Agrupación
El método `ViaticoBloque.AgruparConsecutivos` realiza el procesamiento:
1. Filtra sesiones con fecha válida.
2. Ordena cronológicamente.
3. Agrupa sesiones si la fecha actual es menor o igual a `FechaUltima + 1 día`.

## 2. Vistas de Navegación

El formulario `FormViaticos` ofrece dos perspectivas para el análisis de datos:

### Vista Por Sedes (Clásica)
Estructura: `Docente → Sede → Bloque → Sesión`.
Ideal para revisar la carga académica distribuida geográficamente.

### Vista Por Fechas (Análisis de Conflictos)
Estructura: `Docente → Bloque (Cronológico) → Sesión`.
Esta vista activa el motor de detección de conflictos para identificar solapamientos de agenda.

## 3. Detección de Conflictos de Agenda

El servicio `ViaticoConflictAnalyzer` es el encargado de identificar riesgos en la programación.

### Criterio de Conflicto
Se marca un bloque con un conflicto (⚠️) si se cumplen **todas** las siguientes condiciones:
- Pertenece al mismo docente.
- Ocurre en sedes **distintas**.
- Los intervalos de fechas se solapan: `InicioA <= FinB AND InicioB <= InicioA`.

### Feedback Visual
- **Nodos del Árbol**: Se muestran en color `DarkRed` con fondo `MistyRose`.
- **Tooltips**: Al pasar el mouse sobre un conflicto, se despliega un detalle descriptivo:
  > ⚠️ CONFLICTO DE FECHAS
  > Este bloque en SEDE A se solapa con:
  > 📍 SEDE B (01/04/2026 → 03/04/2026)

## 4. Exportación a Excel y Ruta Multi-Sede

El reporte generado incluye marcas de advertencia (`⚠️`) en la columna de sede para asegurar que el liquidador de viáticos identifique los conflictos incluso fuera de la aplicación.

### Integración Multi-Sede en Asignaturas
La exportación por asignaturas (`AsignaturasExcelExporter.cs`) incluye una lógica avanzada para distribuir los viajes globales de un docente. Ya que un bloque de viático puede abarcar múltiples asignaturas, el algoritmo distribuye los viajes de forma inteligente:
1. Extrae los viáticos de forma **global** (ignorando la asignatura individual) utilizando `ViaticoBloque.AgruparConsecutivos`.
2. Asigna bloques de 4 viajes consecutivos a las filas de las asignaturas que dicta el docente.
3. Si el docente tiene más de 4 viajes, el algoritmo crea una **Fila Duplicada** debajo de su asignatura que contiene solo la información mínima (`N°`, `Territorial`, `CETAP`, `Materia`, `Identificación`) y coloca allí los viajes restantes para evitar duplicar cálculos salariales.
4. Las columnas de ruta se ajustan dinámicamente para cada fila:
   - **Origen:** Residencia del docente.
   - **Destino:** La última ciudad visitada en el conjunto de 4 viáticos de esa fila.
   - **Escala:** Las ciudades intermedias de la ruta.

### Rango de Fechas de Comisión y Regla de Pernocta en Último Día
Al exportar las columnas de comisiones (`FECHA INICIO 1..8` / `FECHA FIN 1..8`) de la plantilla
`FV-FO-033` mediante `AsignaturasExcelExporter`, el rango efectivo se calcula con
`ViaticoBloqueFormatter.ObtenerRangoComision`. El mismo criterio (`FechaFinComision`) ya se aplica
en `GF-FO-023` a través de `ViaticosIndividualExporter`.
- **Fecha Inicio:** Corresponde a `FechaMinima` del bloque de viáticos.
- **Fecha Fin:** Corresponde a `FechaFinComision` (`FechaMaxima` extendida en 1 día de regreso cuando cualquier sesión del bloque presenta `isPernoctaUltimoDia = 1`).
- **Cálculo en Plantilla:** La plantilla institucional deduce los días de comisión a partir del rango (`INT(diff) + MOD`), por lo que reflejar la fecha real de retorno garantiza una liquidación exacta (ej. 2,5 días en vez de 1,5 días con pernocta).

### Formato y Resaltado del Nivel 2 en la Vista Viáticos
En el TreeView de viáticos (`FormVerVinculaciones.cs`), el nodo de Nivel 2 (bloque de viático) presenta la información consolidada del viaje mediante `ViaticoBloqueFormatter.FormatearNodoNivel2`:
- **Estructura:** `📅 [FechaInicio] - [FechaFin] | [Duración] d | [Aéreo/Terrestre] | [Ruta] | $[ValorTransporte] | [DIA HORA] | [Pernocta Ultimo Dia]`
- **Semántica de Duración:** Se utiliza la abreviatura `d` (ej. `1,5 d`, `2,5 d`).
- **Retiro de `Cierre:` y Código de Ruta solo en la vista:** Se eliminó el prefijo redundante `Cierre:` y el código de trayecto (`[CodRuta]`/`SIN-COD`) de la cadena de visualización del nodo para mejorar la legibilidad. La pre-creación en base de datos (`GetOrCreateCodigoTrayecto`) se mantiene intacta y la exportación de `COD_RUTA` a la columna BL de la plantilla `FV-FO-033` vía `AsignaturasExcelExporter` continúa funcionando sin alteración.
- **Resaltado Visual:** Si el bloque posee `TienePernoctaUltimoDia == true` (`isPernoctaUltimoDia = 1`), el nodo se resalta en color `Color.DarkOrange` con el sufijo `| Pernocta Ultimo Dia`; en caso contrario se dibuja en `Color.DarkBlue`.


