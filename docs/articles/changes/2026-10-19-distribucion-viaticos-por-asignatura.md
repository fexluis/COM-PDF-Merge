# Distribución de Viáticos por Asignatura/Sede y Prevención de Desbordes en Plantilla Excel

**Fecha:** 2026-10-19  
**Módulo:** Académico (`ConversorPDF.XLL.Academic`)  
**Rondas:** 143 (Fases 761 a 766), 144 (Fases 767 a 773), 145 (Fases 774 a 776), 146 (Fases 777 a 779) y 147 (Fases 780 a 782)

## Purpose
Resolver el desacoplamiento en la exportación a Excel (`AsignaturasExcelExporter.cs`) donde los bloques de viáticos se asignaban mediante un recorrido plano secuencial (`bloquesGlobales.Skip(bloqueIndex).Take(8)`), provocando que los viáticos de una sede se imprimieran en la fila de otra sede y anulando el gasto de terminal aérea (`BQ`) en filas adicionales. Se implementa un motor de asignación determinista (`AsignaturaViaticoAllocator`) con reglas de negocio R1 a R7 que asigna cada comisión a la asignatura/sede/ventana correspondiente a su vinculación, preserva íntegros los bloques que cruzan múltiples sedes y restringe la generación de filas de continuación exclusivamente como último recurso ante saturación física. La Ronda 144 perfeccionó la distribución en dos pases (F768), propagó la saturación de fila 234 para evitar truncamiento en cascada (F767), blindó el parseo con hora y cultura `es-CO` (F769), diferenció avisos huérfanos sin inflar saturación (F770), completó la LLAVE (Y) con prestaciones en continuación (F771) y priorizó avisos críticos en UI (F772). La Ronda 145 subsanó el conflicto de identificador de bucle `CS0136` (F774), sincronizó las métricas de conteo de continuaciones y avisos de corte (F775), y desacopló la detección de capacidad máxima leyendo dinámicamente `T_Principal8` (F776). La Ronda 146 blindó la liberación de objetos COM en `ObtenerFilaMaximaPlantilla` (F777), condicionó el registro de advertencias a la escritura efectiva del plan (F778), y actualizó la documentación interna del exportador (F779). La Ronda 147 garantizó el registro inmediato de advertencias en `LoggerHelper` aun ante fallos de escritura COM (F780), precisó el alcance del texto de filas omitidas por corte de capacidad (F781), y completó la liberación de objetos COM intermedios (`coleccion` y `filas`) en `ObtenerFilaMaximaPlantilla` (F782).

## Changes
1. **Nuevo Servicio de Dominio (`AsignaturaViaticoAllocator.cs` y `AsignaturaViaticoFila.cs`, Fases 761, 768, 769, 770):**
   - Implementa el algoritmo determinista de distribución con límite estricto de 8 comisiones por fila (`MaxBloquesPorFila = 8`).
   - **Distribución en Dos Pases (Fase 768):** El Pase 1 reserva con prioridad estricta los cupos para las asignaciones naturales (R1) de cada fila, evitando que bloques desbordados (R4) consuman cupos que requerirán bloques cronológicamente posteriores de la misma fila. El Pase 2 procesa pendientes aplicando R2, R4 y R5.
   - **R1 (Fila natural):** Asignación por correspondencia exacta de `IdVinculacion`.
   - **R2 (Bloques multi-sede):** Bloques que cruzan sedes o vinculaciones consecutivas se mantienen enteros (1 viaje, 1 transporte) y se asignan a la fila con más sesiones del bloque que tenga cupo; desempate por ventana `U` más temprana.
   - **R3 (Límite por fila):** Máximo 8 comisiones (columnas `AK–BG`).
   - **R4 (Reasignación por saturación):** Si la fila natural está llena, el bloque pasa a otra fila con cupo priorizando coincidencia de ventana temporal `[U, V]`, presencia de la sede en la ruta del bloque y cercanía a la fecha de inicio.
   - **Distinción de advertencia sin vinculación (Fase 770):** Si un bloque no coincide con ninguna vinculación (`candidatas.Count == 0`), se etiqueta `"sin vinculación coincidente"` y no se contabiliza en `UltimosBloquesReasignados`.
   - **R5 (Fila de continuación como último recurso):** Si ninguna fila tiene cupo, se crea una fila de continuación inmediatamente después de la fila base replicando datos de la asignatura con sus propias comisiones, ruta y tiquetes.
   - **R6 (Validación temporal):** Advertencia de auditoría si la fecha del bloque se encuentra fuera del rango `[U, V]` de la vinculación sin descartar el viático.
   - **R7 (Invarianza total):** Suma de bloques asignados igual a la totalidad de bloques calculados en la vista Viáticos.
   - **Parseo robusto con hora y cultura es-CO (Fase 769):** Inclusión de formatos con hora (`HH:mm:ss`) y uso explícito de `CultureInfo.GetCultureInfo("es-CO")` para cadenas con `/`, previniendo inversiones silenciosas de día y mes.

2. **Refactorización y Blindaje de Exportador Excel (`AsignaturasExcelExporter.cs`, Fases 762, 767, 771, 774, 775, 776, 777, 778, 779, 780, 781, 782):**
   - Agrupación enriquecida por asignatura incluyendo la ventana temporal (`VinculacionFechaInicio`, `VinculacionFechaFin`) en la clave.
   - Sustitución del ciclo plano `Skip/Take(8)` por la invocación a `AsignaturaViaticoAllocator.Distribuir`.
   - **Trazabilidad asegurada ante fallos COM (Fase 780):** Invocación inmediata a `LoggerHelper.Warning` para los avisos del docente antes de iniciar el bucle de escritura, manteniendo el diagnóstico disponible ante cualquier aborto de ejecución.
   - **Claridad en aviso de capacidad (Fase 781):** Texto informativo que delimita las filas omitidas al docente de corte y explicita que los docentes posteriores no fueron procesados.
   - **Liberación exhaustiva de objetos COM (Fases 777, 782):** Bloque `try/finally` en `ObtenerFilaMaximaPlantilla` con llamadas defensivas a `Marshal.ReleaseComObject` para `coleccion`, `tabla`, `rango` y `filas`.
   - **Sincronización de métricas de plan completo (Fase 778):** Agregación de `_ultimosAvisos` y `UltimosBloquesReasignados` postergada tras el bucle de filas y condicionada a `planCompleto == true`.
   - **Actualización de documentación y bitácora (Fase 779):** Actualizado el resumen XML doc de `Exportar` y corregido el identificador en la bitácora `walkthrough.md`.
   - **Resolución de ámbito de variables CS0136 (Fase 774):** Renombrado el identificador del bucle de filas a `indiceFila`, evitando colisiones de ámbito con los bucles internos `for (int i ...)` en comisiones y ruta.
   - **Exactitud en métricas y corte por capacidad (Fase 775):** Conteo de `UltimasFilasContinuacion` condicionado a la escritura efectiva tras el guard de capacidad e inclusión del docente de corte (`docenteCorte`) en el aviso consolidado. Actualizado el comentario obsoleto respecto a la replicación de datos en duplicados.
   - **Detección dinámica de capacidad institucional (Fase 776):** Método `ObtenerFilaMaximaPlantilla` que inspecciona `ws.ListObjects["T_Principal8"]` con fallback a 234 (`FILA_MAXIMA_PLANTILLA_DEFAULT`).
   - **Propagación de saturación física (Fase 767):** La saturación interrumpe el bucle general de docentes y cuantifica las filas omitidas (`filasOmitidas`), evitando generar cientos de alertas repetitivas.
   - **Completitud de LLAVE (Y) en continuación (Fase 771):** Se escriben `Prestaciones` (columna X) y `PuntajeOcasionales` (columna Z) en filas duplicadas, garantizando la consistencia de la fórmula de auditoría `Y = CONCAT(VINCULACIÓN, CATEGORÍA, PRESTACIONES)`.
   - Soporte en filas de continuación: emisión de `U`/`V` (columnas 21-22), `Grupo` (columna 11) y omisión intencional de horas (`L–O`) para no duplicar horas docente en fórmulas de Excel.
   - **Corrección de anulación de tiquetes aéreos (Líneas 477-483 eliminadas):** Se suprimió la sobreescritura incondicional de `RequiereTiquetesAereos = "NO"` y `TerminalAerea = SIN_TERMINAL` en filas duplicadas, permitiendo que las comisiones de continuación liquiden correctamente sus terminales aéreas en la columna `BQ`.
   - Exposición de trazabilidad: propiedades `UltimosAvisos`, `UltimosBloquesReasignados` y `UltimasFilasContinuacion`.

3. **Retroalimentación en Interfaz de Usuario (`FormVerVinculaciones.cs`, Fases 763, 772):**
   - **Priorización de avisos críticos (Fase 772):** La ventana emergente busca en primer lugar avisos de `"continuación"`, `"saturación"` o `"Capacidad"` antes de recurrir al primer aviso genérico (evitando que avisos informativos R6 oculten alertas operativas críticas).
   - Resumen cuantitativo: `"Exportación finalizada. N bloques reasignados / M filas de continuación / K advertencias (ver log)."` con icono `MessageBoxIcon.Warning` si hay advertencias.

4. **Registro en Proyecto de Compilación (`ConversorPDF.XLL.csproj`, Fase 764):**
   - Inclusión explícita de `AsignaturaViaticoFila.cs` y `AsignaturaViaticoAllocator.cs` en el grupo de compilación del proyecto non-SDK.

5. **Suite de Pruebas Unitarias Automatizadas (`AsignaturaViaticoAllocatorTests.cs`, Fases 765, 773):**
   - Cobertura exhaustiva ampliada a 11 escenarios:
     - Caso canónico 15 bloques en 4 sedes (4/6/2/3).
     - Bloques multi-sede indivisibles (R2).
     - Desborde con reasignación (R4).
     - Desborde total con fila de continuación (R5).
     - Invariante global de máximo 8 comisiones.
     - Advertencias fuera de ventana (R6).
     - Parseo multi-formato ISO y latino.
     - Guardas de colecciones vacías.
     - **Caso real dos Calis, Cartago y El Águila con ventanas solapadas (Fase 773):** Validación exacta de 4/6/2/3 sin desbordes ni avisos espurios.
     - **Bloque sin vinculación coincidente (Fase 773):** Validación de aviso descriptivo sin alterar el contador de saturación.

## How to use
1. En Excel, abrir el formulario **Ver Vinculaciones** desde la cinta institucional.
2. Hacer clic en **Exportar Asignaturas a Excel**.
3. El sistema distribuye automáticamente las comisiones en dos pases, garantizando la reserva de cupo para asignaciones naturales.
4. Al finalizar, la ventana emergente resume el estado de la exportación destacando advertencias críticas de capacidad o continuación si se hubieran producido.
5. Las columnas de ruta, gastos terrestres y terminales aéreas (`BQ`) reflejan exactamente los destinos y valores de las comisiones de cada fila.

## Test Coverage (Fases 765, 773)
- `AsignaturaViaticoAllocatorTests.cs`:
  - `Ejemplo15Bloques4Sedes_AsignaPorVinculacion`
  - `BloqueCruzaDosSedes_SeMantieneEnteroEnFilaConMasSesiones`
  - `BloqueCruzaDosSedes_FilaMayoritariaLlena_UsaOtraVinculacionConCupo`
  - `FilaNaturalLlena_ReasignaAOtraFilaConCupo`
  - `TodasLasFilasLlenas_CreaFilaDeContinuacion`
  - `NingunaFilaSuperaOchoComisiones`
  - `BloqueFueraDeVentana_SeEscribeYAdvierte`
  - `VentanaIsoYddMMyyyy_SeParseanAmbos`
  - `SinAsignaturas_SinBloques_DevuelveVacio`
  - `CasoRealDosCalisCartagoElAguila_15Bloques_RespetaVentanaYFila` (Fase 773)
  - `BloqueSinVinculacionCoincidente_AdvierteSinContarSaturacion` (Fase 773)
