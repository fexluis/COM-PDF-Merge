# Comprobación de huellas antes de guardar escalados

## Propósito

Rechazar propuestas calculadas con datos que cambiaron entre evaluación y confirmación. Mantener atomicidad del lote y reconocer la reaplicación de un estado posterior completo sin volver a escribir.

## Cambios

- AgHuellaPlanVentana calcula huella del contexto anterior y del contexto posterior esperado, sin mutar sus entidades. Serialización canónica, SHA-256 y precisión round-trip de valores flotantes.
- AgHuellaEscaladoPersistido, en Infrastructure, captura por contenido aula, relaciones, nodos y catálogos persistidos del cálculo. La lectura al proponer es transaccional; la comparación y la comprobación posterior ocurren en la misma transacción que las escrituras.
- AgEscaladorVentanaAula sella identidad, perfil, ruta y patrones automáticos autorizados. Cambios de propuesta, contexto o BD invalidan el guardado. Exige exactamente un aula y, cuando corresponde, un nodo del perfil; UPDATE con guarda Desde.
- Las postcondiciones incluyen ventana, nodo y patrones. El fallo de cualquier elemento revierte todo el lote y devuelve cero confirmaciones/modificaciones. Los contadores y el sello posterior se publican solo después del commit.
- La reaplicación verifica perfil, contexto compatible y huella SQL posterior completa antes de exigir el estado anterior. Un nodo o patrón incompleto impide confirmar idempotencia.
- La UI captura el perfil al evaluar, vuelve a leer contexto antes de aplicar y recarga la BD después del commit. Se retiró la simulación de cambios en memoria. La notificación posterior al commit tiene un resultado separado de la persistencia.
- Incluye los archivos de planificación fechada recuperados, sus contratos y la integración con fábrica/validador de los que depende la huella. Las versiones de aplicación/instalador preexistentes quedan fuera del commit.

Los cambios en cualquier catálogo incluido invalidan conservadoramente la propuesta aunque mantengan sus conteos. Las decisiones manuales o de dry-run sin conexión no autorizan escritura. El cálculo local de ranuras sigue sin certificar disponibilidad global de docentes o aulas físicas.

## Uso y comprobación

En Diagnóstico u Optimizar, evaluar un aula con demanda que requiere el siguiente escalón y aceptar la propuesta. Si horas, patrones, ventana, bloques, festivos o perfil cambian antes de guardar, se exige recalcular y no se aplica el lote. El contexto recién leído por la UI contrasta también entradas no cubiertas por la persistencia SQL.

La reaplicación con la misma propuesta y postestado completo confirma el aula sin nuevas escrituras. No cambia FIJO/LEGADO ni otros perfiles. Si el guardado se confirmó y falla una notificación o la recarga visual, el mensaje conserva esa distinción.

## Cobertura de pruebas

AgHuellaEscaladoTests: guardado, reaplicación con contexto anterior/recargado, auditoría de ausencia de nuevas escrituras, cambios de contenido con iguales conteos y demanda total, contexto vigente distinto, perfil distinto, propuesta alterada, rollback del lote, error SQL tras actualizar aula, nodo/patrón posterior incompleto, cancelación y ausencia de instantánea. SQLite temporal aislado, sin BD operativa ni Excel.

AgEscaladorVentanaAulaPersistenciaTests: fixtures ajustados para que la demanda del contexto exista en la BD; las aserciones de persistencia y conservación de patrones permanecen. AgPlanificadorRanurasTests cubre la recuperación de sesiones fechadas y compactación.

## Verificación

Solo revisión estática, comprobación de inclusiones en el csproj y git diff --check. Compilación, proyecto, pruebas y smoke NO ejecutados por restricción previa del usuario. Commit y push solicitados expresamente; esto no acredita release ni ausencia de regresiones. Pendiente autorización para Fast dirigido, Full Debug/x64, Full Release/x64 y smoke sobre copia de BD.
